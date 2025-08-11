#   GNU GPLv3
#   <this is an add-on Script/Macro for the geospatial software "Trimble Business Center" aka TBC>
#   <you'll need at least the "Survey Advanced" licence of TBC in order to run this script>
#	<see the ToolTip section below for a brief explanation what the script does>
#	<see the Help-Files for more details>
#   Copyright (C) 2023 Ronny Schneider
#
#   This program is free software: you can redistribute it and/or modify
#   it under the terms of the GNU General Public License as published by
#   the Free Software Foundation, either version 3 of the License, or
#   (at your option) any later version.
#
#   This program is distributed in the hope that it will be useful,
#   but WITHOUT ANY WARRANTY; without even the implied warranty of
#   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#   GNU General Public License for more details.
#
#   You should have received a copy of the GNU General Public License
#   along with this program.  If not, see <https://www.gnu.org/licenses/>

from System.Collections.Generic import List, IEnumerable # import here, otherwise there is a weird issue with Count and Add for lists
import os
exec(open("C:\ProgramData\Trimble\MacroCommands3\SCR Macros\SCR_Imports.py").read())

def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_Remove3DFaces"
    cmdData.CommandName = "SCR_Remove3DFaces"
    cmdData.Caption = "_SCR_Remove3DFaces"
    cmdData.UIForm = "SCR_Remove3DFaces"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "DTM"
        cmdData.ShortCaption = "Remove3D Faces"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.03
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "find and remove 3D Faces below a set area-threshold"
        cmdData.ToolTipTextFormatted = "find and remove 3D Faces below a set area-threshold"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png") 
        cmdData.ImageSmall = b
    except:
        pass


class SCR_Remove3DFaces(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_Remove3DFaces.xaml") as s:
            wpf.LoadComponent (self, s)
        self.currentProject = currentProject
        self.macroFileFolder = macroFileFolder

    def HelpClicked(self, cmd, e):
        webbrowser.open("C:\ProgramData\Trimble\MacroCommands3\SCR Macros\MacroHelp\MacroHelp.htm#" + type(self).__name__)


    def OnLoad(self, cmd, buttons, event):
        self.okBtn = buttons[0]
        buttons[2].Content = "Help"
        buttons[2].Visibility = Visibility.Visible
        buttons[2].Click += self.HelpClicked
        self.Caption = cmd.Command.Caption
        self.threshold.NumberOfDecimals = 6
        self.threshold.MinValue = 0.0

        # get the units for area
        self.aunits = self.currentProject.Units.Area
        self.afp = self.aunits.Properties.Copy()
        areasuffix = self.aunits.Units[self.aunits.DisplayType].Abbreviation
        self.arealabel.Content = "Area Threshold  [" + areasuffix + "]"

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def SetDefaultOptions(self):

        lserial = OptionsManager.GetUint("SCR_Remove3DFaces.layerpicker", 8)
        o = self.currentProject.Concordance.Lookup(lserial) # get the object withj the saved serial number
        if o != None:   # check if we actually got an object back
            if isinstance(o.GetSite(), LayerCollection):    # check if the object is a layer
                self.layerpicker.SetSelectedSerialNumber(lserial, InputMethod(3))
            else:  # if not then set Layer zero                       
                self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:   # if not then set Layer zero               
            self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))

        self.threshold.Value = OptionsManager.GetDouble("SCR_Remove3DFaces.threshold", 0.000001)

        self.checkplanarea.IsChecked = OptionsManager.GetBool("SCR_Remove3DFaces.checkplanarea", True)
        self.checkslopearea.IsChecked = OptionsManager.GetBool("SCR_Remove3DFaces.checkslopearea", False)
        
        self.coloronly.IsChecked = OptionsManager.GetBool("SCR_Remove3DFaces.coloronly", False)
        try:    self.threshcolorpicker.SelectedColor = Color.FromArgb(OptionsManager.GetInt("SCR_Remove3DFaces.threshcolorpicker"))
        except: self.threshcolorpicker.SelectedColor = Color.Magenta


    def SaveOptions(self):

        OptionsManager.SetValue("SCR_Remove3DFaces.layerpicker", self.layerpicker.SelectedSerialNumber)

        OptionsManager.SetValue("SCR_Remove3DFaces.threshold", self.threshold.Value)

        OptionsManager.SetValue("SCR_Remove3DFaces.checkplanarea", self.checkplanarea.IsChecked)
        OptionsManager.SetValue("SCR_Remove3DFaces.checkslopearea", self.checkslopearea.IsChecked)
        
        OptionsManager.SetValue("SCR_Remove3DFaces.coloronly", self.coloronly.IsChecked)

        OptionsManager.SetValue("SCR_Remove3DFaces.threshcolorpicker", self.threshcolorpicker.SelectedColor.ToArgb())


    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''




        self.success.Content = ''
        inputok = True
        try:
            threshold = abs(self.threshold.Value)   # in case of negative value entry, isn't really necessary
            self.threshold.Value = threshold
        except:
            inputok = False

        if inputok:
            self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
            UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
            
            planarea = clr.StrongBox[float]()
            slopearea = clr.StrongBox[float]()
            
            p0 = clr.StrongBox[Point3D]()
            p1 = clr.StrongBox[Point3D]()
            p2 = clr.StrongBox[Point3D]()

            try:
                with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                    layerobject = self.currentProject.Concordance.Lookup(self.layerpicker.SelectedSerialNumber) # get the layer as object
                    layermembers = []   # get the layer members in a separate list
                                        # if we start deleting stuff the original layer objects content is changing constantly
                                        # and the progress counter wouldn't work properly
                    for m in layerobject:
                        layermembers.Add(m.SerialNumber)  # we get serial number list of all the elements on that layer
                    
                    ProgressBar.TBC_ProgressBar.Title = "checking Layermembers"
                    del_count = 0

                    for i in range(0, layermembers.Count):
                        if ProgressBar.TBC_ProgressBar.SetProgress(math.floor((i+1) * 100 / layermembers.Count)):
                            break   # function returns true if user pressed cancel

                        lm = self.currentProject.Concordance.Lookup(layermembers[i])
                        if isinstance(lm, Face3D):
                            p0.Value = lm.Point0
                            p1.Value = lm.Point1
                            p2.Value = lm.Point2

                            Triangle3D.Area(p0, p1, p2, planarea, slopearea)
                            
                            if self.checkplanarea.IsChecked:
                                testarea = planarea.Value
                            if self.checkslopearea.IsChecked:
                                testarea = slopearea.Value
                            
                            if testarea < threshold:
                                
                                if self.coloronly.IsChecked:
                                    del_count += 1
                                    lm.Color = self.threshcolorpicker.SelectedColor
                                else:
                                    del_count += 1
                                    lmsite = lm.GetSite()    # we find out in which container the serial number resides
                                    lmsite.Remove(lm.SerialNumber)   # we delete the object from that container
                    
                    failGuard.Commit()
                    UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
                    self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
            
            except Exception as e:
                tt = sys.exc_info()
                exc_type, exc_obj, exc_tb = sys.exc_info()
                # EndMark MUST be set no matter what
                # otherwise TBC won't work anymore and needs to be restarted
                self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
                UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
                self.error.Content += '\nan Error occurred - Result probably incomplete\n' + str(exc_type) + '\n' + str(exc_obj) + '\nLine ' + str(exc_tb.tb_lineno)

        ProgressBar.TBC_ProgressBar.Title = ""
        if self.coloronly.IsChecked:
            self.success.Content += '\n' + str(del_count) + ' - 3D-Faces re-colored'
        else:
            self.success.Content += '\n' + str(del_count) + ' - 3D-Faces deleted'

        self.SaveOptions()

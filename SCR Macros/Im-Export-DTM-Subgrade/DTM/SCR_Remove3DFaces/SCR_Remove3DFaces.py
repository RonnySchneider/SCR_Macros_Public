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

from System.Collections.Generic import List, IEnumerable
exec(open(r"C:\ProgramData\Trimble\MacroCommands3\SCR Macros\SCR_Imports.py").read())

_OPTIONS = {
    "layerpicker":     8,
    "threshold":       0.000001,
    "checkplanarea":   True,
    "checkslopearea":  False,
    "coloronly":       False,
    "threshcolorpicker": Color.Magenta,
}

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

        cmdData.Version = 1.08
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
        SCROptions.LoadMacroOptions(self, "SCR_Remove3DFaces", _OPTIONS, self.currentProject)

    def SaveOptions(self):
        SCROptions.SaveMacroOptions(self, "SCR_Remove3DFaces", _OPTIONS)


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
            UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
            self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
            
            planarea = clr.StrongBox[float]()
            slopearea = clr.StrongBox[float]()
            

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

                    time1 = datetime.now()
                    for i in range(0, layermembers.Count):

                        if (datetime.now() - time1).seconds > 0.5:
                            if ProgressBar.TBC_ProgressBar.SetProgress(math.floor((i+1) * 100 / layermembers.Count)):
                                break   # function returns true if user pressed cancel
                            time1 = datetime.now()

                        lm = self.currentProject.Concordance.Lookup(layermembers[i])
                        
                        p0_org = p1_org = p2_org = None
                        if isinstance(lm, Face3D):
                            p0_org = lm.Point0
                            p1_org = lm.Point1
                            p2_org = lm.Point2
                        elif isinstance(lm, Linestring):
                            polyseg = lm.ComputePolySeg()
                            polyseg = polyseg.ToWorld()
                            nodes = polyseg.ToPoint3DArray()
                            if nodes.Count == 4 and Point3D.IsDuplicate(nodes[0], nodes[3]):
                                p0_org = nodes[0]
                                p1_org = nodes[1]
                                p2_org = nodes[2]

                        if p0_org and p1_org and p2_org:
                            
                            p0 = clr.StrongBox[Point3D](p0_org)
                            p1 = clr.StrongBox[Point3D](p1_org)
                            p2 = clr.StrongBox[Point3D](p2_org)
                            
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
                    self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
                    UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
            
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

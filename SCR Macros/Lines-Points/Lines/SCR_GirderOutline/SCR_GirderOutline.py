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
    cmdData.Key = "SCR_GirderOutline"
    cmdData.CommandName = "SCR_GirderOutline"
    cmdData.Caption = "_SCR_GirderOutline"
    cmdData.UIForm = "SCR_GirderOutline"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "Lines"
        cmdData.ShortCaption = "Girder Outlines"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.02
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "create Girder Outlines"
        cmdData.ToolTipTextFormatted = "create Girder Outlines"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png") 
        cmdData.ImageSmall = b
    except:
        pass

class SCR_GirderOutline(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_GirderOutline.xaml") as s:
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

        self.lType = clr.GetClrType(IPolyseg)

        optionMenu = SelectionContextMenuHandler()
        # remove options that don't apply here
        optionMenu.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        self.objs.ButtonContextMenu = optionMenu

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        #self.lfp = self.lunits.Properties.Copy()
        linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation
        #self.lfp.AddSuffix = False
        self.girderwidthlabel.Content = "Girder Width [" + linearsuffix + "]"
        self.girderlengthlabel.Content = "Girder Length [" + linearsuffix + "]"
        self.girderskewlabel.Content = "Girder Skew [" + linearsuffix + "]"
        
		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass


    def SetDefaultOptions(self):
        
        settingserial = OptionsManager.GetUint("SCR_GirderOutline.layerpicker", 8) # 8 is FixedSerial for Layer Zero
        o = self.currentProject.Concordance.Lookup(settingserial)
        if o != None:
            if isinstance(o.GetSite(), LayerCollection):    
                self.layerpicker.SetSelectedSerialNumber(settingserial, InputMethod(3))
            else:                       
                self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:                       
            self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))

        self.girderwidth.Distance = OptionsManager.GetDouble("SCR_GirderOutline.girderwidth", 0.500)
        self.girderlength.Distance = OptionsManager.GetDouble("SCR_GirderOutline.girderlength", 20.0000)
        self.girderskew.Distance = OptionsManager.GetDouble("SCR_GirderOutline.girderskew", 0.1000)

        self.leftforward.IsChecked = OptionsManager.GetBool("SCR_GirderOutline.leftforward", True)
        self.rightforward.IsChecked = OptionsManager.GetBool("SCR_GirderOutline.rightforward", False)

    def SaveOptions(self):

        OptionsManager.SetValue("SCR_GirderOutline.layerpicker", self.layerpicker.SelectedSerialNumber)
        OptionsManager.SetValue("SCR_GirderOutline.girderwidth", abs(self.girderwidth.Distance))
        OptionsManager.SetValue("SCR_GirderOutline.girderlength", abs(self.girderlength.Distance))
        OptionsManager.SetValue("SCR_GirderOutline.girderskew", abs(self.girderskew.Distance))

        OptionsManager.SetValue("SCR_GirderOutline.leftforward", self.leftforward.IsChecked)
        OptionsManager.SetValue("SCR_GirderOutline.rightforward", self.rightforward.IsChecked)

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''
        self.success.Content=''
        
        wv = self.currentProject [Project.FixedSerial.WorldView]

        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
      
        try:

            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                for o in self.objs:
                    if isinstance(o, self.lType):

                        orgpoly = o.ComputePolySeg().Clone()
                        
                        orgvec = Vector3D(orgpoly.FirstNode.Point, orgpoly.LastNode.Point)

                        # extend/shrink/center properly in 3D
                        extvec = orgvec.Clone()
                        extvec.Length = (abs(self.girderlength.Distance) - orgvec.Length2D) / 2
                            
                        tmp2 = orgpoly.LastNode.Point + extvec
                        extvec.Negate()
                        tmp1 = orgpoly.FirstNode.Point + extvec

                        # create the proper length CL polyseg
                        clpoly = PolySeg.PolySeg()
                        clpoly.Add(tmp1)
                        clpoly.Add(tmp2)

                        # offset to the left and right
                        leftpoly = clpoly.Clone()
                        leftpoly = leftpoly.Offset(Side.Left, abs(self.girderwidth.Distance) / 2)[1]

                        rightpoly = clpoly.Clone()
                        rightpoly = rightpoly.Offset(Side.Right, abs(self.girderwidth.Distance) / 2)[1]

                        # shift them forwards or backwards
                        shiftvec = orgvec.Clone()
                        shiftvec.Length = abs(self.girderskew.Distance) / 2

                        if self.leftforward.IsChecked:
                            leftpoly.Translate(leftpoly.FirstNode, leftpoly.LastNode, shiftvec)
                            shiftvec.Negate()
                            rightpoly.Translate(rightpoly.FirstNode, rightpoly.LastNode, shiftvec)
                        else:
                            rightpoly.Translate(rightpoly.FirstNode, rightpoly.LastNode, shiftvec)
                            shiftvec.Negate()
                            leftpoly.Translate(leftpoly.FirstNode, leftpoly.LastNode, shiftvec)

                        # add geometry from the left side to the right side
                        leftpoly.Reverse()
                        rightpoly.Add(leftpoly.FirstNode.Point)
                        rightpoly.Add(leftpoly.LastNode.Point)
                        rightpoly.Close(True)
                        # and only draw the right side
                        l = wv.Add(clr.GetClrType(Linestring))
                        l.Layer = self.layerpicker.SelectedSerialNumber
                        l.Append(rightpoly, None, False, False)

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
        
        self.success.Content += '\nDone'
        ProgressBar.TBC_ProgressBar.Title = ""

        Keyboard.Focus(self.girderlength)

        self.SaveOptions()

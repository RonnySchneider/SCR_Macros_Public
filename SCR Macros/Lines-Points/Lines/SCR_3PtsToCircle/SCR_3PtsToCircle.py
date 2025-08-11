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
    cmdData.Key = "SCR_3PtsToCircle"
    cmdData.CommandName = "SCR_3PtsToCircle"
    cmdData.Caption = "_SCR_3PtsToCircle"
    cmdData.UIForm = "SCR_3PtsToCircle"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "Lines"
        cmdData.ShortCaption = "Circle/Center from 3 Points"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.02
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Circle/Center from 3 Points"
        cmdData.ToolTipTextFormatted = "Circle/Center from 3 Points"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_3PtsToCircle(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_3PtsToCircle.xaml") as s:
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

        optionMenu = SelectionContextMenuHandler()
        # remove options that don't apply here
        optionMenu.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        self.objs.ButtonContextMenu = optionMenu

        self.objs.IsEntityValidCallback = self.IsValidPoints

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def SetDefaultOptions(self):
        lserial = OptionsManager.GetUint("SCR_3PtsToCircle.layerpicker", 8)
        o = self.currentProject.Concordance.Lookup(lserial) # get the object withj the saved serial number
        if o != None:   # check if we actually got an object back
            if isinstance(o.GetSite(), LayerCollection):    # check if the object is a layer
                self.layerpicker.SetSelectedSerialNumber(lserial, InputMethod(3))
            else:  # if not then set Layer zero                       
                self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:   # if not then set Layer zero               
            self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))

        self.drawcircle.IsChecked = OptionsManager.GetBool("SCR_3PtsToCircle.drawcircle", False)
        self.drawcenter.IsChecked = OptionsManager.GetBool("SCR_3PtsToCircle.drawcenter", False)

    def SaveOptions(self):
        OptionsManager.SetValue("SCR_3PtsToCircle.layerpicker", self.layerpicker.SelectedSerialNumber)

        OptionsManager.SetValue("SCR_3PtsToCircle.drawcircle", self.drawcircle.IsChecked)
        OptionsManager.SetValue("SCR_3PtsToCircle.drawcenter", self.drawcenter.IsChecked)

    def IsValidPoints(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, CadPoint):
            return True
        if isinstance(o, CoordPoint):
            return True
        return False

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''



        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())

        wv = self.currentProject [Project.FixedSerial.WorldView]

        try:
            # the "with" statement will unroll any changes if something go wrong
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
            
                plist = []

                for p in self.objs:
                    if isinstance(p, CoordPoint) or isinstance(p, CadPoint):
                        plist.Add(clr.StrongBox[Point3D](p.Position))

                if plist.Count != 3:
                    self.success.Content = '\nselect 3 Points'

                else:
                    outcenter = clr.StrongBox[Point3D]()
                    outradius = clr.StrongBox[float]()
                    outdir = clr.StrongBox[bool]()
                    outsmall = clr.StrongBox[bool]()
                    outangle = clr.StrongBox[float]()

                    Arc.GetThreePointArcProperties(plist[0], plist[1], plist[2], outcenter, outradius, outdir, outsmall, outangle)

                    circlecenter = outcenter.Value
                    circlecenter.Z = (plist[0].Value.Z + plist[1].Value.Z + plist[2].Value.Z) / 3

                    if self.drawcenter.IsChecked:
                        cadPoint = wv.Add(clr.GetClrType(CadPoint))
                        cadPoint.Layer = self.layerpicker.SelectedSerialNumber
                        cadPoint.Point0 = circlecenter

                    if self.drawcircle.IsChecked:
                        newcircle = wv.Add(clr.GetClrType(CadCircle))
                        newcircle.Layer = self.layerpicker.SelectedSerialNumber
                        newcircle.CenterPoint = circlecenter
                        newcircle.Radius = outradius.Value

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
        
        Keyboard.Focus(self.objs)

        self.SaveOptions()

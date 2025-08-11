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
    cmdData.Key = "SCR_SimpleCADPoint"
    cmdData.CommandName = "SCR_SimpleCADPoint"
    cmdData.Caption = "_SCR_SimpleCADPoint"
    cmdData.UIForm = "SCR_SimpleCADPoint"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "Points"
        cmdData.ShortCaption = "Simple CAD-Point"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.05
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "create Simple CAD-Point"
        cmdData.ToolTipTextFormatted = "create Simple CAD-Point"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_SimpleCADPoint(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_SimpleCADPoint.xaml") as s:
            wpf.LoadComponent (self, s)
        self.currentProject = currentProject
        self.macroFileFolder = macroFileFolder

    def HelpClicked(self, cmd, e):
        webbrowser.open("C:\ProgramData\Trimble\MacroCommands3\SCR Macros\MacroHelp\MacroHelp.htm#" + type(self).__name__)
        #webbrowser.open("https://docs.google.com/document/d/1qLOWR3lWK97Swg8CfZo1vJOjO05vJQImZllHFRKZyuA/edit#heading=h.gb8w7gj4y4ww")


    def OnLoad(self, cmd, buttons, event):
        self.okBtn = buttons[0]
        buttons[2].Content = "Help"
        buttons[2].Visibility = Visibility.Visible
        buttons[2].Click += self.HelpClicked
        self.Caption = cmd.Command.Caption
        #types = Array [Type] ([CadPoint]) + Array [Type] ([Point3D])    # we fill an array with TBC object types, we could combine different types
        
        self.coordpick1.ShowElevationIf3D = True
        self.coordpick1.ValueChanged += self.coordpick1Changed
        self.coordpick2.ShowElevationIf3D = True
        self.coordpick2.ValueChanged += self.coordpick2Changed
        self.midpointChanged(cmd, event)

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def coordpick1Changed(self, ctrl, e):

        if self.midpoint.IsChecked:
            self.coordpick2.CursorStyle = CursorStyle.CrossHair | CursorStyle.RubberLine
            if self.coordpick1.ResultCoordinateSystem:
                self.coordpick2.AnchorPoint = MousePosition(self.coordpick1.ClickWindow, self.coordpick1.Coordinate, self.coordpick1.ResultCoordinateSystem)
                Keyboard.Focus(self.coordpick2)
            else:
                self.coordpick2.AnchorPoint = None
                
        else:
            if e.Cause == InputMethod.Mouse:     
                self.OkClicked(ctrl, e)

    def coordpick2Changed(self, ctrl, e):

        if self.midpoint.IsChecked:
            if e.Cause == InputMethod.Mouse:     
                self.OkClicked(None, None)

    def midpointChanged(self, sender, e):
        if self.midpoint.IsChecked:
            self.coordpick1.AutoTab = True
        else:
            self.coordpick1.AutoTab = False

    def SetDefaultOptions(self):

        settingserial = OptionsManager.GetUint("SCR_SimpleCADPoint.layerpicker", 8) # 8 is FixedSerial for Layer Zero
        o = self.currentProject.Concordance.Lookup(settingserial)
        if o != None:
            if isinstance(o.GetSite(), LayerCollection):    
                self.layerpicker.SetSelectedSerialNumber(settingserial, InputMethod(3))
            else:                       
                self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:                       
            self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        
        self.manualelev.IsChecked = OptionsManager.GetBool("SCR_SimpleCADPoint.manualelev", False)
        self.elev.Elevation = OptionsManager.GetDouble("SCR_SimpleCADPoint.elev", 0.000)
        self.midpoint.IsChecked = OptionsManager.GetBool("SCR_SimpleCADPoint.midpoint", False)
   
    def SaveOptions(self):

        OptionsManager.SetValue("SCR_SimpleCADPoint.layerpicker", self.layerpicker.SelectedSerialNumber)
        OptionsManager.SetValue("SCR_SimpleCADPoint.manualelev", self.manualelev.IsChecked)
        OptionsManager.SetValue("SCR_SimpleCADPoint.elev", self.elev.Elevation)
        OptionsManager.SetValue("SCR_SimpleCADPoint.midpoint", self.midpoint.IsChecked)
       
    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()
    

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''

        self.success.Content = ""

        wv = self.currentProject [Project.FixedSerial.WorldView]
        wv.PauseGraphicsCache(True)

        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())

        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                p1 = self.coordpick1.Coordinate
                if self.midpoint.IsChecked:
                    p2 = self.coordpick2.Coordinate

                if self.manualelev.IsChecked:
                    p1.Z = self.elev.Elevation
                    if self.midpoint.IsChecked:
                        p2.Z = self.elev.Elevation

                if self.midpoint.IsChecked:
                    p = Point3D.MidPoint(p1, p2)
                else:
                    p = p1

                cadPoint = wv.Add(clr.GetClrType(CadPoint))
                cadPoint.Point0 = p
                cadPoint.Layer = self.layerpicker.SelectedSerialNumber

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
        
        Keyboard.Focus(self.coordpick1)
        self.SaveOptions()           
        wv.PauseGraphicsCache(False)
        
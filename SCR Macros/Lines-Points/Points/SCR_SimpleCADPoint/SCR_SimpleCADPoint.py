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
    "layerpicker": 8, "standardmode": True, "manualelev": False,
    "midpoint": False, "linemode": False,
    "elev": 0.0,
}

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

        cmdData.Version = 1.10
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

        self.ViewOverlay = Guid.NewGuid()
        self.overlayBag = OverlayBag(self.ViewOverlay)

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
        self.coordpick3.ShowElevationIf3D = True
        self.coordpick3.ValueChanged += self.coordpick3Changed
        self.coordpick3.AutoTab = False

        self.midpointChanged(cmd, event)

        self.lType = clr.GetClrType(IPolyseg)
        self.linepicker1.IsEntityValidCallback = self.IsValid
        self.linepicker1.ValueChanged += self.line1Changed

        SCRExpanders.wire_pairs([
            (self.expander_standardmode, self.standardmode),
            (self.expander_linemode, self.linemode),
        ])
		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def drawoverlay(self):

        tt = SCRtest1()
        tt2 = SCROverlayBag.SCRtest2()

        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)
        self.overlayBag = OverlayBag(self.ViewOverlay) # create a new/empty overlaybag

        l1 = self.linepicker1.Entity
        if l1:
            self.overlayBag.AddPolyline(SCROverlayBag.getpolypoints(l1), Color.Red.ToArgb(), 3)

        # if it only needs to be visible in all Planview then remove the Hoops3DViewGUID
        array = Array[Guid]([DisplayWindow.Hoops3DViewGUID, DisplayWindow.HoopsPlanViewGUID])
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.AddOverlayGeometry(array, self.overlayBag)

        return

    def line1Changed(self, ctrl, e):
        l1 = self.linepicker1.Entity
        if l1 != None:
            Keyboard.Focus(self.coordpick3)

        self.drawoverlay()

    def IsValid(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        return False

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

    def coordpick3Changed(self, ctrl, e):

        if self.linemode.IsChecked:
            if e.Cause == InputMethod.Mouse:     
                self.OkClicked(None, None)

    def midpointChanged(self, sender, e):
        if self.midpoint.IsChecked:
            self.coordpick1.AutoTab = True
        else:
            self.coordpick1.AutoTab = False

    def SetDefaultOptions(self):
        SCROptions.LoadMacroOptions(self, "SCR_SimpleCADPoint", _OPTIONS, self.currentProject)

    def SaveOptions(self):
        SCROptions.SaveMacroOptions(self, "SCR_SimpleCADPoint", _OPTIONS)

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def Dispose(self, cmd, disposing):
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)
    

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''

        self.success.Content = ""

        wv = self.currentProject [Project.FixedSerial.WorldView]
        wv.PauseGraphicsCache(True)

        inputok=True

        if self.linemode.IsChecked:
            l1 = self.linepicker1.Entity
            if l1==None: 
                self.success.Content += '\nno Line 1 selected'
                inputok=False

        if inputok:
            UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
            self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)

            try:
                with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                    
                    p = None
                    if self.standardmode.IsChecked:
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

                    elif self.linemode.IsChecked:

                        outPointOnCL1 = clr.StrongBox[Point3D]()
                        station1 = clr.StrongBox[float]()
                        p2 = self.coordpick3.Coordinate

                        polyseg1 = l1.ComputePolySeg()
                        polyseg1 = polyseg1.ToWorld()
                        polyseg1_v = l1.ComputeVerticalPolySeg()
                        
                        polyseg1.FindPointFromPoint(p2, outPointOnCL1, station1)  # with that Point compute the Chainage on Line 1
                        p = outPointOnCL1.Value
                        p.Z = p2.Z

                    if p:
                        cadPoint = wv.Add(clr.GetClrType(CadPoint))
                        cadPoint.Point0 = p
                        cadPoint.Layer = self.layerpicker.SelectedSerialNumber

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
        
        

        self.SaveOptions()           
        wv.PauseGraphicsCache(False)
        
        if self.standardmode.IsChecked:        
            Keyboard.Focus(self.coordpick1)
        elif self.linemode.IsChecked:
            Keyboard.Focus(self.coordpick3)

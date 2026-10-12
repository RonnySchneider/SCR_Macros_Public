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
    "drawtext": False,
    "unitpicker": "Meter", "addunitsuffix": False, "textdecimals": 4.0, "textheight": 0.1,
    "includedist": False, "includeos": False, "includedh": False,
    "includedist3d": False, "includeos3d": False, "includeosperp": False, "includedhperp": False,
    "textlayerpicker": 8,
    "resultunitpicker": "", "resultdecimals": 4.0, # empty unit = project unit
}

# named colours exist in both System.Drawing.Color (overlay) and System.Windows.Media.Colors (UI swatches)
COLOR_LINE = "DodgerBlue"
COLOR_DIST2D = "Magenta"
COLOR_OS2D = "Lime"
COLOR_DH = "Gold"
COLOR_DIST3D = "Orange" # TBC seems to draw Orange above Magenta regardless of the order added
COLOR_OS3D = "Red"
COLOR_OSPERP = "Cyan"
COLOR_DHPERP = "BlueViolet"

def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_MeasureOffsetDistance"
    cmdData.CommandName = "SCR_MeasureOffsetDistance"
    cmdData.Caption = "_SCR_MeasureOffsetDistance"
    cmdData.UIForm = "SCR_MeasureOffsetDistance"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "Lines"
        cmdData.ShortCaption = "Measure Distance/Offset"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.171
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Measure Distance/Offset"
        cmdData.ToolTipTextFormatted = "Measure Distance/Offset"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png") 
        cmdData.ImageSmall = b
    except:
        pass


class SCR_MeasureOffsetDistance(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_MeasureOffsetDistance.xaml") as s:
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
        #self.coordCtl1.ValueChanged += self.Coord1Changed
        #self.coordCtl2.ValueChanged += self.Coord2Changed
        self.coordCtl3.ValueChanged += self.Coord3Changed
        self.coordCtl3.AutoTab = False
        self.btnSwap.Click += self.SwapClicked
        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        # last computed results in metres, re-formatted whenever the result unit or decimals change
        self.resultvalues = []

        for swatch, colorname in [(self.swatchdist2d, COLOR_DIST2D), (self.swatchos2d, COLOR_OS2D), (self.swatchdh, COLOR_DH),
                                  (self.swatchdist3d, COLOR_DIST3D), (self.swatchos3d, COLOR_OS3D),
                                  (self.swatchosperp, COLOR_OSPERP), (self.swatchdhperp, COLOR_DHPERP)]:
            swatch.Background = SolidColorBrush(getattr(Colors, colorname))

        self.coordCtl1.ValueChanged += self.Coord1Changed
        self.coordCtl2.ValueChanged += self.Coord2Changed
        self._overlay_guid = Guid.NewGuid()

        self.textdecimals.MinValue = 0
        self.textdecimals.NumberOfDecimals = 0
        self.textheight.DistanceMin = 0
        self.decdecimals.Click += self.decdecimals_Click
        self.incdecimals.Click += self.incdecimals_Click
        self.resultdecimals.MinValue = 0
        self.resultdecimals.NumberOfDecimals = 0
        self.decresultdecimals.Click += self.decresultdecimals_Click
        self.incresultdecimals.Click += self.incresultdecimals_Click

        SCRExpanders.wire_pairs([
            (self.expander_drawtext, self.drawtext),
        ])
		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass
        self.unitssetup(None, None)

    # def Coord1Changed(self, ctrl, e):
    #     self.coordCtl2.CursorStyle = CursorStyle.CrossHair | CursorStyle.RubberLine
    #     if self.coordCtl1.ResultCoordinateSystem:
    #         self.coordCtl2.AnchorPoint = MousePosition(self.coordCtl1.ClickWindow, self.coordCtl1.Coordinate, self.coordCtl1.ResultCoordinateSystem)
    #     else:
    #         self.coordCtl2.AnchorPoint = None
    # 
    # def Coord2Changed(self, ctrl, e):
    #     self.coordCtl3.CursorStyle = CursorStyle.CrossHair | CursorStyle.RubberLine
    #     if self.coordCtl2.ResultCoordinateSystem:
    #         self.coordCtl3.AnchorPoint = MousePosition(self.coordCtl2.ClickWindow, self.coordCtl2.Coordinate, self.coordCtl2.ResultCoordinateSystem)
    #     else:
    #         self.coordCtl3.AnchorPoint = None

    def SetDefaultOptions(self):
        SCROptions.LoadMacroOptions(self, "SCR_MeasureOffsetDistance", _OPTIONS, self.currentProject)

    def SaveOptions(self):
        SCROptions.SaveMacroOptions(self, "SCR_MeasureOffsetDistance", _OPTIONS)

    def unitssetup(self, sender, e):
        # setup everything for the unit conversions of the text output and the result boxes
        self.textlfp = self.lunits.Properties.Copy() # create a copy in order to set the decimals and enable/disable the suffix
        self.textlfp.AddSuffix = False # disable suffix, we need to set it manually, it would always add the current projects units
        self.resultlfp = self.lunits.Properties.Copy()
        self.resultlfp.AddSuffix = False # the unit is shown in the labels

        self.fillunitpicker(self.unitpicker, LinearType.Meter)
        self.fillunitpicker(self.resultunitpicker, self.lunits.DisplayType)

        self.unitpicker.SelectionChanged += self.unitschanged
        self.textdecimals.ValueChanged += self.unitschanged
        self.resultunitpicker.SelectionChanged += self.resultunitschanged
        self.resultdecimals.ValueChanged += self.resultunitschanged

        self.unitschanged(None, None)
        self.resultunitschanged(None, None)

    def fillunitpicker(self, picker, defaultunit):
        tt = picker.Text # the loaded option, before we fill the list
        for u in self.lunits.Units:
            item = ComboBoxItem()
            item.Content = u.Key
            item.FontSize = 1
            picker.Items.Add(item)
            if u.Key == defaultunit:
                picker.SelectedItem = item
        if picker.SelectedIndex < 0:
            picker.SelectedIndex = 0
        if tt != "":
            picker.Text = tt

    def pickerunit(self, picker):
        # find the enum for the selected LinearType
        unit = LinearType(0)
        for e in range(0, 19):
            if LinearType(e) == picker.SelectedItem.Content:
                unit = LinearType(e)
        return unit

    def unitschanged(self, sender, e):
        self.outputunit = self.pickerunit(self.unitpicker)
        self.textlfp.NumberOfDecimals = int(self.textdecimals.Value)

        self.textheight.DisplayUnit = self.outputunit
        self.textheight.ShowControlIcon(False)
        self.textheight.FormatProperty.AddSuffix = ControlBoolean(1)
        # small values would be displayed as 0 with the default decimals, increase them as needed
        self.textheight.FormatProperty.NumberOfDecimals = self.decimalsneeded(self.textheight.Distance, int(self.textdecimals.Value))

    def resultunitschanged(self, sender, e):
        self.resultunit = self.pickerunit(self.resultunitpicker)
        self.resultlfp.NumberOfDecimals = int(self.resultdecimals.Value)

        linearsuffix = self.lunits.Units[self.resultunit].Abbreviation
        self.textdistancelabel.Content = "Distance along [" + linearsuffix + "]"
        self.textoffsetlabel.Content = "Offset [" + linearsuffix + "]"
        self.textdhlabel.Content = "vertical dh to Line [" + linearsuffix + "]"
        self.textdistance3dlabel.Content = "Distance along [" + linearsuffix + "]"
        self.textoffset3dlabel.Content = "3D Offset [" + linearsuffix + "]"
        self.textoffsetperplabel.Content = "perp. Offset [" + linearsuffix + "]"
        self.textdhperplabel.Content = "perp. dh [" + linearsuffix + "]"

        self.showresults()

    def showresults(self):
        for box, v in self.resultvalues:
            if v is None:
                box.Text = ''
            else:
                box.Text = self.lunits.Format(LinearType.Meter, v, self.resultlfp, self.resultunit)

    def decresultdecimals_Click(self, sender, e):
        if not self.resultdecimals.Value == 0:
            self.resultdecimals.Value -= 1

    def incresultdecimals_Click(self, sender, e):
        self.resultdecimals.Value += 1

    def decimalsneeded(self, v, mindecimals):
        # smallest number of decimals (>= mindecimals, <= 6) that shows the value without rounding it away
        d = mindecimals
        while d < 6 and abs(round(v, d) - v) > 1e-9:
            d += 1
        return d

    def decdecimals_Click(self, sender, e):
        if not self.textdecimals.Value == 0:
            self.textdecimals.Value -= 1

    def incdecimals_Click(self, sender, e):
        self.textdecimals.Value += 1

    def tooutputunit(self, v):
        self.textlfp.AddSuffix = self.addunitsuffix.IsChecked
        return self.lunits.Format(LinearType.Meter, v, self.textlfp, self.outputunit)

    def SwapClicked(self, sender, e):
        c1 = self.coordCtl1.Coordinate
        c2 = self.coordCtl2.Coordinate
        wv = self.currentProject[Project.FixedSerial.WorldView]
        csd = wv.CoordinateSystemDefinition
        self.coordCtl1.SetCoordinate(c2, self.currentProject, csd)
        self.coordCtl2.SetCoordinate(c1, self.currentProject, csd)
        self.OkClicked(None, None)
        Keyboard.Focus(self.coordCtl3)

    def Coord1Changed(self, ctrl, e):
        self.coordCtl2.CursorStyle = CursorStyle.CrossHair | CursorStyle.RubberLine
        if self.coordCtl1.ResultCoordinateSystem:
            self.coordCtl2.AnchorPoint = MousePosition(self.coordCtl1.ClickWindow, self.coordCtl1.Coordinate, self.coordCtl1.ResultCoordinateSystem)
        else:
            self.coordCtl2.AnchorPoint = None
        self.OkClicked(None, None)
        Keyboard.Focus(self.coordCtl2)

    def Coord2Changed(self, ctrl, e):
        self.OkClicked(None, None)
        Keyboard.Focus(self.coordCtl3)

    def Coord3Changed(self, ctrl, e):
        self.OkClicked(None, None)
        Keyboard.Focus(self.coordCtl3)

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content = ''
        self._clear_overlay()

        wv = self.currentProject[Project.FixedSerial.WorldView]

        p1 = self.coordCtl1.Coordinate
        p2 = self.coordCtl2.Coordinate
        p3 = self.coordCtl3.Coordinate

        if p1 is None or p2 is None or p3 is None:
            return

        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)

        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                dx, dy, dz = p2.X - p1.X, p2.Y - p1.Y, p2.Z - p1.Z
                px, py, pz = p3.X - p1.X, p3.Y - p1.Y, p3.Z - p1.Z
                len2d = math.sqrt(dx * dx + dy * dy)
                len3d = math.sqrt(dx * dx + dy * dy + dz * dz)

                if len3d < 1e-9:
                    self.error.Content = 'Point A and B are identical'
                    self.resultvalues = [(box, None) for box, v in self.resultvalues]
                    self.showresults()
                    failGuard.Commit()
                    self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
                    UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
                    return

                # side of the offset point, X/Y is Easting/Northing => cross > 0 means left of A->B
                cross = dx * py - dy * px
                side = -1.0 if cross > 0 else 1.0 # left negative, right positive

                # 2D - project in plan, foot point gets the line elevation at that station
                foot2d = None
                level = None
                if len2d > 1e-9:
                    t2 = (px * dx + py * dy) / (len2d * len2d)
                    foot2d = Point3D(p1.X + t2 * dx, p1.Y + t2 * dy, p1.Z + t2 * dz)
                    level = Point3D(p3.X, p3.Y, foot2d.Z) # offset point at line elevation, splits offset and dh
                    offset2d = math.sqrt((p3.X - foot2d.X) ** 2 + (p3.Y - foot2d.Y) ** 2)
                    dist2dvalue, os2dvalue, dhvalue = t2 * len2d, side * offset2d, p3.Z - foot2d.Z
                else:
                    dist2dvalue, os2dvalue, dhvalue = None, None, None
                    self.error.Content = 'A and B are identical in plan - no 2D result'

                # 3D - shortest distance to the infinite line through A and B
                t3 = (px * dx + py * dy + pz * dz) / (len3d * len3d)
                foot3d = Point3D(p1.X + t3 * dx, p1.Y + t3 * dy, p1.Z + t3 * dz)
                offset3d = math.sqrt((p3.X - foot3d.X) ** 2 + (p3.Y - foot3d.Y) ** 2 + (p3.Z - foot3d.Z) ** 2)
                dist3dvalue, os3dvalue = t3 * len3d, side * offset3d

                # A-B plane - contains A-B and the horizontal direction square to it, undefined if A-B is plumb
                # project the offset point perpendicular onto that plane
                projected = None
                if len2d > 1e-9:
                    # plane normal n = h x u, with u along A-B and h horizontal to the right, points upwards
                    nx = -dx * dz / (len2d * len3d)
                    ny = -dy * dz / (len2d * len3d)
                    nz = len2d / len3d
                    perpdh = px * nx + py * ny + pz * nz # positive above the plane
                    projected = Point3D(p3.X - perpdh * nx, p3.Y - perpdh * ny, p3.Z - perpdh * nz)
                    # shortest distance from the projected point to A-B
                    qx, qy, qz = projected.X - p1.X, projected.Y - p1.Y, projected.Z - p1.Z
                    tq = (qx * dx + qy * dy + qz * dz) / (len3d * len3d)
                    footperp = Point3D(p1.X + tq * dx, p1.Y + tq * dy, p1.Z + tq * dz)
                    offsetperp = math.sqrt((projected.X - footperp.X) ** 2 + (projected.Y - footperp.Y) ** 2 + (projected.Z - footperp.Z) ** 2)
                    osperpvalue, dhperpvalue = side * offsetperp, perpdh
                else:
                    footperp = None
                    osperpvalue, dhperpvalue = None, None

                self.resultvalues = [(self.resultdistance, dist2dvalue), (self.resultoffset, os2dvalue), (self.resultdh, dhvalue),
                                     (self.resultdistance3d, dist3dvalue), (self.resultoffset3d, os3dvalue),
                                     (self.resultoffsetperp, osperpvalue), (self.resultdhperp, dhperpvalue)]
                self.showresults()

                self._draw_overlay(p1, p2, p3, foot2d, level, foot3d, projected, footperp)

                if self.drawtext.IsChecked:
                    textparts = []
                    if foot2d is not None:
                        if self.includedist.IsChecked: textparts.Add("d=" + self.tooutputunit(dist2dvalue))
                        if self.includeos.IsChecked: textparts.Add("dO=" + self.tooutputunit(os2dvalue))
                        if self.includedh.IsChecked: textparts.Add("dH=" + self.tooutputunit(dhvalue))
                    if self.includedist3d.IsChecked: textparts.Add("d3D=" + self.tooutputunit(dist3dvalue))
                    if self.includeos3d.IsChecked: textparts.Add("dO3D=" + self.tooutputunit(os3dvalue))
                    if projected is not None:
                        if self.includeosperp.IsChecked: textparts.Add("dOp=" + self.tooutputunit(osperpvalue))
                        if self.includedhperp.IsChecked: textparts.Add("dHp=" + self.tooutputunit(dhperpvalue))

                    if len(textparts) > 0:
                        t = wv.Add(clr.GetClrType(MText))
                        t.AlignmentPoint = p3
                        t.AttachPoint = AttachmentPoint.BottomLeft
                        t.Rotation = 2 * math.pi/2 - Vector3D(p1, p2).Azimuth - math.pi/2
                        t.Height = abs(self.textheight.Distance)
                        t.Layer = self.textlayerpicker.SelectedSerialNumber
                        t.TextString = '\\P'.join(textparts)

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

    def Dispose(self, cmd, disposing):
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self._overlay_guid)

    def _clear_overlay(self):
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self._overlay_guid)

    def _draw_overlay(self, p1, p2, p3, foot2d, level, foot3d, projected, footperp):
        # 2D and 3D distance along both lie on A-B, the 2D one is drawn wide and the 3D one narrower on top
        # the draw order inside a bag seems to follow the colour, not the order added - hence the colour choice
        bag = OverlayBag(self._overlay_guid)
        line = getattr(Color, COLOR_LINE).ToArgb()
        dist2d = getattr(Color, COLOR_DIST2D).ToArgb()
        os2d = getattr(Color, COLOR_OS2D).ToArgb()
        dh = getattr(Color, COLOR_DH).ToArgb()
        dist3d = getattr(Color, COLOR_DIST3D).ToArgb()
        os3d = getattr(Color, COLOR_OS3D).ToArgb()
        osperp = getattr(Color, COLOR_OSPERP).ToArgb()
        dhperp = getattr(Color, COLOR_DHPERP).ToArgb()

        if foot2d is not None:
            bag.AddPolyline(Array[Point3D]([p1, foot2d]), dist2d, 8)
        bag.AddPolyline(Array[Point3D]([p1, foot3d]), dist3d, 3)
        bag.AddPolyline(Array[Point3D]([p1, p2]), line, 1)

        if foot2d is not None:
            bag.AddPolyline(Array[Point3D]([foot2d, level]), os2d, 3)
            bag.AddPolyline(Array[Point3D]([level, p3]), dh, 3)
        if projected is not None:
            bag.AddPolyline(Array[Point3D]([footperp, projected]), osperp, 3)
            bag.AddPolyline(Array[Point3D]([projected, p3]), dhperp, 3)
        bag.AddPolyline(Array[Point3D]([foot3d, p3]), os3d, 3)

        bag.AddMarker(p1, GraphicMarkerTypes.HollowCircle_IndependentColor, line, "A", 0, 0.0, 1.0)
        bag.AddMarker(p2, GraphicMarkerTypes.HollowCircle_IndependentColor, line, "B", 0, 0.0, 1.0)
        if foot2d is not None:
            bag.AddMarker(foot2d, GraphicMarkerTypes.HollowCircle_IndependentColor, dist2d, "", 0, 0.0, 0.7)
        bag.AddMarker(foot3d, GraphicMarkerTypes.HollowCircle_IndependentColor, dist3d, "", 0, 0.0, 0.7)
        if projected is not None:
            bag.AddMarker(projected, GraphicMarkerTypes.HollowCircle_IndependentColor, osperp, "", 0, 0.0, 0.7)
        bag.AddMarker(p3, GraphicMarkerTypes.HollowCircle_IndependentColor, os3d, "Offset", 0, 0.0, 1.0)
        views = Array[Guid]([DisplayWindow.HoopsPlanViewGUID, DisplayWindow.Hoops3DViewGUID])
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.AddOverlayGeometry(views, bag)

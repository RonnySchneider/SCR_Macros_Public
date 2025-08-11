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
    cmdData.Key = "SCR_PerpDistToFace"
    cmdData.CommandName = "SCR_PerpDistToFace"
    cmdData.Caption = "_SCR_PerpDistToFace"
    cmdData.UIForm = "SCR_PerpDistToFace" # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "DTM"
        cmdData.ShortCaption = "Dist to Face"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.07
        cmdData.MacroAuthor = "SCR"
        cmdData.ToolTipTitle = "3D distance to Face"
        cmdData.ToolTipTextFormatted = "get the perpendicular (or plumb) 3D distance from a point to anywhere on a 3-point plane"
    except:
        pass
    try:
        b = Bitmap(macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


# the name of this class must match name from cmdData.UIForm (above)
class SCR_PerpDistToFace(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader(macroFileFolder + r"\SCR_PerpDistToFace.xaml") as s:
            wpf.LoadComponent(self, s)
        self.currentProject = currentProject
        self.macroFileFolder = macroFileFolder

        self.ViewOverlay = Guid.NewGuid()
        self.overlayBag = OverlayBag(self.ViewOverlay)


    def HelpClicked(self, cmd, e):
        webbrowser.open("C:\ProgramData\Trimble\MacroCommands3\SCR Macros\MacroHelp\MacroHelp.htm#" + type(self).__name__)

    def OnLoad(self, cmd, buttons, event):
        self.okBtn = buttons[0]
        buttons[2].Content = "Help"
        buttons[2].Visibility = Visibility.Visible
        buttons[2].Click += self.HelpClicked
        buttons[0].Content = "Apply"
        buttons[1].Content = "Close"
        self.Caption = cmd.Command.Caption

        self.coordCtl1.ValueChanged += self.Coord1Changed
        self.coordCtl2.ValueChanged += self.Coord2Changed
        self.coordCtl3.ValueChanged += self.Coord3Changed
        self.coordCtl4.ValueChanged += self.Coord4Changed
        self.coordCtl5.ValueChanged += self.Coord5Changed
        self.coordCtl6.ValueChanged += self.Coord6Changed

        self.cadpointType = clr.GetClrType(CadPoint)
        self.coordpointType = clr.GetClrType(CoordPoint)
        self.pointcloudType = clr.GetClrType(PointCloudRegion)

        self.objs.IsEntityValidCallback = self.IsValid
        self.textheight.DistanceMin = 0
		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass
        
        self.unitssetup(None, None)

    def IsValid(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.coordpointType):
            return True
        if isinstance(o, self.cadpointType):
            return True
        if isinstance(o, self.pointcloudType):
            return True

        return False

    def SetDefaultOptions(self):
        self.computeplumb.IsChecked = OptionsManager.GetBool("SCR_PerpDistToFace.computeplumb", False)
        lserial = OptionsManager.GetUint("SCR_PerpDistToFace.layerpicker", 8)
        o = self.currentProject.Concordance.Lookup(lserial)
        if o != None:
            if isinstance(o.GetSite(), LayerCollection):    
                self.layerpicker.SetSelectedSerialNumber(lserial, InputMethod(3))
            else:                       
                self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:                       
            self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        try:    self.outcolorpicker.SelectedColor = Color.FromArgb(OptionsManager.GetInt("SCR_PerpDistToFace.outcolorpicker"))
        except: self.outcolorpicker.SelectedColor = Color.Red

        self.unitpicker.Text = OptionsManager.GetString("SCR_PerpDistToFace.unitpicker", "Meter")
        self.addunitsuffix.IsChecked = OptionsManager.GetBool("SCR_PerpDistToFace.addunitsuffix", False)
        self.textdecimals.Value = OptionsManager.GetDouble("SCR_PerpDistToFace.textdecimalsfloat", 4)
        self.textheight.Distance = OptionsManager.GetDouble("SCR_PerpDistToFace.textheightfloat", 0.1)

        self.checkBox_point.IsChecked = OptionsManager.GetBool("SCR_PerpDistToFace.checkBox_point", False)
        self.checkBox_line.IsChecked = OptionsManager.GetBool("SCR_PerpDistToFace.checkBox_line", True)
        self.checkBox_text.IsChecked = OptionsManager.GetBool("SCR_PerpDistToFace.checkBox_text", True)
        self.projectsinglepoint.IsChecked = OptionsManager.GetBool("SCR_PerpDistToFace.projectsinglepoint", True)
        self.projectmultipoints.IsChecked = OptionsManager.GetBool("SCR_PerpDistToFace.projectmultipoints", False)
        self.projectdirectionalshot.IsChecked = OptionsManager.GetBool("SCR_PerpDistToFace.projectdirectionalshot", False)

    def SaveOptions(self):
        OptionsManager.SetValue("SCR_PerpDistToFace.computeplumb", self.computeplumb.IsChecked)
        OptionsManager.SetValue("SCR_PerpDistToFace.layerpicker", self.layerpicker.SelectedSerialNumber)
        OptionsManager.SetValue("SCR_PerpDistToFace.outcolorpicker", self.outcolorpicker.SelectedColor.ToArgb())

        OptionsManager.SetValue("SCR_PerpDistToFace.unitpicker", self.unitpicker.Text)
        OptionsManager.SetValue("SCR_PerpDistToFace.addunitsuffix", self.addunitsuffix.IsChecked)
        OptionsManager.SetValue("SCR_PerpDistToFace.textdecimalsfloat", abs(self.textdecimals.Value))
        OptionsManager.SetValue("SCR_PerpDistToFace.textheightfloat", self.textheight.Distance)

        OptionsManager.SetValue("SCR_PerpDistToFace.checkBox_point", self.checkBox_point.IsChecked)
        OptionsManager.SetValue("SCR_PerpDistToFace.checkBox_line", self.checkBox_line.IsChecked)
        OptionsManager.SetValue("SCR_PerpDistToFace.checkBox_text", self.checkBox_text.IsChecked)

        OptionsManager.SetValue("SCR_PerpDistToFace.projectsinglepoint", self.projectsinglepoint.IsChecked)
        OptionsManager.SetValue("SCR_PerpDistToFace.projectmultipoints", self.projectmultipoints.IsChecked)
        OptionsManager.SetValue("SCR_PerpDistToFace.projectdirectionalshot", self.projectdirectionalshot.IsChecked)

    def unitssetup(self, sender, e):
        # setup everything for the unit conversions
        self.outputunitenum = 0
        self.textdecimals.NumberOfDecimals = 0

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        self.lfp = self.lunits.Properties.Copy() # create a copy in order to set the decimals and enable/disable the suffix
        self.lfp.AddSuffix = False # disable suffix, we need to set it manually, it would always add the current projects units

        # fill the unitpicker
        for u in self.lunits.Units:
            item = ComboBoxItem()
            item.Content = u.Key
            item.FontSize = 1
            self.unitpicker.Items.Add(item)

        tt = self.unitpicker.Text
        self.unitpicker.SelectedIndex = 0
        if tt != "":
            self.unitpicker.Text = tt
        self.unitpicker.SelectionChanged += self.unitschanged
        self.textdecimals.MinValue = 0.0
        self.textdecimals.ValueChanged += self.unitschanged

        self.lfp.NumberOfDecimals = int(self.textdecimals.Value)
        self.unitschanged(None, None)
    
    def unitschanged(self, sender, e):

        # find the enum for the selected LinearType
        for e in range(0, 19):
            if LinearType(e) == self.unitpicker.SelectedItem.Content:
                self.outputunitenum = e
        
        # loop through all objects of self and set the properties for all DistanceEdits
        # the code is slower than doing it manually for each single one
        # but more convenient since we don't have to worry about how many DistanceEdit Controls we have in the UI
        tt = self.__dict__.items()
        for i in self.__dict__.items():
            if i[1].GetType() == TBCWpf.DistanceEdit().GetType():
                i[1].DisplayUnit = LinearType(self.outputunitenum)
                i[1].ShowControlIcon(False)
                i[1].FormatProperty.AddSuffix = ControlBoolean(1)
                i[1].FormatProperty.NumberOfDecimals = int(self.textdecimals.Value)

    def decdecimals_Click(self, sender, e):
        if not self.textdecimals.Value == 0:
            self.textdecimals.Value -= 1
             # setup the linear format properties
            self.lfp.NumberOfDecimals = int(self.textdecimals.Value)
            self.unitschanged(None, None)

    def incdecimals_Click(self, sender, e):
        self.textdecimals.Value += 1
        # setup the linear format properties
        self.lfp.NumberOfDecimals = int(self.textdecimals.Value)
        self.unitschanged(None, None)

    def tooutputunit(self, v):
        
        self.lfp.AddSuffix = self.addunitsuffix.IsChecked
        return self.lunits.Format(LinearType.Meter, v, self.lfp, LinearType(self.outputunitenum))


    def Coord1Changed(self, ctrl, e):
        self.coordCtl2.CursorStyle = CursorStyle.CrossHair | CursorStyle.RubberLine
        if self.coordCtl1.ResultCoordinateSystem:
            self.coordCtl2.AnchorPoint = MousePosition(self.coordCtl1.ClickWindow, self.coordCtl1.Coordinate, self.coordCtl1.ResultCoordinateSystem)
        else:
            self.coordCtl2.AnchorPoint = None

        if not self.coordCtl1.Coordinate.Is3D :
            self.coordCtl1.StatusMessage = "No valid coordinate defined, must be 3D"
        else:
            self.coordCtl1.StatusMessage = ""

        self.drawoverlay()

    def Coord2Changed(self, ctrl, e):
        self.coordCtl3.CursorStyle = CursorStyle.CrossHair | CursorStyle.RubberLine
        if self.coordCtl2.ResultCoordinateSystem:
            self.coordCtl3.AnchorPoint = MousePosition(self.coordCtl2.ClickWindow, self.coordCtl2.Coordinate, self.coordCtl2.ResultCoordinateSystem)
        else:
            self.coordCtl3.AnchorPoint = None

        if not self.coordCtl2.Coordinate.Is3D :
            self.coordCtl2.StatusMessage = "No valid coordinate defined, must be 3D"
        else:
            self.coordCtl2.StatusMessage = ""

        self.drawoverlay()

    def Coord3Changed(self, ctrl, e):
        self.coordCtl4.CursorStyle = CursorStyle.CrossHair | CursorStyle.RubberLine
        if self.coordCtl3.ResultCoordinateSystem:
            self.coordCtl4.AnchorPoint = MousePosition(self.coordCtl3.ClickWindow, self.coordCtl3.Coordinate, self.coordCtl3.ResultCoordinateSystem)
        else:
            self.coordCtl4.AnchorPoint = None

        if not self.coordCtl3.Coordinate.Is3D :
            self.coordCtl3.StatusMessage = "No valid coordinate defined, must be 3D"
        else:
            self.coordCtl3.StatusMessage = ""

        self.drawoverlay()

    def Coord4Changed(self, ctrl, e):
        # set keyboard focus if change was due to mouse pick
        if e.Cause == InputMethod.Mouse:     
            self.OkClicked(None, None)

        if not self.coordCtl4.Coordinate.Is3D :
            self.coordCtl4.StatusMessage = "No valid coordinate defined, must be 3D"
        else:
            self.coordCtl4.StatusMessage = ""

    def Coord5Changed(self, ctrl, e):
        # set keyboard focus if change was due to mouse pick
        if e.Cause == InputMethod.Mouse:     
            self.OkClicked(None, None)

        if not self.coordCtl5.Coordinate.Is3D :
            self.coordCtl5.StatusMessage = "No valid coordinate defined, must be 3D"
        else:
            self.coordCtl5.StatusMessage = ""

    def Coord6Changed(self, ctrl, e):
        # set keyboard focus if change was due to mouse pick
        if e.Cause == InputMethod.Mouse:     
            self.OkClicked(None, None)

        if not self.coordCtl6.Coordinate.Is3D :
            self.coordCtl6.StatusMessage = "No valid coordinate defined, must be 3D"
        else:
            self.coordCtl6.StatusMessage = ""
        
    def drawoverlay(self):

        wv = self.currentProject [Project.FixedSerial.WorldView]
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)
        self.overlayBag = OverlayBag(self.ViewOverlay) # create a new/empty overlaybag

        if self.coordCtl1.Coordinate.Is3D and self.coordCtl2.Coordinate.Is3D and self.coordCtl3.Coordinate.Is3D:
       
            self.overlayBag.AddPolyline(Array[Point3D]([self.coordCtl1.Coordinate, self.coordCtl2.Coordinate, self.coordCtl3.Coordinate, \
                                                        self.coordCtl1.Coordinate]), Color.Blue.ToArgb(), 5)

            self.overlayBag.AddMarker(self.coordCtl1.Coordinate, GraphicMarkerTypes.HollowCircle_IndependentColor, Color.Magenta.ToArgb(), "   V1", 0, 0, 2.0) # last 2 numbers, markercircle-rotation/scale
            self.overlayBag.AddMarker(self.coordCtl2.Coordinate, GraphicMarkerTypes.HollowCircle_IndependentColor, Color.Magenta.ToArgb(), "   V2", 0, 0, 2.0) # last 2 numbers, markercircle-rotation/scale
            self.overlayBag.AddMarker(self.coordCtl3.Coordinate, GraphicMarkerTypes.HollowCircle_IndependentColor, Color.Magenta.ToArgb(), "   V3", 0, 0, 2.0) # last 2 numbers, markercircle-rotation/scale
            
            # if it only needs to be visible in all Planview then remove the Hoops3DViewGUID
            array = Array[Guid]([DisplayWindow.Hoops3DViewGUID, DisplayWindow.HoopsPlanViewGUID])
            TrimbleOffice.TheOffice.MainWindow.AppViewManager.AddOverlayGeometry(array, self.overlayBag)

        return

    def CancelClicked(self, thisCmd, args):
        thisCmd.CloseUICommand()

    def Dispose(self, cmd, disposing):
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)

    def OkClicked(self, thisCmd, e):

        self.success.Content=''

        wv = self.currentProject[Project.FixedSerial.WorldView]

        pointlist = []
        if self.projectmultipoints.IsChecked:
            for o in self.objs.SelectedMembers(self.currentProject):
                if isinstance(o, self.cadpointType):
                    pointlist.Add(o.Point0)
                elif isinstance(o, self.coordpointType):
                    pointlist.Add(o.Position)
                elif isinstance(o, self.pointcloudType):
                    integration = o.Integration  # = SdePointCloudRegionIntegration
                    selectedid = integration.GetSelectedCloudId() # it seems the selected points form a sub-cloud

                    regiondb = integration.PointCloudDatabase # PointCloudDatabase
                    sdedb = regiondb.Integration # SdePointCloudDatabaseIntegration
                    scanpointlist = sdedb.GetPoints(selectedid)
                    for p in scanpointlist:
                        pointlist.Add(p)

        elif self.projectsinglepoint.IsChecked:
            p4_sel = self.coordCtl4.Coordinate
            if not p4_sel.Is3D:
                self.coordCtl4.StatusMessage = "No valid coordinate defined, must be 3D"
            else:
                self.coordCtl4.StatusMessage = ""
                pointlist.Add(p4_sel)

        inputok=True
        try:
            self.th = self.textheight.Distance
        except:
            self.success.Content='error in Text Height'
            inputok=False
        
        if inputok:
            self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
            UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
            try:
                with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                    
                    if self.coordCtl1.Coordinate.Is3D and self.coordCtl2.Coordinate.Is3D and self.coordCtl3.Coordinate.Is3D:

                        p1_sel = self.coordCtl1.Coordinate
                        p2_sel = self.coordCtl2.Coordinate
                        p3_sel = self.coordCtl3.Coordinate

                        p = Plane3D(p1_sel,p2_sel,p3_sel)
                        p = p[0] # the plane is returned as first element

                            ### v = p.normal
                            ### if v.Length < Point3D.Tolerance:
                            ###     return
                            ### v.Normalize()
                            ### 
                            ### #   we use the normal vector to fill the "Hesse" plane formula
                            ### #   since we use the normalized vector the denominator is 1 and we can ignore it 
                            ### 
                            ### #   v.x * x  +  v.y * y + v.z * z
                            ### #   -----------------------------  =  d  =  nx * x  +  ny * y + nz * z
                            ### #             |v| = 1
                            ### 
                            ### #  now we use our point instead one of the vertexes in the plane formula and the resulting value is the
                            ### #  3D distance to that plane
                            ### 
                            ### #  d3d =  d - nx * x  +  ny * y + nz * z
                            ### 
                            ### 
                            ### d=(p1_sel.X * v.X + p1_sel.Y * v.Y + p1_sel.Z * v.Z)
                            ### d3d = (d - p4_sel.X * v.X - p4_sel.Y * v.Y - p4_sel.Z * v.Z)
                            ### 
                            ### pnew=Point3D(0,0,0)
                            ### pnew.X = p4_sel.X + d3d * v.X
                            ### pnew.Y = p4_sel.Y + d3d * v.Y
                            ### pnew.Z = p4_sel.Z + d3d * v.Z
                        
                        if not self.projectdirectionalshot.IsChecked and p.IsValid:
                            for p4_sel in pointlist:

                                if p4_sel.Is3D:
                                    if self.computeplumb.IsChecked:
                                        pnew = Plane3D.IntersectWithRay(p, p4_sel, Vector3D(0, 0, 1))
                                    else:
                                        pnew = Plane3D.IntersectWithRayPerpendicular(p, p4_sel)
                                    
                                    # if pnew != Point3D.Undefined
                                    self.drawtext(p4_sel, pnew)

                        elif self.projectdirectionalshot.IsChecked and p.IsValid:
                            p5_sel = self.coordCtl5.Coordinate
                            p6_sel = self.coordCtl6.Coordinate

                            if p5_sel.Is3D and p6_sel.Is3D:
                                pnew = Plane3D.IntersectWithRay(p, p5_sel, Vector3D(p5_sel, p6_sel))
                                self.drawtext(p6_sel, pnew)


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
        
        self.SaveOptions()
        if self.projectsinglepoint.IsChecked:
            Keyboard.Focus(self.coordCtl4)
        if self.projectdirectionalshot.IsChecked:
            Keyboard.Focus(self.coordCtl5)
 
    def drawtext(self, pfrom, pto):

        wv = self.currentProject[Project.FixedSerial.WorldView]
        d3d = Vector3D(pfrom, pto).Length

        if self.checkBox_point.IsChecked:
            cadPoint = wv.Add(clr.GetClrType(CadPoint))
            cadPoint.Point0 = pto
            cadPoint.Layer = self.layerpicker.SelectedSerialNumber
            cadPoint.Color = self.outcolorpicker.SelectedColor

        if self.checkBox_line.IsChecked:
            l = wv.Add(clr.GetClrType(Linestring))
            e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
            e.Position = pfrom 
            l.AppendElement(e)
            e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
            e.Position = pto
            l.AppendElement(e)       
            l.Layer = self.layerpicker.SelectedSerialNumber
            l.Color = self.outcolorpicker.SelectedColor

        if self.checkBox_text.IsChecked:
            textObj = wv.Add(clr.GetClrType(MText))
            textObj.TextString = self.tooutputunit(d3d)

            textObj.Height = self.th
            textObj.Point0 = pto
            textObj.Layer = self.layerpicker.SelectedSerialNumber
            textObj.Color = self.outcolorpicker.SelectedColor

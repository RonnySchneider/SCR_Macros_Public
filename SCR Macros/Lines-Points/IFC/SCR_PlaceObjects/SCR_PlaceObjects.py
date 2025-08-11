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
    cmdData.Key = "SCR_PlaceObjects"
    cmdData.CommandName = "SCR_PlaceObjects"
    cmdData.Caption = "_SCR_PlaceObjects"
    cmdData.UIForm = "SCR_PlaceObjects"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "IFC"
        cmdData.ShortCaption = "Place Objects"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.15
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "place Objects"
        cmdData.ToolTipTextFormatted = "place Objects"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png") 
        cmdData.ImageSmall = b
    except:
        pass

class SCR_PlaceObjects(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_PlaceObjects.xaml") as s:
            wpf.LoadComponent (self, s)
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
        self.Caption = cmd.Command.Caption

        self.lType = clr.GetClrType(IPolyseg)

        self.linepicker1.IsEntityValidCallback = self.IsValidAlignment
        self.linepicker1.ValueChanged += self.line1Changed
        self.startstation.ValueChanged += self.line1Changed
        self.endstation.ValueChanged += self.line1Changed
        self.chooserange.Checked += self.line1Changed
        self.chooserange.Unchecked += self.line1Changed

        types = Array [Type] (SurfaceTypeLists.AllWithCutFillMap)    # we fill an array with TBC object types, we could combine different types
        self.surfacepicker.FilterByEntityTypes = types    # we fill the dropdownlist by applying that types array as filter
        self.surfacepicker.AllowNone = False              # our list shall not show an empty field

        self.objs.IsEntityValidCallback = self.IsValidObject

        self.cadpointType = clr.GetClrType(CadPoint)
        self.coordpointType = clr.GetClrType(CoordPoint)
        self.pointcloudType = clr.GetClrType(PointCloudRegion)

        self.point_objs.IsEntityValidCallback = self.IsValidPointObject
        #### get the units for linear distance
        ###self.lunits = self.currentProject.Units.Linear
        ####self.lfp = self.lunits.Properties.Copy()
        ###linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation
        ####self.lfp.AddSuffix = False
        ###self.girderwidthlabel.Content = "Girder Width [" + linearsuffix + "]"
        ###self.girderlengthlabel.Content = "Girder Length [" + linearsuffix + "]"
        ###self.girderskewlabel.Content = "Girder Skew [" + linearsuffix + "]"

        self.coordpick1.ShowElevationIf3D = True
        self.coordpick2.ShowElevationIf3D = True

        self.interval.FormatProperty.AddSuffix = ControlBoolean(1)
        self.horizontaloffset.FormatProperty.AddSuffix = ControlBoolean(1)
        self.elevationoffset.FormatProperty.AddSuffix = ControlBoolean(1)

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

        self.ignoreGotFocus = False
        self.selectionControls = [self.point_objs, self.objs]

    def SetDefaultOptions(self):
        wv = self.currentProject [Project.FixedSerial.WorldView]
        
        self.relayer.IsChecked = OptionsManager.GetBool("SCR_PlaceObjects.relayer", False)
        lserial = OptionsManager.GetUint("SCR_PlaceObjects.layerpicker", 8)
        o = self.currentProject.Concordance.Lookup(lserial) # get the object withj the saved serial number
        if o != None:   # check if we actually got an object back
            if isinstance(o.GetSite(), LayerCollection):    # check if the object is a layer
                self.layerpicker.SetSelectedSerialNumber(lserial, InputMethod(3))
            else:  # if not then set Layer zero                       
                self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:   # if not then set Layer zero               
            self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))

        self.linemode.IsChecked = OptionsManager.GetBool("SCR_PlaceObjects.linemode", True)
        self.chooserange.IsChecked = OptionsManager.GetBool("SCR_PlaceObjects.chooserange", False)
        self.startstation.SetStation(OptionsManager.GetDouble("SCR_PlaceObjects.startstation", 0.000), self.currentProject)
        self.endstation.SetStation(OptionsManager.GetDouble("SCR_PlaceObjects.endstation", 0.000), self.currentProject)
        self.interval.Distance = OptionsManager.GetDouble("SCR_PlaceObjects.interval", 1.0000)
        self.horizontaloffset.Distance = OptionsManager.GetDouble("SCR_PlaceObjects.horizontaloffset", 0.0000)
        self.tiltobject.IsChecked = OptionsManager.GetBool("SCR_PlaceObjects.tiltobject", False)
        self.pointmode.IsChecked = OptionsManager.GetBool("SCR_PlaceObjects.pointmode", False)
        self.generalorientation.IsChecked = OptionsManager.GetBool("SCR_PlaceObjects.generalorientation", True)
        self.genrot.Angle = OptionsManager.GetDouble("SCR_PlaceObjects.genrot", 0.0)
        self.squaretoline.IsChecked = OptionsManager.GetBool("SCR_PlaceObjects.squaretoline", False)
        self.lineelevation.IsChecked = OptionsManager.GetBool("SCR_PlaceObjects.lineelevation", True)
        self.surfaceelevation.IsChecked = OptionsManager.GetBool("SCR_PlaceObjects.surfaceelevation", False)
        # Select surface
        try:    self.surfacepicker.SelectIndex(OptionsManager.GetInt("SCR_PlaceObjects.surfacepicker", 0))
        except: self.surfacepicker.SelectIndex(0)
        self.elevationoffset.Elevation = OptionsManager.GetDouble("SCR_PlaceObjects.elevationoffset", 0.0000)

        x = OptionsManager.GetDouble("SCR_PlaceObjects.coordpick1_x", 0.0)
        y = OptionsManager.GetDouble("SCR_PlaceObjects.coordpick1_y", 0.0)
        z = OptionsManager.GetDouble("SCR_PlaceObjects.coordpick1_z", 0.0)
        self.coordpick1.SetCoordinate(Point3D(x, y, z), self.currentProject, wv.CoordinateSystemDefinition)
        x = OptionsManager.GetDouble("SCR_PlaceObjects.coordpick2_x", 0.0)
        y = OptionsManager.GetDouble("SCR_PlaceObjects.coordpick2_y", 0.0)
        z = OptionsManager.GetDouble("SCR_PlaceObjects.coordpick2_z", 0.0)
        self.coordpick2.SetCoordinate(Point3D(x, y, z), self.currentProject, wv.CoordinateSystemDefinition)

    def SaveOptions(self):

        OptionsManager.SetValue("SCR_PlaceObjects.relayer", self.relayer.IsChecked)
        OptionsManager.SetValue("SCR_PlaceObjects.relayer", self.layerpicker.SelectedSerialNumber)
        OptionsManager.SetValue("SCR_PlaceObjects.linemode", self.linemode.IsChecked)
        OptionsManager.SetValue("SCR_PlaceObjects.chooserange", self.chooserange.IsChecked)
        OptionsManager.SetValue("SCR_PlaceObjects.startstation", self.startstation.Distance)
        OptionsManager.SetValue("SCR_PlaceObjects.endstation", self.endstation.Distance)
        OptionsManager.SetValue("SCR_PlaceObjects.interval", abs(self.interval.Distance))
        OptionsManager.SetValue("SCR_PlaceObjects.horizontaloffset", abs(self.horizontaloffset.Distance))
        OptionsManager.SetValue("SCR_PlaceObjects.tiltobject", self.tiltobject.IsChecked)
        OptionsManager.SetValue("SCR_PlaceObjects.pointmode", self.pointmode.IsChecked)
        OptionsManager.SetValue("SCR_PlaceObjects.generalorientation", self.generalorientation.IsChecked)
        OptionsManager.SetValue("SCR_PlaceObjects.genrot", self.genrot.Angle)
        OptionsManager.SetValue("SCR_PlaceObjects.squaretoline", self.squaretoline.IsChecked)
        OptionsManager.SetValue("SCR_PlaceObjects.lineelevation", self.lineelevation.IsChecked)
        OptionsManager.SetValue("SCR_PlaceObjects.surfaceelevation", self.surfaceelevation.IsChecked)
        OptionsManager.SetValue("SCR_PlaceObjects.surfacepicker", self.surfacepicker.SelectedIndex)
        OptionsManager.SetValue("SCR_PlaceObjects.elevationoffset", self.elevationoffset.Elevation)

        OptionsManager.SetValue("SCR_PlaceObjects.coordpick1_x", self.coordpick1.Coordinate.X)
        OptionsManager.SetValue("SCR_PlaceObjects.coordpick1_y", self.coordpick1.Coordinate.Y)
        OptionsManager.SetValue("SCR_PlaceObjects.coordpick1_z", self.coordpick1.Coordinate.Z)
        OptionsManager.SetValue("SCR_PlaceObjects.coordpick2_x", self.coordpick2.Coordinate.X)
        OptionsManager.SetValue("SCR_PlaceObjects.coordpick2_y", self.coordpick2.Coordinate.Y)
        OptionsManager.SetValue("SCR_PlaceObjects.coordpick2_z", self.coordpick2.Coordinate.Z)

    def drawoverlay(self):

        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)
        self.overlayBag = OverlayBag(self.ViewOverlay) # create a new/empty overlaybag

        l1 = self.linepicker1.Entity

        if l1 != None:
            self.overlayBag.AddPolyline(self.getpolypoints(l1), Color.Green.ToArgb(), 4)

            if self.chooserange.IsChecked:
                self.overlayBag.AddPolyline(self.getclippedpolypoints(l1, self.startstation.Distance, self.endstation.Distance), Color.Orange.ToArgb(), 2)

            for p in self.getarrowlocations(l1, 10): # returns list with location and perp right azimuth [Point3D, perpVector3D.Value.Azimuth]
                self.overlayBag.AddMarker(p[0], GraphicMarkerTypes.Arrow_IndependentColor, Color.Orange.ToArgb(), "", 0, math.pi - p[1], 3.0)



        # if it only needs to be visible in all Planview then remove the Hoops3DViewGUID
        array = Array[Guid]([DisplayWindow.Hoops3DViewGUID, DisplayWindow.HoopsPlanViewGUID])
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.AddOverlayGeometry(array, self.overlayBag)

        return

    def getarrowlocations(self, l1, intervals):

        pts = []

        polyseg = l1.ComputePolySeg().Clone()
        polyseg = polyseg.ToWorld()
        polyseg_v = l1.ComputeVerticalPolySeg().Clone()
        
        interval = polyseg.ComputeStationing() / intervals
        
        computestation = interval
        while computestation < polyseg.ComputeStationing():
            outSegment = clr.StrongBox[Segment]()
            out_t = clr.StrongBox[float]()
            outPointOnCL = clr.StrongBox[Point3D]()
            perpVector3D = clr.StrongBox[Vector3D]()
            outdeflectionAngle = clr.StrongBox[float]()
            
            polyseg.FindPointFromStation(computestation, outSegment, out_t, outPointOnCL, perpVector3D, outdeflectionAngle)
            p = outPointOnCL.Value
            if polyseg_v != None:
                p.Z = polyseg_v.ComputeVerticalSlopeAndGrade(computestation)[1]

            pts.Add([p, perpVector3D.Value.Azimuth])

            computestation += interval

        return pts

    def getpolypoints(self, l):

        if l != None:
            polyseg = l.ComputePolySeg().Clone()
            polyseg = polyseg.ToWorld()
            polyseg_v = l.ComputeVerticalPolySeg().Clone()
            if not polyseg_v and not polyseg.AllPointsAre3D:
                polyseg_v = PolySeg.PolySeg()
                polyseg_v.Add(Point3D(polyseg.BeginStation,0,0))
                polyseg_v.Add(Point3D(polyseg.ComputeStationing(), 0, 0))
            chord = polyseg.Linearize(0.001, 0.001, 50, polyseg_v, False)

        return chord.ToPoint3DArray()
    

    # need to Clone the Polyseg upon creation
    # otherwise something weird is going on
    # and the functions above (getpoly and arrow) may return clipped sections
    # as if the ClipStationRange somehow transpires through to the original Linestring
    # although the Linestring stays unchanged on screen
    # no clue what is going on
    def getclippedpolypoints(self, l, startstation, endstation):
        
        if l != None:
            
            if startstation > endstation: startstation, endstation = endstation, startstation 
            polyseg = l.ComputePolySeg().Clone()
            polyseg = polyseg.ToWorld()
            if math.isnan(endstation):
                endstation = polyseg.ComputeStationing()
            polyseg.ClipStationRange(startstation, endstation, True)
            polyseg.Trim()
            tt = polyseg.ComputeStationing()
            
            polyseg_v = l.ComputeVerticalPolySeg().Clone()
            if not polyseg_v:
                polyseg_v = PolySeg.PolySeg()
                polyseg_v.Add(Point3D(polyseg.BeginStation,0,0))
                polyseg_v.Add(Point3D(polyseg.ComputeStationing(), 0, 0))

            minmax = CommandHelper.GetMinMaxElevations(l)

            polyseg_v.Clip(Limits3D(Point3D(startstation, minmax[1]-1, 0), Point3D(endstation, minmax[2]+1, 0)), Side.Out)
            polyseg_v.Trim()

            chord = polyseg.Linearize(0.001, 0.001, 50, polyseg_v, False)

        return chord.ToPoint3DArray()

    def IsValidAlignment(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        return False

    def IsValidPointObject(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.cadpointType):
            return True
        if isinstance(o, self.coordpointType):
            return True
        if isinstance(o, self.pointcloudType):
            return True
        return False

    def IsValidObject(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        tt = o.GetType()
        if isinstance(o, IFCMesh):
            return True
        if isinstance(o, BIMEntity):
            return True
        if isinstance(o, Linestring):
            return True
        return False

    def Selection_PreviewGotFocus(self, sender, e):
        self.ignoreGotFocus = True
        for ctrl in self.selectionControls:
            value = (sender == ctrl)
            ctrl.ProcessGlobalSelectionChanges = value
            ctrl.UpdateTextOnSelectionChange = value
        pass

    def Selection_ValueChanged(self, sender, e):
        #There is a bug in the Selection control, it resets the selection on the "GotFocus" event and then sets it back to the proper selection
        if self.ignoreGotFocus:
            self.ignoreGotFocus = False
            return

    def line1Changed(self, ctrl, e):
        l1=self.linepicker1.Entity
        if l1 != None:
            self.startstation.StationProvider = l1
            self.endstation.StationProvider = l1
        self.drawoverlay()

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def Dispose(self, cmd, disposing):
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content = ''
        self.success.Content = ''
        
        wv = self.currentProject [Project.FixedSerial.WorldView]

        inputok = True

        if self.linemode.IsChecked:
            l1 = self.linepicker1.Entity

            if l1 == None: 
                self.error.Content += '\nno Line selected'
                inputok = False

            #if o_org == None: 
            #    self.error.Content += '\nno Object selected'
            #    inputok = False

            if self.interval.Distance == 0.0: 
                self.error.Content += '\nenter valid interval'
                inputok = False

            try:
                startstation = self.startstation.Distance
            except:
                self.error.Content += '\nStart Chainage error'
                inputok=False
            try:
                endstation = self.endstation.Distance
            except:
                self.error.Content += '\nEnd Chainage error'
                inputok=False

            if startstation > endstation: startstation, endstation = endstation, startstation   # swap values if necessary

        if self.pointmode.IsChecked and self.squaretoline.IsChecked:
            l2 = self.linepicker2.Entity

            if l2 == None: 
                self.error.Content += '\nno Line selected'
                inputok = False


        if math.isnan(self.elevationoffset.Elevation):
            self.elevationoffset.Elevation = 0.0
      
        if inputok:

            UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
            self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)

            outSegment1 = clr.StrongBox[Segment]()
            out_t = clr.StrongBox[float]()
            outPointOnCL1 = clr.StrongBox[Point3D]()
            perpVector3D = clr.StrongBox[Vector3D]()
            outdeflectionAngle = clr.StrongBox[float]()
            
            surface = wv.Lookup(self.surfacepicker.SelectedSerial)

            try:

                with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                    ps1 = self.coordpick1.Coordinate
                    ps2 = self.coordpick2.Coordinate

                    if self.linemode.IsChecked:

                        polyseg1 = l1.ComputePolySeg()
                        polyseg1 = polyseg1.ToWorld()
                        polyseg1_v = l1.ComputeVerticalPolySeg()

                        if self.chooserange.IsChecked:
                            computestartstation = startstation
                            computeendstation = endstation
                        else:
                            computestartstation = polyseg1.BeginStation
                            computeendstation = polyseg1.ComputeStationing()

                        # compile chainage list
                        chainages = []
                        # add the interval range on the first line
                        if abs(self.interval.Distance) != 0:
                            while computestartstation <= computeendstation:    # as long as we aren't at the end of the line 1
                                chainages.Add(computestartstation)

                                if computestartstation == computeendstation: break
                                computestartstation += abs(self.interval.Distance)
                                if computestartstation > computeendstation: computestartstation = computeendstation
                        j = 0
                        for computestation in chainages:    # go through the chainage list and compute the transformation matrix if possible
                            j += 1
                            ProgressBar.TBC_ProgressBar.Title = "Processing Chainage " + str(j) + ' / ' + str(chainages.Count)
                            if ProgressBar.TBC_ProgressBar.SetProgress(j * 100 // chainages.Count):
                                break
                            
                            # compute a point on line 1 and get the perpendicular vector
                            polyseg1.FindPointFromStation(computestation, outSegment1, out_t, outPointOnCL1, perpVector3D, outdeflectionAngle) # compute point and vector on line 1

                            p1 = outPointOnCL1.Value

                            if polyseg1_v != None:
                                p1.Z = polyseg1_v.ComputeVerticalSlopeAndGrade(computestation)[1] + self.elevationoffset.Elevation
                                # if the station is exactly on a vertical node the slope will be 0
                                s = polyseg1_v.FirstSegment
                                while s is not None:
                                    if computestation >= s.BeginPoint.X and computestation < s.EndPoint.X:
                                        lineslope = Vector3D(s.BeginPoint, s.EndPoint)
                                    s = polyseg1_v.Next(s)

                            # original pick system
                            ovf = Vector3D(ps1, ps2)
                            ovr = ovf.Clone()
                            ovr.Rotate90(Side.Right)
                            # target system along line
                            tvr = perpVector3D.Value
                            tvf = tvr.Clone()
                            tvf.Rotate90(Side.Left)

                            if ovf.Is2D: ovf.Z = 0
                            if ovr.Is2D: ovr.Z = 0
                            if tvf.Is2D: tvf.Z = 0
                            if tvr.Is2D: tvr.Z = 0

                            svn = None
                            if self.surfaceelevation.IsChecked:

                                outPoint = clr.StrongBox[Point3D]()
                                outPrimitive = clr.StrongBox[Primitive]()
                                outInt = clr.StrongBox[Int32]()
                                outByte = clr.StrongBox[Byte]()
                                
                                if surface.PickSurface(p1, outPrimitive, outInt, outByte, outPoint):
                                    p1.Z = outPoint.Value.Z + self.elevationoffset.Elevation

                                    if self.surfaceelevation.IsChecked:
                                        # compute the normal vector at that location
                                        verticelist = []
                                        outstring = ''

                                        i = outInt.Value
                                        if surface.GetTriangleMaterial(i) != surface.NullMaterialIndex():
                                            verticelist.Add(surface.GetVertexPoint(surface.GetTriangleIVertex(i,0)))
                                            verticelist.Add(surface.GetVertexPoint(surface.GetTriangleIVertex(i,1)))
                                            verticelist.Add(surface.GetVertexPoint(surface.GetTriangleIVertex(i,2)))
                                        
                                        if verticelist.Count != 3:
                                            continue
                                        p = Plane3D(verticelist[0], verticelist[1], verticelist[2])[0]

                                        svn = p.normal # surface vector normal
                                        svr = svn.Clone() # surface vector right
                                        svr.Horizon = 0
                                        svr.Rotate90(Side.Right)
                                        svf = svn.Clone()   #surface vector front
                                        svf.Rotate(svr, -math.pi/2)

                                        slv = Vector3D(1,0,0) # slope vector in direction of line but with surface slope
                                        slv.Azimuth = tvf.Azimuth
                                        prof = PolySeg.PolySeg()
                                        prof.Add(p1)
                                        prof.Add(p1 + slv)
                                        prof = surface.Profile(prof, False)
                                        slv = Vector3D(prof.FirstSegment.BeginPoint, prof.FirstSegment.EndPoint)

                                else:
                                    self.error.Content += "\ncouldn't compute a surface elevation for all locations"

                            #BuildTransformMatrix(Trimble.Vce.Geometry.Point3D fromPnt, Trimble.Vce.Geometry.Point3D toPnt, double rotate, double scaleX, double scaleY, double scaleZ)
                            # this is the direction of the alignment at the paste point

                            if self.tiltobject.IsChecked:
                                if svn and self.surfaceelevation.IsChecked:

                                    tt = Vector3D.AngleBetweenSigned(svf, slv, svn)
                                    svf.Rotate(svn, tt)
                                    svr.Rotate(svn, tt)

                                    tvf = svf.Clone()
                                    tvr = svr.Clone()

                                else:
                                    if lineslope.X == 0:
                                        tvf.Horizon = 0
                                    else:
                                        tvf.Horizon = math.atan(lineslope.Y / lineslope.X) # lineslope is a 2D profile view vector
                            
                            targetrot = Spinor3D.ComputeRotation(ovf, ovr, tvf, tvr)
                            
                            offsetvector = tvr.Clone()
                            if math.isnan(self.horizontaloffset.Distance):
                                offsetvector.Length = 0
                            else:
                                offsetvector.Length = self.horizontaloffset.Distance
                            p1 = p1 + offsetvector
                            targetmatrix = Matrix4D.BuildTransformMatrix(Vector3D(ps1), Vector3D(ps1, p1), targetrot, Vector3D(1,1,1))
                            
                            #linevec.Z = 0
                            #targetrot = Spinor3D.ComputeRotation(Vector3D(ps1, ps2), Vector3D(0, 0, 1), linevec, Vector3D(0, 0, 1))
                            #targetmatrix = Matrix4D.BuildTransformMatrix(Vector3D(ps1), Vector3D(ps1, p1), targetrot, Vector3D(1,1,1))

                            self.transform_obj(targetmatrix)
                            #self.reloadprojectexplorer()

                    elif self.pointmode.IsChecked:
                        pointlist = []

                        for o in self.point_objs.SelectedMembers(self.currentProject):
                            if isinstance(o, self.cadpointType):
                                pointlist.Add(o.Point0)
                            #elif isinstance(o, self.point3dType):
                            #    tt=o
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
                        
                        for p in pointlist:
                            
                            if not p.Is3D:
                                p.Z = 0 + self.elevationoffset.Elevation
                            else:
                                p.Z += self.elevationoffset.Elevation

                            if self.surfaceelevation.IsChecked:
                                surfp = surface.PickSurface(p)
                                if surfp[0]:
                                    p.Z = surfp[1].Z + self.elevationoffset.Elevation
                                else:
                                    self.error.Content += "\ncouldn't compute a surface elevation for all locations"
                            
                            #BuildTransformMatrix(Trimble.Vce.Geometry.Point3D fromPnt, Trimble.Vce.Geometry.Point3D toPnt, double rotate, double scaleX, double scaleY, double scaleZ)
                            # this is the direction of the alignment at the paste point
                            if self.generalorientation.IsChecked:
                                rotation = Vector3D(ps1, ps2).Azimuth - self.genrot.Angle
                            else:
                                polyseg2 = l2.ComputePolySeg()
                                polyseg2 = polyseg2.ToWorld()
                                outPointOnCL = clr.StrongBox[Point3D]()
                                outstation = clr.StrongBox[float]()
                                if polyseg2.FindPointFromPoint(p, outPointOnCL, outstation):
                                    rotation = Vector3D(ps1, ps2).Azimuth - Vector3D(p, outPointOnCL.Value).Azimuth
            
                                    #l = wv.Add(clr.GetClrType(Linestring))
                                    #e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                    #e.Position = p 
                                    #l.AppendElement(e)
                                    #e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                    #e.Position = outPointOnCL.Value
                                    #l.AppendElement(e)

                                    
                            targetmatrix = Matrix4D.BuildTransformMatrix(ps1, p, rotation, 1.0, 1.0, 1.0)

                            self.transform_obj(targetmatrix)
                            #self.reloadprojectexplorer()

                        #cadPoint = wv.Add(clr.GetClrType(CadPoint))
                        #cadPoint.Point0 = p1

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

        self.SaveOptions()

    def transform_obj(self, targetmatrix):

        self.wv = self.currentProject [Project.FixedSerial.WorldView]
        self.dp = self.currentProject.CreateDuplicator()

        for o_org in self.objs:
            if isinstance(o_org, IFCMesh) :
                # create a copy of the selected object
                o_new = self.wv.Add(clr.GetClrType(o_org.GetType()))
                o_new.CopyBody(self.currentProject.Concordance, self.currentProject.TransactionManager, o_org, self.dp)                        
                if self.relayer.IsChecked:
                    o_new.Layer = self.layerpicker.SelectedSerialNumber
                o_new.Transform(TransformData(targetmatrix, None))

            elif isinstance(o_org, Linestring):

                polyseg = o_org.ComputePolySeg()
                polyseg = polyseg.ToWorld()
                polyseg_v = o_org.ComputeVerticalPolySeg()
                polyseg = polyseg.Linearize(0.0001, 0.0001, 1000, polyseg_v, False)

                l = self.wv.Add(clr.GetClrType(Linestring))
                l.Layer = o_org.Layer
                l.Name = o_org.Name
                l.Color = o_org.Color

                for p in polyseg.Point3Ds():
                    e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                    e.Position = targetmatrix.TransformPoint(p)
                    l.AppendElement(e)
                    

            elif isinstance(o_org, BIMEntity):
                
                newbimsite = o_org.GetSite()
                newbimentity = newbimsite.Add(BIMEntity)
                newbimentity.CopyBody(self.currentProject.Concordance, self.currentProject.TransactionManager, o_org, self.dp) 
                newbimentity.Name = o_org.Name + ' - Copy'
                newbimentity.Transform(TransformData(targetmatrix, None)) # need to transform the bimentity instead trying to transform the meshinstance

                self.recursivebuildandtransform(o_org, newbimentity, targetmatrix)
                # recursively rebuild the bimenity tree and 
                
                tt = o_org.GetChildEntities()

                tt2 = tt


        return
                
    
    def recursivebuildandtransform(self, prev_e, new_e, targetmatrix):

        for e in prev_e:
        
            if isinstance(e, BIMEntity):
                newbimsite = new_e
                newbimentity = newbimsite.Add(BIMEntity)
                newbimentity.CopyBody(self.currentProject.Concordance, self.currentProject.TransactionManager, e, self.dp) 
                newbimentity.Transform(TransformData(targetmatrix, None)) # need to transform the bimentity instead trying to transform the meshinstance
                self.recursivebuildandtransform(e, newbimentity, targetmatrix)

            elif isinstance(e, ShellMeshInstance):
                meshdata = e.GetShellMeshData()
                newmeshinstance = new_e.Add(ShellMeshInstance)
                #meshmatrix = Matrix4D.Multiply(e.GetSite().AffineTransformation, targetmatrix)
                newmeshinstance.CreateShell(0, meshdata.SerialNumber, e.Color, e.Transform)

                tt = 2


        return

    def reloadprojectexplorer(self):

        # now update the project explorer manually - haven't found a better way yet
        te = ExplorerData.TheExplorer
        
        ## find the bimnodecollection - node
        for i in te.Items.AllItems:
            if isinstance(i, BimEntityCollectionNode):
                bimcn = i
        
        bimcn.Expanded = True # must expand it, otherwise the allitems list won't be populated
        
        tt = [] # need temp list since we may run into an error while removing data from a live enumerator
        for i in bimcn.Items.AllItems:
            if isinstance(i, BimEntityNode):
                tt.Add(i)
        
        # now call removeitem on all bimnodes - that basically empties the bim collection node in the project explorer
        for i in tt:
            bimcn.RemoveItem(i)
        
        # trigger repopulate of all child items from bim collection node level
        # this will recreate the items list based on the project tree database we altered with addobserver/removeobserver
        bimcn.Populate()
        
        # re-expand the project explorer tree
        for i in bimcn.Items.AllItems:
            if isinstance(i, BimEntityNode) or isinstance(i, BimGroupNode):
                i.Expanded = True

        # redraw the project explorer ui
        te.UpdateData()

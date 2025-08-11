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
    cmdData.Key = "SCR_ElevateToPlane"
    cmdData.CommandName = "SCR_ElevateToPlane"
    cmdData.Caption = "_SCR_ElevateToPlane"
    cmdData.UIForm = "SCR_ElevateToPlane" # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "DTM"
        cmdData.ShortCaption = "Elevate to Plane"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.11
        cmdData.MacroAuthor = "SCR"
        cmdData.ToolTipTitle = "Elevate to Plane"
        cmdData.ToolTipTextFormatted = "Elevate Points or Lines to Plane"
    except:
        pass
    try:
        b = Bitmap(macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


# the name of this class must match name from cmdData.UIForm (above)
class SCR_ElevateToPlane(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader(macroFileFolder + r"\SCR_ElevateToPlane.xaml") as s:
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

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        #self.lfp = self.lunits.Properties.Copy()
        self.linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation

        self.toleranceheader.Content = 'vertical threshold between interpolated and true elevations [' + self.linearsuffix + ']'

        self.vertol.DistanceType = DistanceType.Z
        self.vertol.DistanceMin = 0.000001

        self.coordpick1.ValueChanged += self.Coord1Changed
        self.coordpick2.ValueChanged += self.Coord2Changed
        self.coordpick3.ValueChanged += self.Coord3Changed

        self.cadpointType = clr.GetClrType(CadPoint)
        self.coordpointType = clr.GetClrType(CoordPoint)
        self.lType = clr.GetClrType(IPolyseg)

        self.objs.IsEntityValidCallback = self.IsValid

        self.linepicker1.IsEntityValidCallback = self.IsValid2
        self.linepicker1.ValueChanged += self.lineChanged

        UIEvents.AfterDataProcessing += self.backgroundupdateend
		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass
        
        #self.unitssetup(None, None)

    def backgroundupdateend(self, sender, e):
        try:
            if sender.CommandName == "Undo" or sender.CommandName == "Redo":
                self.drawoverlay()
        except:
            pass

    def IsValid(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.coordpointType):
            return True
        if isinstance(o, self.cadpointType):
            return True
        if isinstance(o, self.lType):
            return True

        return False

    def IsValid2(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True

        return False

    def lineChanged(self, ctrl, e):
        l1 = self.linepicker1.Entity
        if l1 != None:
            self.startstation.StationProvider = l1
            self.endstation.StationProvider = l1

        self.drawoverlay()

    def SetDefaultOptions(self):
        
        self.vertol.Distance = OptionsManager.GetDouble("SCR_ElevateToPlane.vertol", 0.0001)

        lserial = OptionsManager.GetUint("SCR_ElevateToPlane.layerpicker", 8)
        o = self.currentProject.Concordance.Lookup(lserial)
        if o != None:
            if isinstance(o.GetSite(), LayerCollection):    
                self.layerpicker.SetSelectedSerialNumber(lserial, InputMethod(3))
            else:                       
                self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:                       
            self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))

        wv = self.currentProject[Project.FixedSerial.WorldView]
        x = OptionsManager.GetDouble("SCR_ElevateToPlane.coordpick1_x", 0.0)
        y = OptionsManager.GetDouble("SCR_ElevateToPlane.coordpick1_y", 0.0)
        z = OptionsManager.GetDouble("SCR_ElevateToPlane.coordpick1_z", 0.0)
        self.coordpick1.SetCoordinate(Point3D(x, y, z), self.currentProject, wv.CoordinateSystemDefinition)
        x = OptionsManager.GetDouble("SCR_ElevateToPlane.coordpick2_x", 0.0)
        y = OptionsManager.GetDouble("SCR_ElevateToPlane.coordpick2_y", 0.0)
        z = OptionsManager.GetDouble("SCR_ElevateToPlane.coordpick2_z", 0.0)
        self.coordpick2.SetCoordinate(Point3D(x, y, z), self.currentProject, wv.CoordinateSystemDefinition)
        x = OptionsManager.GetDouble("SCR_ElevateToPlane.coordpick3_x", 0.0)
        y = OptionsManager.GetDouble("SCR_ElevateToPlane.coordpick3_y", 0.0)
        z = OptionsManager.GetDouble("SCR_ElevateToPlane.coordpick3_z", 0.0)
        self.coordpick3.SetCoordinate(Point3D(x, y, z), self.currentProject, wv.CoordinateSystemDefinition)

        self.multiobjs.IsChecked = OptionsManager.GetBool("SCR_ElevateToPlane.multiobjs", True)
        self.linewithlimits.IsChecked = OptionsManager.GetBool("SCR_ElevateToPlane.linewithlimits", False)

    def SaveOptions(self):
        OptionsManager.SetValue("SCR_ElevateToPlane.vertol", self.vertol.Distance)
        OptionsManager.SetValue("SCR_ElevateToPlane.layerpicker", self.layerpicker.SelectedSerialNumber)

        OptionsManager.SetValue("SCR_ElevateToPlane.coordpick1_x", self.coordpick1.Coordinate.X)
        OptionsManager.SetValue("SCR_ElevateToPlane.coordpick1_y", self.coordpick1.Coordinate.Y)
        OptionsManager.SetValue("SCR_ElevateToPlane.coordpick1_z", self.coordpick1.Coordinate.Z)
        OptionsManager.SetValue("SCR_ElevateToPlane.coordpick2_x", self.coordpick2.Coordinate.X)
        OptionsManager.SetValue("SCR_ElevateToPlane.coordpick2_y", self.coordpick2.Coordinate.Y)
        OptionsManager.SetValue("SCR_ElevateToPlane.coordpick2_z", self.coordpick2.Coordinate.Z)
        OptionsManager.SetValue("SCR_ElevateToPlane.coordpick3_x", self.coordpick3.Coordinate.X)
        OptionsManager.SetValue("SCR_ElevateToPlane.coordpick3_y", self.coordpick3.Coordinate.Y)
        OptionsManager.SetValue("SCR_ElevateToPlane.coordpick3_z", self.coordpick3.Coordinate.Z)
        OptionsManager.SetValue("SCR_ElevateToPlane.multiobjs", self.multiobjs.IsChecked)
        OptionsManager.SetValue("SCR_ElevateToPlane.linewithlimits", self.linewithlimits.IsChecked)

    def Coord1Changed(self, ctrl, e):
        self.coordpick2.CursorStyle = CursorStyle.CrossHair | CursorStyle.RubberLine
        if self.coordpick1.ResultCoordinateSystem:
            self.coordpick2.AnchorPoint = MousePosition(self.coordpick1.ClickWindow, self.coordpick1.Coordinate, self.coordpick1.ResultCoordinateSystem)
        else:
            self.coordpick2.AnchorPoint = None

        if not self.coordpick1.Coordinate.Is3D :
            self.coordpick1.StatusMessage = "No valid coordinate defined, must be 3D"
        else:
            self.coordpick1.StatusMessage = ""

        self.drawoverlay()

    def Coord2Changed(self, ctrl, e):
        self.coordpick3.CursorStyle = CursorStyle.CrossHair | CursorStyle.RubberLine
        if self.coordpick2.ResultCoordinateSystem:
            self.coordpick3.AnchorPoint = MousePosition(self.coordpick2.ClickWindow, self.coordpick2.Coordinate, self.coordpick2.ResultCoordinateSystem)
        else:
            self.coordpick3.AnchorPoint = None

        if not self.coordpick2.Coordinate.Is3D :
            self.coordpick2.StatusMessage = "No valid coordinate defined, must be 3D"
        else:
            self.coordpick2.StatusMessage = ""

        self.drawoverlay()

    def Coord3Changed(self, ctrl, e):
        #self.coordCtl4.CursorStyle = CursorStyle.CrossHair | CursorStyle.RubberLine
        #if self.coordpick3.ResultCoordinateSystem:
        #    self.coordCtl4.AnchorPoint = MousePosition(self.coordpick3.ClickWindow, self.coordpick3.Coordinate, self.coordpick3.ResultCoordinateSystem)
        #else:
        #    self.coordCtl4.AnchorPoint = None

        if not self.coordpick3.Coordinate.Is3D :
            self.coordpick3.StatusMessage = "No valid coordinate defined, must be 3D"
        else:
            self.coordpick3.StatusMessage = ""

        self.drawoverlay()

        
    def drawoverlay(self):

        wv = self.currentProject [Project.FixedSerial.WorldView]
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)
        self.overlayBag = OverlayBag(self.ViewOverlay) # create a new/empty overlaybag

        if self.coordpick1.Coordinate.Is3D and self.coordpick2.Coordinate.Is3D and self.coordpick3.Coordinate.Is3D:
       
            self.overlayBag.AddPolyline(Array[Point3D]([self.coordpick1.Coordinate, self.coordpick2.Coordinate, self.coordpick3.Coordinate, \
                                                        self.coordpick1.Coordinate]), Color.Blue.ToArgb(), 5)

            self.overlayBag.AddMarker(self.coordpick1.Coordinate, GraphicMarkerTypes.HollowCircle_IndependentColor, Color.Magenta.ToArgb(), "   V1", 0, 0, 2.0) # last 2 numbers, markercircle-rotation/scale
            self.overlayBag.AddMarker(self.coordpick2.Coordinate, GraphicMarkerTypes.HollowCircle_IndependentColor, Color.Magenta.ToArgb(), "   V2", 0, 0, 2.0) # last 2 numbers, markercircle-rotation/scale
            self.overlayBag.AddMarker(self.coordpick3.Coordinate, GraphicMarkerTypes.HollowCircle_IndependentColor, Color.Magenta.ToArgb(), "   V3", 0, 0, 2.0) # last 2 numbers, markercircle-rotation/scale
            
        
        l1 = self.linepicker1.Entity
        if self.linewithlimits.IsChecked and l1 != None:
            self.overlayBag.AddPolyline(self.getpolypoints(l1), Color.Green.ToArgb(), 4)

        # if it only needs to be visible in all Planview then remove the Hoops3DViewGUID
        array = Array[Guid]([DisplayWindow.Hoops3DViewGUID, DisplayWindow.HoopsPlanViewGUID])
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.AddOverlayGeometry(array, self.overlayBag)

        return

    def getpolypoints(self, l):

        if l != None:
            polyseg = l.ComputePolySeg()
            polyseg = polyseg.ToWorld()
            polyseg_v = l.ComputeVerticalPolySeg()
            if not polyseg_v and not polyseg.AllPointsAre3D:
                polyseg_v = PolySeg.PolySeg()
                polyseg_v.Add(Point3D(polyseg.BeginStation,0,0))
                polyseg_v.Add(Point3D(polyseg.ComputeStationing(), 0, 0))
            chord = polyseg.Linearize(0.001, 0.001, 50, polyseg_v, False)

        return chord.ToPoint3DArray()

    def CancelClicked(self, thisCmd, args):
        thisCmd.CloseUICommand()

    def Dispose(self, cmd, disposing):
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)
        UIEvents.AfterDataProcessing -= self.backgroundupdateend

    def OkClicked(self, thisCmd, e):

        self.success.Content=''

        wv = self.currentProject[Project.FixedSerial.WorldView]

        if self.coordpick1.Coordinate.Is3D and self.coordpick2.Coordinate.Is3D and self.coordpick3.Coordinate.Is3D:
        
            self.p = Plane3D(self.coordpick1.Coordinate, self.coordpick2.Coordinate, self.coordpick3.Coordinate)[0] # the plane is returned as first element

            startstation = self.startstation.Distance
            endstation = self.endstation.Distance
            if endstation < startstation: startstation, endstation = endstation, startstation

            self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
            UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
            try:
                with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                    
                    if self.multiobjs.IsChecked:
                    
                        for o in self.objs:
    
                            if isinstance(o, self.coordpointType) or isinstance(o, self.cadpointType):
                        
                                o.Position = self.planeelevation(o.Position)

                            elif isinstance(o, self.lType):

                                # we are in multi objs
                                # so, remove all vertical information
                                self.removeverticalsfromlinestring(o, 0, 0)

                                polyseg = o.ComputePolySeg()
                                polyseg = polyseg.ToWorld()
                                #polyseg_v = o.ComputeVerticalPolySeg()

                                # drape the whole polyline onto the plane
                                polyseg_v_new = self.drapepolyseg(polyseg)
                                if isinstance(o, Linestring):
                                    # add the new vpi information to the original linestring
                                    self.addvpitolinestring(polyseg_v_new, o)
                                else:
                                    # draw new line    
                                    l = wv.Add(clr.GetClrType(Linestring))
                                    l.Layer = o.Layer
                                    l.Color = o.Color
                                    l.Append(polyseg, polyseg_v_new, False, False)
                    
                    elif self.linewithlimits.IsChecked and not (startstation and endstation == 0):

                        l = self.linepicker1.Entity

                        if not isinstance(l, Linestring):

                            el = 0 if math.isnan(l.Elevation) else l.Elevation
                           
                            polyseg = l.ComputePolySeg()
                            polyseg = polyseg.ToWorld()
                            polyseg_v = l.ComputeVerticalPolySeg()
                            if not polyseg_v and not polyseg.AllPointsAre3D:
                                polyseg_v = PolySeg.PolySeg()
                                polyseg_v.Add(Point3D(polyseg.BeginStation, el, 0))
                                polyseg_v.Add(Point3D(polyseg.ComputeStationing(), el, 0))
                            # draw new line    
                            lnew = wv.Add(clr.GetClrType(Linestring))
                            lnew.Layer = l.Layer
                            lnew.Color = l.Color
                            lnew.Append(polyseg, polyseg_v, False, False)
                            l = lnew

                        # for a single line remove vertical information between start and endstation
                        self.removeverticalsfromlinestring(l, startstation, endstation)

                        # compute a horizontal polyseg
                        polyseg = l.ComputePolySeg()
                        polyseg = polyseg.ToWorld()
                        # break it at the given stations
                        # Breaks a polyseg at the given stations and makes either inside or outside invisible
                        if polyseg.ClipStationRange(startstation, endstation, True): #true, keep segments between start/end
                            # drape only that part
                            polyseg_v_new = self.drapepolyseg(polyseg)
                            # only add this new drape information into the existing linestring
                            self.addvpitolinestring(polyseg_v_new, l)


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
        
        self.drawoverlay()
        self.SaveOptions()

    def addvpitolinestring(self, polyseg_v, l):

        for p in polyseg_v.Point3Ds():

            e = VerticalElementFactory.Create(clr.GetClrType(IVerticalNoCurve),  clr.GetClrType(IStationVerticalLocation))
            e.StationDistance = p.X
            e.Elevation = p.Y
            l.AddVerticalElement(e)

    def removeverticalsfromlinestring(self, o, startstation, endstation):

        # if it's a linestring we try to keep the original horizontal information
        # i.e. smoothed segment information gets lost when converted to a polyseg
        # so, remove only the vertical information and add just the new vertical polyseg
        if isinstance(o, Linestring):

            # remove all existing vertical information
            # from a potential vertical
            if self.multiobjs.IsChecked:
                for i in reversed(range(0, o.VerticalElementCount)):
                    o.RemoveVerticalElementAt(i)

            else: # for a single line only if it's between the selected stations
                vertinfo = o.GetVerticalElements()
                for i in reversed(range(0, vertinfo.Count)):
                    if vertinfo[i].StationDistance >= startstation and vertinfo[i].StationDistance < endstation:
                        o.RemoveVerticalElementAt(i)

            # remove all existing vertical information
            # from the horizontal elements and disconnect from PointID's if necessary
            polyseg = o.ComputePolySeg()
            for i in range(0, o.ElementCount):

                e = o.ElementAt(i, True) # True is Clone the element
                estation = self.computechainage(polyseg, e.Position)

                if self.multiobjs.IsChecked or (self.linewithlimits and estation >= startstation and estation < endstation):
                    etype = e.GetType()
                    if not "PointId" in etype.Name and not "Dependent" in etype.Name:
                        tt = e.Position
                        tt.To2D()
                        e.Position = tt
                        o.ReplaceElementAt(e, i)
                    else:
                        # create new elements in order to disconnect from points
                        if etype.Name == "PointIdStraightElement":
                            enew = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                        if etype.Name == "DependentSmoothCurveElement":
                            enew = ElementFactory.Create(clr.GetClrType(ISmoothCurveSegment), clr.GetClrType(IXYZLocation))
                        if etype.Name == "DependentBestFitArcElement":
                            enew = ElementFactory.Create(clr.GetClrType(IBestFitArcSegment), clr.GetClrType(IXYZLocation))
                        if etype.Name == "PointIdArcElement":
                            enew = ElementFactory.Create(clr.GetClrType(IArcSegment), clr.GetClrType(IXYZLocation))
                            enew.LargeSmall = e.LargeSmall
                            enew.LeftRight = e.LeftRight
                            enew.Radius = e.Radius
                        if etype.Name == "PointIdPIArcElement":
                            enew = ElementFactory.Create(clr.GetClrType(IPIArcSegment), clr.GetClrType(IXYZLocation))
                            enew.Radius = e.Radius
                        if etype.Name == "DependentTangentArcElement":
                            enew = ElementFactory.Create(clr.GetClrType(ITangentArcSegment), clr.GetClrType(IXYZLocation))
                            enew.TangentType = e.TangentType
                    
                        # that's the same for all of them
                        p = self.currentProject.Concordance.Lookup(e.LocationSerialNumber).Position
                        p.To2D()
                        enew.Position = p
                        o.ReplaceElementAt(enew, i)

    def computechainage(self, polyseg, p):
        
        outPointOnCL = clr.StrongBox[Point3D]()
        station = clr.StrongBox[float]()

        polyseg.FindPointFromPoint(p, outPointOnCL, station)

        return station.Value

    def drapepolyseg(self, polyseg):

        polyseg_v_new = PolySeg.PolySeg()
        s = polyseg.FirstSegment
        while s is not None:

            if s.Visible: # in case of single line the polyseg.ClipStationRange will set the unwanted section to invisible, but keep the stationing, nice side effect 

                if s.Type == SegmentType.Line: # basically keep/copy the old information to the new polyseg
                
                    polyseg_v_new.Add(Point3D(s.BeginStation, self.planeelevation(s.BeginPoint).Z))
                    polyseg_v_new.Add(Point3D(s.EndStation, self.planeelevation(s.EndPoint).Z))

                else:   # add additional vertical information if necessary                                 
                    polyseg_v_new.Add(Point3D(s.BeginStation, self.planeelevation(s.BeginPoint).Z))

                    inter_elevs = [] # tuples with t and elevation
                    # add start and end of segment to that list
                    inter_elevs.Add([0, self.planeelevation(s.BeginPoint).Z])
                    inter_elevs.Add([1, self.planeelevation(s.EndPoint).Z])
                    
                    hadtoadd = True # in order to get the test started
                    failsafe = 0
                    # now compute an interpolated elevation and compare with plane elevation in the middle between current elevation information
                    # if the difference is outside the threshold we add additional vertical information
                    while hadtoadd:
                        failsafe += 1
                        hadtoadd = False
                        new_inter_elevs = []
                        
                        for i in range(0, inter_elevs.Count - 1):
                            
                            t0 = inter_elevs[i][0]
                            t1 = inter_elevs[i + 1][0]

                            current_t = (inter_elevs[i + 1][0] + inter_elevs[i][0]) / 2
                            ptest_z = (inter_elevs[i + 1][1] + inter_elevs[i][1]) / 2 # interpolated z
                            ptestonplane_z = self.planeelevation(s.ComputePoint(current_t)[1]).Z # plane elevation at current t
                            
                            #tt = abs(ptest_z - ptestonplane_z)
                            if abs(ptest_z - ptestonplane_z) > self.vertol.Distance:
                                new_inter_elevs.Add([current_t, ptestonplane_z])
                    
                        if new_inter_elevs.Count > 0:
                            hadtoadd = True
                            for e in new_inter_elevs:
                                inter_elevs.Add(e)
                        
                            inter_elevs = list(set(tuple(element) for element in inter_elevs))
                            inter_elevs.sort()

                            #tt = 1

                        if failsafe > 1000:
                            hadtoadd = False

                    for e in inter_elevs:
                        #tt = s.Station(e[0])
                        polyseg_v_new.Add(Point3D(s.Station(e[0]), e[1]))

                
            s = polyseg.Next(s)

        return polyseg_v_new

    def planeelevation(self, pold):

        if not pold.Is3D:
            pold.Z = 0
        pnew = Plane3D.IntersectWithRay(self.p, pold, Vector3D(0, 0, 1))

        return pnew

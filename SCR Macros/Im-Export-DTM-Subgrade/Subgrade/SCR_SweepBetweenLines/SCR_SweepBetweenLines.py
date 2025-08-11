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
    cmdData.Key = "SCR_SweepBetweenLines"
    cmdData.CommandName = "SCR_SweepBetweenLines"
    cmdData.Caption = "_SCR_SweepBetweenLines"
    cmdData.UIForm = "SCR_SweepBetweenLines"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "Subgrade"
        cmdData.ShortCaption = "Sweep between Lines"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.12
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Create Cross Connections between two Lines"
        cmdData.ToolTipTextFormatted = "i.e. sweep around the corner at T-Intersections, or extend Subgrade"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_SweepBetweenLines(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_SweepBetweenLines.xaml") as s:
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
        #types = Array [Type] ([CadPoint]) + Array [Type] ([Point3D])    # we fill an array with TBC object types, we could combine different types
        self.linepicker1.IsEntityValidCallback = self.IsValid
        self.linepicker1.ValueChanged += self.lineChanged
        self.linepicker2.IsEntityValidCallback = self.IsValid
        self.linepicker2.ValueChanged += self.lineChanged
        self.startstation.ValueChanged += self.lineChanged
        self.endstation.ValueChanged += self.lineChanged
        self.chooserange.Checked += self.lineChanged
        self.chooserange.Unchecked += self.lineChanged
        self.lType = clr.GetClrType(IPolyseg)

        types = Array [Type] (SurfaceTypeLists.AllWithCutFillMap)    # we fill an array with TBC object types, we could combine different types
        self.designsurfacepicker.FilterByEntityTypes = types    # we fill the dropdownlist by applying that types array as filter
        self.designsurfacepicker.AllowNone = False              # our list shall not show an empty field

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass
        self.unitssetup(None, None)

    def SetDefaultOptions(self):
        lserial = OptionsManager.GetUint("SCR_SweepBetweenLines.layerpicker", 8)
        o = self.currentProject.Concordance.Lookup(lserial) # get the object withj the saved serial number
        if o != None:   # check if we actually got an object back
            if isinstance(o.GetSite(), LayerCollection):    # check if the object is a layer
                self.layerpicker.SetSelectedSerialNumber(lserial, InputMethod(3))
            else:  # if not then set Layer zero                       
                self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:   # if not then set Layer zero               
            self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        
        self.perptoline1.IsChecked = OptionsManager.GetBool("SCR_SweepBetweenLines.perptoline1", False)
        self.chooserange.IsChecked = OptionsManager.GetBool("SCR_SweepBetweenLines.chooserange", False)
        self.startstation.SetStation(OptionsManager.GetDouble("SCR_SweepBetweenLines.startstation", 0.000), self.currentProject)
        self.endstation.SetStation(OptionsManager.GetDouble("SCR_SweepBetweenLines.endstation", 0.000), self.currentProject)

        self.maxlinedistance.Distance = OptionsManager.GetDouble("SCR_SweepBetweenLines.maxlinedistance", 10.000)
        self.interval.Distance = OptionsManager.GetDouble("SCR_SweepBetweenLines.interval", 1.000)
        self.eloffset.Distance = OptionsManager.GetDouble("SCR_SweepBetweenLines.eloffset", 0.000)
        self.usesideslope.IsChecked = OptionsManager.GetBool("SCR_SweepBetweenLines.usesideslope", False)
        self.sideslope.Value = OptionsManager.GetDouble("SCR_SweepBetweenLines.sideslope", 0.000)
        self.slope1.IsChecked = OptionsManager.GetBool("SCR_SweepBetweenLines.slope1", True)
        self.slope2.IsChecked = OptionsManager.GetBool("SCR_SweepBetweenLines.slope2", False)

        self.extendgivenamount.IsChecked = OptionsManager.GetBool("SCR_SweepBetweenLines.extendgivenamount", True)
        self.extend.Distance = OptionsManager.GetDouble("SCR_SweepBetweenLines.extend", 1.000)


        self.extendtodtm.IsChecked = OptionsManager.GetBool("SCR_SweepBetweenLines.extendtodtm", True)
        try:    self.designsurfacepicker.SelectIndex(OptionsManager.GetInt("SCR_SweepBetweenLines.designsurfacepicker", 0))
        except: self.designsurfacepicker.SelectIndex(0)
        self.connectinputlines.IsChecked = OptionsManager.GetBool("SCR_SweepBetweenLines.connectinputlines", True)
        self.benchpattern.IsChecked = OptionsManager.GetBool("SCR_SweepBetweenLines.benchpattern", False)
        self.drawtextensions.IsChecked = OptionsManager.GetBool("SCR_SweepBetweenLines.drawtextensions", True)
        self.connectextensions.IsChecked = OptionsManager.GetBool("SCR_SweepBetweenLines.connectextensions", True)
        self.adddtmprofile.IsChecked = OptionsManager.GetBool("SCR_SweepBetweenLines.adddtmprofile", True)
        self.dtmprofilelength.Distance = OptionsManager.GetDouble("SCR_SweepBetweenLines.dtmprofilelength", 10.000)

    def SaveOptions(self):
        OptionsManager.SetValue("SCR_SweepBetweenLines.layerpicker", self.layerpicker.SelectedSerialNumber)
        OptionsManager.SetValue("SCR_SweepBetweenLines.perptoline1", self.perptoline1.IsChecked)

        OptionsManager.SetValue("SCR_SweepBetweenLines.chooserange", self.chooserange.IsChecked)
        OptionsManager.SetValue("SCR_SweepBetweenLines.startstation", self.startstation.Distance)
        OptionsManager.SetValue("SCR_SweepBetweenLines.endstation", self.endstation.Distance)

        OptionsManager.SetValue("SCR_SweepBetweenLines.maxlinedistance", self.maxlinedistance.Distance)
        OptionsManager.SetValue("SCR_SweepBetweenLines.interval", self.interval.Distance)
        OptionsManager.SetValue("SCR_SweepBetweenLines.eloffset", self.eloffset.Distance)
        OptionsManager.SetValue("SCR_SweepBetweenLines.usesideslope", self.usesideslope.IsChecked)
        OptionsManager.SetValue("SCR_SweepBetweenLines.sideslope", self.sideslope.Value)
        OptionsManager.SetValue("SCR_SweepBetweenLines.slope1", self.slope1.IsChecked)
        OptionsManager.SetValue("SCR_SweepBetweenLines.slope2", self.slope2.IsChecked)

        OptionsManager.SetValue("SCR_SweepBetweenLines.extendgivenamount", self.extendgivenamount.IsChecked)
        OptionsManager.SetValue("SCR_SweepBetweenLines.extend", self.extend.Distance)

        OptionsManager.SetValue("SCR_SweepBetweenLines.extendtodtm", self.extendtodtm.IsChecked)
        try:    # if nothing is selected it would throw an error
            OptionsManager.SetValue("SCR_SweepBetweenLines.designsurfacepicker", self.designsurfacepicker.SelectedIndex)
        except:
            pass
        OptionsManager.SetValue("SCR_SweepBetweenLines.connectinputlines", self.connectinputlines.IsChecked)
        OptionsManager.SetValue("SCR_SweepBetweenLines.benchpattern", self.benchpattern.IsChecked)
        OptionsManager.SetValue("SCR_SweepBetweenLines.drawtextensions", self.drawtextensions.IsChecked)
        OptionsManager.SetValue("SCR_SweepBetweenLines.connectextensions", self.connectextensions.IsChecked)
        OptionsManager.SetValue("SCR_SweepBetweenLines.adddtmprofile", self.adddtmprofile.IsChecked)
        OptionsManager.SetValue("SCR_SweepBetweenLines.dtmprofilelength", self.dtmprofilelength.Distance)
        
    def IsValid(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        return False

    def lineChanged(self, ctrl, e):
        l1 = self.linepicker1.Entity
        if l1 != None:
            self.startstation.StationProvider = l1
            self.endstation.StationProvider = l1
        self.drawoverlay()

    def drawoverlay(self):

        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)
        self.overlayBag = OverlayBag(self.ViewOverlay) # create a new/empty overlaybag

        l1 = self.linepicker1.Entity
        l2 = self.linepicker2.Entity

        if l1 != None:
            self.overlayBag.AddPolyline(self.getpolypoints(l1), Color.Red.ToArgb(), 4)

            if self.chooserange.IsChecked:
                self.overlayBag.AddPolyline(self.getclippedpolypoints(l1, self.startstation.Distance, self.endstation.Distance), Color.Orange.ToArgb(), 2)

            for p in self.getarrowlocations(l1, 10): # returns list with location and perp right azimuth [Point3D, perpVector3D.Value.Azimuth]
                self.overlayBag.AddMarker(p[0], GraphicMarkerTypes.Arrow_IndependentColor, Color.Orange.ToArgb(), "", 0, math.pi - p[1], 3.0)

        if l2 != None:
            self.overlayBag.AddPolyline(self.getpolypoints(l2), Color.Green.ToArgb(), 4)

            #for p in self.getarrowlocations(l2, 10): # returns list with location and perp right azimuth [Point3D, perpVector3D.Value.Azimuth]
            #    self.overlayBag.AddMarker(p[0], GraphicMarkerTypes.Arrow_IndependentColor, Color.Orange.ToArgb(), "", 0, math.pi - p[1], 3.0)

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

    def unitssetup(self, sender, e):
        # setup everything for the unit conversions

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        self.lfp = self.lunits.Properties.Copy() # create a copy in order to set the decimals and enable/disable the suffix
        self.lfp.AddSuffix = False # disable suffix, we need to set it manually, it would always add the current projects units

        self.unitschanged(None, None)
    
    def unitschanged(self, sender, e):

        # loop through all objects of self and set the properties for all DistanceEdits
        # the code is slower than doing it manually for each single one
        # but more convenient since we don't have to worry about how many DistanceEdit Controls we have in the UI
        tt = self.__dict__.items()
        for i in self.__dict__.items():
            if i[1].GetType() == TBCWpf.DistanceEdit().GetType():
                i[1].DisplayUnit = LinearType(self.lunits.DisplayType)
                i[1].ShowControlIcon(False)
                i[1].FormatProperty.AddSuffix = ControlBoolean(1)
                #i[1].FormatProperty.NumberOfDecimals = self.lfp.NumberOfDecimals

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def Dispose(self, cmd, disposing):
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)


    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''


        self.success.Content += ''

        wv = self.currentProject [Project.FixedSerial.WorldView]
        wv.PauseGraphicsCache(True)

        self.success.Content=''
        # self.label_benchmark.Content = ''

        # start_t = timer ()
        dtmintersectpoints = DynArray()

        inputok=True

        maxlinedistance = abs(self.maxlinedistance.Distance)

        eloffset = self.eloffset.Distance

        interval = self.interval.Distance
        if interval == 0.0:
            self.error.Content += '\nInterval must not be Zero'
            inputok=False

        extend = self.extend.Distance
        dtmprofilelength = self.dtmprofilelength.Distance
 
        l1=self.linepicker1.Entity
        l2=self.linepicker2.Entity
        if l1==None: 
            self.error.Content += '\nno Line 1 selected'
            inputok=False
        if l2==None:
            self.error.Content += '\nno Line 2 selected'
            inputok=False

        startstation = self.startstation.Distance
        endstation = self.endstation.Distance
        
        if inputok:
            self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
            UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())

            if startstation > endstation: startstation, endstation = endstation, startstation   # swap values if necessary

            try:    
                # the "with" statement will unroll any changes if something go wrong
                with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                    polyseg1=l1.ComputePolySeg()
                    polyseg1=polyseg1.ToWorld()
                    polyseg1_v=l1.ComputeVerticalPolySeg()

                    polyseg2=l2.ComputePolySeg()
                    polyseg2=polyseg2.ToWorld()
                    polyseg2_v=l2.ComputeVerticalPolySeg()

                    if self.chooserange.IsChecked:
                        computestartstation = startstation
                        computeendstation = endstation
                    else:
                        computestartstation = polyseg1.BeginStation
                        computeendstation = polyseg1.ComputeStationing()

                    outSegment=clr.StrongBox[Segment]()
                    outPointOnCL1=clr.StrongBox[Point3D]()
                    outPointOnCL2=clr.StrongBox[Point3D]()
                    perpVector3D=clr.StrongBox[Vector3D]()
                    out_t=clr.StrongBox[float]()
                    outdeflectionAngle=clr.StrongBox[float]()
                    
                    station2=clr.StrongBox[float]()
                    p1=Point3D()

                    intersections=Intersections()

                    if self.extendtodtm.IsChecked and self.designsurfacepicker.SelectedSerial!=0:
                        surface = wv.Lookup(self.designsurfacepicker.SelectedSerial) # we get our selected surface as object
                        nTri = surface.NumberOfTriangles    # we need the number of triangles in the surface to count through them

                    dtmintersectcount = 0
                    li = 0

                    #start_t = timer()
                    while computestartstation <= computeendstation:    # as long as we aren't at the end of the line
                        
                        dtmintersect = False
                        # compute a point on line 1
                        polyseg1.FindPointFromStation(computestartstation, outSegment, out_t, outPointOnCL1, perpVector3D, outdeflectionAngle) # compute point and vector on line 1
                        
                        p1 = outPointOnCL1.Value
                        if polyseg1_v!=None:
                            p1.Z=polyseg1_v.ComputeVerticalSlopeAndGrade(computestartstation)[1]
                        
                        # depending on checkbox compute perpendicular to line 1 or 2
                        foundp2 = False
                        if self.perptoline1.IsChecked:
                            perpseg = PolySeg.PolySeg(SegmentLine(p1, perpVector3D.Value))
                            perpseg.Extend(10000.0)
                            #tt = perpseg.Length()
                            #perpseg.ExtendEnd(False, False, 10000.0, True)

                            if perpseg.Intersect(polyseg2, True, intersections):
                            ### # we have to go through all segments of line 2 and check if we intersect
                            ### # we have to do it that way since we don't know if and when our perpvector
                            ### # would actually intersect the second line
                            ### s = polyseg2.FirstSegment   # we get us the first segment of line 2
                            ### while s is not None:
                            ###     foundvalidintersection = False
                            ###     # we use the Segment.Intersect function since we can use True to extend the lines
                            ###     # unfortunately it extends both lines indefintely, which we don't really want for
                            ###     # the segments of line2
                            ###     if perpseg.Intersect(s, True, intersections):
                            ###         if intersections.Count == 1: # if we only get one intersection
                            ###             # we check if we are actually inside the segment of line 2
                            ###             # and not on the extension
                            ###             if intersections[0].T2 >= 0 and intersections[0].T2 <=1: 
                            ###                 ip = intersections[0].Point
                            ###                 foundvalidintersection = True
                            ###         else:
                                i2 = 0
                                foundvalidintersection = False
                                for i in intersections: # if we get multiple intersections we want the shortest one
                                    ### if i.T2 >= 0 and i.T2 <=1:    
                                    foundvalidintersection = True
                                    if i2 == 0:
                                        ip = i.Point
                                        i2 += 1
                                    else:
                                        if p1.Distance2D(i.Point) < p1.Distance2D(ip):
                                            ip = i.Point
                              
                                        
                                if foundvalidintersection:        
                                    polyseg2.FindPointFromPoint(ip, outPointOnCL2, station2)
                                    p2=outPointOnCL2.Value
                                    if polyseg2_v != None:
                                        p2.Z=polyseg2_v.ComputeVerticalSlopeAndGrade(station2.Value)[1]
                                    foundp2 = True
                            
                            intersections.Clear()
                            ###     s = polyseg2.Next(s)
                        else:
                            if polyseg2.FindPointFromPoint(p1,outPointOnCL2,station2):
                                p2 = outPointOnCL2.Value
                                if polyseg2_v != None:  
                                    p2.Z = polyseg2_v.ComputeVerticalSlopeAndGrade(station2.Value)[1]
                                foundp2 = True


                        if foundp2:

                            if self.usesideslope.IsChecked:

                                if self.slope1.IsChecked:
                                    p2.Z = p1.Z + ((self.sideslope.Value / 100) * Vector3D(p1, p2).Length2D)
                                if self.slope2.IsChecked:
                                    p1.Z = p2.Z + ((self.sideslope.Value / 100) * Vector3D(p1, p2).Length2D)

                            p1.Z += eloffset
                            p2.Z += eloffset        # add the elevation offset
                            v = Vector3D(p1, p2)
                            if v.Length > maxlinedistance:    # we test that we don't intersect too far away, double intersections in tight curves
                                self.success.Content += '\nmax. Distance between Line 1 and 2 was hit'
                                if computestartstation==computeendstation: break
                                computestartstation+=interval
                                if computestartstation>computeendstation: computestartstation=computeendstation
                                continue
                            v.Length = extend # we extend the line by the manual amount
                            p3 = p2 + v   # p1 on line 1; p2 on line 2; p3 manual line extension; p4 dtm intersection
                            p4 = Point3D()
                            # now we try to find an intersection with the DTM
                            if self.extendtodtm.IsChecked and self.designsurfacepicker.SelectedSerial!=0:
                                #if self.drawtextensions.IsChecked or self.connectextensions.IsChecked:
                                tiepoint = clr.StrongBox[Point3D]()
                                # slope in Computetie is zenith angle with upwards=0
                                # Vector3D.Horizon is positive above the horizon and negative below
                                if surface.ComputeTie(p2, v, math.pi/2 - (v.Horizon), maxlinedistance, tiepoint):
                                    if Point3D.Distance(p2,tiepoint.Value)[0] < maxlinedistance:
                                        dtmintersect = True
                                        p4 = tiepoint.Value
                                        dtmintersectpoints.Add(p4)  # add the DTM intersection point to the list

                            # draw a line between the input lines
                            if self.connectinputlines.IsChecked:  
                                l = wv.Add(clr.GetClrType(Linestring))
                                l.Layer=self.layerpicker.SelectedSerialNumber
                                if p2.Z > p1.Z: l.Color = Color.Red
                                else: l.Color = Color.Black
                                #e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                #e.Position = p1
                                #l.AppendElement(e)
                                #e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                #e.Position = p2
                                #l.AppendElement(e) 

                                polyseg = PolySeg.PolySeg()
                                if self.benchpattern.IsChecked and li % 2 != 0:
                                    if p2.Z > p1.Z:
                                        polyseg.Add(p2)
                                    else:
                                        polyseg.Add(p1)
                                    
                                    polyseg.Add(Point3D.MidPoint(p1, p2))
                                
                                else:
                                    if p2.Z > p1.Z:
                                        polyseg.Add(p2)
                                        polyseg.Add(p1)
                                    else:
                                        polyseg.Add(p1)
                                        polyseg.Add(p2)

                                l.Append(polyseg, None, False, False)
                            
                            # p1 on line 1; p2 on line 2; p3 manual line extension; p4 dtm intersection
                            # draw a line between p2 and p3 or p3 and p4
                            drawextension=False
                            if self.extendtodtm.IsChecked and dtmintersect == True and self.drawtextensions.IsChecked:
                                drawextension=True
                                p1draw=p2
                                p2draw=p4
                                if p2.Z>p1.Z: drawcolor = Color.Aquamarine
                                else: drawcolor = Color.Lime
                            if self.extendtodtm.IsChecked and dtmintersect == False and self.extendgivenamount.IsChecked: 
                                drawextension=True
                                p1draw=p2
                                p2draw=p3
                                if p2.Z>p1.Z: drawcolor = Color.Magenta
                                else: drawcolor = Color.Blue
                            if self.extendtodtm.IsChecked==False and self.extendgivenamount.IsChecked: 
                                drawextension=True
                                p1draw=p2
                                p2draw=p3
                                if p2.Z>p1.Z: drawcolor = Color.Magenta
                                else: drawcolor = Color.Blue

                            # draw the extension line
                            if drawextension==True:
                                l = wv.Add(clr.GetClrType(Linestring))
                                l.Layer=self.layerpicker.SelectedSerialNumber
                                e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                e.Position = p1draw
                                l.AppendElement(e)
                                e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                e.Position = p2draw
                                l.AppendElement(e) 
                                l.Color = drawcolor
                            
                            # draw the DTM profile
                            if (self.adddtmprofile.IsChecked and dtmintersect == True):
                                v = Vector3D(p1,p2)
                                v.To2D()
                                v.Length = dtmprofilelength
                                p5 = p4 + v                   
                                seg = SegmentLine(p4, p5)
                                polysegdtm = PolySeg.PolySeg(seg)
                                profile = surface.Profile(polysegdtm,True)
                                l = wv.Add(clr.GetClrType(Linestring))
                                l.Layer = self.layerpicker.SelectedSerialNumber
                                l.Color = Color.Yellow
                                for i in range(0, profile.NumberOfNodes):
                                    drawpoint = profile[i].Point
                                    e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                    e.Position = drawpoint
                                    l.AppendElement(e)

                        li += 1
                        if computestartstation==computeendstation: break
                        computestartstation+=interval
                        if computestartstation>computeendstation: computestartstation=computeendstation
                    
                    # draw the connection line along the dtm hits
                    if (self.extendtodtm.IsChecked and self.designsurfacepicker.SelectedSerial!=0 and 
                        dtmintersectpoints.Count>0 and self.connectextensions.IsChecked):
                        for i in range(0, dtmintersectpoints.Count):
                            if i==0:
                                l = wv.Add(clr.GetClrType(Linestring))
                                l.Layer=self.layerpicker.SelectedSerialNumber
                                e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                e.Position = dtmintersectpoints[i]
                                l.AppendElement(e)
                            else:
                                e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                e.Position = dtmintersectpoints[i]
                                l.AppendElement(e)

                    #end_t=timer()
                    #self.error.Content += '\nOverall - ' + str(timedelta(seconds=end_t-start_t))

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

        wv.PauseGraphicsCache(False)

        self.SaveOptions()

    


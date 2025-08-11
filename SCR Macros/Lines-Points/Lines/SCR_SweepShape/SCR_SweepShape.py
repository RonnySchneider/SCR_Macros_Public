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
    cmdData.Key = "SCR_SweepShape"
    cmdData.CommandName = "SCR_SweepShape"
    cmdData.Caption = "_SCR_SweepShape"
    cmdData.UIForm = "SCR_SweepShape"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "Lines"
        cmdData.ShortCaption = "Sweep Shape"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.18
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Sweep Shape along Line"
        cmdData.ToolTipTextFormatted = "Sweep Shape along Line"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_SweepShape(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_SweepShape.xaml") as s:
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

        self.linepicker1.IsEntityValidCallback = self.IsValid
        self.linepicker2.IsEntityValidCallback = self.IsValid
        self.linepicker3.IsEntityValidCallback = self.IsValid

        self.linepicker1.ValueChanged += self.line1Changed
        self.linepicker2.ValueChanged += self.lineselectionChanged
        self.linepicker3.ValueChanged += self.lineselectionChanged

        self.startstation.ValueChanged += self.line1Changed
        self.endstation.ValueChanged += self.line1Changed
        self.chooserange.Checked += self.line1Changed
        self.chooserange.Unchecked += self.line1Changed
        
        #self.interval.NumberOfDecimals = 4
        #self.eloffset.NumberOfDecimals = 4
        #self.hortol.NumberOfDecimals = 4
        #self.vertol.NumberOfDecimals = 4
        #self.nodespacing.NumberOfDecimals = 4

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        self.lfp = self.lunits.Properties.Copy()
        self.linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation

        self.labelinterval.Content = 'Interval [' + self.linearsuffix + ']'
        self.headertolerance.Header = 'define Shape Chording Tolerances [' + self.linearsuffix + ']'

        self.lType = clr.GetClrType(IPolyseg)

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass


    def SetDefaultOptions(self):
        lserial = OptionsManager.GetUint("SCR_SweepShape.layerpicker", 8)
        o = self.currentProject.Concordance.Lookup(lserial) # get the object withj the saved serial number
        if o != None:   # check if we actually got an object back
            if isinstance(o.GetSite(), LayerCollection):    # check if the object is a layer
                self.layerpicker.SetSelectedSerialNumber(lserial, InputMethod(3))
            else:  # if not then set Layer zero                       
                self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:   # if not then set Layer zero               
            self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))

        self.chooserange.IsChecked = OptionsManager.GetBool("SCR_SweepShape.chooserange", False)
        self.startstation.SetStation(OptionsManager.GetDouble("SCR_SweepShape.startstation", 0.000), self.currentProject)
        self.endstation.SetStation(OptionsManager.GetDouble("SCR_SweepShape.endstation", 0.000), self.currentProject)

        self.interval.Distance = OptionsManager.GetDouble("SCR_SweepShape.interval", 1.000)
        #self.eloffset.Value = OptionsManager.GetDouble("SCR_SweepShape.eloffset", 0.000)
        self.includegradechanges.IsChecked = OptionsManager.GetBool("SCR_SweepShape.includegradechanges", True)
        self.drawxlines.IsChecked = OptionsManager.GetBool("SCR_SweepShape.drawxlines", False)
        self.drawshapes.IsChecked = OptionsManager.GetBool("SCR_SweepShape.drawshapes", False)
        self.drawpoints.IsChecked = OptionsManager.GetBool("SCR_SweepShape.drawpoints", False)
        self.drawlonglines.IsChecked = OptionsManager.GetBool("SCR_SweepShape.drawlonglines", False)
        self.draw3dshell.IsChecked = OptionsManager.GetBool("SCR_SweepShape.draw3dshell", False)

        self.hortol.Distance = OptionsManager.GetDouble("SCR_SweepShape.hortol", 0.001)
        #self.vertol.Value = OptionsManager.GetDouble("SCR_SweepShape.vertol", 0.001)
        self.nodespacing.Distance = OptionsManager.GetDouble("SCR_SweepShape.nodespacing", 10)

    def SaveOptions(self):
        OptionsManager.SetValue("SCR_SweepShape.layerpicker", self.layerpicker.SelectedSerialNumber)

        OptionsManager.SetValue("SCR_SweepShape.chooserange", self.chooserange.IsChecked)
        OptionsManager.SetValue("SCR_SweepShape.startstation", self.startstation.Distance)
        OptionsManager.SetValue("SCR_SweepShape.endstation", self.endstation.Distance)

        OptionsManager.SetValue("SCR_SweepShape.interval", self.interval.Distance)
        #OptionsManager.SetValue("SCR_SweepShape.eloffset", self.eloffset.Value)
        OptionsManager.SetValue("SCR_SweepShape.includegradechanges", self.includegradechanges.IsChecked)
        OptionsManager.SetValue("SCR_SweepShape.drawxlines", self.drawxlines.IsChecked)
        OptionsManager.SetValue("SCR_SweepShape.drawshapes", self.drawshapes.IsChecked)
        OptionsManager.SetValue("SCR_SweepShape.drawpoints", self.drawpoints.IsChecked)
        OptionsManager.SetValue("SCR_SweepShape.drawlonglines", self.drawlonglines.IsChecked)
        OptionsManager.SetValue("SCR_SweepShape.draw3dshell", self.draw3dshell.IsChecked)
        
        OptionsManager.SetValue("SCR_SweepShape.hortol", self.hortol.Distance)
        #OptionsManager.SetValue("SCR_SweepShape.vertol", self.vertol.Value)
        OptionsManager.SetValue("SCR_SweepShape.nodespacing", self.nodespacing.Distance)

    def drawoverlay(self):

        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)
        self.overlayBag = OverlayBag(self.ViewOverlay) # create a new/empty overlaybag

        l1 = self.linepicker1.Entity
        l2 = self.linepicker2.Entity
        l3 = self.linepicker3.Entity

        if l1 != None:
            self.overlayBag.AddPolyline(self.getpolypoints(l1), Color.Red.ToArgb(), 4)

            if self.chooserange.IsChecked:
                self.overlayBag.AddPolyline(self.getclippedpolypoints(l1, self.startstation.Distance, self.endstation.Distance), Color.Orange.ToArgb(), 2)

            for p in self.getarrowlocations(l1, 10): # returns list with location and perp right azimuth [Point3D, perpVector3D.Value.Azimuth]
                self.overlayBag.AddMarker(p[0], GraphicMarkerTypes.Arrow_IndependentColor, Color.Orange.ToArgb(), "", 0, math.pi - p[1], 3.0)

        if l2 != None:
            self.overlayBag.AddPolyline(self.getpolypoints(l2), Color.Green.ToArgb(), 4)

        if l3 != None:
            self.overlayBag.AddPolyline(self.getpolypoints(l3), Color.Blue.ToArgb(), 4)

        # if it only needs to be visible in all Planview then remove the Hoops3DViewGUID
        array = Array[Guid]([DisplayWindow.Hoops3DViewGUID, DisplayWindow.HoopsPlanViewGUID])
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.AddOverlayGeometry(array, self.overlayBag)

        return

    def getarrowlocations(self, l1, intervals):

        pts = []

        polyseg = l1.ComputePolySeg()
        polyseg = polyseg.ToWorld()
        polyseg_v = l1.ComputeVerticalPolySeg()
        
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
            polyseg = l.ComputePolySeg()
            polyseg = polyseg.ToWorld()
            polyseg_v = l.ComputeVerticalPolySeg()
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

    def IsValid(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        return False

    def chooserangeChanged(self, sender, e):
        if self.chooserange.IsChecked==True:
            self.startstation.IsEnabled=True
            self.endstation.IsEnabled=True
        else:
            self.startstation.IsEnabled=False
            self.endstation.IsEnabled=False

    def lineselectionChanged(self, ctrl, e):

        self.drawoverlay()

    def line1Changed(self, ctrl, e):
        l1 = self.linepicker1.Entity
        if l1 != None:
            self.startstation.StationProvider = l1
            self.endstation.StationProvider = l1
            if not self.chooserange.IsChecked:
                Keyboard.Focus(self.linepicker2)

        self.drawoverlay()

    def createnamedpointChanged(self, sender, e):
        if self.createnamedpoint.IsChecked:
            self.createnamedpointpanel.IsEnabled = True
        else:
            self.createnamedpointpanel.IsEnabled = False

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def Dispose(self, cmd, disposing):
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''

        self.success.Content += ''

        wv = self.currentProject [Project.FixedSerial.WorldView]
        dp = self.currentProject.CreateDuplicator()

        self.success.Content=''

        inputok=True

        l1 = self.linepicker1.Entity
        l2 = self.linepicker2.Entity
        l3 = self.linepicker3.Entity
        
        if l1==None: 
            self.success.Content += '\nno Line 1 selected'
            inputok=False
        if l2==None:
            self.success.Content += '\nno Line 2 selected'
            inputok=False
        if l3==None:
            self.success.Content += '\nno Shape selected'
            inputok=False
        
        if l1 == l2 == l3: inputok = False

        try:
            startstation = self.startstation.Distance
        except:
            self.success.Content += '\nStart Chainage error'
            inputok=False

        try:
            endstation = self.endstation.Distance
        except:
            self.success.Content += '\nEnd Chainage error'
            inputok=False
        
        if inputok:
            self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
            UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
            outSegment1 = clr.StrongBox[Segment]()
            outPointOnCL1 = clr.StrongBox[Point3D]()
            outPointOnCL2 = clr.StrongBox[Point3D]()
            perpVector3D = clr.StrongBox[Vector3D]()
            outaz = clr.StrongBox[Vector3D]()
            out_t = clr.StrongBox[float]()
            out_offset = clr.StrongBox[float]()
            out_side = clr.StrongBox[Side]()
            outdeflectionAngle = clr.StrongBox[float]()
            station1 = clr.StrongBox[float]()
            station2 = clr.StrongBox[float]()
            outelevation1 = clr.StrongBox[float]()
            outslope1 = clr.StrongBox[Vector3D]()
            p1 = Point3D()

            chainage_trans = Array[float]([0.000]) + Array[Matrix4D]([Matrix4D()])
            all_chainage_trans = DynArray()

            if startstation > endstation: startstation, endstation = endstation, startstation   # swap values if necessary
             
            try:
                # the "with" statement will unroll any changes if something go wrong
                with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                    polyseg1 = l1.ComputePolySeg()
                    polyseg1 = polyseg1.ToWorld()
                    polyseg1_v = l1.ComputeVerticalPolySeg()

                    polyseg2 = l2.ComputePolySeg()
                    polyseg2 = polyseg2.ToWorld()
                    polyseg2_v = l2.ComputeVerticalPolySeg()

                    polyseg3 = l3.ComputePolySeg()
                    polyseg3 = polyseg3.ToWorld()
                    polyseg3 = polyseg3.Linearize(abs(self.hortol.Distance), 10000,
                                abs(self.nodespacing.Distance), None, False)

                    if self.chooserange.IsChecked:
                        computestartstation = startstation
                        computeendstation = endstation
                    else:
                        computestartstation = polyseg1.BeginStation
                        computeendstation = polyseg1.ComputeStationing()

                    ps1 = self.coordpick1.Coordinate
                    ps1.Z = 0
                    ps2 = self.coordpick2.Coordinate
                    ps2.Z = 0
                    
                    # compile chainage list
                    chainages = []
                    # add grade changes on line 2
                    if self.includegradechanges.IsChecked:
                        for n in polyseg2_v.Nodes():    # go through all the nodes in the profile of line 2
                        # we are working with a profile, so X is the Chainage and Y is the Elevation
                            try:
                                if n.Visible:
                                    polyseg2.FindPointFromStation(n.Point.X, outPointOnCL2)  # compute the world XY from the Chainage on Line 2
                                    polyseg1.FindPointFromPoint(outPointOnCL2.Value, outPointOnCL1, station1)  # with that Point compute the Chainage on Line 1
                                    # check if it falls within our range
                                    if station1.Value >= computestartstation and station1.Value <= computeendstation:
                                        chainages.Add(station1.Value)
                            except:
                                break

                    rounddecimals = self.currentProject.Units.Station.Properties.NumberOfDecimals
                    # add the interval range on the first line
                    if self.interval.Distance != 0:
                        while computestartstation <= computeendstation:    # as long as we aren't at the end of the line 1
                            # avoid duplicates in the chainage list, but keep full precision for the grade breaks
                            canadd = True
                            for c in chainages: 
                                if round(c, rounddecimals) - round(computestartstation, rounddecimals) == 0:
                                    canadd = False
                            if canadd == True: chainages.Add(computestartstation)

                            if computestartstation == computeendstation: break
                            computestartstation += self.interval.Distance
                            if computestartstation > computeendstation: computestartstation = computeendstation
           

                    intersections=Intersections()

                    chainages = list(set(chainages)) # remove duplicates, there shouldn't actually be any
                    chainages.sort()  # sort list
                    tt = chainages.Count
                    j = 0
                    if self.drawshapes.IsChecked or self.drawpoints.IsChecked:
                        ProgressBar.TBC_ProgressBar.Title = "create Shapes/Points"
                    else:
                        ProgressBar.TBC_ProgressBar.Title = "compute Transformation Values"

                    
                    # prepare arrays for IFCMesh
                    if self.draw3dshell.IsChecked:
                        validshellvertices = []  
                        if polyseg3.IsClosed:
                            nodecount = polyseg3.NumberOfNodes - 1 # the last node is the same as the first
                            drawlastnode = False
                        else:
                            nodecount = polyseg3.NumberOfNodes
                            drawlastnode = True

                    for computestation in chainages:    # go through the chainage list and compute the transformation matrix if possible
                        # compute a point on line 1 and get the perpendicular vector
                        polyseg1.FindPointFromStation(computestation, outSegment1, out_t, outPointOnCL1, perpVector3D, outdeflectionAngle) # compute point and vector on line 1

                        p1 = outPointOnCL1.Value
                        foundp2 = False
                        finalp2 = Point3D.Zero
                        
                        if polyseg1_v != None:
                            polyseg1_v.ComputeVerticalSlopeAndGrade(computestation, outelevation1, outslope1)
                            if outslope1.Value.Length == 0: # we have to make sure we've got a valid slope value back
                                polyseg1_v.ComputeVerticalSlopeAndGrade(computestation + 0.0005, outelevation1, outslope1) # if not, we try a slightly different chainage
                            if outslope1.Value.Length != 0: # if we have a slope value we continue to try to find a point in line 2
                                p1.Z = outelevation1.Value
                        
                                perpseg = SegmentLine(p1, perpVector3D.Value)

                                # we have to go through all segments of line 2 and check if we intersect
                                # we have to do it that way since we don't know if and when our perpvector
                                # would actually intersect the second line
                                s = polyseg2.FirstSegment   # we get us the first segment of line 2
                                while s is not None:
                                    foundvalidintersection = False
                                    # we use the Segment.Intersect function since we can use True to extend the lines
                                    # unfortunately it extends both lines indefintely, which we don't really want for
                                    # the segments of line2
                                    if perpseg.Intersect(s, True, intersections):
                                        if intersections.Count == 1: # if we only get one intersection
                                            # we check if we are actually inside the segment of line 2
                                            # and not on the extension
                                            if intersections[0].T2 >= 0 and intersections[0].T2 <=1: 
                                                ip = intersections[0].Point
                                                foundvalidintersection = True
                                        else:
                                            i2 = 0
                                            for i in intersections: # if we get multiple intersections (on arcs) we want the shortest one
                                                if i.T2 >= 0 and i.T2 <=1:    
                                                    foundvalidintersection = True
                                                    if i2 == 0:
                                                        ip = i.Point
                                                        i2 += 1
                                                    else:
                                                        if p1.Distance2D(i.Point) < p1.Distance2D(ip):
                                                            ip = i.Point
                                                
                                        if foundvalidintersection:        
                                            polyseg2.FindPointFromPoint(ip, outPointOnCL2, station2)
                                            p2 = outPointOnCL2.Value
                                            if polyseg2_v != None:
                                                p2.Z=polyseg2_v.ComputeVerticalSlopeAndGrade(station2.Value)[1]
                                            foundp2 = True
                                    
                                    # in tight corners we only want the first intersection        
                                    if foundp2 == True and \
                                       (finalp2 == Point3D.Zero or Vector3D(p1, p2).Length2D < Vector3D(p1, finalp2).Length2D):
                                        finalp2 = p2

                                    intersections.Clear()
                                    s = polyseg2.Next(s)


                        if foundp2:
                            p2 = finalp2

                            #p1.Z += self.eloffset.Value
                            #p2.Z += self.eloffset.Value        # add the elevation offset

                            # we need to know on which side p2/l2 are
                            polyseg1.FindPointFromPoint(p2, outSegment1, out_t, outPointOnCL1, station1, outaz, out_offset, out_side)
                            
                            # we create a horizontal vector in direction of l1 at p1
                            l1_dir = Vector3D(p1, Point3D(p2.X, p2.Y, p1.Z)) # here it is still the sideways vector to line 2
                            # rotate it sideways about Z axis -> HORIZONTAL vector along line 1
                            if out_side.Value == Side.Right:
                                l1_dir.Rotate90(Side.Left)
                            else:
                                l1_dir.Rotate90(Side.Right)
                            l1_dir.Length = 2
                            # we create a copy and adjust the slope of that new vector to represent the slope at P1 + 90
                            # Vector3D.Horizon is positive above the horizon and negative below
                            l1_dir_slope = l1_dir.Clone()
                            l1_dir_slope.Horizon = math.atan(outslope1.Value.Y / outslope1.Value.X) + math.pi/2
                            # we rotate the new vector around the horizontal sideways vector
                            # Rotate(Trimble.Vce.Geometry.Vector3D rotationAxis, double rotationAngle)
                            if out_side.Value == Side.Right:
                                l1_dir_slope.Rotate(l1_dir, -1 * Vector3D(p1,p2).Horizon)
                            else:
                                l1_dir_slope.Rotate(l1_dir, Vector3D(p1,p2).Horizon)

                            ## test draw
                            #l = wv.Add(clr.GetClrType(Linestring))
                            #e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                            #e.Position = p1 
                            #l.AppendElement(e)
                            #e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                            #e.Position = p1 + l1_dir_slope
                            #l.AppendElement(e)
                            #l.Layer = self.layerpicker.SelectedSerialNumber

                            vps1ps2_90 = Vector3D(ps1, ps2)
                            vps1ps2_90.Rotate90(Side.Left)

                            vp1p2 = Vector3D(p1, p2)
                            if out_side.Value == Side.Left:
                                vp1p2.Negate()

                            targetrot = Spinor3D.ComputeRotation(Vector3D(ps1, ps2), vps1ps2_90, vp1p2, l1_dir_slope)
                            #BuildTransformMatrix(Vector3D fromPoint, Vector3D translation, Spinor3D rotation, Vector3D scale)
                            targetmatrix = Matrix4D.BuildTransformMatrix(Vector3D(ps1), Vector3D(ps1, p1), targetrot, Vector3D(1,1,1))

                            chainage_trans[0] = computestation
                            chainage_trans[1] = targetmatrix
                            all_chainage_trans.Add(chainage_trans.Clone())

                            tt = polyseg3.Nodes()

                            # draw shapes and/or points
                            if self.drawshapes.IsChecked or self.drawpoints.IsChecked or self.draw3dshell.IsChecked:
                                
                                if self.drawshapes.IsChecked:
                                    l3_new = wv.Add(clr.GetClrType(Linestring))
                                    l3_new.Color = l3.Color
                                    l3_new.Layer = self.layerpicker.SelectedSerialNumber
                                
                                if self.drawpoints.IsChecked == False: # if we don't want to draw all points
                                    #we create a temp point for the transformation and delete it afterwards
                                    cadPoint = wv.Add(clr.GetClrType(CadPoint))
                                
                                for n in polyseg3.Nodes():

                                    try:
                                        if n.Visible:
                                            if (n.Length > 0.0) or (n.Length == 0.0 and drawlastnode):
                                            
                                                pnew = n.Point
                                                pnew.Z = 0
                                                if self.drawpoints.IsChecked: # if we want to draw all points
                                                    cadPoint = wv.Add(clr.GetClrType(CadPoint))
                                                    cadPoint.Layer = self.layerpicker.SelectedSerialNumber

                                                cadPoint.Point0 = pnew
                                                
                                                cadPoint.Transform(TransformData(targetmatrix, Matrix4D(Vector3D.Zero)))
                                                
                                                if self.draw3dshell.IsChecked:
                                                    validshellvertices.Add(Point3D(cadPoint.Point0))

                                                if self.drawshapes.IsChecked: 
                                                    e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                                    e.Position = cadPoint.Point0
                                                    l3_new.AppendElement(e)
                                    except:
                                        break


                                if self.drawpoints.IsChecked == False:
                                    osite = cadPoint.GetSite()    # we find out in which container the serial number reside
                                    osite.Remove(cadPoint.SerialNumber)   # we delete the object from that container

                            # draw lines between p1 and p2
                            if self.drawxlines.IsChecked:
                                l = wv.Add(clr.GetClrType(Linestring))
                                e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                e.Position = p1 
                                l.AppendElement(e)
                                e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                e.Position = p2
                                l.AppendElement(e)
                                l.Layer = self.layerpicker.SelectedSerialNumber
                                l.Color = Color.Blue

                        # update progress bar                                                                             
                        j += 1                                                                                            
                        if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(j * 100 / chainages.Count)):                            
                            break   # function returns true if user pressed cancel                                        

                    # after going trough all the chainages                                                                

                    if self.draw3dshell.IsChecked:
                        ifcvertices = Array[Point3D]([Point3D()] * validshellvertices.Count)
                        for i in range(0, validshellvertices.Count):
                            ifcvertices[i] = validshellvertices[i]
                        ifcverticeindex = 0
                        if polyseg3.IsClosed:
                            ifcfacelist = Array[Int32]([Int32()] * (2 * 4 * nodecount * math.floor(((validshellvertices.Count / nodecount) - 1)))) # define size of facelist array
                        else:
                            ifcfacelist = Array[Int32]([Int32()] * (2 * 4 * (nodecount - 1) * math.floor(((validshellvertices.Count / nodecount) - 1)))) # define size of facelist array

                        ifcnormals = Array[Point3D]([Point3D()]*0)
                                                                                                                              
                        # prepare the IFCFacelist                                                                                   
                        ifcfacelistindex = 0                                                                                    
                        for c in range(0, math.floor((validshellvertices.Count / nodecount)) - 1):                                          #         11----10
                            for f in range(0, nodecount - 1):                                                                       #         /|    | 
                                                                                                                                #        / 8----9
                                ifcfacelist[ifcfacelistindex + 0] = 3 # always 3 vertices per triangle                          #       /    
                                ifcfacelist[ifcfacelistindex + 1] = (c * nodecount) + f               # 0                       #      7----6
                                ifcfacelist[ifcfacelistindex + 2] = (c * nodecount) + f + 1           # 1                       #     /|    |
                                ifcfacelist[ifcfacelistindex + 3] = ((c + 1) * nodecount) + f + 1     # 5                       #    / 4----5
                                ifcfacelistindex += 4                                                                           #   /        
                                                                                                                                #  3----2    
                                ifcfacelist[ifcfacelistindex + 0] = 3 # always 3 vertices per triangle                          #  |    |    
                                ifcfacelist[ifcfacelistindex + 1] = (c * nodecount) + f               # 0                       #  0----1    
                                ifcfacelist[ifcfacelistindex + 2] = ((c + 1) * nodecount) + f + 1     # 5
                                ifcfacelist[ifcfacelistindex + 3] = ((c + 1) * nodecount) + f         # 4
                                ifcfacelistindex += 4

                            if polyseg3.IsClosed:
                                ifcfacelist[ifcfacelistindex + 0] = 3 # always 3 vertices per triangle
                                ifcfacelist[ifcfacelistindex + 1] = (c  * nodecount) + 0                   # 0
                                ifcfacelist[ifcfacelistindex + 2] = ((c + 1) * nodecount) + 0              # 4
                                ifcfacelist[ifcfacelistindex + 3] = ((c + 1) * nodecount) + nodecount - 1  # 7
                                ifcfacelistindex += 4


                                ifcfacelist[ifcfacelistindex + 0] = 3 # always 3 vertices per triangle
                                ifcfacelist[ifcfacelistindex + 1] = (c  * nodecount) + 0                   # 0
                                ifcfacelist[ifcfacelistindex + 2] = ((c + 1) * nodecount) + nodecount - 1  # 7
                                ifcfacelist[ifcfacelistindex + 3] = (c * nodecount) + nodecount - 1        # 3
                                ifcfacelistindex += 4



                        ### mesh = wv.Add(clr.GetClrType(IFCMesh))
                        ### 
                        ### mesh.CreateShell(0, Guid.NewGuid(), ifcvertices, ifcfacelist, ifcnormals, Color.Red.ToArgb(), Matrix4D())
                        ### mesh._TriangulatedFaceList = ifcfacelist # otherwise GetTriangulatedFaceList() in the perpDist macro will not get the necessary values
                        self.createifcobject("SweepedShapes", "Shape1", "Shape1", ifcvertices, ifcfacelist, ifcnormals, Color.Red.ToArgb(), Matrix4D(), self.layerpicker.SelectedSerialNumber)



                    # draw long lines
                    j = 0
                    ProgressBar.TBC_ProgressBar.Title = "create longitudinal Lines"
                    if self.drawlonglines.IsChecked:
                        cadPoint = wv.Add(clr.GetClrType(CadPoint))
                        for n in polyseg3.Nodes():
                            try:
                                if n.Visible and n.Length > 0.0:
                                    
                                    l = wv.Add(clr.GetClrType(Linestring))
                                    l.Layer = self.layerpicker.SelectedSerialNumber
                                    for t in all_chainage_trans:
                                        
                                        pnew = n.Point
                                        pnew.Z = 0

                                        cadPoint.Point0 = pnew
                                        
                                        cadPoint.Transform(TransformData(t[1], Matrix4D(Vector3D.Zero)))
                                     
                                        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                        e.Position = cadPoint.Point0
                                        l.AppendElement(e)

                                        # update progress bar
                                        j += 1
                                        if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(j * 100 / (polyseg3.NumberOfNodes * all_chainage_trans.Count))):
                                            break   # function returns true if user pressed cancel
                            
                            except:
                                break
                        osite = cadPoint.GetSite()    # we find out in which container the serial number reside
                        osite.Remove(cadPoint.SerialNumber)   # we delete the object from that container


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
        Keyboard.Focus(self.linepicker1)

    
    def createifcobject(self, ifcprojectname, ifcshellgroupname, ifcshellname, ifcvertices, ifcfacelist, ifcnormals, color, transform, layer):
  
        # get the BIM entity collection from the world view - if it doesn't exist yet it is created
        bimEntityColl = BIMEntityCollection.ProvideEntityCollection(self.currentProject, True)
        shellMeshDataColl = ShellMeshDataCollection.ProvideShellMeshDataCollection(self.currentProject, True)

        # theoretical IFC hierarchy in TBC
        # IFCPROJECT
        #   IFCSITE
        #       IFCBUILDING
        #           IFCBUILDINGSTOREY
        #               IFCBUILDINGELEMENTPROXY
        
        # test if provided project name exists, if yes then add to it
        found = False
    
        for e in bimEntityColl:
            if e.Name == ifcprojectname:
                bimprojectEntity = e
                found = True
                continue
        if not found:
            bimprojectEntity = bimEntityColl.Add(clr.GetClrType(BIMEntity))
            bimprojectEntity.EntityType = "IFCPROJECT"
            bimprojectEntity.Description = ifcprojectname # must be set, otherwise export fails
            bimprojectEntity.BIMGuid = Guid.NewGuid()

        #bimsiteEntity = bimprojectEntity.Add(clr.GetClrType(BIMEntity))
        #bimsiteEntity.EntityType = "IFCSITE"
        #bimsiteEntity.Description = "TestSite" # must be set, otherwise export fails
        #bimsiteEntity.BIMGuid = Guid.NewGuid()

        #bimbuildingEntity = bimprojectEntity.Add(clr.GetClrType(BIMEntity))
        #bimbuildingEntity.EntityType = "IFCBUILDING"
        #bimbuildingEntity.Description = "TestBuilding" # must be set, otherwise export fails
        #bimbuildingEntity.BIMGuid = Guid.NewGuid()

        #bimbuildingstoreyEntity = bimbuildingEntity.Add(clr.GetClrType(BIMEntity))
        #bimbuildingstoreyEntity.EntityType = "IFCBUILDINGSTOREY"
        #bimbuildingstoreyEntity.Description = "TestStorey" # must be set, otherwise export fails
        #bimbuildingstoreyEntity.BIMGuid = Guid.NewGuid()
        
        # find unique group name
        while found == True:
            found = False
            for e in bimprojectEntity:
                if e.Name == ifcshellgroupname:
                    found = True
            if found:
                ifcshellgroupname = PointHelper.Helpers.Increment(ifcshellgroupname, None, True)
        
        # if we create at least IFCPROJECT and IFCBUILDINGELEMENTPROXY (which includes the IFCMesh)
        # we can export it properly into an IFC file but keep the meshes separate
        # upon import TBC will structure it as follows
        # IFCPROJECT --> IFCPROJECT
        # everything will end up under "Default Site - IFCSITE"
        # the IFCBUILDINGELEMENTPROXY with the subsequent IFCMesh get combined into separate BIM Objects

        bimobjectEntity = bimprojectEntity.Add(clr.GetClrType(BIMEntity))
        bimobjectEntity.EntityType = "IFCBUILDINGELEMENTPROXY"
        bimobjectEntity.Description = ifcshellgroupname
        bimobjectEntity.BIMGuid = Guid.NewGuid()
        bimobjectEntity.Mode = DisplayMode(1 + 2 + 64 + 128 + 512 + 4096)
        # MUST set layer, otherwise the properties manager will falsly show layer "0"
        # but the object is invisible until manually changing the layer
        bimobjectEntity.Layer = layer 

        shellmeshdata = shellMeshDataColl.AddShellMeshData(self.currentProject)
        shellmeshdata.CreateShellMeshData(ifcvertices, ifcfacelist, ifcnormals)
        shellmeshdata.SetVolumeCalculationShell(ifcvertices, ifcfacelist)
        #shellmeshdata.Description = ifcshellgroupname
        #shellmeshdata = shellMeshDataColl.TryCreateFromShell3D(None, 0)

        meshInstance = bimobjectEntity.Add(clr.GetClrType(ShellMeshInstance))
        meshInstance.CreateShell(0, shellmeshdata.SerialNumber, color, transform)




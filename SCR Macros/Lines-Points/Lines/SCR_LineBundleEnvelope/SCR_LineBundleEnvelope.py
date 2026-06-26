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
    "layerpicker": 8, "chooserange": False,
    "startstation": 0.0, "endstation": 0.0,
    "interval": 1.0, "includebundlenodes": False,
    "drawrectangle": False, "drawpoints": False,
    "drawlonglines": False, "draw3dshell": False,
}

def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_LineBundleEnvelope"
    cmdData.CommandName = "SCR_LineBundleEnvelope"
    cmdData.Caption = "_SCR_LineBundleEnvelope"
    cmdData.UIForm = "SCR_LineBundleEnvelope"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "Lines"
        cmdData.ShortCaption = "Line-Bundle Envelope"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.21
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Envelope around a Bundle of Lines"
        cmdData.ToolTipTextFormatted = "Envelope around a Bundle of Lines"

    except:
        pass
    try:
        b = Bitmap(macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_LineBundleEnvelope(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader(macroFileFolder + r"\SCR_LineBundleEnvelope.xaml") as s:
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
        self.Caption = cmd.Command.Caption

        self.linepicker1.IsEntityValidCallback = self.IsValid
        self.linepicker1.ValueChanged += self.lineChanged

        self.startstation.ValueChanged += self.lineChanged
        self.endstation.ValueChanged += self.lineChanged
        self.chooserange.Checked += self.lineChanged
        self.chooserange.Unchecked += self.lineChanged

        self.objs.IsEntityValidCallback = self.IsValid
        
        #self.interval.NumberOfDecimals = 4
        
        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        #self.lfp = self.lunits.Properties.Copy()
        self.linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation

        self.labelinterval.Content = 'Interval [' + self.linearsuffix + ']'
        self.interval.DistanceMin = 0

        self.lType = clr.GetClrType(IPolyseg)

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def drawoverlay(self):

        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)
        self.overlayBag = OverlayBag(self.ViewOverlay) # create a new/empty overlaybag

        l1 = self.linepicker1.Entity

        if l1:
            self.overlayBag.AddPolyline(SCROverlayBag.getpolypoints(l1), Color.Green.ToArgb(), 4)

            if self.chooserange.IsChecked:
                self.overlayBag.AddPolyline(SCROverlayBag.getclippedpolypoints(l1, self.startstation.Distance, self.endstation.Distance), Color.Orange.ToArgb(), 2)

            for p in SCROverlayBag.getarrowlocations(l1, 10): # returns list with location and perp right azimuth [Point3D, perpVector3D.Value.Azimuth]
                self.overlayBag.AddMarker(p[0], GraphicMarkerTypes.Arrow_IndependentColor, Color.Orange.ToArgb(), "", 0, math.pi - p[1], 3.0)

        # if it only needs to be visible in all Planview then remove the Hoops3DViewGUID
        array = Array[Guid]([DisplayWindow.Hoops3DViewGUID, DisplayWindow.HoopsPlanViewGUID])
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.AddOverlayGeometry(array, self.overlayBag)

        return

    def SetDefaultOptions(self):
        SCROptions.LoadMacroOptions(self, "SCR_LineBundleEnvelope", _OPTIONS, self.currentProject)

    def SaveOptions(self):
        SCROptions.SaveMacroOptions(self, "SCR_LineBundleEnvelope", _OPTIONS)

    def IsValid(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        return False

    def chooserangeChanged(self, sender, e):
        if self.chooserange.IsChecked == True:
            self.startstation.IsEnabled = True
            self.endstation.IsEnabled = True
        else:
            self.startstation.IsEnabled = False
            self.endstation.IsEnabled = False

    def lineChanged(self, ctrl, e):
        l1 = self.linepicker1.Entity
        if l1 != None:
            self.startstation.StationProvider = l1
            self.endstation.StationProvider = l1

        self.drawoverlay()

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand()

    def Dispose(self, cmd, disposing):
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''



        wv = self.currentProject[Project.FixedSerial.WorldView]

        self.success.Content = ''

        inputok = True

        l1 = self.linepicker1.Entity
        
        if l1 == None: 
            self.success.Content += '\nno Line 1 selected'
            inputok = False
        
        try:
            startstation = self.startstation.Distance
        except:
            self.success.Content += '\nStart Chainage error'
            inputok = False

        try:
            endstation = self.endstation.Distance
        except:
            self.success.Content += '\nEnd Chainage error'
            inputok = False
        
        interval = abs(self.interval.Distance)

        if inputok:
            UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
            self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)

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
            rounddecimals = self.currentProject.Units.Station.Properties.NumberOfDecimals
            intersections = Intersections()

            all_limits = DynArray()

            if startstation > endstation: startstation, endstation = endstation, startstation   # swap values if necessary
                
            try:
                # the "with" statement will unroll any changes if something go
                # wrong
                with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

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
                    
                    if self.includebundlenodes.IsChecked or interval != 0:  
                        if self.includebundlenodes.IsChecked:
                            o_count = 0
                            for o in self.objs:
                                if isinstance(o, self.lType) and o != l1:
                                    o_count += 1
                                    ProgressBar.TBC_ProgressBar.Title = "add Bundle Nodes for Line " + str(o_count) + " / " + str(self.objs.Count)
                                    
                                    polyseg2 = o.ComputePolySeg()
                                    polyseg2 = polyseg2.ToWorld()
                                    polyseg2_v = o.ComputeVerticalPolySeg()
                                    
                                    node_count = polyseg2.ToPoint3DArray().Count + polyseg2_v.NumberOfNodes
                                    j = 0
                                    # add nodes on line 2 to the chainage list
                                    for np in polyseg2.ToPoint3DArray():
                                        try:
                                            if polyseg1.FindPointFromPoint(np, outPointOnCL1, station1):  # with that Point compute the Chainage on Line 1
                                                # check if it falls within our
                                                                                                                                                    # range
                                                if station1.Value >= computestartstation and station1.Value <= computeendstation:
                                                    chainages.Add(round(station1.Value, rounddecimals))

                                                j += 1
                                                if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(j * 100 / (node_count))):
                                                    break   # function returns true if user pressed cancel
                                        

                                                    ## check if we've got that
                                                    ## station already
                                                    #canadd = True
                                                    #for c in chainages:
                                                    #    if round(c, rounddecimals)
                                                    #    - round(station1.Value,
                                                    #    rounddecimals) == 0:
                                                    #        canadd = False
                                                    #if canadd == True:
                                                    #chainages.Add(station1.Value)
                                    
                                        except:
                                            break

                                    # add grade changes on line 2
                                    for n in polyseg2_v:    # go through all the nodes in the profile of line 2
                                    # we are working with a profile, so X is the
                                    # Chainage and Y is the Elevation
                                        try:
                                            if n.Visible:
                                                polyseg2.FindPointFromStation(n.Point.X, outPointOnCL2)  # compute the world XY from the Chainage on Line 2
                                                polyseg1.FindPointFromPoint(outPointOnCL2.Value, outPointOnCL1, station1)  # with that Point compute the Chainage on Line 1
                                                # check if it falls within our
                                                                                                                                                                     # range
                                                if station1.Value >= computestartstation and station1.Value <= computeendstation:
                                                    chainages.Add(round(station1.Value, rounddecimals))

                                                j += 1
                                                if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(j * 100 / (node_count))):
                                                    break   # function returns true if user pressed cancel

                                                    ## check if we've got that
                                                    ## station already
                                                    #canadd = True
                                                    #for c in chainages:
                                                    #    if round(c, rounddecimals)
                                                    #    - round(station1.Value,
                                                    #    rounddecimals) == 0:
                                                    #        canadd = False
                                                    #if canadd == True:
                                                    #chainages.Add(station1.Value)
                                                    #chainages.Add(station1.Value)
                                        except:
                                            break

                        ## add the interval range on the first line to the chainage
                        ## list
                        if interval != 0:
                            while computestartstation <= computeendstation:    # as long as we aren't at the end of the line 1
                                # avoid duplicates in the chainage list, but keep
                                                                                                       # full precision for the grade breaks
                                #canadd = True
                                                                                                       #for c in chainages:
                                #    if round(c, rounddecimals) -
                                                                                                       #    round(computestartstation, rounddecimals) ==
                                                                                                       #    0:
                                #        canadd = False
                                #if canadd == True:
                                chainages.Add(computestartstation)
                            
                                if computestartstation == computeendstation: break
                                computestartstation += interval
                                if computestartstation > computeendstation: computestartstation = computeendstation
                        
                    else:
                        self.success.Content += '\nEnter Interval > 0 or use Bundle Kinks'
                        
                    
                    chainages = list(set(chainages)) # remove duplicates
                    chainages.sort()  # sort list

                    chainage_progress_count = 0

                    if chainages.Count > 0:
                        for computestation in chainages:    # go through the chainage list and compute perpendicular onto each line

                            p2_count = 0    

                            # 0 - chainage, 1 - max_osl, 2 - max_osr, 3 - min_z, 4
                            # - max_z, 5 - point bottom left, 6 - top left, 7 - top
                            # right, 8 - bottom right
                            limit_at_station = Array[float](5 * [0.000]) + Array[Point3D](4 * [Point3D()])

                            limit_at_station[0] = computestation

                            chainage_progress_count += 1
                            ProgressBar.TBC_ProgressBar.Title = "computing Envelope at Chainage " + str(chainage_progress_count) + " / " + str(chainages.Count)
                            if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(chainage_progress_count * 100 / chainages.Count)):
                                break   # function returns true if user pressed cancel

                            # compute a point on line 1 and get the perpendicular
                            # vector
                            polyseg1.FindPointFromStation(computestation, outSegment1, out_t, outPointOnCL1, perpVector3D, outdeflectionAngle) # compute point and vector on line 1
                        
                            p1 = outPointOnCL1.Value
                            if polyseg1_v != None:
                                polyseg1_v.ComputeVerticalSlopeAndGrade(computestation, outelevation1, outslope1)
                            p1.Z = outelevation1.Value

                            perpseg = PolySeg.PolySeg(SegmentLine(p1, perpVector3D.Value))
                            perpseg.Extend(10000.0)

                            for o in self.objs: # go through all the other lines
                                if isinstance(o, self.lType):

                                    polyseg2 = o.ComputePolySeg()
                                    polyseg2 = polyseg2.ToWorld()
                                    polyseg2_v = o.ComputeVerticalPolySeg()

                                    foundp2 = False
                                    finalp2 = Point3D.Zero
                                    if perpseg.Intersect(polyseg2, True, intersections):
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
                            
                                    ##### we have to go through all segments of line 2
                                    ##### and check if we intersect
                                    ##### we have to do it that way since we don't know
                                    ##### if and when our perpvector
                                    ##### would actually intersect the second line
                                    ####s = polyseg2.FirstSegment   # we get us the first segment of line 2
                                    ####while s is not None:
                                    ####    foundvalidintersection = False
                                    ####    # we use the Segment.Intersect function
                                    ####    # since we can use True to extend the lines
                                    ####    # unfortunately it extends both lines
                                    ####    # indefintely, which we don't really want
                                    ####    # for
                                    ####    # the segments of line2
                                    ####    if perpseg.Intersect(s, True, intersections):
                                    ####        if intersections.Count == 1: # if we only get one intersection
                                    ####            # we check if we are actually
                                    ####                                                                                # inside the segment of line 2
                                    ####                                                                                # and not on the extension
                                    ####            if intersections[0].T2 >= 0 and intersections[0].T2 <= 1: 
                                    ####                ip = intersections[0].Point
                                    ####                foundvalidintersection = True
                                    ####        else:
                                    ####            i2 = 0
                                    ####            for i in intersections: # if we get multiple intersections (on arcs) we want the shortest one
                                    ####                if i.T2 >= 0 and i.T2 <= 1:    
                                    ####                    foundvalidintersection = True
                                    ####                    if i2 == 0:
                                    ####                        ip = i.Point
                                    ####                        i2 += 1
                                    ####                    else:
                                    ####                        if p1.Distance2D(i.Point) < p1.Distance2D(ip):
                                    ####                            ip = i.Point
                                    ####                
                                    ####        if foundvalidintersection:        
                                    ####            polyseg2.FindPointFromPoint(ip, outPointOnCL2, station2)
                                    ####            p2 = outPointOnCL2.Value
                                    ####            if polyseg2_v != None:
                                    ####                p2.Z = polyseg2_v.ComputeVerticalSlopeAndGrade(station2.Value)[1]
                                    ####            foundp2 = True
                                    ####    
                                    ####    # in tight corners we only want the first
                                    ####    # intersection
                                    ####    if foundp2 == True and \
                                    ####       (finalp2 == Point3D.Zero or Vector3D(p1, p2).Length2D < Vector3D(p1, finalp2).Length2D):
                                    ####        finalp2 = p2
                                    ####
                                    ####    intersections.Clear()
                                    ####    s = polyseg2.Next(s)
                        
                                    # if we found a point on the bundle line
                                    if foundp2:
                                        #p2 = finalp2
                                        p2_count += 1
                                        # we need to know on which side p2/l2 are
                                        polyseg1.FindPointFromPoint(p2, outSegment1, out_t, outPointOnCL1, station1, outaz, out_offset, out_side)
                                        if out_side.Value == Side.Right:
                                            p2_os = out_offset.Value
                                        else:
                                            p2_os = -1 * out_offset.Value
                                        
                                        if p2.X == 0:
                                            tt = p2

                                        if p2_count == 1:
                                            limit_at_station[1] = p2_os
                                            limit_at_station[2] = p2_os
                                            limit_at_station[3] = p2.Z
                                            limit_at_station[4] = p2.Z
                                            limit_at_station[5] = p2
                                            limit_at_station[7] = p2
                                # 0 - chainage, 1 - max_osl, 2 - max_osr, 3 -
                                # min_z, 4 - max_z, 5 - point bottom left, 6 - top
                                # left, 7 - top right, 8 - bottom right
                                        else:
                                            if p2_os < limit_at_station[1]:
                                                limit_at_station[1] = p2_os
                                                limit_at_station[5] = p2
                                            if p2_os > limit_at_station[2]:
                                                limit_at_station[2] = p2_os
                                                limit_at_station[7] = p2
                                            if p2.Z < limit_at_station[3]: limit_at_station[3] = p2.Z
                                            if p2.Z > limit_at_station[4]: limit_at_station[4] = p2.Z
                            
                            # compute the corners
                            # bottom left
                            bl = Point3D(limit_at_station[5].X, limit_at_station[5].Y, limit_at_station[3])
                            limit_at_station[5] = bl
                            # top left
                            tl = Point3D(limit_at_station[5].X, limit_at_station[5].Y, limit_at_station[4])
                            limit_at_station[6] = tl
                            # top right
                            tr = Point3D(limit_at_station[7].X, limit_at_station[7].Y, limit_at_station[4])
                            limit_at_station[7] = tr
                            # top left
                            br = Point3D(limit_at_station[7].X, limit_at_station[7].Y, limit_at_station[3])
                            limit_at_station[8] = br

                            if p2_count > 0:
                                all_limits.Add(limit_at_station.Clone())

                        # tt = all_limits

                        # 0 - chainage, 1 - max_osl, 2 - max_osr, 3 - min_z, 4 -
                        # max_z, 5 - point bottom left, 6 - top left, 7 - top
                        # right, 8 - bottom right
                        
                        # draw rectangle
                        if self.drawrectangle.IsChecked:
                            limit_progress_count = 0
                            for limit in all_limits:

                                limit_progress_count += 1
                                ProgressBar.TBC_ProgressBar.Title = "draw Rectangle " + str(limit_progress_count) + " / " + str(all_limits.Count)
                                if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(chainage_progress_count * 100 / chainages.Count)):
                                    break   # function returns true if user pressed cancel

                                new_l = wv.Add(clr.GetClrType(Linestring))

                                e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                e.Position = limit[5]
                                new_l.AppendElement(e)
                                e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                e.Position = limit[6]
                                new_l.AppendElement(e)
                                e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                e.Position = limit[7]
                                new_l.AppendElement(e)
                                e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                e.Position = limit[8]
                                new_l.AppendElement(e)
                                e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                e.Position = limit[5]
                                new_l.AppendElement(e)

                                new_l.Layer = self.layerpicker.SelectedSerialNumber


                        if self.drawpoints.IsChecked: # if we want to draw all points
                            limit_progress_count = 0
                            for limit in all_limits:

                                limit_progress_count += 1
                                ProgressBar.TBC_ProgressBar.Title = "draw Points " + str(limit_progress_count) + " / " + str(all_limits.Count)
                                if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(chainage_progress_count * 100 / chainages.Count)):
                                    break   # function returns true if user pressed cancel

                                cadPoint = wv.Add(clr.GetClrType(CadPoint))
                                cadPoint.Layer = self.layerpicker.SelectedSerialNumber
                                cadPoint.Point0 = limit[5]
                                cadPoint = wv.Add(clr.GetClrType(CadPoint))
                                cadPoint.Layer = self.layerpicker.SelectedSerialNumber
                                cadPoint.Point0 = limit[6]
                                cadPoint = wv.Add(clr.GetClrType(CadPoint))
                                cadPoint.Layer = self.layerpicker.SelectedSerialNumber
                                cadPoint.Point0 = limit[7]
                                cadPoint = wv.Add(clr.GetClrType(CadPoint))
                                cadPoint.Layer = self.layerpicker.SelectedSerialNumber
                                cadPoint.Point0 = limit[8]

                        if self.drawlonglines.IsChecked: # draw the longitudinal lines
                            new_l1 = wv.Add(clr.GetClrType(Linestring))
                            new_l2 = wv.Add(clr.GetClrType(Linestring))
                            new_l3 = wv.Add(clr.GetClrType(Linestring))
                            new_l4 = wv.Add(clr.GetClrType(Linestring))
                            new_l1.Layer = self.layerpicker.SelectedSerialNumber
                            new_l2.Layer = self.layerpicker.SelectedSerialNumber
                            new_l3.Layer = self.layerpicker.SelectedSerialNumber
                            new_l4.Layer = self.layerpicker.SelectedSerialNumber
                            
                            limit_progress_count = 0
                            for limit in all_limits:

                                limit_progress_count += 1
                                ProgressBar.TBC_ProgressBar.Title = "draw longitudinal Lines " + str(limit_progress_count) + " / " + str(all_limits.Count)
                                if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(chainage_progress_count * 100 / chainages.Count)):
                                    break   # function returns true if user pressed cancel

                                e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                e.Position = limit[5]
                                new_l1.AppendElement(e)
                                
                                e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                e.Position = limit[6]
                                new_l2.AppendElement(e)

                                e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                e.Position = limit[7]
                                new_l3.AppendElement(e)

                                e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                e.Position = limit[8]
                                new_l4.AppendElement(e)

                        if self.draw3dshell.IsChecked: # draw the 3D-Shell lines
                            # prepare the empty arrays
                            ifcvertices = Array[Point3D]([Point3D()] * (all_limits.Count * 4))
                            ifcfacelist = Array[Int32]([Int32()] * (((all_limits.Count - 1) * 4  * 2 * 4) + 16)) # 4 sided * 2 triangles * 4 entries + 16 at the start/end
                            ifcnormals = Array[Point3D]([Point3D()]*0)
                            
                            ifcverticeindex = 0                                                              #         11----10
                            for limit in all_limits:                                                         #         /|    | 
                                ifcvertices[ifcverticeindex + 0] = limit[5]                                  #        / 8----9
                                ifcvertices[ifcverticeindex + 1] = limit[6]                                  #       /    
                                ifcvertices[ifcverticeindex + 2] = limit[7]                                  #      7----6
                                ifcvertices[ifcverticeindex + 3] = limit[8]                                  #     /|    |
                                ifcverticeindex += 4                                                         #    / 4----5      
                                                                                                             #   /          
                            ifcfacelistindex = 0                                                             #  3----2      
                            nodecount = 4                                                                    #  |    |      
                            for c in range(0, all_limits.Count - 1):                                         #  0----1      
                                
                                if c == 0: # close front of shape
                                    ifcfacelist[ifcfacelistindex + 0] = 3 # always 3 vertices per triangle         
                                    ifcfacelist[ifcfacelistindex + 1] = (c * nodecount) + 0               # 0      
                                    ifcfacelist[ifcfacelistindex + 2] = (c * nodecount) + 3               # 3
                                    ifcfacelist[ifcfacelistindex + 3] = (c * nodecount) + 2               # 2
                                    ifcfacelistindex += 4

                                    ifcfacelist[ifcfacelistindex + 0] = 3 # always 3 vertices per triangle         
                                    ifcfacelist[ifcfacelistindex + 1] = (c * nodecount) + 0               # 0      
                                    ifcfacelist[ifcfacelistindex + 2] = (c * nodecount) + 2               # 2
                                    ifcfacelist[ifcfacelistindex + 3] = (c * nodecount) + 1               # 1      
                                    ifcfacelistindex += 4

                                if c == all_limits.Count - 2: # close back of shape
                                    ifcfacelist[ifcfacelistindex + 0] = 3 # always 3 vertices per triangle         
                                    ifcfacelist[ifcfacelistindex + 1] = ((c + 1) * nodecount) + 0               # 4      
                                    ifcfacelist[ifcfacelistindex + 2] = ((c + 1) * nodecount) + 1               # 5
                                    ifcfacelist[ifcfacelistindex + 3] = ((c + 1) * nodecount) + 2               # 6
                                    ifcfacelistindex += 4
                                
                                    ifcfacelist[ifcfacelistindex + 0] = 3 # always 3 vertices per triangle         
                                    ifcfacelist[ifcfacelistindex + 1] = ((c + 1) * nodecount) + 0               # 4      
                                    ifcfacelist[ifcfacelistindex + 2] = ((c + 1) * nodecount) + 2               # 6
                                    ifcfacelist[ifcfacelistindex + 3] = ((c + 1) * nodecount) + 3               # 7      
                                    ifcfacelistindex += 4

                                for f in range(0, nodecount - 1):                                                  
                                                                                                                   
                                    ifcfacelist[ifcfacelistindex + 0] = 3 # always 3 vertices per triangle         
                                    ifcfacelist[ifcfacelistindex + 1] = (c * nodecount) + f               # 0      
                                    ifcfacelist[ifcfacelistindex + 2] = (c * nodecount) + f + 1           # 1      
                                    ifcfacelist[ifcfacelistindex + 3] = ((c + 1) * nodecount) + f + 1     # 5      
                                    ifcfacelistindex += 4

                                    ifcfacelist[ifcfacelistindex + 0] = 3 # always 3 vertices per triangle
                                    ifcfacelist[ifcfacelistindex + 1] = (c * nodecount) + f               # 0
                                    ifcfacelist[ifcfacelistindex + 2] = ((c + 1) * nodecount) + f + 1     # 5
                                    ifcfacelist[ifcfacelistindex + 3] = ((c + 1) * nodecount) + f         # 4
                                    ifcfacelistindex += 4


                                # close the shape
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

                            self.createifcobject("LineBundles", "Bundle1", "Bundle1", ifcvertices, ifcfacelist, ifcnormals, Color.Red.ToArgb(), Matrix4D(), self.layerpicker.SelectedSerialNumber)
                            #mesh._TriangulatedFaceList = ifcfacelist # otherwise GetTriangulatedFaceList() in the perpDist macro will not get the necessary values



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

        self.success.Content += '\nDone'
        ProgressBar.TBC_ProgressBar.Title = ""
        self.SaveOptions()

    
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
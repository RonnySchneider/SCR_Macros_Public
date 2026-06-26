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
    "creationmode": True, "layerpicker": 8,
    "createnamedpoint": False, "pointnametext": "", "pointnamestart": 0.0, "pointnameincrement": 0.0,
    "chooserange": False, "startstation": 0.0, "endstation": 0.0,
    "interval": 1.0, "eloffset": 0.0, "includegradechanges": True,
    "drawxlines": False, "benchpattern": False, "drawpoints": True,
    "inspectionmode": False, "refstation": 0.0,
}

def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_CreatePointsOnXlines"
    cmdData.CommandName = "SCR_CreatePointsOnXlines"
    cmdData.Caption = "_SCR_CreatePointsOnXlines"
    cmdData.UIForm = "SCR_CreatePointsOnXlines"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "Points"
        cmdData.ShortCaption = "Points on Xlines"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.29
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Create Points where a Xline crosses another line"
        cmdData.ToolTipTextFormatted = "Create Points where a Xline crosses another line"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_CreatePointsOnXlines(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_CreatePointsOnXlines.xaml") as s:
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
        self.linepicker4.IsEntityValidCallback = self.IsValid

        #self.interval.NumberOfDecimals = 4
        self.interval.DistanceMin = 0.000000001
        #self.eloffset.NumberOfDecimals = 4
        self.pointnamestart.NumberOfDecimals = 0
        self.pointnameincrement.NumberOfDecimals = 0

        self.lType = clr.GetClrType(IPolyseg)

        SCRExpanders.wire_pairs([
            (self.expander_creationmode, self.creationmode),
            (self.expander_inspectionmode, self.inspectionmode),
        ])
		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

        self.linepicker1.ValueChanged += self.lineChanged
        self.linepicker2.ValueChanged += self.lineChanged
        self.linepicker3.ValueChanged += self.lineChanged
        self.linepicker4.ValueChanged += self.lineChanged
        self.linepicker4.ValueChanged += self.autorun
        self.linepicker4.AutoTab = False

        self.startstation.ValueChanged += self.lineChanged
        self.endstation.ValueChanged += self.lineChanged
        self.chooserange.Checked += self.lineChanged
        self.chooserange.Unchecked += self.lineChanged

        self.creationmode.Checked += self.lineChanged
        self.inspectionmode.Checked += self.lineChanged

        self.refstation.ValueChanged += self.refstationChanged
        self.refstation.AutoTab = False

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        self.lfp = self.lunits.Properties.Copy()
        linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation
        self.lfp.AddSuffix = False
        self.inspectionresultel.Content = "Elevation on Line 2 [" + linearsuffix + "]: "
        self.inspectionresultos.Content = "Offset to Line 2 [" + linearsuffix + "]: "
        self.labelinterval.Content = "Interval [" + linearsuffix + "]: "
        self.labeleloffset.Content = "Elevation Offset [" + linearsuffix + "]: "


    def SetDefaultOptions(self):
        SCROptions.LoadMacroOptions(self, "SCR_CreatePointsOnXlines", _OPTIONS, self.currentProject)

    def SaveOptions(self):
        SCROptions.SaveMacroOptions(self, "SCR_CreatePointsOnXlines", _OPTIONS)

    def IsValid(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        return False

    def lineChanged(self, ctrl, e):
        
        if self.creationmode.IsChecked:
            l1 = self.linepicker1.Entity
            if l1 != None:
                self.startstation.StationProvider = l1
                self.endstation.StationProvider = l1

        else: #inspection mode
            l3 = self.linepicker3.Entity
            if l3 != None:
                self.refstation.StationProvider = l3
        self.drawoverlay()

    def refstationChanged(self, ctrl, e):
        if e.Cause == InputMethod.Mouse:
            self.OkClicked(None, None)
        Keyboard.Focus(self.refstation)

    def drawoverlay(self):

        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)
        self.overlayBag = OverlayBag(self.ViewOverlay) # create a new/empty overlaybag

        l1 = self.linepicker1.Entity
        l2 = self.linepicker2.Entity
        l3 = self.linepicker3.Entity
        l4 = self.linepicker4.Entity

        if self.creationmode.IsChecked:
            if l1:
                self.overlayBag.AddPolyline(SCROverlayBag.getpolypoints(l1), Color.Green.ToArgb(), 4)
        
                if self.chooserange.IsChecked:
                    self.overlayBag.AddPolyline(SCROverlayBag.getclippedpolypoints(l1, self.startstation.Distance, self.endstation.Distance), Color.Orange.ToArgb(), 2)

                for p in SCROverlayBag.getarrowlocations(l1, 10): # returns list with location and perp right azimuth [Point3D, perpVector3D.Value.Azimuth]
                    self.overlayBag.AddMarker(p[0], GraphicMarkerTypes.Arrow_IndependentColor, Color.Orange.ToArgb(), "", 0, math.pi - p[1], 3.0)

            if l2:
                self.overlayBag.AddPolyline(SCROverlayBag.getpolypoints(l2), Color.Blue.ToArgb(), 4)
                for p in SCROverlayBag.getarrowlocations(l2, 10): # returns list with location and perp right azimuth [Point3D, perpVector3D.Value.Azimuth]
                    self.overlayBag.AddMarker(p[0], GraphicMarkerTypes.Arrow_IndependentColor, Color.Orange.ToArgb(), "", 0, math.pi - p[1], 3.0)
        else:
            if l3:
                self.overlayBag.AddPolyline(SCROverlayBag.getpolypoints(l3), Color.Red.ToArgb(), 4)
                for p in SCROverlayBag.getarrowlocations(l3, 10): # returns list with location and perp right azimuth [Point3D, perpVector3D.Value.Azimuth]
                    self.overlayBag.AddMarker(p[0], GraphicMarkerTypes.Arrow_IndependentColor, Color.Orange.ToArgb(), "", 0, math.pi - p[1], 3.0)
            if l4:
                self.overlayBag.AddPolyline(SCROverlayBag.getpolypoints(l4), Color.Magenta.ToArgb(), 4)
                for p in SCROverlayBag.getarrowlocations(l4, 10): # returns list with location and perp right azimuth [Point3D, perpVector3D.Value.Azimuth]
                    self.overlayBag.AddMarker(p[0], GraphicMarkerTypes.Arrow_IndependentColor, Color.Orange.ToArgb(), "", 0, math.pi - p[1], 3.0)

        # if it only needs to be visible in all Planview then remove the Hoops3DViewGUID
        array = Array[Guid]([DisplayWindow.Hoops3DViewGUID, DisplayWindow.HoopsPlanViewGUID])
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.AddOverlayGeometry(array, self.overlayBag)

        return

    def autorun(self, sender, e):
        if e.Cause == InputMethod.Mouse:     
            self.OkClicked(None, None)
        Keyboard.Focus(self.linepicker4)

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def Dispose(self, cmd, disposing):
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''
        self.success.Content=''

        wv = self.currentProject [Project.FixedSerial.WorldView]
        wv.PauseGraphicsCache(True)

        inputok=True

        if self.creationmode.IsChecked:
            l1 = self.linepicker1.Entity
            l2 = self.linepicker2.Entity
        else: # inspection
            l1 = self.linepicker3.Entity
            l2 = self.linepicker4.Entity

        if l1==None: 
            self.success.Content += '\nno Line 1 selected'
            inputok=False
        if l2==None:
            self.success.Content += '\nno Line 2 selected'
            inputok=False
        
        if l1 == l2: inputok = False

        if self.creationmode.IsChecked:
            try:
                startstation = self.startstation.Distance
            except:
                self.success.Content += '\nStart Chainage error'
                inputok=False
        else: # inspection
            try:
                startstation = self.refstation.Distance
            except:
                self.success.Content += '\nReference Chainage error'
                inputok=False

        if self.creationmode.IsChecked:
            try:
                endstation = self.endstation.Distance
            except:
                self.success.Content += '\nEnd Chainage error'
                inputok=False
        else: # inspection
            try:
                endstation = startstation
            except:
                self.success.Content += '\nReference Chainage error'
                inputok=False

        
        if inputok:
            UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
            self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)

            outPointOnCL2 = clr.StrongBox[Point3D]()
            station1 = clr.StrongBox[float]()
            station2 = clr.StrongBox[float]()
            p1 = Point3D()

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

                    if self.chooserange.IsChecked or self.inspectionmode.IsChecked:
                        computestartstation = startstation
                        computeendstation = endstation
                    else:
                        computestartstation = polyseg1.BeginStation
                        computeendstation = polyseg1.ComputeStationing()

                    # compile chainage list
                    chainages = []
                    # add grade changes on line 2
                    if self.includegradechanges.IsChecked and not self.inspectionmode.IsChecked:
                        if polyseg2_v != None:
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
                        else:
                            for n in polyseg2.Nodes():
                                try:
                                    if n.Visible:
                                        polyseg1.FindPointFromPoint(n.Point, outPointOnCL1, station1)  # with that Point compute the Chainage on Line 1
                                        # check if it falls within our range
                                        if station1.Value >= computestartstation and station1.Value <= computeendstation:
                                            chainages.Add(station1.Value)
                                except:
                                    break

                    rounddecimals = self.currentProject.Units.Station.Properties.NumberOfDecimals
                    # add the interval range on the first line
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
           


                    chainages = list(set(chainages)) # remove duplicates, there shouldn't actually be any
                    chainages.sort()  # sort list

                    intersections = Intersections()
                    j = 0
                    outSegment = clr.StrongBox[Segment]()
                    out_t = clr.StrongBox[float]()
                    outPointOnCL1 = clr.StrongBox[Point3D]()
                    perpVector3D = clr.StrongBox[Vector3D]()
                    outdeflectionAngle = clr.StrongBox[float]()

                    #start_t = timer()
                    for c in range(chainages.Count):    # go through the chainage list
                        computestation = chainages[c]
                        # compute a point on line 1 and get the perpendicular vector
                        polyseg1.FindPointFromStation(computestation, outSegment, out_t, outPointOnCL1, perpVector3D, outdeflectionAngle) # compute point and vector on line 1
                        
                        p1 = outPointOnCL1.Value
                        if polyseg1_v != None:
                            p1.Z = polyseg1_v.ComputeVerticalSlopeAndGrade(computestation)[1]
                        
                        perpseg = PolySeg.PolySeg(SegmentLine(p1, perpVector3D.Value))
                        perpseg.Extend(10000.0)
                        #finalp2 = Point3D.Zero
                        
                        foundp2 = False
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
                            for i in intersections: # if we get multiple intersections (on arcs) we want the shortest one
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
                                p2 = outPointOnCL2.Value
                                if polyseg2_v != None:
                                    p2.Z=polyseg2_v.ComputeVerticalSlopeAndGrade(station2.Value)[1]
                                foundp2 = True
                            
                            ### # in tight corners we only want the first intersection        
                            ### if foundp2 == True and \
                            ###    (finalp2 == Point3D.Zero or Vector3D(p1, p2).Length2D < Vector3D(p1, finalp2).Length2D):
                            ###     finalp2 = p2

                        intersections.Clear()
                            ### s = polyseg2.Next(s)


                        if foundp2:
                            #p2 = finalp2

                            if self.creationmode.IsChecked:
                                p1.Z += self.eloffset.Distance
                                p2.Z += self.eloffset.Distance        # add the elevation offset
                                
                                if self.drawpoints.IsChecked:
                                    if self.createnamedpoint.IsChecked:
                                        pnew_wv = CoordPoint.CreatePoint(self.currentProject, self.pointnametext.Text + '{0:g}'.format(round(self.pointnamestart.Value, 0)))
                                        pnew_wv.AddPosition(p2)
                                        pnew_wv.Layer = self.layerpicker.SelectedSerialNumber
                                        self.pointnamestart.Value += self.pointnameincrement.Value
                                    else:
                                        cadPoint = wv.Add(clr.GetClrType(CadPoint))
                                        cadPoint.Layer = self.layerpicker.SelectedSerialNumber
                                        cadPoint.Point0 = p2

                                # draw lines between p1 and p2
                                if self.drawxlines.IsChecked:
                                    l = wv.Add(clr.GetClrType(Linestring))
                                    l.Layer = self.layerpicker.SelectedSerialNumber
                                    #e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                    #e.Position = p1 
                                    #l.AppendElement(e)
                                    #e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                    #e.Position = p2
                                    #l.AppendElement(e)

                                    polyseg = PolySeg.PolySeg()
                                    if self.benchpattern.IsChecked and c % 2 != 0:
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

                            else:
                                self.inspectel.Content = self.lunits.Format(p2.Z, self.lfp)
                                self.inspectos.Content = self.lunits.Format(Vector3D(p1, p2).Length2D, self.lfp)
                        else:
                            if self.inspectionmode.IsChecked:
                                self.inspectel.Content = "no solution found"
                                self.inspectos.Content = ""
                        
                        ## update progress bar
                        #j += 1
                        #ProgressBar.TBC_ProgressBar.Title = "compute Point: " + str(j) + "/" + str(chainages.Count)
                        #if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(j * 100 / chainages.Count)):
                        #    break   # function returns true if user pressed cancel

                    #end_t=timer()
                    #self.error.Content += '\nOverall - ' + str(timedelta(seconds=end_t-start_t))

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
        
        #Keyboard.Focus(self.refstation)

        wv.PauseGraphicsCache(False)

        self.SaveOptions()

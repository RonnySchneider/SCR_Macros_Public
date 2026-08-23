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
    "drawMin":              True,
    "layerPickerMin":       8,              # 8 = Layer Zero FixedSerial
    "colourPickerMin":      Color.Red,
    "drawMax":              True,
    "layerPickerMax":       8,
    "colourPickerMax":      Color.Green,
    "drawGutter":           False,
    "layerPickerGutter":    8,
    "colourPickerGutter":   Color.Blue,
    "circleSelect":         False,
    "squareSelect":         True,
    "pickerBearing":        0.0,
    "apertureRadius":       0.25,
    "apertureHeight":       3.0,
    "colourPickerAperture": Color.Lime,
}
def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_PointcloudExtractor"
    cmdData.CommandName = "SCR_PointcloudExtractor"
    cmdData.Caption = "_SCR_PointcloudExtractor"
    cmdData.UIForm = "SCR_PointcloudExtractor"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "Pointcloud"
        cmdData.ShortCaption = "Extract from Cloud"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.12
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Extract from Cloud"
        cmdData.ToolTipTextFormatted = "Extract from Cloud"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_PointcloudExtractor(StackPanel): # this inherits from the WPF StackPanel control
    BENCHMARKING = False

    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_PointcloudExtractor.xaml") as s:
            wpf.LoadComponent (self, s)
        self.currentProject = currentProject
        self.macroFileFolder = macroFileFolder

        self.ViewOverlay = Guid.NewGuid()
        self.overlayBag = OverlayBag(self.ViewOverlay)

        self.myLineHigh = 0 # serial number of linestring
        self.myLineLow = 0 # serial number of linestring
        self.myLineGutter = 0 # serial number of linestring

        self.currminp = Point3D()
        self.currmaxp = Point3D()
        self.currGutterp = Point3D()

        self.activeViewFilter = None
        self.activeForm = None

    def HelpClicked(self, _cmd, _e):
        webbrowser.open("C:\ProgramData\Trimble\MacroCommands3\SCR Macros\MacroHelp\MacroHelp.htm#" + type(self).__name__)

    def OnLoad(self, cmd, buttons, event):

        self.okBtn = buttons[0]
        buttons[2].Content = "Help"
        buttons[2].Visibility = Visibility.Visible
        buttons[2].Click += self.HelpClicked
        self.Caption = cmd.Command.Caption

        self.coordCtl.ValueChanged += self.CoordChanged
        self.coordCtl.ShowGdiCursor += self.drawoverlay

        self.btnNewMin.Click += self.BtnNewClicked
        self.btnNewMax.Click += self.BtnNewClicked
        self.btnNewGutter.Click += self.BtnNewClicked

        self.apertureRadius.ValueChanged += self.radiuschange
        self.circleSelect.Checked += self.radiuschange
        self.squareSelect.Checked += self.radiuschange

        self.pickerBearing.ValueChanged += self.bearingChanged
        self.btnResetBearing.Click += self.btnResetBearingClicked

        SCRExpanders.wire_pairs([
            (self.expander_drawMin, self.drawMin),
            (self.expander_drawMax, self.drawMax),
            (self.expander_drawGutter, self.drawGutter),
        ])
        try:
            self.SetDefaultOptions()
        except:
            pass

        self.visiblesdecloud = None
        self.visiblesdecloudprefiltered = None
        self.cachewaspaused = False

        self.activeForm = TrimbleOffice.TheOffice.MainWindow.AppViewManager.ActiveView
        try:
            self.activeViewFilter = self.currentProject.Concordance[self.activeForm.ViewFilter]  
            self.activeViewFilter.FilterChanged += self.viewfiltersettingschangeevent
        except:
            pass

        TrimbleOffice.TheOffice.MainWindow.AppViewManager.ViewActivated += self.activeviewchanged

        UIEvents.UIViewFilterChanged += self.activeviewchanged # i.e. if View Filter for a window changed

        self.activeviewchanged(None, None) # get the active window on startup
        self.radiuschange(None, None)
        self.viewfiltersettingschangeevent(None, None)

        # We want to start cmd with focus on coordinate control.  It's a bit
        # tricky to set focus in the OnLoad method.  The following lines do
        # that.
        def SetFocusToControl():
            Keyboard.Focus(self.coordCtl)
            
        #if self.nameCtl.AutoTabExtSkip:
        #    Dispatcher.BeginInvoke(Dispatcher.CurrentDispatcher, Action(SetFocusToControl))

    def activeviewchanged(self, sender, e):

        oldviewfilter = None

        if self.activeViewFilter:
            self.activeViewFilter.FilterChanged -= self.viewfiltersettingschangeevent
            oldviewfilter = self.activeViewFilter

        self.activeForm = TrimbleOffice.TheOffice.MainWindow.AppViewManager.ActiveView
        
        self.activeViewFilter = None
        try:
            self.activeViewFilter = self.currentProject.Concordance[self.activeForm.ViewFilter]   
            self.activeViewFilter.FilterChanged += self.viewfiltersettingschangeevent
        except:
            pass
        
        if self.activeViewFilter != None and oldviewfilter != None and oldviewfilter.SerialNumber != self.activeViewFilter.SerialNumber:
            self.viewfiltersettingschangeevent(None, None) # a full view filter change is the same as if a layer was ticked on/off in the view filter manager

    def viewfiltersettingschangeevent(self, _sender, _e):
        
        # get rid of potentially existing SdeCloud object
        try:
            self.visiblesdecloud.Dispose()
            self.visiblesdecloud = None
        except: self.visiblesdecloud = None
        try:
            self.visiblesdecloudprefiltered.Dispose()
            self.visiblesdecloudprefiltered = None
        except: self.visiblesdecloudprefiltered = None

        try:
            if self.activeViewFilter and ExtractionTools().PointCloudVisible() and not self.activeForm.View.CachePaused:
                self.visiblesdecloud = ExtractionTools().GetVisiblePoints(self.currentProject)
                self.cachewaspaused = False

        except Exception:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            self.debug2.Content += '\nan Error occurred - Result probably incomplete\n' + str(exc_type) + '\n' + str(exc_obj) + '\nLine ' + str(exc_tb.tb_lineno)

        
        # with point clouds it can be that the view is still loading and the Cache paused
        # can't use wait or sleep in a while loop
        # the cache won't turn unpaused as long as the python script is running    
        if self.activeForm and self.activeForm.View.CachePaused:
            self.cachewaspaused = True

        ProgressBar.TBC_ProgressBar.Title = ""
        if self.visiblesdecloud:
            self.debug.Content = self.visiblesdecloud.PointCount.ToString("N0") + ' points on screen'
        else:
            self.debug.Content = 'currently no cloud points retrieved from DB'
            if self.cachewaspaused:
                self.debug.Content += '\nCache was still paused - will try again upon next Mouse-Move'

        #self.debug.Content +='\n' + str((datetime.now() - octreestart).seconds) + " Seconds"
        #tt = 1

        #self.activeViewFilter.FilterChanged += self.viewfiltersettingschangeevent

    def radiuschange(self, _sender, _e):

        shape = PolySeg.PolySeg()
        r = self.apertureRadius.Distance
        
        if self.circleSelect.IsChecked:

            shape.Add(ArcSegment(Point3D(-r, 0, 0), Point3D(r, 0, 0), r))
            shape.Add(ArcSegment(Point3D(r, 0, 0), Point3D(-r, 0, 0), r))

        elif self.squareSelect.IsChecked:

            
            
            shape.Add(List[Point3D]([Point3D(-r, -r, 0),
                                     Point3D( r, -r, 0),
                                     Point3D( r,  r, 0),
                                     Point3D(-r,  r, 0)]))
            shape.Close(True)

        self.aperturepolychord = shape.Linearize(0.002, 0.002, 50, None, False)

        self.viewfiltersettingschangeevent(None, None)


    def moveshapetosnap(self, p, nv):

        if nv == Vector3D(0,0,1):
            td = Matrix4D.BuildTransformMatrix(Point3D(0, 0, 0), p, self.pickerBearing.Direction, 1, 1, 1)
        else:
            # will only work if Normal != 0, 0, 1
            vx = nv.Clone()
            # rotate 90 degrees around world-Z axis and make it level with world horizon
            vx.RotateAboutZ(math.pi/2)
            vx.Horizon = 0
            # compute matrix to line up UCS-X,Z with World-X,Z axis
            rottozero = Spinor3D.ComputeRotation(vx, nv, Vector3D(1,0,0), Vector3D(0,0,1))
            # transformation to 0, 0, 0
            matrixtozero = Matrix4D.BuildTransformMatrix(Vector3D(p), Vector3D(p, Point3D(0, 0, 0)), rottozero, Vector3D(1,1,1))
            td = Matrix4D.Inverse(matrixtozero)

            td = td.RotateZ(0.2)

        newchord = self.aperturepolychord.Clone()
        newchord.Transform(td)
        return newchord.ToPoint3DArray()

    def prefilter(self, closesttocursor): # has been checked before the call that it exists and is 3D
        
        r5 = 10 * self.apertureRadius.Distance
        sdefilterorigin5 = SdePoint3D(closesttocursor.X - r5, closesttocursor.Y - r5, closesttocursor.Z - r5)
        sdefilterbox5 = SdeFilterByBox(sdefilterorigin5, SdeVector3D(2*r5 ,0, 0), SdeVector3D(0, 2*r5, 0), SdeVector3D(0, 0, 2*r5))

        r1 = self.apertureRadius.Distance
        aperturefilterbox = SdeRect3D(closesttocursor.X - r1, closesttocursor.Y - r1, closesttocursor.Z - r1, 2*r1, 2*r1, 2*r1)

        resetprefilter = False
        errorcleanup = False

        # if we have a prefiltered cloud and the bounding box for it we can try to evaluate if the aperture is still in it
        if self.visiblesdecloudprefiltered and self.oldprefilterbox != None:

            try:
                resetprefilter = not self.oldprefilterbox.Contains(aperturefilterbox)
            except:
                resetprefilter = False
                errorcleanup = True

        # if we haven't produced an exception yet, and we don't have a prefilter yet or not to reset the prefilter
        if not errorcleanup and (not self.visiblesdecloudprefiltered or resetprefilter):
            try:
                exclusion = clr.StrongBox[SdeCloud]()
                self.visiblesdecloudprefiltered = SdeCloud(self.visiblesdecloud, sdefilterbox5, SdePointSource.Full, exclusion, None, IntPtr(0))
                self.oldprefilterbox = SdeRect3D(closesttocursor.X - r5, closesttocursor.Y - r5, closesttocursor.Z - r5, 2*r5, 2*r5, 2*r5)
                exclusion.Dispose()
            except:
                errorcleanup = True
        
        # cleanup if we produced any error
        if errorcleanup: 
            try: 
                self.visiblesdecloudprefiltered.Dispose()
                self.visiblesdecloudprefiltered = None
            except: self.visiblesdecloudprefiltered = None
            
            try: exclusion.Dispose()
            except: pass


    def bench(self, label, t):
        if self.BENCHMARKING and t is not None:
            self.timestr += '\n' + label + ': ' + str((datetime.now() - t).microseconds)
        return datetime.now() if self.BENCHMARKING else None

    def drawoverlay(self, sender, e):

        try:
            if self.cachewaspaused:
                self.viewfiltersettingschangeevent(None, None)
                return

            if self.activeViewFilter and self.visiblesdecloud:

                #wv = self.currentProject [Project.FixedSerial.WorldView]
                #wv.PauseGraphicsCache(True)
                TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)
                self.overlayBag = OverlayBag(self.ViewOverlay) # create a new/empty overlaybag

                closesttocursor = None

                self.timestr = '\nMicroseconds:' if self.BENCHMARKING else ''
                if isinstance(self.activeForm, clr.GetClrType(Hoops3dView)):
                    t = datetime.now() if self.BENCHMARKING else None
                    closesttocursor = self.activeForm.View.PointCloudPick(e.MousePosition.X, e.MousePosition.Y, False)
                    t = self.bench('Mouselocation', t)

                elif isinstance(self.activeForm, clr.GetClrType(Hoops2dView)):
                    t = datetime.now() if self.BENCHMARKING else None

                    apertureSize = InputSettings.PickAperture
                    pm = self.activeForm.View.ViewCache.SdeViewCache.PickingManager
                    pa = PickAperature(self.activeForm.View, SystemDrawingPoint(e.MousePosition.X - apertureSize, e.MousePosition.Y - apertureSize), SystemDrawingPoint(e.MousePosition.X + apertureSize, e.MousePosition.Y + apertureSize))
                    tt = pm.SelectPointCloudPosition(pa, False, False) # entityLocation, fastMode
                    #Point3D localPoint = ViewCache.SdeViewCache.PickingManager.SelectPointCloudPosition(new GraphicsEngineHoops.PickAperature(this, new Point(pixelX - aperatureSize, pixelY - aperatureSize), new Point(pixelX + aperatureSize, pixelY + aperatureSize)), false, fastMode);
                    if not tt.IsUndefined and tt.Is3D:
                        tt.X += self.activeForm.View.ViewCache.Transform.TranslateX
                        tt.Y += self.activeForm.View.ViewCache.Transform.TranslateY
                        tt.Z += self.activeForm.View.ViewCache.Transform.TranslateZ
                    
                    closesttocursor = Point3D(tt)
                    
                    if self.BENCHMARKING:
                        self.timestr += '\n2D-MouseXY: ' + str(closesttocursor.X) + "," + str(closesttocursor.Y) + "," + str(closesttocursor.Z)
                    t = self.bench('2D-Mouselocation', t)

                if closesttocursor and closesttocursor.Is3D:
                    
                    self.prefilter(closesttocursor)

                    self.overlayBag.AddPolyline(self.moveshapetosnap(closesttocursor, Vector3D(0,0,1)), self.colourPickerAperture.SelectedColor.ToArgb(), 2)

                    self.currGutterp = Point3D()

                    #if self.visiblesdecloudprefiltered:
                    #    bl = Point3D(self.visiblesdecloudprefiltered.BoundingBox.Location.X, self.visiblesdecloudprefiltered.BoundingBox.Location.Y, closesttocursor.Z)
                    #    br = bl + Vector3D(self.visiblesdecloudprefiltered.BoundingBox.SizeX, 0, 0)
                    #    tr = br + Vector3D(0, self.visiblesdecloudprefiltered.BoundingBox.SizeY, 0)
                    #    tl = tr - Vector3D(self.visiblesdecloudprefiltered.BoundingBox.SizeX, 0, 0)
                    #
                    #    self.overlayBag.AddPolyline(Array[Point3D]([bl, br, tr, tl, bl]), self.colourPickerPrefilter.SelectedColor.ToArgb(), 2)

                    # casting with an extension like ToSde() doesn't work in IronPython
                    # either import
                    # from System.Windows.Media.Media3D import Point3D as SdePoint3D, Vector3D as SdeVector3D
                    # and build the objects manually
                    # or import Point3DExtensions
                    # problem with those is that they are spread over multiple assemblies, you need to add the proper reference
                    # Point3DExtensions.ToSde is in Namespace "Trimble.Vce.Geometry" but inside "Trimble.Vce.Scanning"
                    # if you reference "Trimble.Vce.Geometry" only you won't have access to the methods in Trimble.Vce.Scanning->Namespace:Trimble.Vce.Geometry.Point3DExtensions
                    # but once you reference "Trimble.Vce.Scanning" it's enough to import from "Trimble.Vce.Geometry"

                    filteredpoints = []

                    sdefilterorigin = None
                    sdefilter = None
                    r = self.apertureRadius.Distance
                    h = self.apertureHeight.Distance
                    if self.circleSelect.IsChecked:
                        sdefilterorigin = Point3DExtensions.ToSde(closesttocursor)
                        sdefilter = SdeFilterByRadius(sdefilterorigin, r)
                    elif self.squareSelect.IsChecked:
                        # or build the correct type objects manually, as for instance the vectors below; is probably simpler than finding the assembly the extensions are hidden in
                        v_x = Vector3D(2*r ,0, 0)
                        v_x.Azimuth = math.pi/2 - self.pickerBearing.Direction
                        v_y = v_x.Clone()
                        v_y.Rotate90(Side.Left)
                        tmp = closesttocursor - v_x/2 - v_y/2

                        sdefilterorigin = SdePoint3D(tmp.X, tmp.Y, tmp.Z - h/2)
                        sdefilter = SdeFilterByBox(sdefilterorigin, SdeVector3D(v_x.X, v_x.Y, v_x.Z), SdeVector3D(v_y.X, v_y.Y, v_y.Z), SdeVector3D(0, 0, h))

                    
                    minmax = None
                    #sampling = SdeSpatialSamplingParameters()
                    #sampling.ResolutionInMeter = r / 2.5

                    if self.visiblesdecloudprefiltered and self.visiblesdecloudprefiltered.PointCount > 0 and sdefilter:
                    
                        exclusion = clr.StrongBox[SdeCloud]()
                        exclusion2 = clr.StrongBox[SdeCloud]()
                        
                        t = datetime.now() if self.BENCHMARKING else None
                        filtered = SdeCloud(self.visiblesdecloudprefiltered, sdefilter, SdePointSource.Full, exclusion, None, IntPtr(0))
                        #filteredsampled = SdeCloud(filtered,sampling , None, IntPtr(0))
                        t = self.bench('\nFilterByBox from prefiltered', t)
                        if self.BENCHMARKING:
                            self.timestr += '\nfiltering from prefiltered#: ' + str(self.visiblesdecloudprefiltered.PointCount)

                        #time1 = datetime.now()
                        #filtered2 = SdeCloud(self.visiblesdecloud, sdefilter, SdePointSource.Full, exclusion2, None, IntPtr(0))
                        ##filteredsampled = SdeCloud(filtered,sampling , None, IntPtr(0))
                        #timestr += '\n\nFilterByBox from all visible: ' + str((datetime.now() - time1).microseconds)
                        #timestr += '\nfiltering fromd#: ' + str(self.visiblesdecloud.PointCount)
                        #exclusion2.Dispose() # !!!!!! super important
                        #filtered2.Dispose()

                        
                        if filtered.PointCount > 0:
                            t = datetime.now() if self.BENCHMARKING else None
                            minmax = filtered.GetMinMaxPoints(SdeVector3D(0,0,1), None, IntPtr(0))
                            t = self.bench('\nget MinMax', t)

                            if self.drawGutter.IsChecked and self.squareSelect.IsChecked:
                                try:
                                    vy_len = v_y.Length2D
                                    vy_ux = v_y.X / vy_len
                                    vy_uy = v_y.Y / vy_len
                                    profile = []
                                    for p in filtered.GetEnumerator():
                                        dx = p.Coordinates.X - closesttocursor.X
                                        dy = p.Coordinates.Y - closesttocursor.Y
                                        profile.append((dx * vy_ux + dy * vy_uy, p.Coordinates.Z))
                                    profile.sort(key=lambda q: q[0])
                                    N = 10
                                    off0, off1 = profile[0][0], profile[-1][0]
                                    bw = (off1 - off0) / N if off1 > off0 else 1.0
                                    bins = [[] for _ in range(N)]
                                    for off, z in profile:
                                        bins[min(int((off - off0) / bw), N - 1)].append(z)
                                    means = [(off0 + (i + 0.5) * bw, sum(b) / len(b)) for i, b in enumerate(bins) if b]
                                    if len(means) >= 2:
                                        best_off, best_z, max_dz = None, None, 0.0
                                        for i in range(len(means) - 1):
                                            dz = abs(means[i + 1][1] - means[i][1])
                                            if dz > max_dz:
                                                max_dz = dz
                                                best_off = (means[i][0] + means[i + 1][0]) / 2.0
                                                best_z   = (means[i][1] + means[i + 1][1]) / 2.0
                                        if best_off is not None:
                                            self.currGutterp = Point3D(
                                                closesttocursor.X + best_off * vy_ux,
                                                closesttocursor.Y + best_off * vy_uy,
                                                best_z)
                                except:
                                    pass

                            #try:
                            #    pass
                            #    #time1 = datetime.now()
                            #    #for p in filtered.GetEnumerator():
                            #    #    self.overlayBag.AddMarker(Point3D(p.Coordinates.X, p.Coordinates.Y, p.Coordinates.Z), GraphicMarkerTypes.BigDot_IndependentColor, Color.Orange.ToArgb(), "", 0, 0, 1.0)
                            #    #timestr += '\nadd to Overlay: ' + str((datetime.now() - time1).microseconds)
                            #except Exception as e:
                            #    tt = sys.exc_info()
                            #    exc_type, exc_obj, exc_tb = sys.exc_info()
                            #    self.error.Content += '\nan Error occurred - Result probably incomplete\n' + str(exc_type) + '\n' + str(exc_obj) + '\nLine ' + str(exc_tb.tb_lineno)
                            #finally:
                            #    try: exclusion.Dispose() # !!!!!! super important
                            #    except: pass
                            #    try: filtered.Dispose()
                            #    except: pass
                            #    #filteredsampled.Dispose()

                        exclusion.Dispose() # !!!!!! super important
                        filtered.Dispose()


                    if minmax:
                        simpleelcompare = True

                        if simpleelcompare:

                            self.currminp = Point3D(minmax.Item1.Coordinates.X, minmax.Item1.Coordinates.Y, minmax.Item1.Coordinates.Z)
                            self.currmaxp = Point3D(minmax.Item2.Coordinates.X, minmax.Item2.Coordinates.Y, minmax.Item2.Coordinates.Z)

                            t = datetime.now() if self.BENCHMARKING else None
                            self.overlayBag.AddMarker(self.currminp, GraphicMarkerTypes.BigDot_IndependentColor, self.colourPickerMin.SelectedColor.ToArgb(), "", 0, 0, 3.0)
                            self.overlayBag.AddMarker(self.currmaxp, GraphicMarkerTypes.BigDot_IndependentColor, self.colourPickerMax.SelectedColor.ToArgb(), "", 0, 0, 3.0)
                            if self.drawGutter.IsChecked and not self.currGutterp.IsUndefined and self.currGutterp.Is3D:
                                self.overlayBag.AddMarker(self.currGutterp, GraphicMarkerTypes.BigDot_IndependentColor, self.colourPickerGutter.SelectedColor.ToArgb(), "", 0, 0, 3.0)
                            t = self.bench('add to Overlay minmax', t)

                            if self.myLineLow and self.drawMin.IsChecked:
                                line = self.currentProject.Concordance[self.myLineLow]
                                if line and line.ElementCount:
                                    lastPos = line.ElementAt(line.ElementCount - 1).Position
                                    self.overlayBag.AddPolyline(Array[Point3D]([lastPos, self.currminp]), self.colourPickerMin.SelectedColor.ToArgb(), 1)

                            if self.myLineHigh and self.drawMax.IsChecked:
                                line = self.currentProject.Concordance[self.myLineHigh]
                                if line and line.ElementCount:
                                    lastPos = line.ElementAt(line.ElementCount - 1).Position
                                    self.overlayBag.AddPolyline(Array[Point3D]([lastPos, self.currmaxp]), self.colourPickerMax.SelectedColor.ToArgb(), 1)

                            if self.myLineGutter and self.drawGutter.IsChecked and not self.currGutterp.IsUndefined and self.currGutterp.Is3D:
                                line = self.currentProject.Concordance[self.myLineGutter]
                                if line and line.ElementCount:
                                    lastPos = line.ElementAt(line.ElementCount - 1).Position
                                    self.overlayBag.AddPolyline(Array[Point3D]([lastPos, self.currGutterp]), self.colourPickerGutter.SelectedColor.ToArgb(), 1)

                            ExploreObjectControlHelper.DrawCursorText(e.MessagingView, e.MousePosition, 10, "Max: " + str("{:.{}f}".format(self.currmaxp.Z, 4)) + "\nMin: " + str("{:.{}f}".format(self.currminp.Z, 4)) + self.timestr)

                        else: # compute best fit plane first


                            rwcloudpoints = []
                            point3dcloudpoints = []
                            #for p in filteredpoints:
                            #    self.overlayBag.AddMarker(p, GraphicMarkerTypes.BigDot_IndependentColor, Color.Blue.ToArgb(), "", 0, 0, 1.0)
                            #    rwcloudpoints.Add(RwPoint3D(p.X, p.Y, p.Z))
                            #    #point3dcloudpoints.Add(p)
                            rwcloudpoints = [RwPoint3D(p.Coordinates.X, p.Coordinates.Y, p.Coordinates.Z) for p in filteredpoints]

                            rwplane = None
                            rwplane = RwPlane3D.FitPlaneTo3DPoints(rwcloudpoints)
                            if rwplane:
                                centerp = Point3D(rwplane.Point.X, rwplane.Point.Y, rwplane.Point.Z)
                                v = Vector3D(rwplane.NormalVector.X, rwplane.NormalVector.Y, rwplane.NormalVector.Z)

                                self.overlayBag.AddPolyline(self.moveshapetosnap(closesttocursor, v), Color.Yellow.ToArgb(), 1)

                                p = Plane3D(v, centerp)

                                # Vector3D.Horizon is positive above the horizon and negative below
                                pv = Vector3D(filteredpoints[0], Plane3D.IntersectWithRayPerpendicular(p, filteredpoints[0]))
                                mini = 0
                                maxi = 0
                                if pv.Horizon < 0:
                                    mind = pv.Length
                                    maxd = pv.Length
                                else:
                                    mind = -pv.Length
                                    maxd = -pv.Length
                                for i in range(filteredpoints.Count):
                                    pv = Vector3D(filteredpoints[i], Plane3D.IntersectWithRayPerpendicular(p, filteredpoints[i]))
                                    if pv.Horizon < 0 and pv.Length > maxd:
                                        maxi = i
                                        maxd = pv.Length
                                    elif pv.Horizon > 0 and -pv.Length < mind:
                                        mini = i
                                        mind = -pv.Length

                                self.overlayBag.AddMarker(filteredpoints[mini], GraphicMarkerTypes.BigDot_IndependentColor, self.colourPickerMin.SelectedColor.ToArgb(), "", 0, 0, 3.0)
                                self.overlayBag.AddMarker(filteredpoints[maxi], GraphicMarkerTypes.BigDot_IndependentColor, self.colourPickerMax.SelectedColor.ToArgb(), "", 0, 0, 3.0)
                                    
                                ExploreObjectControlHelper.DrawCursorText(e.MessagingView, e.MousePosition, 15, "Max: " + str("{:.{}f}".format(maxd, 4) + "\nMin: " + str("{:.{}f}".format(mind, 4))))

                    else:
                        ExploreObjectControlHelper.DrawCursorText(e.MessagingView, e.MousePosition, 10, self.timestr)

                # if it only needs to be visible in all Planview then remove the Hoops3DViewGUID
                array = Array[Guid]([DisplayWindow.Hoops3DViewGUID, DisplayWindow.HoopsPlanViewGUID])
                TrimbleOffice.TheOffice.MainWindow.AppViewManager.AddOverlayGeometry(array, self.overlayBag)

            #wv.PauseGraphicsCache(False)
            
        except Exception as _err:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            self.error.Content += '\nan Error occurred - Result probably incomplete\n' + str(exc_type) + '\n' + str(exc_obj) + '\nLine ' + str(exc_tb.tb_lineno)

        return

    def SetDefaultOptions(self):

        self.apertureRadius.ValueChanged -= self.radiuschange
        self.circleSelect.Checked -= self.radiuschange
        self.squareSelect.Checked -= self.radiuschange

        SCROptions.LoadProjectOptions(self, "SCR_PointcloudExtractor", _OPTIONS, self.currentProject)

        self.apertureRadius.ValueChanged += self.radiuschange
        self.circleSelect.Checked += self.radiuschange
        self.squareSelect.Checked += self.radiuschange

    def SaveOptions(self, _sender=None, _e=None):

        SCROptions.SaveProjectOptions(self, "SCR_PointcloudExtractor", _OPTIONS, self.currentProject)


    def Dispose(self, _thisCmd, _disposing):

        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)

    def CoordChanged(self, _ctrl, e):
        # set keyboard focus if change was due to mouse pick
        if e.Cause == InputMethod.Mouse or e.Cause == InputMethod.Snap:
            self.OkClicked(None, None)

    def CancelClicked(self, thisCmd, args):
        try: self.visiblesdecloud.Dispose()
        except: pass
        try: self.visiblesdecloudprefiltered.Dispose()
        except: pass

        self.cleanupeventtriggers()
        self.SaveOptions(None, None)
        thisCmd.CloseUICommand()

    def cleanupeventtriggers(self):

        TrimbleOffice.TheOffice.MainWindow.AppViewManager.ViewActivated -= self.activeviewchanged
        UIEvents.UIViewFilterChanged -= self.activeviewchanged
        if self.activeViewFilter:
            self.activeViewFilter.FilterChanged -= self.viewfiltersettingschangeevent

        self.coordCtl.ShowGdiCursor -= self.drawoverlay
        self.apertureRadius.ValueChanged -= self.radiuschange
        self.circleSelect.Checked -= self.radiuschange
        self.squareSelect.Checked -= self.radiuschange
        #self.activeForm.View.MouseWheel -= self.mousewheel

    def bearingChanged(self, _ctrl, _e):
        Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Input, Action(lambda: Keyboard.Focus(self.coordCtl)))

    def btnResetBearingClicked(self, _sender, _e):
        self.pickerBearing.Direction = 0

    def BtnNewClicked(self, sender, e):
        if sender == self.btnNewMin:
            self.myLineLow = 0
        elif sender == self.btnNewMax:
            self.myLineHigh = 0
        elif sender == self.btnNewGutter:
            self.myLineGutter = 0
        else:
            self.myLineLow = 0
            self.myLineHigh = 0
            self.myLineGutter = 0
        Keyboard.Focus(self.coordCtl)


    def appendPointToLine(self, lineSerial, coord, layerpicker, colourpicker):

        line = None

        if not lineSerial or lineSerial == 0:
            wv = self.currentProject[Project.FixedSerial.WorldView]
            line = wv.Add(clr.GetClrType(Linestring))

            newSeg = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
            newSeg.Position = coord
            line.AppendElement(newSeg)
            line.Layer = layerpicker.SelectedSerialNumber # we only need to set it once, upon creation

        else:
            line = self.currentProject.Concordance[lineSerial]
            if line:
                line.Color = colourpicker.SelectedColor
                elemCount = line.ElementCount
                if elemCount:
                    lastElem = line.ElementAt(elemCount - 1)
                    if Point3D.IsDuplicate2D(clr.Reference[Point3D](coord), clr.Reference[Point3D](lastElem.Position)):
                        lastElem.Position = coord
                        line.ReplaceElementAt(lastElem, elemCount - 1)
                    else:
                        if Control.ModifierKeys & Keys.Control:
                            newSeg = ElementFactory.Create(clr.GetClrType(ITangentArcSegment), clr.GetClrType(IXYZLocation))
                            newSeg.TangentType = ArcTangentType.TangentTangent
                        else:
                            newSeg = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                        newSeg.Position = coord
                        line.AppendElement(newSeg)

        return line.SerialNumber


    def OkClicked(self, thisCmd, _e):

        ctrl = self.coordCtl
        coordLow  = self.currminp
        coordHigh = self.currmaxp

        if not self.drawMin.IsChecked and not self.drawMax.IsChecked and not self.drawGutter.IsChecked:
            return False

        if not ctrl.ResultCoordinateSystem:
            self.coordCtl.StatusMessage = "No coordinate defined"
            return False

        self.coordCtl.StatusMessage = ""

        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(Client.CommandGranularity.Command, self.Caption)

        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                if self.drawMin.IsChecked and not coordLow.IsUndefined and coordLow.Is3D:
                    self.myLineLow  = self.appendPointToLine(self.myLineLow,  coordLow,  self.layerPickerMin, self.colourPickerMin)

                if self.drawMax.IsChecked and not coordHigh.IsUndefined and coordHigh.Is3D:
                    self.myLineHigh = self.appendPointToLine(self.myLineHigh, coordHigh, self.layerPickerMax, self.colourPickerMax)

                if self.drawGutter.IsChecked and not self.currGutterp.IsUndefined and self.currGutterp.Is3D:
                    self.myLineGutter = self.appendPointToLine(self.myLineGutter, self.currGutterp, self.layerPickerGutter, self.colourPickerGutter)

                failGuard.Commit()

        except Exception as _err:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            self.error.Content += '\nan Error occurred - Result probably incomplete\n' + str(exc_type) + '\n' + str(exc_obj) + '\nLine ' + str(exc_tb.tb_lineno)

        finally:
            self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
            UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())

        Keyboard.Focus(self.coordCtl)
        return
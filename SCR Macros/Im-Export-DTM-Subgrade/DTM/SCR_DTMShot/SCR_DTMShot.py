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
    "shootifc":       True,
    "shootdtm":       False,
    "shootplane":     False,
    "layerpicker":    8,
    "surfacepicker":  0,
    "usepoints":      True,
    "extendline":     False,
    "breakline":      False,
    "drawlinepoint":  False,
    "singleentity":   True,
    "polygonselect":  False,
}

def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_DTMShot"
    cmdData.CommandName = "SCR_DTMShot"
    cmdData.Caption = "_SCR_DTMShot"
    cmdData.UIForm = "SCR_DTMShot"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "DTM"
        cmdData.ShortCaption = "DTM Shot"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.12
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "intersect a DTM with a 2-Point Vector"
        cmdData.ToolTipTextFormatted = "intersect a DTM with a 2-Point Vector"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class _PolyPixelFilter(IMessageFilter):
    # Intercepts Windows messages at the application pump level (before DispatchMessage)
    # to capture the pixel coordinates of every polygon/rectangle vertex as the user draws
    # in a HOOPS view (2D or 3D).  We cannot use WinForms MouseDown events because HOOPS
    # consumes them at WndProc level before they bubble up.
    #
    # HOOPS selection protocol:
    #   - Rectangle: single drag — DOWN at corner 1, UP at corner 2 (only recorded if the
    #     mouse moved >5 px, i.e. it was a genuine drag and not a jittery click).
    #   - Polygon: each vertex is a single left-click (WM_LBUTTONDOWN only); a double-click
    #     closes the polygon (WM_LBUTTONDBLCLK, not captured here).
    #
    # HWND filtering: the inner Hoops view WinForms control hosts one or more native child
    # windows for HOOPS rendering.  Clicks land on those child HWNDs, not on the WinForms
    # control itself.  Control.FromChildHandle() walks the native parent chain to find the
    # owning WinForms control, so we compare that against view_hwnd to accept exactly the
    # clicks inside the active view and ignore everything else (toolbar buttons, macro panel,
    # other views).  When the user switches to a different Hoops view, _get_active_hwnd()
    # is called to adopt the new view's HWND on the first click there.
    #
    # Accumulation: pts grows across all selection events; OkClicked reads and resets it.
    # Exception: OnPickerValueChanged clears pts when the selection drops to zero and no
    # draw is in progress (_draw_active=False), discarding stray deselect clicks.

    WM_LBUTTONDOWN = 0x0201
    WM_LBUTTONUP   = 0x0202

    def __init__(self):
        self.pts = []
        self._last_down = None
        self.view_hwnd = None  # Handle of inner Hoops view; None = accept all (not yet known)
        self._draw_active = False  # True between LBUTTONDOWN and matching LBUTTONUP
        self._get_active_hwnd = None  # () -> int|None; set by macro to query current view

    @staticmethod
    def _parse(lp):
        # lParam of WM_LBUTTONDOWN/UP packs client-relative (x, y) as two signed 16-bit words.
        x = lp & 0xFFFF
        y = (lp >> 16) & 0xFFFF
        if x > 32767: x -= 65536
        if y > 32767: y -= 65536
        return x, y

    def PreFilterMessage(self, m):
        try:
            if m.Msg not in (self.WM_LBUTTONDOWN, self.WM_LBUTTONUP):
                return False
            if self.view_hwnd is not None:
                # Accept only clicks whose native window belongs to the active Hoops view.
                owner = System.Windows.Forms.Control.FromChildHandle(m.HWnd)
                if owner is None:
                    return False
                click_hwnd = int(owner.Handle)
                if click_hwnd != int(self.view_hwnd):
                    # HWND mismatch — the user may have switched to a different Hoops view.
                    # Check whether this click landed on the currently active view; if so,
                    # adopt its HWND so subsequent events are filtered correctly.
                    active_hwnd = None
                    if self._get_active_hwnd is not None:
                        try:
                            active_hwnd = self._get_active_hwnd()
                        except Exception:
                            pass
                    if active_hwnd is not None and click_hwnd == active_hwnd:
                        self.view_hwnd = active_hwnd
                    else:
                        return False
            if m.Msg == self.WM_LBUTTONDOWN:
                x, y = self._parse(m.LParam.ToInt32())
                self._last_down = (x, y)
                self._draw_active = True
                self.pts.append((x, y))
            elif m.Msg == self.WM_LBUTTONUP:
                x, y = self._parse(m.LParam.ToInt32())
                self._draw_active = False
                if self._last_down is not None:
                    dx = x - self._last_down[0]
                    dy = y - self._last_down[1]
                    if dx * dx + dy * dy > 25:  # >5 px distance = genuine drag, not a click
                        self.pts.append((x, y))
        except Exception:
            pass
        return False


class SCR_DTMShot(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_DTMShot.xaml") as s:
            wpf.LoadComponent (self, s)
        self.currentProject = currentProject
        self.macroFileFolder = macroFileFolder

        self.ViewOverlay = Guid.NewGuid()
        self.overlayBag = OverlayBag(self.ViewOverlay)
        self.view3d = None
        self._inner_view = None
        self._poly_filter = _PolyPixelFilter()
        self._poly_filter._get_active_hwnd = self._get_active_inner_hwnd
        self._last_2d_intersections = []
        System.Windows.Forms.Application.AddMessageFilter(self._poly_filter)

    def HelpClicked(self, cmd, e):
        webbrowser.open("C:\ProgramData\Trimble\MacroCommands3\SCR Macros\MacroHelp\MacroHelp.htm#" + type(self).__name__)


    def OnLoad(self, cmd, buttons, event):
        self.okBtn = buttons[0]
        buttons[2].Content = "Help"
        buttons[2].Visibility = Visibility.Visible
        buttons[2].Click += self.HelpClicked
        self.Caption = cmd.Command.Caption
        
        wv = self.currentProject [Project.FixedSerial.WorldView]

        self.ignoreGotFocus = False
        self.selectionControls = [self.ifcs, self.linepickermulti]

        self.ifcs.IsEntityValidCallback = self.IsValidIFC

        types = Array[Type](SurfaceTypeLists.AllWithCutFillMap)+Array[Type]([clr.GetClrType(ProjectedSurface)])
        #types.extend (Array[Type]([clr.GetClrType(ProjectedSurface)]))
        self.surfacepicker.FilterByEntityTypes = types
        self.surfacepicker.AllowNone = False

        self.coordpick1.ShowElevationIf3D = True
        self.coordpick2.ShowElevationIf3D = True
        self.coordpick2.ValueChanged += self.CoordPick2Changed
        self.coordpick2.AutoTab = False

        self.coordCtl1.ValueChanged += self.Coord1Changed
        self.coordCtl2.ValueChanged += self.Coord2Changed
        self.coordCtl3.ValueChanged += self.Coord3Changed


        self.lType = clr.GetClrType(Linestring) # don't allow 2D CadLines
        self.linepicker1.ValueChanged += self.lineChanged
        self.linepickermulti.ValueChanged += self.OnPickerValueChanged
        self.linepicker1.AutoTab = False
        self.linepicker1.IsEntityValidCallback = self.IsValidLine

        SCRExpanders.wire_pairs([
            (self.expander_usepoints, self.usepoints),
            (self.expander_extendline, self.extendline),
        ])
		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            active = TrimbleOffice.TheOffice.MainWindow.AppViewManager.ActiveView
            if isinstance(active, (clr.GetClrType(Hoops3dView), clr.GetClrType(Hoops2dView))):
                inner = active.View
                if inner is not None:
                    self._inner_view = inner
                    self._poly_filter.view_hwnd = inner.Handle
        except:
            pass

        try:
            self.SetDefaultOptions()
        except:
            pass

    def _get_active_inner_hwnd(self):
        try:
            active = TrimbleOffice.TheOffice.MainWindow.AppViewManager.ActiveView
            if isinstance(active, (clr.GetClrType(Hoops3dView), clr.GetClrType(Hoops2dView))):
                inner = active.View
                if inner is not None:
                    return int(inner.Handle)
        except Exception:
            pass
        return None

    def CoordPick2Changed(self, ctrl, e):
        if e.Cause == InputMethod.Mouse:     
            self.OkClicked(None, None)

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
        
        if not self.coordCtl3.Coordinate.Is3D :
            self.coordCtl3.StatusMessage = "No valid coordinate defined, must be 3D"
        else:
            self.coordCtl3.StatusMessage = ""

        self.drawoverlay()
        
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
        if self._poly_filter:
            System.Windows.Forms.Application.RemoveMessageFilter(self._poly_filter)
            self._poly_filter = None
        self._inner_view = None

    def OnPickerValueChanged(self, sender, e):
        try:
            active = TrimbleOffice.TheOffice.MainWindow.AppViewManager.ActiveView
            if isinstance(active, (clr.GetClrType(Hoops3dView), clr.GetClrType(Hoops2dView))):
                self.view3d = active
                self._inner_view = active.View
                if self._inner_view is not None:
                    self._poly_filter.view_hwnd = self._inner_view.Handle
            else:
                self.view3d = None
                self._inner_view = None
            # Cache MostRecentPickIntersections while non-empty as a fallback for 2D
            # in case TransformWorldToScreen is not available on the 2D inner view.
            try:
                ix = list(sender.MostRecentPickIntersections)
                if ix:
                    self._last_2d_intersections = ix
            except Exception:
                pass
            # If the picker just cleared via a deselect click (not mid-draw), discard
            # any pts captured so far — they are stray clicks that would pollute the
            # next selection's polygon/rectangle shape.
            # Guard: if _draw_active is True the user has already pressed the button
            # to start the new selection (TBC clears the old selection internally
            # before the new one lands), so we must NOT clear pts in that case.
            try:
                if (sender.SelectedSerials.Count == 0
                        and self._poly_filter is not None
                        and not self._poly_filter._draw_active):
                    self._poly_filter.pts = []
            except Exception:
                pass
        except Exception:
            pass

    def IsValidLine(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        return False

    def IsValidIFC(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, BIMEntity):
            return True
        if isinstance(o, Shell3D):
            return True
        return False

    def lineChanged(self, ctrl, e):
        self.OkClicked(None, None)

    def SetDefaultOptions(self):
        SCROptions.LoadMacroOptions(self, "SCR_DTMShot", _OPTIONS, self.currentProject)

    def SaveOptions(self):
        SCROptions.SaveMacroOptions(self, "SCR_DTMShot", _OPTIONS)

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def Selection_PreviewGotFocus(self, sender, e):
        self.ignoreGotFocus = True
        for ctrl in self.selectionControls:
            value = (sender == ctrl)
            ctrl.ProcessGlobalSelectionChanges = value
            ctrl.UpdateTextOnSelectionChange = value

    def Selection_ValueChanged(self, sender, e):
        if self.ignoreGotFocus:
            self.ignoreGotFocus = False
            return

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''

        self.success.Content = ''

        if not self.usepoints.IsChecked and not self.extendline.IsChecked:
            self.error.Content = 'select an operation first'
            return

        wv = self.currentProject [Project.FixedSerial.WorldView]
        wv.PauseGraphicsCache(True)

        inputok=True

        if inputok:
            UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
            self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
            try:
                with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                    
                    dtmintersection = None
                    self.vertexlist = self.createvertexlist()

                    if self.usepoints.IsChecked:
                        p1 = self.coordpick1.Coordinate
                        p2 = self.coordpick2.Coordinate
                        shot = Vector3D(p1, p2)
                        
                        #if isinstance(surface, ProjectedSurface):

                        dtmintersection = self.intersectdtm(self.vertexlist, p1, p2, False)

                        #else: # standard surface is easy
                        #    tiepoint = clr.StrongBox[Point3D]()     # we create us the variable the ComputeTie wants for the output
                        #    if surface.ComputeTie(p1, shot, math.pi/2 - (shot.Horizon), 10000, tiepoint): # we compute the surface intersection
                        #        dtmintersection = tiepoint.Value

                        if dtmintersection:
                            cadPoint = wv.Add(clr.GetClrType(CadPoint))
                            cadPoint.Layer = self.layerpicker.SelectedSerialNumber
                            cadPoint.Point0 = dtmintersection

                    else: # extend line

                        if self.singleentity.IsChecked:

                            l1 = self.linepicker1.Entity

                            if l1:
                                self.extendpolyseg(l1, [self.linepicker1.PickPointProjected])

                        elif self.polygonselect.IsChecked:
                            activeForm = TrimbleOffice.TheOffice.MainWindow.AppViewManager.ActiveView

                            if self.linepickermulti.SelectedSerials.Count > 0:
                                # Pixel-space end-detection: _poly_filter.pts holds all
                                # WM_LBUTTONDOWN/UP events captured while the user drew the
                                # selection shape on the active view.  We snapshot and reset here.
                                # Rectangle selection produces exactly 2 pts (diagonal corners);
                                # synthesise the 4-corner rect so the same ray-cast works for both.
                                # The linestring endpoint that falls INSIDE the shape is the end the
                                # user targeted; we extend from that end toward the DTM surface.
                                inner = self._inner_view
                                pix_pts = list(self._poly_filter.pts)
                                self._poly_filter.pts = []

                                if len(pix_pts) == 2:  # rectangle: synthesise 4 corners
                                    x1, y1 = pix_pts[0]
                                    x2, y2 = pix_pts[1]
                                    pix_pts = [(x1,y1),(x2,y1),(x2,y2),(x1,y2)]

                                handled = False

                                if inner and len(pix_pts) >= 3:
                                    if isinstance(activeForm, clr.GetClrType(Hoops3dView)):
                                        # 3D view: HOOPS provides a direct world → pixel projection.
                                        try:
                                            for sn in self.linepickermulti.SelectedSerials:
                                                ent = self.currentProject.Concordance[sn]
                                                if not isinstance(ent, self.lType):
                                                    continue
                                                try:
                                                    ps = ent.ComputePolySeg()
                                                    if ps is None:
                                                        continue
                                                    ps = ps.ToWorld()
                                                    start_w = ps.FirstSegment.BeginPoint
                                                    end_w   = ps.LastSegment.EndPoint
                                                except:
                                                    continue
                                                s_px = inner.TransformWorldToScreen(start_w)
                                                e_px = inner.TransformWorldToScreen(end_w)
                                                s_in = self._px_in_polygon(s_px.X, s_px.Y, pix_pts)
                                                e_in = self._px_in_polygon(e_px.X, e_px.Y, pix_pts)
                                                if s_in and not e_in:
                                                    self.extendpolyseg(ent, [], from_start=True)
                                                elif not s_in and e_in:
                                                    self.extendpolyseg(ent, [], from_start=False)
                                            handled = True
                                        except Exception:
                                            pass

                                    elif isinstance(activeForm, clr.GetClrType(Hoops2dView)):
                                        # 2D planview primary: MostRecentPickIntersections gives
                                        # the world coordinates where the selection fence crosses
                                        # each entity.  Lines that cross the boundary appear here;
                                        # lines fully enclosed by the shape do NOT appear in MRPI.
                                        intersections = list(self.linepickermulti.MostRecentPickIntersections) \
                                                        or self._last_2d_intersections
                                        handled_serials = set()
                                        for ix in intersections:
                                            ent = self.currentProject.Concordance[ix.Item1]
                                            if not isinstance(ent, self.lType):
                                                continue
                                            ipoints = [i.Point for i in ix.Item2]
                                            if ipoints:
                                                self.extendpolyseg(ent, ipoints)
                                                handled_serials.add(ix.Item1)
                                        # Any selected entity not in MRPI is fully enclosed by the
                                        # polygon.  Use GetViewExtents to derive the exact world
                                        # bounds of the visible area, convert the pixel polygon to
                                        # a world polygon, then test which endpoint falls inside.
                                        remaining = [sn for sn in self.linepickermulti.SelectedSerials
                                                     if sn not in handled_serials]
                                        if remaining and len(pix_pts) >= 3:
                                            try:
                                                _min_ref = clr.Reference[Point3D]()
                                                _max_ref = clr.Reference[Point3D]()
                                                activeForm.GetViewExtents(_min_ref, _max_ref)
                                                _mn = _min_ref.Value
                                                _mx = _max_ref.Value
                                                _vw = float(inner.Width)
                                                _vh = float(inner.Height)
                                                _dx = _mx.X - _mn.X
                                                _dy = _mx.Y - _mn.Y
                                                world_poly = [
                                                    (_mn.X + (px / _vw) * _dx,
                                                     _mx.Y - (py / _vh) * _dy)
                                                    for px, py in pix_pts
                                                ]
                                                for sn in remaining:
                                                    ent = self.currentProject.Concordance[sn]
                                                    if not isinstance(ent, self.lType):
                                                        continue
                                                    try:
                                                        ps = ent.ComputePolySeg()
                                                        if ps is None:
                                                            continue
                                                        ps = ps.ToWorld()
                                                        start_w = ps.FirstSegment.BeginPoint
                                                        end_w   = ps.LastSegment.EndPoint
                                                    except:
                                                        continue
                                                    s_in = self._px_in_polygon(
                                                        start_w.X, start_w.Y, world_poly)
                                                    e_in = self._px_in_polygon(
                                                        end_w.X, end_w.Y, world_poly)
                                                    if s_in and not e_in:
                                                        self.extendpolyseg(ent, [], from_start=True)
                                                    elif not s_in and e_in:
                                                        self.extendpolyseg(ent, [], from_start=False)
                                            except Exception:
                                                pass
                                        handled = True
                                self._last_2d_intersections = []

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
        self.SaveOptions()
        if self.polygonselect.IsChecked:
            GlobalSelection.Clear()

        if self.usepoints.IsChecked:
            Keyboard.Focus(self.coordpick1)
        elif self.polygonselect.IsChecked:
            Keyboard.Focus(self.linepickermulti)
        else:
            Keyboard.Focus(self.linepicker1)

        wv.PauseGraphicsCache(False)


    def _px_in_polygon(self, x, y, pts):
        """Ray-casting point-in-polygon test in pixel space. pts is list of (px, py) tuples."""
        n = len(pts)
        inside = False
        j = n - 1
        for i in range(n):
            xi, yi = pts[i]
            xj, yj = pts[j]
            if ((yi > y) != (yj > y)) and (x < (xj - xi) * (y - yi) / float(yj - yi) + xi):
                inside = not inside
            j = i
        return inside

    def extendpolyseg(self, l1, pickpoints, from_start=None):

        polyseg = l1.ComputePolySeg()
        polyseg = polyseg.ToWorld()
        polyseg_v = l1.ComputeVerticalPolySeg()
        polyseg = polyseg.Linearize(0.0001, 0.0001, 1000, polyseg_v, False)

        if self.extendline.IsChecked and not self.breakline.IsChecked:

            if from_start is not None:
                pickedstartofline = from_start
            else:
                total = polyseg.ComputeStationing()
                min_station = float('inf')
                max_station = float('-inf')
                for pickpoint in pickpoints:
                    found, pout, pstation = polyseg.FindPointFromPoint(pickpoint)
                    if found:
                        if pstation < min_station: min_station = pstation
                        if pstation > max_station: max_station = pstation

                if min_station == float('inf'):
                    return

                pickedstartofline = min_station < (total - max_station)

            if pickedstartofline:
                p1 = polyseg.FirstSegment.EndPoint
                p2 = polyseg.FirstSegment.BeginPoint
            else:
                p1 = polyseg.LastSegment.BeginPoint
                p2 = polyseg.LastSegment.EndPoint
    
            #if isinstance(surface, ProjectedSurface):
    
            dtmintersection = self.intersectdtm(self.vertexlist, p1, p2, False)
    
            #else: # standard surface is easy
            #    tiepoint = clr.StrongBox[Point3D]()     # we create us the variable the ComputeTie wants for the output
            #    # slope in Computetie is zenith angle with upwards=0
            #    # Vector3D.Horizon is positive above the horizon and negative below
            #    if surface.ComputeTie(p1, shot, math.pi/2 - (shot.Horizon), 10000, tiepoint): # we compute the surface intersection
            #        dtmintersection = tiepoint.Value
    
            if dtmintersection:
                try:
                    if pickedstartofline:
                        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                        e.Position = dtmintersection
                        # if the old segment is longer than the distance to the new intersection
                        if Vector3D(p1, p2).Length > Vector3D(p1, dtmintersection).Length:
                            # replace as new start
                            l1.ReplaceElementAt(e, 0)
                        else:
                            # add as new start
                            l1.InsertElementAt(e, 0)
                    else:
                        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                        e.Position = dtmintersection
                        # if the old segment is longer than the distance to the new intersection
                        if Vector3D(p1, p2).Length > Vector3D(p1, dtmintersection).Length:
                            # need to replace last element
                            l1.ReplaceElementAt(e, l1.ElementCount - 1)
                        else:
                            # add extra element
                            l1.AppendElement(e)
                except Exception:
                    return

                if self.drawlinepoint.IsChecked:
                    cadPoint = wv.Add(clr.GetClrType(CadPoint))
                    cadPoint.Layer = self.layerpicker.SelectedSerialNumber
                    cadPoint.Point0 = dtmintersection
    
        if self.extendline.IsChecked and self.breakline.IsChecked:
            s = polyseg.FirstSegment
            while s is not None:
                dtmintersection = None
                if s.Visible:
                    dtmintersection = self.intersectdtm(self.vertexlist, s.BeginPoint, s.EndPoint, True)
                
                       
                if dtmintersection:
                    pch = polyseg.FindPointFromPoint(dtmintersection)
                    if not pch[2] == 0:
                        if self.drawlinepoint.IsChecked:
                            cadPoint = wv.Add(clr.GetClrType(CadPoint))
                            cadPoint.Layer = self.layerpicker.SelectedSerialNumber
                            cadPoint.Point0 = dtmintersection
                        
                        l1 = l1.BreakAtStation(pch[2]) # result is a new linestring after the break point                            
                        # compute a new polyseg for the rest of the original linestring
                        # and repeat the search for segments which intersect the surface
                        polyseg = l1.ComputePolySeg()
                        polyseg = polyseg.ToWorld()
                        polyseg_v = l1.ComputeVerticalPolySeg()
                        polyseg = polyseg.Linearize(0.0001, 0.0001, 1, polyseg_v, False)
                        s = polyseg.FirstSegment
                    else:
                        s = polyseg.Next(s)
                s = polyseg.Next(s)

    
    def PointInsideTriangle3D(self, vertex1, p2, p3, p4):
        # see also https://blackpawn.com/texts/pointinpoly/
    
        a = Vector3D(p2.X-vertex1.X, p2.Y-vertex1.Y, p2.Z-vertex1.Z)     # v1
        b = Vector3D(p3.X-vertex1.X, p3.Y-vertex1.Y, p3.Z-vertex1.Z)     # v0
        w = Vector3D(p4.X-vertex1.X, p4.Y-vertex1.Y, p4.Z-vertex1.Z)     # v2
    
        aa = Vector3D.DotProduct(a,a)[0]  # dot11
        ab = Vector3D.DotProduct(a,b)[0]  # dot10
        bb = Vector3D.DotProduct(b,b)[0]  # dot00
        wa = Vector3D.DotProduct(w,a)[0]  # dot21
        wb = Vector3D.DotProduct(w,b)[0]  # dot20
        
        # dot10 * dot10 - dot11 * dot00
        d = ab * ab - aa * bb
        if d == 0:
            # in case the three triangle vertices are in one line we can't compute it and ignore that "triangle"
            inside = False
        else:
            inside = True
                    # v = dot10 * dot20 - dot00 * dot21
            s = round((ab * wb - bb * wa) / d, 6) # rounding that value a bit down gives us better results when very close to the triangle side, what we are
            if s < 0 or s > 1:                    # otherwise we might miss a value
                inside = False
                    # u = dot10 * dot21 - dot11 * dot20
            t = round((ab * wa - aa * wb) / d, 6)
            if t < 0 or (s + t) > 1:
               inside = False
    
        return inside

    def intersectdtm(self, vertexlist, p1, p2, isectonsegmentonly):
        # isectonsegmentonly - is a bool value; for breaking lines we only want the locations when the intersection is between t=0 and t=1
        
        shot = Vector3D(p1, p2)

        # setup the line
        seg1 = SegmentLine(p1, p2)

        # prepare variables
        out_t = clr.StrongBox[float]()
        outPointOnCL = clr.StrongBox[Point3D]()
        testside = clr.StrongBox[Side]()

        shortest_intersection = None

        for i in range(0, vertexlist.Count, 3):
        
            p = Plane3D(vertexlist[i], vertexlist[i+1], vertexlist[i+2])[0] # the plane is returned as first element
        
            if not p.IsValid:
                continue
        
            pnew = Plane3D.IntersectWithRay(p, p1, shot)
        
            if self.shootplane.IsChecked: # in case of plane mode we do want the single solution we'll have

                shortest_intersection = pnew

            else:
                # we only want results were the intersecting point is within the tested triangle
                # otherwise we get hundreds of false results which don't lie on the DTM
                
                #if Triangle2D.IsPointInside(v1,v2,v3,pnew)[0] == True:
                ### if PointInsideTriangle3D(v1,v2,v3,pnew):
                if pnew != Point3D.Undefined and self.PointInsideTriangle3D(vertexlist[i], vertexlist[i+1], vertexlist[i+2], pnew):
                    if not Point3D.IsDuplicate(pnew, p2, 0.000001)[0]:
                        # project the point - only if it's in front of us we want it
                        if seg1.ProjectPoint(pnew, out_t, outPointOnCL, testside):
                            if not isectonsegmentonly and out_t.Value > 0.0:
                                if not shortest_intersection:
                                    shortest_intersection = pnew
                                else:
                                    if Vector3D(p1, pnew).Length < Vector3D(p1, shortest_intersection).Length:
                                        shortest_intersection = pnew

                            if isectonsegmentonly and out_t.Value > 0.0 and out_t.Value <= 1.0:
                                if not shortest_intersection:
                                    shortest_intersection = pnew
                                else:
                                    if Vector3D(p1, pnew).Length < Vector3D(p1, shortest_intersection).Length:
                                        shortest_intersection = pnew
        

                            
        #tt1 = Vector3D(p1, shortest_intersection).Length
        #tt2 = Vector3D(p2, shortest_intersection).Length
                                    
        if not isectonsegmentonly:
            if shortest_intersection and Vector3D(p1, shortest_intersection).Length > 0 and Vector3D(p2, shortest_intersection).Length > 0:
                return shortest_intersection
            else:
                return None
        else:
            if shortest_intersection and Vector3D(p1, shortest_intersection).Length > 0 and Vector3D(p1, shortest_intersection).Length <= Vector3D(p1, p2).Length:
                return shortest_intersection
            else:
                return None
        
    def createvertexlist(self):
        # create a list of triangle vertices
        vertexlist = []
        if self.shootdtm.IsChecked:

            surface = self.currentProject.Concordance.Lookup(self.surfacepicker.SelectedSerial)
            nTri = surface.NumberOfTriangles

            if isinstance(surface,ProjectedSurface):
                projected=True
            else:
                projected=False
                
            for i in range(nTri):
                if surface.GetTriangleMaterial(i) == surface.NullMaterialIndex(): continue
                if projected==True:
                    vertexlist.Add(surface.TransformPointToWorldDelegate(surface.GetVertexPoint(surface.GetTriangleIVertex(i,0))))
                    vertexlist.Add(surface.TransformPointToWorldDelegate(surface.GetVertexPoint(surface.GetTriangleIVertex(i,1))))
                    vertexlist.Add(surface.TransformPointToWorldDelegate(surface.GetVertexPoint(surface.GetTriangleIVertex(i,2))))
                else:
                    vertexlist.Add(surface.GetVertexPoint(surface.GetTriangleIVertex(i,0)))
                    vertexlist.Add(surface.GetVertexPoint(surface.GetTriangleIVertex(i,1)))
                    vertexlist.Add(surface.GetVertexPoint(surface.GetTriangleIVertex(i,2)))

        elif self.shootplane.IsChecked:

            if self.coordCtl1.Coordinate.Is3D and self.coordCtl2.Coordinate.Is3D and self.coordCtl3.Coordinate.Is3D:

                vertexlist.Add(self.coordCtl1.Coordinate)
                vertexlist.Add(self.coordCtl2.Coordinate)
                vertexlist.Add(self.coordCtl3.Coordinate)

        elif self.shootifc.IsChecked: # if we use IFCs

            for o in self.ifcs.SelectedMembers(self.currentProject):
                # o = self.currentProject.Concordance.Lookup(sn)
                verticesGlobal = []
                
                # create Point3D List of vertices, not in any order yet
                if  isinstance(o, Shell3D): # in case it is an IFC Mesh we get us the coordinates
                    try: #2023.11
                        vertexIndices = o.GetTriangulatedFaceList() # this works with self created 3DShells - i.e. SweepShape, or Linebundle
                        verticesLocal = o.GetVertex() # vertices as Point3Ds
                        for i in range(0, vertexIndices.Count, 4):
                            verticesGlobal.Add(o.GlobalTransformation.TransformPoint(verticesLocal[vertexIndices[i + 1]]))
                            verticesGlobal.Add(o.GlobalTransformation.TransformPoint(verticesLocal[vertexIndices[i + 2]]))
                            verticesGlobal.Add(o.GlobalTransformation.TransformPoint(verticesLocal[vertexIndices[i + 3]]))

                    except: # 2024.00
                        tt = o.GetTrianglesForInspection()
                        for t in tt:
                            verticesGlobal.Add(t.pointA)
                            verticesGlobal.Add(t.pointB)
                            verticesGlobal.Add(t.pointC)

                elif isinstance(o, BIMEntity):
                    verticesGlobal = []
                    for shellMeshInstance in o.GetGeometry():
                        shellMeshData = shellMeshInstance.GetShellMeshData()
                        
                        try: #2023.11
                            # DEPENDING ON THE TYPE OF IFC THE DIFFERENT METHODS RETURN EMPTY LISTS
                            vertexIndices = shellMeshData.GetTriangulatedFaceList() # this works for the bridge IFC
                            if vertexIndices.Count == 0:
                                vertexIndices = shellMeshData.GetFaces()    # this works for the geotech, but not the bridges
                                verticesLocal = shellMeshData.GetVertex() # vertices as Point3Ds

                            for i in range(0, vertexIndices.Count, 4):
                                verticesGlobal.Add(o.GlobalTransformation.TransformPoint(verticesLocal[vertexIndices[i + 1]]))
                                verticesGlobal.Add(o.GlobalTransformation.TransformPoint(verticesLocal[vertexIndices[i + 2]]))
                                verticesGlobal.Add(o.GlobalTransformation.TransformPoint(verticesLocal[vertexIndices[i + 3]]))
                        
                        except: # 2024.00
                            tt = shellMeshData.GetTrianglesForInspectionInternal(shellMeshInstance.GlobalTransformation)
                            for t in tt:
                                verticesGlobal.Add(t.pointA)
                                verticesGlobal.Add(t.pointB)
                                verticesGlobal.Add(t.pointC)

                for i in range(0, verticesGlobal.Count, 3):

                    # it can be that the IFC contains "triangles" where all three points are on one line, we don't want those
                    # that would lead to a division by zero in the algorithm that checks if the perpendicular solution is within the triangle
                    # checking it here is doubling up some computations, but it could also mess up the plane and normal vector creation
                    vertex1 = verticesGlobal[i+0]
                    vertex2 = verticesGlobal[i+1]
                    vertex3 = verticesGlobal[i+2]

                    a = Vector3D(vertex1, vertex2)
                    b = Vector3D(vertex1, vertex3)
                    
                    aa = Vector3D.DotProduct(a,a)[0]
                    ab = Vector3D.DotProduct(a,b)[0]
                    bb = Vector3D.DotProduct(b,b)[0]
                    
                    d = round(ab * ab - aa * bb, 14) # otherwise it could still happen that "triangles" slip through

                    if d != 0:
                        vertexlist.Add(vertex1)
                        vertexlist.Add(vertex2)
                        vertexlist.Add(vertex3)

        return vertexlist
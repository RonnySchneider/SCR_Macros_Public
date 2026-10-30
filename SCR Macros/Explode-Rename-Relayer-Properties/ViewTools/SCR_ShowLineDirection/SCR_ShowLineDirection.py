exec(open(r"C:\ProgramData\Trimble\MacroCommands3\SCR Macros\SCR_Imports.py").read())

# ── module-level state ────────────────────────────────────────────────────────
# TBC runs this file once for Startup() and again for the UI panel, giving two
# separate module instances. The overlay guid and the event handler references are
# therefore kept in AppDomain data slots, so any instance can unsubscribe and clear
# what another one wired up. Only the instance that subscribed last (the "owner")
# draws, so the flags in effect are always the ones applied most recently.

_SHARED_PREFIX = "SCR_ShowLineDirection.shared."


def _shared_get(name):
    return AppDomain.CurrentDomain.GetData(_SHARED_PREFIX + name)


def _shared_set(name, value):
    AppDomain.CurrentDomain.SetData(_SHARED_PREFIX + name, value)


_SETTINGS_KEY          = "SCR_ShowLineDirection.enabled"
_SETTINGS_KEY_INTERVAL = "SCR_ShowLineDirection.arrow_interval"
_SETTINGS_KEY_COLOR    = "SCR_ShowLineDirection.arrow_color"
_SETTINGS_KEY_SEGTYPE  = "SCR_ShowLineDirection.show_seg_type"
_SETTINGS_KEY_SEGWIDTH = "SCR_ShowLineDirection.seg_type_width"
_SETTINGS_KEY_TTFAIL   = "SCR_ShowLineDirection.show_tt_fail"
_SETTINGS_KEY_TTWIDTH  = "SCR_ShowLineDirection.tt_fail_width"

_overlay_guid = _shared_get("overlay_guid")
if _overlay_guid is None:
    _overlay_guid = Guid.NewGuid()
    _shared_set("overlay_guid", _overlay_guid)
_module_id           = Guid.NewGuid()   # identifies this module instance as the drawing owner
_proj_loaded_handler = None   # keeps UIEvents.UIProjectLoaded delegate alive
_proj_unload_handler = None   # keeps UIEvents.BeforeUIProjectUnloaded delegate alive
_active_project      = None
_redraw_pending      = False  # debounce flag — only one redraw queued at a time
_arrow_interval      = 5.0   # number of arrows per line
_arrow_color         = Color.Lime
_enabled             = False
_show_seg_type       = False
_seg_type_width      = 3
_show_tt_fail        = False
_tt_fail_width       = 5
_active_panel        = None

_LINE_TYPE = clr.GetClrType(IPolyseg)

_OPTIONS = {
    "enabled": False,
}

# ── TBC registration ──────────────────────────────────────────────────────────

def Setup(cmdData, macroFileFolder):
    cmdData.Key         = "SCR_ShowLineDirection"
    cmdData.CommandName = "SCR_ShowLineDirection"
    cmdData.Caption     = "_SCR_ShowLineDirection"
    cmdData.UIForm      = "SCR_ShowLineDirection"
    cmdData.HelpFile    = "Macros.chm"
    cmdData.HelpTopic   = "22602"
    cmdData.HasStartup  = True  # TBC calls Startup() on every application launch

    try:
        cmdData.DefaultTabKey         = "SCR Expld-SNR-Relay-Prop"
        cmdData.DefaultTabGroupKey    = "View Tools"
        cmdData.ShortCaption          = "Line Direction"
        cmdData.DefaultRibbonToolSize = 3
        cmdData.Version     = 0.016
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo   = "Shows directional arrows on selected linework. " \
                              "State is saved per project and restored on TBC restart."
        cmdData.ToolTipTitle         = "Show Line Direction Arrows"
        cmdData.ToolTipTextFormatted = "Overlays direction arrows on any selected line. " \
                                      "Persists across sessions via project settings."
    except:
        pass
    try:
        b = Bitmap(macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


def Startup(*args):
    """Called by TBC after all commands are loaded. Wires project-lifecycle events."""
    global _proj_loaded_handler, _proj_unload_handler

    _proj_loaded_handler = lambda sender, e: _on_project_loaded(sender)
    _proj_unload_handler = lambda sender, e: _on_project_unloading()

    UIEvents.UIProjectLoaded         += _proj_loaded_handler
    UIEvents.BeforeUIProjectUnloaded += _proj_unload_handler

    # A project may already be open when Startup fires (e.g. auto-open on restart).
    project = TrimbleOffice.TheOffice.CurrentProject
    if project is not None:
        _on_project_loaded(project)


# ── project lifecycle ─────────────────────────────────────────────────────────

def _on_project_loaded(project):
    global _active_project, _arrow_interval, _arrow_color, _enabled, _show_seg_type, _seg_type_width, _show_tt_fail, _tt_fail_width
    if project is None:
        return
    _active_project = project
    settings = ConstructionCommandsSettings.ProvideObject(project)
    try:
        _arrow_interval = settings.GetDouble(_SETTINGS_KEY_INTERVAL, 5.0)
    except:
        pass
    try:
        _arrow_color = Color.FromArgb(int(settings.GetDouble(_SETTINGS_KEY_COLOR, float(Color.Lime.ToArgb()))))
    except:
        pass
    try:
        _enabled = settings.GetBoolean(_SETTINGS_KEY, False)
    except:
        pass
    try:
        _show_seg_type = settings.GetBoolean(_SETTINGS_KEY_SEGTYPE, False)
    except:
        pass
    try:
        _seg_type_width = int(settings.GetDouble(_SETTINGS_KEY_SEGWIDTH, 3.0))
    except:
        pass
    try:
        _show_tt_fail = settings.GetBoolean(_SETTINGS_KEY_TTFAIL, False)
    except:
        pass
    try:
        _tt_fail_width = int(settings.GetDouble(_SETTINGS_KEY_TTWIDTH, 5.0))
    except:
        pass
    if _enabled or _show_seg_type or _show_tt_fail:
        _subscribe_selection(project)


def _on_project_unloading():
    _unsubscribe_selection()
    _clear_overlay()
    global _active_project
    _active_project = None


# ── GlobalSelection wiring ────────────────────────────────────────────────────

def _subscribe_selection(project):
    """Replaces any existing subscription (from this or another module instance) with this instance's handlers."""
    _unsubscribe_selection()
    selection_handler      = lambda sender, e: _redraw(project)
    data_processed_handler = lambda sender, e: _redraw(project)
    calc_project_handler   = lambda sender, e: _redraw(project)
    GlobalSelection.SelectionEdited              += selection_handler
    UIEvents.AfterDataProcessing                 += data_processed_handler
    CalculateProjectEvents.CalculateProjectEvent += calc_project_handler
    _shared_set("selection_handler",      selection_handler)
    _shared_set("data_processed_handler", data_processed_handler)
    _shared_set("calc_project_handler",   calc_project_handler)
    _shared_set("owner",                  _module_id)


def _unsubscribe_selection():
    selection_handler      = _shared_get("selection_handler")
    data_processed_handler = _shared_get("data_processed_handler")
    calc_project_handler   = _shared_get("calc_project_handler")
    if selection_handler is not None:
        GlobalSelection.SelectionEdited              -= selection_handler
    if data_processed_handler is not None:
        UIEvents.AfterDataProcessing                 -= data_processed_handler
    if calc_project_handler is not None:
        CalculateProjectEvents.CalculateProjectEvent -= calc_project_handler
    _shared_set("selection_handler",      None)
    _shared_set("data_processed_handler", None)
    _shared_set("calc_project_handler",   None)
    _shared_set("owner",                  None)


# ── overlay drawing ───────────────────────────────────────────────────────────

def _redraw(project):
    """Debounced entry point. Schedules one _do_redraw per batch of SelectionEdited events."""
    global _redraw_pending
    if _redraw_pending:
        return
    _redraw_pending = True
    Dispatcher.CurrentDispatcher.BeginInvoke(
        DispatcherPriority.Background,
        Action(lambda: _do_redraw(project))
    )


def _do_redraw(project):
    global _redraw_pending
    _redraw_pending = False
    # disabled before this deferred call fired, or another module instance took over
    if _shared_get("owner") != _module_id:
        return
    _clear_overlay()
    bag = OverlayBag(_overlay_guid)

    draw_error = ""
    try:
        lines = [snap for snap in GlobalSelection.SelectedMembers(project) if isinstance(snap, _LINE_TYPE)]
        for snap in lines:
            if _show_seg_type:
                try:
                    _draw_seg_types(bag, snap)
                except Exception as ex:
                    draw_error = str(ex)
            if _show_tt_fail:
                _draw_arc_failures(bag, snap)
        # arrows last so they sit on top of the other highlighting
        if _enabled:
            for snap in lines:
                _draw_arrows(bag, snap)
    except Exception as ex:
        draw_error = str(ex)
    if _active_panel is not None:
        try:
            _active_panel.error.Content = draw_error
        except:
            pass

    views = Array[Guid]([DisplayWindow.HoopsPlanViewGUID, DisplayWindow.Hoops3DViewGUID])
    TrimbleOffice.TheOffice.MainWindow.AppViewManager.AddOverlayGeometry(views, bag)


def _arrow_locations(line, count):
    """
    count is the maximum number of arrows. If the line has fewer segments than that,
    one arrow is placed at the mid point of each segment instead. Same return format
    as SCROverlayBag.getarrowlocations: [Point3D, right-perpendicular azimuth].
    Lines in a tilted UCS always use the evenly spaced helper.
    """
    if hasattr(line, "Normal") and line.Normal.Horizon < math.pi / 2 - 1e-9:
        return SCROverlayBag.getarrowlocations(line, count)

    polyseg = line.ComputePolySeg().Clone().ToWorld()
    polyseg.ComputeStationing()
    mids = []
    seg = polyseg.FirstSegment
    while seg is not None:
        if seg.EndStation - seg.BeginStation > 1e-9:
            mids.append((seg.BeginStation + seg.EndStation) * 0.5)
        seg = polyseg.Next(seg)
    if len(mids) >= count:
        return SCROverlayBag.getarrowlocations(line, count)

    polyseg_v = line.ComputeVerticalPolySeg()
    out_seg  = clr.StrongBox[Segment]()
    out_t    = clr.StrongBox[float]()
    out_pt   = clr.StrongBox[Point3D]()
    out_vec  = clr.StrongBox[Vector3D]()
    out_defl = clr.StrongBox[float]()
    pts = []
    for sta in mids:
        if not polyseg.FindPointFromStation(sta, out_seg, out_t, out_pt, out_vec, out_defl):
            continue
        p = out_pt.Value
        if polyseg_v is not None:
            p.Z = polyseg_v.ComputeVerticalSlopeAndGrade(sta)[1]
        pts.append([p, out_vec.Value.Azimuth])
    return pts


def _draw_arrows(bag, line):
    try:
        count = max(int(_arrow_interval), 1)
        color = _arrow_color.ToArgb()
        for p in _arrow_locations(line, count):
            bag.AddMarker(p[0], GraphicMarkerTypes.Arrow_IndependentColor,
                          color, "", 0, math.pi - p[1], 3.0)
    except:
        pass


def _seg_type_color(e, etype_name):
    n = etype_name.lower()
    if "straight"   in n: return Color.Cyan.ToArgb()
    if "smooth"     in n: return Color.Magenta.ToArgb()
    if "spiral"     in n: return Color.Yellow.ToArgb()
    if "tangentarc" in n:
        try:
            tt = e.TangentType
            if   tt == ArcTangentType.TangentTangent: return Color.MediumPurple.ToArgb()
            elif tt == ArcTangentType.Begin:           return Color.LimeGreen.ToArgb()
            elif tt == ArcTangentType.End:             return Color.DodgerBlue.ToArgb()
        except:
            pass
        return Color.MediumPurple.ToArgb()
    if "arc"        in n: return Color.Orange.ToArgb()
    return Color.White.ToArgb()


def _draw_seg_types(bag, line):
    try:
        ec = line.ElementCount
    except:
        return
    if ec <= 0:
        return

    for i in range(ec):
        e = None
        try:
            e     = line.ElementAt(i)
            color = _seg_type_color(e, e.GetType().Name)
            ep    = line.GetElementPolysegAt(i)   # world coords
            ep_w  = ep.ToWorld()                  # ensure world (may be no-op)
            chord = ep_w.Linearize(0.001, 0.001, 100, None, False)
            pts   = chord.ToPoint3DArray()
            if pts is not None and len(pts) >= 2:
                bag.AddPolyline(pts, color, _seg_type_width)
        except Exception as ex:
            raise Exception("seg[%d] %s: %s" % (i, e.GetType().Name if e is not None else "?", str(ex)))


def _neighbour_segment(line, i, step, ec, closed):
    """
    Returns the last segment of the nearest previous element (step=-1) or the first
    segment of the nearest next element (step=1), skipping elements without geometry
    (e.g. the start point element). Wraps around on closed lines, else returns None.
    """
    j = i + step
    for k in range(ec):
        if j < 0 or j >= ec:
            if not closed:
                return None
            j %= ec
        if j == i:
            return None
        try:
            ps  = line.GetElementPolysegAt(j).ToWorld()
            seg = ps.LastSegment if step < 0 else ps.FirstSegment
            if seg is not None:
                return seg
        except:
            pass
        j += step
    return None


def _tangent_tangent_check(seg_in, seg_out, a, b):
    """
    Mirrors Linestring.DetermineTangentTangentData (decompiled from TBC 2026.1).
    The tangent at the arc begin a (end of the previous segment) and the tangent at
    the arc end b (start of the next segment) are intersected as infinite lines.
    TBC rejects the arc if the intersection lies ahead of both points or behind both
    points, if the tangent length is zero, or if the begin tangent runs through b.
    Unequal tangent lengths are not a failure - TBC uses the shorter one and adds a
    straight on the longer side. Parallel tangents give a semicircle.
    Returns (failed, pi), pi being the tangent intersection or None if parallel.
    """
    ok, din  = seg_in.Slope(1.0)
    ok, dout = seg_out.Slope(0.0)
    la = math.hypot(din.X, din.Y)
    lb = math.hypot(dout.X, dout.Y)
    if la < 1e-12 or lb < 1e-12:
        return True, None
    dax, day = din.X / la, din.Y / la
    dbx, dby = dout.X / lb, dout.Y / lb
    wx, wy = b.X - a.X, b.Y - a.Y
    if math.hypot(wx, wy) < 1e-6:
        return True, None
    det = dax * dby - day * dbx
    if abs(det) < 1e-9:
        return abs(wx * day - wy * dax) < 1e-6, None
    t1 = (wx * dby - wy * dbx) / det   # a + t1 * din  = pi
    t2 = (wx * day - wy * dax) / det   # b + t2 * dout = pi
    pi = Point3D(a.X + t1 * dax, a.Y + t1 * day, a.Z)
    failed = t1 * t2 >= 0 or abs(t1) < 1e-6 or abs(t2) < 1e-6
    return failed, pi


def _flag_element(bag, ep, color, marker):
    """Thick line plus flags at the start, middle and end of an element polyseg."""
    pts = ep.Linearize(0.001, 0.001, 100, None, False).ToPoint3DArray()
    if pts is None or len(pts) < 2:
        return
    bag.AddPolyline(pts, color, _tt_fail_width)
    for idx in sorted({0, len(pts) // 2, len(pts) - 1}):
        bag.AddMarker(pts[idx], marker, color, "", 0, 0, 1.0)


def _draw_arc_failures(bag, line):
    """
    Highlights arc elements TBC can't build as specified:
    - tangent-tangent arcs without a valid tangent intersection (red), plus the
      tangent lines to the intersection (thin orange)
    - radius arcs whose radius is smaller than half the chord (orange-red); TBC
      can't fit that radius between the two points and draws a different arc
    """
    try:
        ec = line.ElementCount
    except:
        return
    closed = False
    try:
        closed = bool(line.ComputePolySeg().IsClosed)
    except:
        pass
    for i in range(ec):
        try:
            e     = line.ElementAt(i)
            etype = e.GetType().Name.lower()
            ep    = line.GetElementPolysegAt(i).ToWorld()
            first = ep.FirstSegment
            last  = ep.LastSegment
            if first is None or last is None:
                continue

            if "tangentarc" in etype:
                if e.TangentType != ArcTangentType.TangentTangent:
                    continue
                seg_in  = _neighbour_segment(line, i, -1, ec, closed)
                seg_out = _neighbour_segment(line, i,  1, ec, closed)
                pi = None
                if seg_in is None or seg_out is None:
                    failed = True   # nothing to be tangent to
                else:
                    failed, pi = _tangent_tangent_check(seg_in, seg_out, seg_in.EndPoint, last.EndPoint)
                if not failed:
                    continue
                _flag_element(bag, ep, Color.Red.ToArgb(), GraphicMarkerTypes.Flag_Red)
                if pi is not None:
                    a, b  = seg_in.EndPoint, last.EndPoint
                    chord = math.hypot(b.X - a.X, b.Y - a.Y)
                    # skip near-parallel tangents whose intersection is far off screen
                    if math.hypot(pi.X - a.X, pi.Y - a.Y) < 50 * chord:
                        bag.AddPolyline(Array[Point3D]([a, pi, b]), Color.Orange.ToArgb(), 1)

            elif isinstance(e, IArcSegment):
                a, b   = first.BeginPoint, last.EndPoint
                chord  = math.hypot(b.X - a.X, b.Y - a.Y)
                radius = abs(e.Radius)
                if radius > 0 and radius < chord / 2.0 - 1e-6:
                    _flag_element(bag, ep, Color.OrangeRed.ToArgb(), GraphicMarkerTypes.Flag_Orange)
        except:
            pass


def _clear_overlay():
    try:
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(_overlay_guid)
    except:
        pass


# ── public toggle (called by UI and indirectly by Startup) ───────────────────

def set_enabled(project, on, interval=None, color=None, show_seg_type=None, seg_type_width=None, show_tt_fail=None, tt_fail_width=None):
    """Persist the on/off state, interval, colour, segment-type and tangent-arc settings to the project and wire/unwire the selection handler."""
    global _arrow_interval, _arrow_color, _enabled, _show_seg_type, _seg_type_width, _show_tt_fail, _tt_fail_width
    _enabled = on
    settings = ConstructionCommandsSettings.ProvideObject(project)
    settings.SetBoolean(_SETTINGS_KEY, on)
    if interval is not None:
        _arrow_interval = float(interval)
        try:
            settings.SetDouble(_SETTINGS_KEY_INTERVAL, _arrow_interval)
        except:
            pass
    if color is not None:
        _arrow_color = color
        try:
            settings.SetDouble(_SETTINGS_KEY_COLOR, float(color.ToArgb()))
        except:
            pass
    if show_seg_type is not None:
        _show_seg_type = bool(show_seg_type)
        try:
            settings.SetBoolean(_SETTINGS_KEY_SEGTYPE, _show_seg_type)
        except:
            pass
    if seg_type_width is not None:
        _seg_type_width = max(1, min(5, int(seg_type_width)))
        try:
            settings.SetDouble(_SETTINGS_KEY_SEGWIDTH, float(_seg_type_width))
        except:
            pass
    if show_tt_fail is not None:
        _show_tt_fail = bool(show_tt_fail)
        try:
            settings.SetBoolean(_SETTINGS_KEY_TTFAIL, _show_tt_fail)
        except:
            pass
    if tt_fail_width is not None:
        _tt_fail_width = max(1, min(5, int(tt_fail_width)))
        try:
            settings.SetDouble(_SETTINGS_KEY_TTWIDTH, float(_tt_fail_width))
        except:
            pass
    if _enabled or _show_seg_type or _show_tt_fail:
        _subscribe_selection(project)
        _redraw(project)
    else:
        _unsubscribe_selection()
        _clear_overlay()


# ── UI panel ──────────────────────────────────────────────────────────────────

class SCR_ShowLineDirection(StackPanel):

    def __init__(self, currentProject, macroFileFolder):
        with StreamReader(macroFileFolder + r"\SCR_ShowLineDirection.xaml") as s:
            wpf.LoadComponent(self, s)
        self.currentProject  = currentProject
        self.macroFileFolder = macroFileFolder

    def HelpClicked(self, cmd, e):
        webbrowser.open(r"C:\ProgramData\Trimble\MacroCommands3\SCR Macros\MacroHelp\MacroHelp.htm#" + type(self).__name__)

    def OnLoad(self, cmd, buttons, event):
        global _active_panel
        _active_panel = self
        buttons[0].Content    = "Apply"
        buttons[0].Click     += lambda s, e: self._apply()
        buttons[1].Content    = "Close"
        buttons[2].Content    = "Help"
        buttons[2].Visibility = Visibility.Visible
        buttons[2].Click     += self.HelpClicked
        self.Caption = cmd.Command.Caption
        # restore saved state before wiring events to avoid spurious triggers
        settings = ConstructionCommandsSettings.ProvideObject(self.currentProject)
        self.enabled.IsChecked = settings.GetBoolean(_SETTINGS_KEY, False)
        try:
            self.arrow_interval.Value = settings.GetDouble(_SETTINGS_KEY_INTERVAL, 5.0)
        except:
            self.arrow_interval.Value = 5.0
        self.arrow_colour.SelectedColor = Color.Lime
        try:
            self.arrow_colour.SelectedColor = Color.FromArgb(int(settings.GetDouble(_SETTINGS_KEY_COLOR, float(Color.Lime.ToArgb()))))
        except:
            pass
        self.show_seg_type.IsChecked = settings.GetBoolean(_SETTINGS_KEY_SEGTYPE, False)
        try:
            self.seg_type_width.Value = settings.GetDouble(_SETTINGS_KEY_SEGWIDTH, 3.0)
        except:
            self.seg_type_width.Value = 3.0
        self.show_tt_fail.IsChecked = settings.GetBoolean(_SETTINGS_KEY_TTFAIL, False)
        try:
            self.tt_fail_width.Value = settings.GetDouble(_SETTINGS_KEY_TTWIDTH, 5.0)
        except:
            self.tt_fail_width.Value = 5.0
        self.enabled.Checked              += lambda s, e: self._apply()
        self.enabled.Unchecked             += lambda s, e: self._apply()
        self.arrow_interval.ValueChanged   += lambda s, e: self._apply()
        self.arrow_colour.ValueChanged     += lambda s, e: self._apply()
        self.show_seg_type.Checked         += lambda s, e: self._apply()
        self.show_seg_type.Unchecked       += lambda s, e: self._apply()
        self.seg_type_width.ValueChanged   += lambda s, e: self._apply()
        self.show_tt_fail.Checked          += lambda s, e: self._apply()
        self.show_tt_fail.Unchecked        += lambda s, e: self._apply()
        self.tt_fail_width.ValueChanged    += lambda s, e: self._apply()
        self._apply()

    def _refresh_status(self):
        shown = []
        if self.enabled.IsChecked:
            shown.append("arrows")
        if self.show_seg_type.IsChecked:
            shown.append("segment types")
        if self.show_tt_fail.IsChecked:
            shown.append("failed arcs")
        if shown:
            self.status.Content = "Active — " + " + ".join(shown) + " shown"
        else:
            self.status.Content = "Inactive"

    def _apply(self):
        self.error.Content = ""
        try:
            interval        = int(self.arrow_interval.Value)
            color           = self.arrow_colour.SelectedColor
            show_seg_type   = bool(self.show_seg_type.IsChecked)
            seg_type_width  = int(self.seg_type_width.Value)
            show_tt_fail    = bool(self.show_tt_fail.IsChecked)
            tt_fail_width   = int(self.tt_fail_width.Value)
            set_enabled(self.currentProject, bool(self.enabled.IsChecked), interval, color, show_seg_type, seg_type_width, show_tt_fail, tt_fail_width)
            self._refresh_status()
        except Exception as ex:
            self.error.Content = str(ex)
            try:
                _unsubscribe_selection()
            except:
                pass
            try:
                _clear_overlay()
            except:
                pass
            self.enabled.IsChecked = False
            self._refresh_status()

    def CancelClicked(self, sender, args):
        sender.CloseUICommand()

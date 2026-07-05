from System.Collections.Generic import List, IEnumerable
exec(open(r"C:\ProgramData\Trimble\MacroCommands3\SCR Macros\SCR_Imports.py").read())

# ── Persisted UI state ────────────────────────────────────────────────────────

_OPTIONS = {
    "sidepicker":          0,      # 0=Left, 1=Centre, 2=Right
    "designspeed":         40.0,   # reserved — not yet exposed in UI
    "friction":            0.30,   # reserved — not yet exposed in UI
    "stepwidth":           0.05,   # simulation step along the path (metres)
    "reverse_line":        False,  # simulate in the reverse direction along the selected line
    "tp_buffer":           False,  # zero curvature over ±6 m at reverse curves tighter than R40
    "draw_front_corners":  True,
    "draw_mid_sweep":      True,
    "draw_rear_corners":   True,
    "draw_rear_axle":      True,
    "draw_envelope":       True,
    "draw_outline":        False,
    "outline_interval":    5.0,    # metres between body-outline footprints
    "color_front_corners": Color.FromArgb(0, 0, 255),
    "color_mid_sweep":     Color.Purple,
    "color_rear_corners":  Color.Red,
    "color_rear_axle":     Color.Yellow,
    "color_envelope":      Color.Lime,
    "color_outline":       Color.Magenta,
    "first_last_only":     False,  # draw track lines for first and last trailer unit only
    "layer_prefix":        "VPath 1",  # layer name prefix and layer group name
    "rb_layers":           True,   # False = draw temporary view overlay instead
    "auto_refresh_overlay": True,
    "overlay_linewidth":    1,
}

_VEHICLE_KEY = "SCR_VPath.vehicle"   # OptionsManager key for last-used vehicle

# Minimum horizontal curve radii (metres) per AUSTROADS vehicle type.
# Tuple: (absolute_minimum, desirable_minimum)
_RADIUS_TABLE = {
    "DESIGN SEMITRAILER":     (12.5, 15.0),
    "DESIGN TRUCK/BUS":       (12.5, 15.0),
    "DESIGN CAR":             (7.5,  10.0),
    "DESIGN LOW LOADER":      (15.0, 20.0),
    "23m B-DOUBLE":           (15.0, 20.0),
    "33m TYPE 1 ROAD TRAIN":  (15.0, 20.0),
    "49m TYPE 2 ROAD TRAIN":  (15.0, 20.0),
    "12.2m TOURIST BUS":      (12.5, 15.0),
    "ARTICULATED BUS":        (10.6, 15.0),
    "TRUCK AND TRAILER":      (12.5, 15.0),
    "CAR AND CARAVAN":        (7.5,  10.0),
}


# ── TBC registration ──────────────────────────────────────────────────────────

def Setup(cmdData, macroFileFolder):
    cmdData.Key         = "SCR_VPath"
    cmdData.CommandName = "SCR_VPath"
    cmdData.Caption     = "_SCR_VPath"
    cmdData.UIForm      = "SCR_VPath"
    cmdData.HelpFile    = "Macros.chm"
    cmdData.HelpTopic   = "22602"

    try:
        cmdData.DefaultTabKey         = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey    = "Lines"
        cmdData.ShortCaption          = "VPath"
        cmdData.DefaultRibbonToolSize = 3
        cmdData.Version     = 0.24
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo   = "Computes AUSTROADS vehicle swept path along a selected line."
        cmdData.ToolTipTitle         = "Vehicle Swept Path"
        cmdData.ToolTipTextFormatted = "Simulates the swept path of an AUSTROADS vehicle along a selected line."
    except:
        pass
    try:
        b = Bitmap(macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


# ── UI panel ──────────────────────────────────────────────────────────────────

class SCR_VPath(StackPanel):
    """TBC macro panel for AUSTROADS vehicle swept-path simulation.

    Two output modes:
      Layers  — creates permanent TBC linework inside a transaction.
      Overlay — draws temporary coloured polylines directly into the viewport.

    A simulation result cache (self.sim_cache) avoids redundant re-simulation
    when only the draw toggles change in overlay mode.  The cache is keyed by
    all parameters that affect the simulation output; changing any of them
    clears the cache and triggers a full re-simulation.
    """

    def __init__(self, currentProject, macroFileFolder):
        with StreamReader(macroFileFolder + r"\SCR_VPath.xaml") as s:
            wpf.LoadComponent(self, s)
        self.currentProject  = currentProject
        self.macroFileFolder = macroFileFolder

        self.ViewOverlay = Guid.NewGuid()
        self.overlayBag  = OverlayBag(self.ViewOverlay)

    # ── Lifecycle ─────────────────────────────────────────────────────────────

    def OnLoad(self, cmd, buttons, event):
        self.okBtn = buttons[0]
        buttons[2].Content    = "Help"
        buttons[2].Visibility = Visibility.Visible
        buttons[2].Click     += self.HelpClicked
        self.Caption = cmd.Command.Caption

        self.lType = clr.GetClrType(IPolyseg)

        # OK button label reflects the active output mode.
        def update_ok_label(s=None, e=None):
            if bool(self.rb_overlay.IsChecked):
                self.okBtn.Content = "Refresh Overlay"
            else:
                self.okBtn.Content = "Create Linework"

        self.linepicker.IsEntityValidCallback = self.IsValid
        self.linepicker.ValueChanged         += self.lineChanged
        self.rb_overlay.Checked              += lambda s, e: (update_ok_label(), self._schedule_sweep())
        self.rb_layers.Checked               += update_ok_label
        update_ok_label()

        # Drawing-only toggles: just redraw the overlay from the existing cache.
        def on_draw_option_changed(s, e):
            if bool(self.rb_overlay.IsChecked):
                self._schedule_sweep()
        for cb in [self.draw_front_corners, self.draw_mid_sweep, self.draw_rear_corners,
                   self.draw_rear_axle, self.draw_envelope, self.first_last_only, self.draw_outline]:
            cb.Checked   += on_draw_option_changed
            cb.Unchecked += on_draw_option_changed
        for cp in [self.color_front_corners, self.color_mid_sweep, self.color_rear_corners,
                   self.color_rear_axle, self.color_envelope, self.color_outline]:
            cp.ValueChanged += on_draw_option_changed
        self.outline_interval.ValueChanged += on_draw_option_changed

        def on_linewidth_changed(s, e):
            self.linewidth_label.Text = str(int(self.overlay_linewidth.Value))
            if bool(self.rb_overlay.IsChecked):
                self._schedule_sweep()
        self.overlay_linewidth.ValueChanged += on_linewidth_changed


        # Simulation parameters: clear cache and re-simulate.
        def on_sim_param_changed(s, e):
            self.sim_cache = None
            if bool(self.rb_overlay.IsChecked):
                self._schedule_sweep()
        self.sidepicker.SelectionChanged    += on_sim_param_changed
        self.vehiclepicker.SelectionChanged += on_sim_param_changed
        self.reverse_line.Checked   += on_sim_param_changed
        self.reverse_line.Unchecked += on_sim_param_changed

        self.sim_cache        = None
        self._sweep_pending   = False
        self._ui_dispatcher   = Dispatcher.CurrentDispatcher
        # Keep a reference to the handler so it can be removed on Dispose.
        self._compute_handler = lambda sender, e: self._on_entity_edited()
        ModelEvents.EntityEdited += self._compute_handler

        self._load_vehicles()

        try:
            self.SetDefaultOptions()
        except:
            pass

    def Dispose(self, cmd, disposing):
        if self._compute_handler is not None:
            ModelEvents.EntityEdited -= self._compute_handler
            self._compute_handler = None
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)

    def HelpClicked(self, cmd, e):
        webbrowser.open(r"C:\ProgramData\Trimble\MacroCommands3\SCR Macros\MacroHelp\MacroHelp.htm#" + type(self).__name__)

    def CancelClicked(self, cmd, args):
        self.SaveOptions()
        cmd.CloseUICommand()

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.SaveOptions()
        self._run_sweep(validate=True)

    # ── Options persistence ───────────────────────────────────────────────────

    def SetDefaultOptions(self):
        SCROptions.LoadMacroOptions(self, "SCR_VPath", _OPTIONS, self.currentProject)
        self.rb_overlay.IsChecked = not bool(self.rb_layers.IsChecked)
        saved = OptionsManager.GetString(_VEHICLE_KEY, "")
        if saved:
            for i in range(self.vehiclepicker.Items.Count):
                item = self.vehiclepicker.Items[i]
                if isinstance(item, ComboBoxItem) and item.IsEnabled and str(item.Content) == saved:
                    self.vehiclepicker.SelectedIndex = i
                    break

    def SaveOptions(self):
        SCROptions.SaveMacroOptions(self, "SCR_VPath", _OPTIONS)
        selected = self.vehiclepicker.SelectedItem
        if isinstance(selected, ComboBoxItem) and selected.IsEnabled:
            OptionsManager.SetValue(_VEHICLE_KEY, str(selected.Content))

    # ── Vehicle list ──────────────────────────────────────────────────────────

    def _load_vehicles(self):
        """Parse Vehicles_custom.xml and populate the vehicle combo box."""
        self._vehicle_nodes = {}
        xmlpath = os.path.join(self.macroFileFolder, "Vehicles_custom.xml")
        try:
            tree = ET.parse(xmlpath)
            current_cat = None
            for v in tree.getroot().find("Vehicle_Definitions").findall("Vehicle"):
                name = v.get("Name", "")
                cat  = v.get("Category", "")
                if not name:
                    continue
                self._vehicle_nodes[name] = v
                if cat != current_cat:
                    current_cat = cat
                    hdr = ComboBoxItem()
                    hdr.Content    = cat
                    hdr.IsEnabled  = False
                    hdr.FontWeight = FontWeights.Bold
                    self.vehiclepicker.Items.Add(hdr)
                item = ComboBoxItem()
                item.Content = name
                self.vehiclepicker.Items.Add(item)
        except Exception as ex:
            self.error.Text = "Could not load Vehicles_custom.xml: " + str(ex)

    def _get_selected_units(self):
        """Return the parsed unit list for the currently selected vehicle, or None."""
        selected = self.vehiclepicker.SelectedItem
        if not isinstance(selected, ComboBoxItem) or not selected.IsEnabled:
            return None
        v_elem = self._vehicle_nodes.get(str(selected.Content))
        return _parse_vehicle_units(v_elem) if v_elem is not None else None

    # ── Entity validation ─────────────────────────────────────────────────────

    def IsValid(self, serial):
        """Accept only polyseg (alignment) entities in the line picker."""
        o = self.currentProject.Concordance[serial]
        return isinstance(o, self.lType)

    # ── Overlay helpers ───────────────────────────────────────────────────────

    def _add_line_to_bag(self, line, bag):
        """Draw the reference line (green) and direction arrows (orange) into an OverlayBag.
        Respects the reverse_line checkbox."""
        rev = bool(self.reverse_line.IsChecked)
        pts = SCROverlayBag.getpolypoints(line)
        if rev:
            pts = Array[Point3D](list(reversed(list(pts))))
        bag.AddPolyline(pts, Color.Green.ToArgb(), 3)
        arrows = SCROverlayBag.getarrowlocations(line, 10)
        if rev:
            arrows = list(reversed(arrows))
        for p in arrows:
            angle = math.pi - p[1] + (math.pi if rev else 0)
            bag.AddMarker(p[0], GraphicMarkerTypes.Arrow_IndependentColor,
                          Color.Orange.ToArgb(), "", 0, angle, 3.0)

    def _show_line_direction_active(self):
        try:
            settings = ConstructionCommandsSettings.ProvideObject(self.currentProject)
            return settings.GetBoolean("SCR_ShowLineDirection.enabled", False)
        except Exception:
            return False

    def _drawoverlay(self):
        """Redraw the reference-line overlay (green path + orange direction arrows)."""
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)
        self.overlayBag = OverlayBag(self.ViewOverlay)
        line = self.linepicker.Entity
        if line and not self._show_line_direction_active():
            self._add_line_to_bag(line, self.overlayBag)
        views = Array[Guid]([DisplayWindow.Hoops3DViewGUID, DisplayWindow.HoopsPlanViewGUID])
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.AddOverlayGeometry(views, self.overlayBag)

    def _on_entity_edited(self):
        """Called when any model entity changes; may arrive on a background thread.
        Invalidate the cache immediately (assignment is thread-safe), then marshal
        the UI check and sweep scheduling onto the UI dispatcher."""
        self.sim_cache = None
        def on_ui_thread():
            if bool(self.rb_overlay.IsChecked) and bool(self.auto_refresh_overlay.IsChecked):
                self._schedule_sweep()
        self._ui_dispatcher.BeginInvoke(DispatcherPriority.Background, Action(on_ui_thread))

    def lineChanged(self, ctrl, e):
        self.sim_cache = None
        if bool(self.rb_overlay.IsChecked):
            if bool(self.auto_refresh_overlay.IsChecked):
                self._schedule_sweep()
        else:
            self._drawoverlay()

    # ── Deferred sweep scheduling ─────────────────────────────────────────────

    def _schedule_sweep(self):
        """Post a deferred overlay redraw on the dispatcher (coalesces rapid changes)."""
        if self._sweep_pending:
            return
        self._sweep_pending = True
        Dispatcher.CurrentDispatcher.BeginInvoke(
            DispatcherPriority.Background,
            Action(self._deferred_sweep)
        )

    def _deferred_sweep(self):
        self._sweep_pending = False
        if bool(self.rb_layers.IsChecked):
            return
        self._run_sweep()

    # ── Main sweep entry point ────────────────────────────────────────────────

    def _run_sweep(self, validate=False):
        """Run or redraw the swept path.

        Layers mode  — always runs the full simulation inside a TBC transaction.
        Overlay mode — uses the simulation cache when possible; only calls
                       run_simulation() when the cache key does not match.
        """
        use_layers = bool(self.rb_layers.IsChecked)
        self.error.Text = ""

        line = self.linepicker.Entity
        if line is None:
            if validate:
                self.error.Text = "Please select an input line."
            elif not use_layers:
                self._drawoverlay()
            return

        units = self._get_selected_units()
        if not units:
            if validate:
                self.error.Text = "Please select at least one vehicle unit."
            elif not use_layers:
                self._drawoverlay()
            return

        step_size = max(0.01, float(self.stepwidth.Value))
        tp_buffer = bool(self.tp_buffer.IsChecked)
        side      = int(self.sidepicker.SelectedIndex)
        ln_name   = IName.Name.__get__(line)
        v_name    = str(self.vehiclepicker.SelectedItem.Content)
        pfx       = ln_name + " - " + v_name

        wv = self.currentProject[Project.FixedSerial.WorldView]

        # ── Layer name helpers ────────────────────────────────────────────────
        LYR_SUFFIXES = {
            "AR":      "Front Right AR",
            "AL":      "Front Left AL",
            "MR":      "Mid Right MR",
            "ML":      "Mid Left ML",
            "CR":      "Rear Axle Right CR",
            "CL":      "Rear Axle Left CL",
            "RR":      "Rear Right RR",
            "RL":      "Rear Left RL",
            "Outline": "Body Outline",
        }
        lyr_prefix = self.layer_prefix.Text.strip() or "VPath 1"
        lyr_group  = None  # assigned inside the transaction below

        def layer(tag, unit_no=1):
            if not use_layers:
                return None
            lyr = Layer.FindOrCreateLayer(
                self.currentProject,
                "{} Unit {} {}".format(lyr_prefix, unit_no, LYR_SUFFIXES[tag])
            )
            if lyr.LayerGroupSerial == 0 and lyr_group is not None:
                lyr.LayerGroupSerial = lyr_group.SerialNumber
            return lyr.SerialNumber

        # ── RDP decimation ────────────────────────────────────────────────────
        # 5 mm in layers mode gives accurate linework.
        # 25 mm in overlay mode speeds up drawing with no perceptible visual loss.
        dec_tol = 0.005 if use_layers else 0.025
        def decimate(pts, tol=None):
            return _decimate(pts, dec_tol if tol is None else tol)

        # ── Linework / overlay drawing ────────────────────────────────────────
        def make_ls(name, xyzs, layer_sn, color=None, weight=None, closed=False):
            """Create one linestring — either as a TBC Linestring entity or an overlay polyline."""
            if len(xyzs) < 2:
                return
            if use_layers:
                ls = wv.Add(clr.GetClrType(Linestring))
                ls.Name   = name
                ls.Layer  = layer_sn
                ls.Closed = closed
                if color  is not None: ls.Color  = color
                if weight is not None: ls.Weight = weight
                for x, y, z in xyzs:
                    el = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                    el.Position = Point3D(x, y, z)
                    ls.AppendElement(el)
            else:
                coords = list(xyzs)
                if closed and coords[0] != coords[-1]:
                    coords.append(coords[0])
                pts  = Array[Point3D]([Point3D(x, y, z) for x, y, z in coords])
                argb = color.ToArgb() if color is not None else Color.White.ToArgb()
                pw   = int(self.overlay_linewidth.Value)
                self.overlayBag.AddPolyline(pts, argb, pw)

        # ── Envelope helpers ──────────────────────────────────────────────────
        ## Region.Union envelope (replaced by _alpha_hull below)
        # def band_region(rp, lp):
        #     if len(rp) < 2 or len(lp) < 2:
        #         return None
        #     ps = PolySeg.PolySeg()
        #     for pt in rp:            ps.Add(Point3D(pt[0], pt[1], pt[2]))
        #     for pt in reversed(lp):  ps.Add(Point3D(pt[0], pt[1], pt[2]))
        #     ps.Add(Point3D(rp[0][0], rp[0][1], rp[0][2]))
        #     return Region(ps)
        # def build_envelope(pm_bands, trailer_list):
        #     env_regions = List[Region]()
        #     for r in pm_bands:
        #         if r is not None: env_regions.Add(r)
        #     for td in trailer_list:
        #         for r in [band_region(td['env_ar'], td['env_al']),
        #                    band_region(td['env_mr'], td['env_ml']),
        #                    band_region(td['env_cr'], td['env_cl']),
        #                    band_region(td['env_rr'], td['env_rl'])]:
        #             if r is not None: env_regions.Add(r)
        #     if env_regions.Count == 0:
        #         return []
        #     union_list = Region.Union(env_regions, HoleProcessingOptions.RemoveChildHoles)
        #     result = []
        #     for ri, r in enumerate(union_list):
        #         sfx = "" if union_list.Count == 1 else " {}".format(ri + 1)
        #         for seg in r.AreaPolysegSegments():
        #             bps_pts = decimate([(p.X, p.Y, p.Z) for p in seg.ToPoint3DArray()])
        #             if len(bps_pts) >= 2:
        #                 result.append((bps_pts, sfx))
        #     return result

        # ── Progress reporting ────────────────────────────────────────────────
        prog_last = [0.0]
        def progress_fn(pct):
            now = time.time()
            if now - prog_last[0] >= 1.0:
                prog_last[0] = now
                ProgressBar.TBC_ProgressBar.SetProgress(pct)

        # ── Simulation ────────────────────────────────────────────────────────
        def run_simulation():
            """Sample the path, run the kinematic simulation, decimate results,
            optionally compute the envelope union, then update self.sim_cache."""
            # Sample the path at step_size intervals using FindPointFromStation so
            # arc curvature is preserved (ToPoint3DArray() only gives chord endpoints).
            polyseg = line.ComputePolySeg().ToWorld()
            if bool(self.reverse_line.IsChecked):
                polyseg.Reverse()
            total_length = polyseg.ComputeStationing()
            out_seg  = clr.StrongBox[Segment]()
            out_t    = clr.StrongBox[float]()
            out_pt   = clr.StrongBox[Point3D]()
            out_vec  = clr.StrongBox[Vector3D]()
            out_defl = clr.StrongBox[float]()
            path_pts = []
            headings = []
            station  = 0.0
            while True:
                s = min(station, total_length)
                if polyseg.FindPointFromStation(s, out_seg, out_t, out_pt, out_vec, out_defl):
                    p = out_pt.Value; v = out_vec.Value
                    path_pts.append((p.X, p.Y, p.Z))
                    headings.append(math.atan2(v.f, -v.g))
                if s >= total_length:
                    break
                station += step_size

            if len(path_pts) < 2:
                raise ValueError("Line is too short to simulate.")
            n_pts = len(path_pts)

            # Smooth headings using centred chord direction (reduces noise from
            # FindPointFromStation tangent vectors on tightly-segmented alignments).
            for i in range(1, n_pts - 1):
                dx = path_pts[i + 1][0] - path_pts[i - 1][0]
                dy = path_pts[i + 1][1] - path_pts[i - 1][1]
                headings[i] = math.atan2(dy, dx)
            headings[0]         = headings[1]
            headings[n_pts - 1] = headings[n_pts - 2]

            ar_pts, al_pts, cr_pts, cl_pts, rr_pts, rl_pts, mr_pts, ml_pts, unit_tracks, mta_events = _simulate_sweep(
                path_pts, headings, units, None, tp_buffer, step_size, side,
                progress_fn=progress_fn)

            # ── Trailer corner points ─────────────────────────────────────────
            # Computed here (not inside _simulate_sweep) so we can store both the
            # display-decimated and envelope-decimated variants in one pass.
            all_trailers = []
            for tj in range(1, len(units)):
                tu      = units[tj]
                tfo     = tu['FO']; tw = tu['W']; ttr = tu['TR']
                troh    = tu['LR'] - tu['LA']
                tla_eff = tu['LA'] + (tu['DA1'] + tu['DA2'] + tu['DA3']) * 0.5
                t_ar = []; t_al = []; t_mr = []; t_ml = []
                t_cr = []; t_cl = []; t_rr = []; t_rl = []
                for si, (fax, fay, rax, ray, th) in enumerate(unit_tracks[tj]):
                    tc    = math.cos(th); ts = math.sin(th)
                    z     = path_pts[si][2]
                    bog_x = fax - tla_eff * tc
                    bog_y = fay - tla_eff * ts
                    ar = (fax + tc*tfo + ts*tw,   fay + ts*tfo - tc*tw,   z)
                    al = (fax + tc*tfo - ts*tw,   fay + ts*tfo + tc*tw,   z)
                    rr = (rax - tc*troh + ts*ttr, ray - ts*troh - tc*ttr, z)
                    rl = (rax - tc*troh - ts*ttr, ray - ts*troh + tc*ttr, z)
                    t_ar.append(ar); t_al.append(al)
                    t_mr.append(((ar[0]+rr[0])*0.5, (ar[1]+rr[1])*0.5, z))
                    t_ml.append(((al[0]+rl[0])*0.5, (al[1]+rl[1])*0.5, z))
                    t_cr.append((bog_x + ts*ttr,   bog_y - tc*ttr,        z))
                    t_cl.append((bog_x - ts*ttr,   bog_y + tc*ttr,        z))
                    t_rr.append(rr); t_rl.append(rl)
                all_trailers.append({
                    'tj': tj,
                    # Display data decimated at the mode tolerance (0.005 or 0.025).
                    'ar': decimate(t_ar), 'al': decimate(t_al),
                    'mr': decimate(t_mr), 'ml': decimate(t_ml),
                    'cr': decimate(t_cr), 'cl': decimate(t_cl),
                    'rr': decimate(t_rr), 'rl': decimate(t_rl),
                    # Hull input: tight tol + max_seg ensures straight-section coverage.
                    # max_edge not yet computed here; filled in below after max_edge is known.
                    '_hull_raw': (t_ar, t_al, t_mr, t_ml, t_cr, t_cl, t_rr, t_rl),
                })

            # Decimate PM display data.
            dec_ar = decimate(ar_pts); dec_al = decimate(al_pts)
            dec_mr = decimate(mr_pts); dec_ml = decimate(ml_pts)
            dec_cr = decimate(cr_pts); dec_cl = decimate(cl_pts)
            dec_rr = decimate(rr_pts); dec_rl = decimate(rl_pts)

            # ── Envelope hull ─────────────────────────────────────────────────
            # Always computed and cached; draw_envelope checkbox only controls visibility.
            first_last_only = bool(self.first_last_only.IsChecked)
            trailer_subset = ([all_trailers[0], all_trailers[-1]]
                              if first_last_only and len(all_trailers) > 1
                              else all_trailers)
            max_w    = max(u['W'] * 2.0 for u in units)
            max_edge = 1.5 * max_w
            # Build hull-decimated PM lines: tight RDP + max_seg keeps straight-section points.
            hull_ar = _decimate(ar_pts, 0.005, max_edge); hull_al = _decimate(al_pts, 0.005, max_edge)
            hull_mr = _decimate(mr_pts, 0.005, max_edge); hull_ml = _decimate(ml_pts, 0.005, max_edge)
            hull_cr = _decimate(cr_pts, 0.005, max_edge); hull_cl = _decimate(cl_pts, 0.005, max_edge)
            hull_rr = _decimate(rr_pts, 0.005, max_edge); hull_rl = _decimate(rl_pts, 0.005, max_edge)
            # Fill in hull-decimated trailer lines now that max_edge is known.
            for td in all_trailers:
                raw = td.pop('_hull_raw')
                keys = ['hull_ar','hull_al','hull_mr','hull_ml','hull_cr','hull_cl','hull_rr','hull_rl']
                for k, raw_pts in zip(keys, raw):
                    td[k] = _decimate(raw_pts, 0.005, max_edge)
            hull_polylines = [hull_ar, hull_al, hull_mr, hull_ml,
                              hull_cr, hull_cl, hull_rr, hull_rl]
            for td in trailer_subset:
                for key in ['hull_ar','hull_al','hull_mr','hull_ml','hull_cr','hull_cl','hull_rr','hull_rl']:
                    hull_polylines.append(td[key])
            rings, hull_diag = _alpha_hull(hull_polylines, max_edge)
            env_polys = []
            for i, ring in enumerate(rings):
                sfx = "" if len(rings) == 1 else " {}".format(i + 1)
                pts = decimate(ring)
                if len(pts) >= 2:
                    env_polys.append((pts, sfx))

            # ── Minimum path radius ───────────────────────────────────────────
            min_radius = float('inf')
            for i in range(n_pts - 1):
                dh = headings[i + 1] - headings[i]
                if dh >  math.pi: dh -= 2.0 * math.pi
                if dh < -math.pi: dh += 2.0 * math.pi
                if abs(dh) > 1e-9:
                    r = step_size / abs(dh)
                    if r < min_radius:
                        min_radius = r

            if min_radius < float('inf'):
                rmsg = io.StringIO()
                rmsg.write("Min path radius: {:.1f}m\n".format(min_radius))
                vwords = set(w for w in v_name.upper().replace('-', ' ').split() if len(w) > 1)
                candidates = []
                for key, (r_abs, r_des) in _RADIUS_TABLE.items():
                    kwords = set(w for w in key.upper().replace('-', ' ').split() if len(w) > 1)
                    overlap = len(vwords & kwords)
                    if overlap > 0:
                        candidates.append((overlap, key, r_abs, r_des))
                candidates.sort(key=lambda x: -x[0])
                for _, key, r_abs, r_des in candidates:
                    if min_radius >= r_des:
                        status = "OK (>= desirable {:.0f}m)".format(r_des)
                    elif min_radius >= r_abs:
                        status = "Below desirable {:.0f}m, above absolute {:.0f}m".format(r_des, r_abs)
                    else:
                        status = "BELOW absolute minimum {:.0f}m".format(r_abs)
                    rmsg.write("  {}: {}\n".format(key, status))
                self.radius_info.Text = rmsg.getvalue()
            else:
                self.radius_info.Text = ""
            if mta_events:
                mmsg = io.StringIO()
                for (ua, ub), ch in sorted(mta_events.items(), key=lambda x: x[1]):
                    mmsg.write("MTA clamped: Unit {} - Unit {} @ {:.1f}m\n".format(ua + 1, ub + 1, ch))
                self.error.Text = mmsg.getvalue()
            else:
                self.error.Text = ""

            self.sim_cache = {
                'key':        (line.SerialNumber, v_name, side, step_size, tp_buffer, use_layers,
                               bool(self.reverse_line.IsChecked)),
                'ar':         dec_ar,    'al':    dec_al,
                'mr':         dec_mr,    'ml':    dec_ml,
                'cr':         dec_cr,    'cl':    dec_cl,
                'rr':         dec_rr,    'rl':    dec_rl,
                'hull_ar':    hull_ar,   'hull_al': hull_al,
                'hull_mr':    hull_mr,   'hull_ml': hull_ml,
                'hull_cr':    hull_cr,   'hull_cl': hull_cl,
                'hull_rr':    hull_rr,   'hull_rl': hull_rl,
                'max_edge':   max_edge,
                'ar_pts':     ar_pts,    'al_pts': al_pts,   # full-res, for outline step-indexing
                'rl_pts':     rl_pts,    'rr_pts': rr_pts,
                'trailers':   all_trailers,
                'body_length': _body_length(ar_pts, al_pts,
                                            all_trailers[-1]['rr'] if all_trailers else rr_pts),
                'env_polys':  env_polys,
                'unit_tracks': unit_tracks,
                'path_pts':   path_pts,
                'headings':   headings,
                'units':      units,
                'pfx':        pfx,
                'step_size':  step_size,
                'n_pts':      n_pts,
            }

        # ── Drawing from cache ────────────────────────────────────────────────
        def draw_from_cache():
            """Emit all linework / overlay geometry from the current simulation cache."""
            c   = self.sim_cache
            pfx = c['pfx']
            ss  = c['step_size']
            n   = c['n_pts']
            ar_r = c['ar_pts']; al_r = c['al_pts']   # full-res for outline indexing
            rr_r = c['rr_pts']; rl_r = c['rl_pts']
            first_last_only = bool(self.first_last_only.IsChecked)
            last_tj = len(c['units']) - 1

            # ── PM track lines ────────────────────────────────────────────────
            c_fc = self.color_front_corners.SelectedColor
            c_ms = self.color_mid_sweep.SelectedColor
            c_ra = self.color_rear_axle.SelectedColor
            c_rc = self.color_rear_corners.SelectedColor
            if bool(self.draw_front_corners.IsChecked):
                make_ls(pfx + " (AR)", c['ar'], layer("AR", 1), color=c_fc)
                make_ls(pfx + " (AL)", c['al'], layer("AL", 1), color=c_fc)
            if bool(self.draw_mid_sweep.IsChecked):
                make_ls(pfx + " (MR)", c['mr'], layer("MR", 1), color=c_ms)
                make_ls(pfx + " (ML)", c['ml'], layer("ML", 1), color=c_ms)
            if bool(self.draw_rear_axle.IsChecked):
                make_ls(pfx + " (CR)", c['cr'], layer("CR", 1), color=c_ra)
                make_ls(pfx + " (CL)", c['cl'], layer("CL", 1), color=c_ra)
            if bool(self.draw_rear_corners.IsChecked):
                make_ls(pfx + " (RR)", c['rr'], layer("RR", 1), color=c_rc)
                make_ls(pfx + " (RL)", c['rl'], layer("RL", 1), color=c_rc)

            # ── Trailer track lines ───────────────────────────────────────────
            drawn_trailers = []
            for td in c['trailers']:
                tj = td['tj']
                if first_last_only and tj < last_tj:
                    continue
                tpfx = pfx + " T{}".format(tj)
                if bool(self.draw_front_corners.IsChecked):
                    make_ls(tpfx + " (AR)", td['ar'], layer("AR", tj + 1), color=c_fc)
                    make_ls(tpfx + " (AL)", td['al'], layer("AL", tj + 1), color=c_fc)
                if bool(self.draw_mid_sweep.IsChecked):
                    make_ls(tpfx + " (MR)", td['mr'], layer("MR", tj + 1), color=c_ms)
                    make_ls(tpfx + " (ML)", td['ml'], layer("ML", tj + 1), color=c_ms)
                if bool(self.draw_rear_axle.IsChecked):
                    make_ls(tpfx + " (CR)", td['cr'], layer("CR", tj + 1), color=c_ra)
                    make_ls(tpfx + " (CL)", td['cl'], layer("CL", tj + 1), color=c_ra)
                if bool(self.draw_rear_corners.IsChecked):
                    make_ls(tpfx + " (RR)", td['rr'], layer("RR", tj + 1), color=c_rc)
                    make_ls(tpfx + " (RL)", td['rl'], layer("RL", tj + 1), color=c_rc)
                drawn_trailers.append(td)

            # ── Envelope ──────────────────────────────────────────────────────
            if bool(self.draw_envelope.IsChecked):
                if use_layers:
                    env_lyr = Layer.FindOrCreateLayer(self.currentProject, lyr_prefix + " Swept Envelope")
                    if env_lyr.LayerGroupSerial == 0 and lyr_group is not None:
                        env_lyr.LayerGroupSerial = lyr_group.SerialNumber
                    layer_env = env_lyr.SerialNumber
                    draw_hull_polylines = [c['hull_ar'], c['hull_al'], c['hull_mr'], c['hull_ml'],
                                           c['hull_cr'], c['hull_cl'], c['hull_rr'], c['hull_rl']]
                    for td in drawn_trailers:
                        for key in ['hull_ar','hull_al','hull_mr','hull_ml','hull_cr','hull_cl','hull_rr','hull_rl']:
                            draw_hull_polylines.append(td[key])
                    draw_rings, _ = _alpha_hull(draw_hull_polylines, c['max_edge'])
                    draw_env_polys = []
                    for i, ring in enumerate(draw_rings):
                        sfx = "" if len(draw_rings) == 1 else " {}".format(i + 1)
                        pts = decimate(ring, 0.005)
                        if len(pts) >= 2:
                            draw_env_polys.append((pts, sfx))
                    env_polys_out = draw_env_polys
                else:
                    layer_env    = None
                    env_polys_out = c['env_polys']
                for bps_pts, sfx in env_polys_out:
                    make_ls(pfx + " Envelope" + sfx, bps_pts, layer_env,
                            color=self.color_envelope.SelectedColor, weight=150, closed=True)

            # ── Body outline footprints ───────────────────────────────────────
            if self.draw_outline.IsChecked:
                interval   = max(float(self.outline_interval.Distance), ss)
                path_len   = (n - 1) * ss
                ol_indices = [0]
                pos = interval
                while pos < path_len:
                    ol_indices.append(min(int(round(pos / ss)), n - 1))
                    pos += interval
                if ol_indices[-1] != n - 1:
                    ol_indices.append(n - 1)
                overall_length = c['body_length']
                fo = c['units'][0]['FO']
                threshold = path_len + fo - overall_length
                while len(ol_indices) >= 2 and (ol_indices[-2] + 1) * ss + fo >= threshold:
                    del ol_indices[-2]

                def apply_tmpl(tmpl, cx, cy, co, si, z):
                    return [(cx + lx*co - ly*si, cy + lx*si + ly*co, z) for lx, ly in tmpl]

                def make_circle(cx, cy, z, r=0.05):
                    if use_layers:
                        circ = wv.Add(clr.GetClrType(CadCircle))
                        circ.Layer = layer_ol
                        circ.CenterPoint = Point3D(cx, cy, z)
                        circ.Radius = r
                    else:
                        N   = 16
                        pts = Array[Point3D]([Point3D(cx + r * math.cos(2*math.pi*k/N),
                                                      cy + r * math.sin(2*math.pi*k/N), z)
                                              for k in range(N + 1)])
                        self.overlayBag.AddPolyline(pts, self.color_outline.SelectedColor.ToArgb(), 1)

                for i in ol_indices:
                    dist_lbl = "{:.0f}m".format(i * ss)
                    z = ar_r[i][2]
                    make_ls(pfx + " Outline@" + dist_lbl,
                             [ar_r[i], al_r[i], rl_r[i], rr_r[i], ar_r[i]], layer("Outline", 1), color=self.color_outline.SelectedColor)
                    for j, utrk in enumerate(c['unit_tracks']):
                        if i >= len(utrk):
                            continue
                        uj       = c['units'][j]
                        layer_ol = layer("Outline", j + 1)
                        fax, fay, rax, ray, bh = utrk[i]
                        bc  = math.cos(bh); bs = math.sin(bh)
                        ulbl = pfx + " U{} @{}".format(j, dist_lbl)
                        if j == 0:
                            dkp   = uj['DKP']; dt = uj['DT']
                            kpr_x = fax + dkp*bs;  kpr_y = fay - dkp*bc
                            kpl_x = fax - dkp*bs;  kpl_y = fay + dkp*bc
                            make_ls(ulbl + " FA", [(kpr_x, kpr_y, z), (kpl_x, kpl_y, z)], layer_ol, color=self.color_outline.SelectedColor)
                            make_circle(kpr_x, kpr_y, z)
                            make_circle(kpl_x, kpl_y, z)
                            fh   = c['headings'][i]
                            fc   = math.cos(fh); fs = math.sin(fh)
                            ftw2 = uj['FTW'] / 2.0
                            tmpl_f = [(-uj['FTD']/2, -ftw2), (uj['FTD']/2, -ftw2),
                                       (uj['FTD']/2,   ftw2), (-uj['FTD']/2,  ftw2),
                                       (-uj['FTD']/2, -ftw2)]
                            f_lats = [dt - ftw2 - k * (uj['FTW'] + uj['FTS'])
                                       for k in range(max(1, uj['FTC']))]
                            for k, lat in enumerate(f_lats):
                                ks = str(k + 1)
                                make_ls(ulbl + " FTR" + ks, apply_tmpl(tmpl_f, kpr_x + lat*fs, kpr_y - lat*fc, fc, fs, z), layer_ol, color=self.color_outline.SelectedColor)
                                make_ls(ulbl + " FTL" + ks, apply_tmpl(tmpl_f, kpl_x - lat*fs, kpl_y + lat*fc, fc, fs, z), layer_ol, color=self.color_outline.SelectedColor)
                        else:
                            make_circle(fax, fay, z)
                            tfo  = uj['FO']; tw = uj['W']; ttr = uj['TR']
                            troh = uj['LR'] - uj['LA']
                            t_ar = (fax + bc*tfo + bs*tw,   fay + bs*tfo - bc*tw,   z)
                            t_al = (fax + bc*tfo - bs*tw,   fay + bs*tfo + bc*tw,   z)
                            t_rr = (rax - bc*troh + bs*ttr, ray - bs*troh - bc*ttr,  z)
                            t_rl = (rax - bc*troh - bs*ttr, ray - bs*troh + bc*ttr,  z)
                            make_ls(ulbl + " Outline", [t_ar, t_al, t_rl, t_rr, t_ar], layer_ol, color=self.color_outline.SelectedColor)
                        da = [0.0]
                        if uj['DA1'] > 0: da.append(da[-1] + uj['DA1'])
                        if uj['DA2'] > 0: da.append(da[-1] + uj['DA2'])
                        if uj['DA3'] > 0: da.append(da[-1] + uj['DA3'])
                        axle_params = [
                            (uj['ATC1'], uj['ATD1'], uj['ATW1'], uj['ATS1']),
                            (uj['ATC2'], uj['ATD2'], uj['ATW2'], uj['ATS2']),
                            (uj['ATC3'], uj['ATD3'], uj['ATW3'], uj['ATS3']),
                            (uj['ATC4'], uj['ATD4'], uj['ATW4'], uj['ATS4']),
                        ]
                        for ai, aoff in enumerate(da):
                            ax = rax - aoff * bc
                            ay = ray - aoff * bs
                            tr = uj['TR']
                            make_ls(ulbl + " RA{}".format(ai + 1),
                                     [(ax + tr*bs, ay - tr*bc, z),
                                      (ax - tr*bs, ay + tr*bc, z)], layer_ol, color=self.color_outline.SelectedColor)
                            atc, atd, atw, ats = axle_params[ai]
                            if atc < 1:
                                continue
                            atw2   = atw / 2.0
                            tmpl_r = [(-atd/2, -atw2), (atd/2, -atw2),
                                       (atd/2,   atw2), (-atd/2,  atw2),
                                       (-atd/2, -atw2)]
                            r_lats = [tr - atw2 - k * (atw + ats)
                                       for k in range(max(1, atc))]
                            for k, lat in enumerate(r_lats):
                                ks = str(k + 1)
                                make_ls(ulbl + " A{}R{}".format(ai+1, ks), apply_tmpl(tmpl_r, ax + lat*bs, ay - lat*bc, bc, bs, z), layer_ol, color=self.color_outline.SelectedColor)
                                make_ls(ulbl + " A{}L{}".format(ai+1, ks), apply_tmpl(tmpl_r, ax - lat*bs, ay + lat*bc, bc, bs, z), layer_ol, color=self.color_outline.SelectedColor)

        # ── Cache validation and dispatch ─────────────────────────────────────
        sim_key = (line.SerialNumber, v_name, side, step_size, tp_buffer, use_layers,
                   bool(self.reverse_line.IsChecked))
        cache_valid = self.sim_cache is not None and self.sim_cache.get('key') == sim_key

        def run_sim():
            ProgressBar.TBC_ProgressBar.Title = "Computing swept path..."
            try:
                run_simulation()
                draw_from_cache()
            finally:
                ProgressBar.TBC_ProgressBar.Title = ""

        if use_layers:
            committed = False
            UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
            self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
            try:
                with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                    lg_col    = LayerGroupCollection.GetLayerGroupCollection(self.currentProject, True)
                    lyr_group = lg_col.FindOrCreateLayerGroup(lyr_prefix)
                    run_sim()
                    failGuard.Commit()
                    committed = True
            except ValueError as ve:
                self.error.Text = str(ve)
                return
            except Exception:
                exc_type, exc_obj, exc_tb = sys.exc_info()
                self.error.Text = str(exc_type) + "\n" + str(exc_obj) + "\nLine " + str(exc_tb.tb_lineno)
                return
            finally:
                self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
                UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
            if committed:
                self.layer_prefix.Text = PointHelper.Helpers.Increment(self.layer_prefix.Text, None, True)
        else:
            TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)
            self.overlayBag = OverlayBag(self.ViewOverlay)
            if not self._show_line_direction_active():
                self._add_line_to_bag(line, self.overlayBag)
            try:
                if not cache_valid:
                    run_sim()
                else:
                    draw_from_cache()
            except ValueError as ve:
                self.error.Text = str(ve)
                return
            except Exception:
                exc_type, exc_obj, exc_tb = sys.exc_info()
                self.error.Text = str(exc_type) + "\n" + str(exc_obj) + "\nLine " + str(exc_tb.tb_lineno)
                return
            views = Array[Guid]([DisplayWindow.Hoops3DViewGUID, DisplayWindow.HoopsPlanViewGUID])
            TrimbleOffice.TheOffice.MainWindow.AppViewManager.AddOverlayGeometry(views, self.overlayBag)


# ── Pure-Python geometry helpers ──────────────────────────────────────────────
# Placed after the class so the TBC registration and UI code are at the top.
# Python resolves these names at call time, not at class-definition time,
# so the order has no effect on correctness.

def _angle_diff(a, b):
    """Signed angular difference a − b, normalised to (−π, +π]."""
    d = a - b
    return (d + math.pi) % (2 * math.pi) - math.pi


def _body_corners_2d(px, py, ph, unit):
    """World (x, y) of all four body corners for a unit centred at (px, py, ph).
    Returns [FL, FR, RR, RL] — front-left, front-right, rear-right, rear-left."""
    c, s = math.cos(ph), math.sin(ph)
    fo, w, lr = unit['FO'], unit['W'], unit['LR']
    return [(px + c*lx - s*ly, py + s*lx + c*ly)
            for lx, ly in [(fo, w), (fo, -w), (-lr, -w), (-lr, w)]]


def _hitch_2d(px, py, ph, unit):
    """World (x, y) of the tow-hitch of unit, given its front-axle (px, py) and heading ph."""
    lt = unit['LT']
    c, s = math.cos(ph), math.sin(ph)
    return (px - c*lt, py - s*lt)


def _step_trailer(ox, oy, oh, hx, hy, unit):
    """Advance one kinematic trailer step.

    The hitch point moves from (ox, oy) to (hx, hy).  The old heading oh is
    used to seed the new heading, which is then optionally clamped to MTA.

    Returns (hx, hy, new_heading, clamped).
    """
    # Bogie centre is mid-way through the axle group; it drives the kinematics.
    da_total = unit['DA1'] + unit['DA2'] + unit['DA3']
    la = unit['LA'] + da_total * 0.5

    # Pull the bogie centre onto a circle of radius LA centred on the new hitch.
    old_ra_x = ox - la * math.cos(oh)
    old_ra_y = oy - la * math.sin(oh)
    dx = old_ra_x - hx
    dy = old_ra_y - hy
    dist = math.sqrt(dx * dx + dy * dy)
    if dist > 1e-6:
        new_ra_x = hx + la * dx / dist
        new_ra_y = hy + la * dy / dist
    else:
        new_ra_x = old_ra_x
        new_ra_y = old_ra_y

    # New heading = direction from the new RA to the new hitch.
    nh = math.atan2(hy - new_ra_y, hx - new_ra_x)

    # Apply maximum turning angle (MTA=0 means unlimited).
    mta = unit.get('MTA', 90.0)
    clamped = False
    if mta > 0:
        diff = _angle_diff(nh, oh)
        if abs(diff) > math.radians(mta):
            nh = oh + math.copysign(math.radians(mta), diff)
            clamped = True

    return (hx, hy, nh, clamped)


def _parse_vehicle_units(v_elem):
    """Parse all <unit> children of a <Vehicle> XML element into a list of dicts."""
    # Top-level DKP / DT / MSA are defaults for unit 0 when not overridden per-unit.
    top_dkp = float(v_elem.findtext("DKP") or 0)
    top_dt  = float(v_elem.findtext("DT")  or 0)
    top_msa = float(v_elem.findtext("MSA") or 35)

    units = []
    for i, u in enumerate(v_elem.findall("unit")):
        def g(tag, default=0.0):
            el = u.find(tag)
            try:    return float(el.text) if el is not None else float(default)
            except: return float(default)
        def gi(tag, attr, default=0):
            el = u.find(tag)
            try:    return int(el.get(attr, str(default))) if el is not None else int(default)
            except: return int(default)
        def gf(tag, attr, default=0.0):
            el = u.find(tag)
            try:    return float(el.get(attr, str(default))) if el is not None else float(default)
            except: return float(default)

        units.append({
            # Steering / path offset for unit 0 (prime mover) only.
            'DKP': g("DKP", top_dkp) if i == 0 else 0.0,
            'DT':  g("DT",  top_dt)  if i == 0 else 0.0,
            # Geometry
            'MSA': g("MSA", top_msa),   # maximum steering angle (degrees)
            'MTA': g("MTA", 90.0),      # maximum turning angle between units (degrees)
            'FO':  g("FO"),   # front overhang from front axle
            'W':   g("W"),    # body half-width
            'TR':  g("TR"),   # rear track half-width (tyre outer edge)
            'LR':  g("LR"),   # rear overhang from rear axle to body rear
            'LT':  g("LT"),   # tow-hitch distance behind front axle
            'LA':  g("LA"),   # first rear axle distance behind front axle
            # Additional axle spacing within the axle group
            'DA1': g("DA1"), 'DA2': g("DA2"), 'DA3': g("DA3"),
            # Tyre footprint for outlines
            'TS1': g("TS1"), 'TS2': g("TS2"), 'TS3': g("TS3"), 'TS4': g("TS4"),
            'AL1': g("AL1"), 'AL2': g("AL2"), 'AL3': g("AL3"), 'AL4': g("AL4"),
            # Front tyre footprint attributes (FO element attributes)
            'FTC': gi("FO", "FTC"),  'FTD': gf("FO", "FTD"),
            'FTW': gf("FO", "FTW"), 'FTS': gf("FO", "FTS"),
            # Rear axle group tyre footprint attributes
            'ATC1': gi("LA",  "ATC1"), 'ATD1': gf("LA",  "ATD1"),
            'ATW1': gf("LA",  "ATW1"), 'ATS1': gf("LA",  "ATS1"),
            'ATC2': gi("DA1", "ATC2"), 'ATD2': gf("DA1", "ATD2"),
            'ATW2': gf("DA1", "ATW2"), 'ATS2': gf("DA1", "ATS2"),
            'ATC3': gi("DA2", "ATC3"), 'ATD3': gf("DA2", "ATD3"),
            'ATW3': gf("DA2", "ATW3"), 'ATS3': gf("DA2", "ATS3"),
            'ATC4': gi("DA3", "ATC4"), 'ATD4': gf("DA3", "ATD4"),
            'ATW4': gf("DA3", "ATW4"), 'ATS4': gf("DA3", "ATS4"),
        })
    return units


# ── Concave hull via TBC Gem (Gem.ModelType.Dtm + MaxOuterLength) ────────────
#
# Gem(Gem.ModelType.Dtm) triangulates in-memory without a worldview surface.
# gem.MaxOuterLength is the built-in shrinkwrap threshold: outer triangle edges
# longer than this value are removed during ConstructSurface(), producing the
# concave (alpha-shape) boundary.  gem.ExtractOutlines() returns the boundary
# as a flat Point3D list where every two consecutive points form one CCW edge
# segment (i.e. pairs (0,1), (2,3), …).  We chain those pairs into rings.
#     adj[0] ← across edge v1-v2
#     adj[1] ← across edge v2-v0
#     adj[2] ← across edge v0-v1
#   n_k = -1 means no neighbour (outer boundary edge)
#
# edge_map[(vA, vB)] = ti  stores which triangle owns directed edge vA→vB.
# A directed edge vA→vB is owned by the triangle on the LEFT of vA→vB.
# The adjacency lookup is therefore:
#   adj[f] of T = edge_map.get((T[(f+2)%3], T[(f+1)%3]), -1)
# (the reverse direction of the face-f edge)

def _decimate(pts, tol=0.025, max_seg=None):
    """Ramer-Douglas-Peucker line simplification (iterative, O(n log n)).

    max_seg: if set, re-enables original points so no output segment exceeds
             this length — prevents long straights losing all intermediate points.
    """
    n = len(pts)
    if n < 3:
        return list(pts)
    t2   = tol * tol
    keep = [True] * n
    stk  = [(0, n - 1)]
    while stk:
        s, e = stk.pop()
        if e - s < 2:
            continue
        sx = pts[s][0]; sy = pts[s][1]
        ex = pts[e][0]; ey = pts[e][1]
        dx = ex - sx;   dy = ey - sy
        d2 = dx*dx + dy*dy
        mx2 = 0; mi = s
        for i in range(s + 1, e):
            if not keep[i]:
                continue
            if d2 == 0:
                rx = pts[i][0] - sx; ry = pts[i][1] - sy
                cd2 = rx*rx + ry*ry
            else:
                tc = ((pts[i][0] - sx)*dx + (pts[i][1] - sy)*dy) / d2
                if tc < 0.0:  tc = 0.0
                elif tc > 1.0: tc = 1.0
                rx = sx + tc*dx - pts[i][0]
                ry = sy + tc*dy - pts[i][1]
                cd2 = rx*rx + ry*ry
            if cd2 > mx2:
                mx2 = cd2; mi = i
        if mx2 > t2:
            stk.append((s,  mi))
            stk.append((mi, e))
        else:
            for i in range(s + 1, e):
                keep[i] = False
    if max_seg is not None:
        ms2 = max_seg * max_seg
        prev = 0
        for i in range(1, n):
            if not keep[i]:
                continue
            dx = pts[i][0] - pts[prev][0]; dy = pts[i][1] - pts[prev][1]
            if dx*dx + dy*dy > ms2:
                accum = 0.0; last = prev
                for k in range(prev + 1, i):
                    dx2 = pts[k][0] - pts[last][0]; dy2 = pts[k][1] - pts[last][1]
                    accum += math.sqrt(dx2*dx2 + dy2*dy2)
                    if accum >= max_seg:
                        keep[k] = True; last = k; accum = 0.0
            prev = i
    return [pts[i] for i in range(n) if keep[i]]


def _body_length(ar_pts, al_pts, rear_right_pts):
    """Perpendicular distance from rear-right[0] to the front bumper line (al[0]→ar[0])."""
    flx, fly = al_pts[0][0], al_pts[0][1]
    frx, fry = ar_pts[0][0], ar_pts[0][1]
    rrx, rry = rear_right_pts[0][0], rear_right_pts[0][1]
    bdx = frx - flx; bdy = fry - fly
    bl  = math.sqrt(bdx*bdx + bdy*bdy)
    if bl == 0:
        return 0.0
    vx = rrx - flx; vy = rry - fly
    return abs(vx*bdy - vy*bdx) / bl


def _alpha_hull(polylines, max_edge):
    """
    Concave hull of a set of polylines using TBC's Gem Delaunay triangulation.

    polylines : list of polylines, each a list of (x, y, z) tuples.
                Points within each polyline are connected as breaklines so the
                Gem respects the polyline topology.
    max_edge  : Gem.MaxOuterLength — outer triangle edges longer than this are
                trimmed away during ConstructSurface(), producing the concave hull.
    Returns (rings, diag).
    rings : list of closed rings, each a list of (x, y, z) tuples.
    diag  : dict with diagnostic counts for the timing display.
    """
    n_pts = sum(len(pl) for pl in polylines)
    diag  = {'pts_in': n_pts, 'n_tri': 0, 'outline_pts': 0, 'rings': 0}

    if n_pts < 3:
        return [], diag

    gem = Gem(Gem.ModelType.Dtm)
    gem.MaxOuterLength = max_edge

    for polyline in polylines:
        iv_prev = -1
        for (x, y, z) in polyline:
            pt      = Point3D(x, y, 1.0)
            iv, _   = gem.AddVertex(False, GemVertexType.Control, 0, pt)
            if iv_prev >= 0:
                gem.AddBreakline(False, 0, iv_prev, iv)
            iv_prev = iv

    diag['pre_v']  = gem.NumberOfVertices
    diag['pre_bl'] = gem.NumberOfBreaklines

    gem.ConstructSurface()

    diag['n_tri']         = gem.NumberOfTriangles
    diag['n_tri_present'] = gem.NumberOfTrianglesPresent
    diag['n_v_present']   = gem.NumberOfVerticesPresent
    diag['max_outer']     = gem.MaxOuterLength

    # ExtractOutlines returns a flat List<Point3D> where consecutive pairs
    # form CCW outer edge segments: (pts[0],pts[1]), (pts[2],pts[3]), ...
    # Chain overlapping endpoints into closed rings.
    outlines = gem.ExtractOutlines()
    pts_list = list(outlines)
    diag['outline_pts'] = len(pts_list)

    if len(pts_list) < 4:
        return [], diag

    # Build edge adjacency: (x,y) rounded to mm → next (x,y)
    def key(p):
        return (round(p.X, 3), round(p.Y, 3))

    nxt = {}
    for i in range(0, len(pts_list) - 1, 2):
        a = pts_list[i]; b = pts_list[i + 1]
        nxt[key(a)] = (b.X, b.Y, b.Z)

    # Walk into closed rings
    rings   = []
    visited = set()
    for start_key in list(nxt.keys()):
        if start_key in visited:
            continue
        ring = []; cur_key = start_key
        while cur_key not in visited and cur_key in nxt:
            visited.add(cur_key)
            x, y, z = nxt[cur_key]
            ring.append((x, y, 0.0))
            cur_key = key(pts_list[0].__class__(x, y, z))  # re-derive key
            # simpler: key from the coords we just stored
            cur_key = (round(x, 3), round(y, 3))
        if len(ring) >= 3:
            rings.append(ring)

    diag['rings'] = len(rings)
    return rings, diag


def _simulate_sweep(path_pts, headings, units, kappa_max=None,
                    tp_buffer=False, step_size=0.5, side=0, progress_fn=None):
    """Ackermann + kinematic-chain swept-path simulation.

    path_pts  : list of (x, y, z) sampled from the reference line via FindPointFromStation.
                Must be spaced exactly step_size apart.
    headings  : list of float — tangent heading at each sample (radians, world frame).
    units     : list of unit dicts from _parse_vehicle_units().
    tp_buffer : if True, zero curvature over ±6 m at any reverse curve tighter than R40.
    step_size : sample spacing along the path (metres).
    side      : 0=left-kerb reference, 1=centreline, 2=right-kerb reference.
    progress_fn : optional callable(pct) for progress bar updates.

    The prime mover (unit 0) uses exact Ackermann steering:
      body_heading = atan2(FA_y − RA_y, FA_x − RA_x)
    iterated to convergence so the path-tracking offset (DKP/DT) is exact.

    Each trailer unit is advanced kinematically: its bogie is pulled to remain
    on a circle of radius LA centred on the new hitch point.

    Returns a 10-tuple:
      ar_pts, al_pts  — PM front-right / front-left body corners
      cr_pts, cl_pts  — PM rear-axle right / left tyre edges
      rr_pts, rl_pts  — PM rear-body right / left corners
      mr_pts, ml_pts  — PM body mid-right / mid-left (midpoint of front and rear corners)
      unit_tracks     — list of per-unit [(fa_x, fa_y, ra_x, ra_y, heading), ...] for all steps
      mta_events      — dict {(unit_a, unit_b): first_chainage} for MTA clamp occurrences
    All point lists contain world (x, y, z) tuples, one per path sample.
    """
    n = len(path_pts)
    if n < 2 or not units:
        return [], [], [], [], [], [], [], [], [], []

    u0       = units[0]
    TR0      = u0['TR']
    LA0      = u0['LA']                                               # first rear axle, for display
    LA0_eff  = LA0 + (u0['DA1'] + u0['DA2'] + u0['DA3']) * 0.5      # bogie centre, for kinematics
    FO0      = u0['FO']
    W0       = u0['W']
    LR0      = u0['LR']
    rear_oh  = LR0 - LA0   # body overhang behind the first rear axle

    # ── Path curvature κ at each step (centred finite difference of heading) ──
    kappas = [0.0] * n
    for i in range(1, n - 1):
        dh = _angle_diff(headings[i + 1], headings[i - 1])
        dx = path_pts[i + 1][0] - path_pts[i - 1][0]
        dy = path_pts[i + 1][1] - path_pts[i - 1][1]
        ds = math.sqrt(dx * dx + dy * dy)
        kappas[i] = dh / ds if ds > 1e-9 else 0.0
    kappas[0]  = kappas[1]  if n > 2 else 0.0
    kappas[-1] = kappas[-2] if n > 2 else 0.0

    # ── TP buffer: blank curvature over ±6 m at reverse curves with R < 40 m ──
    KAPPA_THRESHOLD = 1.0 / 40.0
    TP_BUF_METRES   = 6.0
    if tp_buffer and step_size > 0.0:
        half_buf = max(1, int(round(TP_BUF_METRES / step_size / 2.0)))
        i = 1
        while i < n:
            k_prev = kappas[i - 1]
            k_curr = kappas[i]
            sign_change  = (k_prev > 0 and k_curr < 0) or (k_prev < 0 and k_curr > 0)
            tight_enough = abs(k_prev) > KAPPA_THRESHOLD or abs(k_curr) > KAPPA_THRESHOLD
            if sign_change and tight_enough:
                lo = max(0, i - half_buf)
                hi = min(n, i + half_buf)
                for j in range(lo, hi): kappas[j] = 0.0
                i = hi
            else:
                i += 1

    # ── FA recovery from path point ───────────────────────────────────────────
    # The input path can be offset from the PM front axle by DKP (body-perpendicular)
    # and DT (wheel-perpendicular).  Body heading is unknown until FA is placed,
    # so we iterate convergence using the previous step's body heading as a seed.
    # On straights this is exact; on bends the one-step lag is negligible.
    dt_sign = (1.0, 0.0, -1.0)[side]
    dkp     = u0['DKP']
    dt      = u0['DT']

    def fa_from_path(px, py, ph, bh):
        return (px - dt_sign * (dkp * math.sin(bh) + dt * math.sin(ph)),
                py + dt_sign * (dkp * math.cos(bh) + dt * math.cos(ph)))

    # ── Initialise PM and trailer chain ──────────────────────────────────────
    h_init   = headings[0]
    fa0_x, fa0_y = fa_from_path(path_pts[0][0], path_pts[0][1], h_init, h_init)
    ra_x = fa0_x - LA0_eff * math.cos(h_init)
    ra_y = fa0_y - LA0_eff * math.sin(h_init)
    body_h_prev = h_init

    # Seed trailer chain straight behind the PM along the initial heading.
    trailer_states = []
    prev = (fa0_x, fa0_y, h_init)
    for i in range(1, len(units)):
        hx, hy = _hitch_2d(prev[0], prev[1], prev[2], units[i - 1])
        trailer_states.append((hx, hy, h_init))
        prev = trailer_states[-1]

    ar_pts = []; al_pts = []; cr_pts = []; cl_pts = []
    rr_pts = []; rl_pts = []; mr_pts = []; ml_pts = []
    unit_tracks = [[] for _ in units]
    mta_events  = {}   # {(unit_a, unit_b): first chainage where MTA was clamped}

    prog_t = time.time()
    for step in range(n):
        if progress_fn is not None:
            now = time.time()
            if now - prog_t >= 1.0:
                prog_t = now
                progress_fn(step * 100 // n)

        z   = path_pts[step][2]
        px  = path_pts[step][0]
        py  = path_pts[step][1]
        ph  = headings[step]

        # Converge body heading and FA with 4 Newton rounds (error < 0.01 mm).
        body_h = body_h_prev
        for it in range(4):
            fa_x, fa_y = fa_from_path(px, py, ph, body_h)
            body_h = math.atan2(fa_y - ra_y, fa_x - ra_x)
        fa_x, fa_y = fa_from_path(px, py, ph, body_h)

        # Bogie-centre position is the kinematic seed for the next iteration.
        # First-rear-axle position is used for display (corner and outline drawing).
        ra_x      = fa_x - LA0_eff * math.cos(body_h)
        ra_y      = fa_y - LA0_eff * math.sin(body_h)
        ra_disp_x = fa_x - LA0 * math.cos(body_h)
        ra_disp_y = fa_y - LA0 * math.sin(body_h)
        c = math.cos(body_h); s = math.sin(body_h)

        # ── Advance trailer chain ─────────────────────────────────────────────
        if trailer_states:
            chainage = step * step_size
            hitch_x, hitch_y = _hitch_2d(fa_x, fa_y, body_h, units[0])
            new_ts = []
            ox, oy, oh = trailer_states[0]
            hx, hy, nh, clamped = _step_trailer(ox, oy, oh, hitch_x, hitch_y, units[1])
            if clamped and (0, 1) not in mta_events:
                mta_events[(0, 1)] = chainage
            new_ts.append((hx, hy, nh))
            for i in range(1, len(trailer_states)):
                px2, py2, ph2 = new_ts[-1]
                nhx, nhy      = _hitch_2d(px2, py2, ph2, units[i])
                ox, oy, oh    = trailer_states[i]
                hx, hy, nh, clamped = _step_trailer(ox, oy, oh, nhx, nhy, units[i + 1])
                if clamped and (i, i + 1) not in mta_events:
                    mta_events[(i, i + 1)] = chainage
                new_ts.append((hx, hy, nh))
            trailer_states = new_ts

        # ── Record per-unit track data ────────────────────────────────────────
        unit_tracks[0].append((fa_x, fa_y, ra_disp_x, ra_disp_y, body_h))
        for jt, ts in enumerate(trailer_states):
            th = ts[2]; tc = math.cos(th); ts_s = math.sin(th)
            unit_tracks[jt + 1].append((ts[0], ts[1],
                                        ts[0] - units[jt + 1]['LA'] * tc,
                                        ts[1] - units[jt + 1]['LA'] * ts_s,
                                        th))

        # ── PM swept-path corner points ───────────────────────────────────────
        # Front corners use the front-axle position and body heading.
        ar_pts.append((fa_x      + c*FO0 + s*W0,    fa_y      + s*FO0 - c*W0,    z))
        al_pts.append((fa_x      + c*FO0 - s*W0,    fa_y      + s*FO0 + c*W0,    z))
        # Rear axle tyre edges use the kinematic bogie-centre position.
        cr_pts.append((ra_x      + s*TR0,            ra_y      - c*TR0,            z))
        cl_pts.append((ra_x      - s*TR0,            ra_y      + c*TR0,            z))
        # Rear body corners use the display-axle position and full body width W0.
        rr_pts.append((ra_disp_x - c*rear_oh + s*W0, ra_disp_y - s*rear_oh - c*W0, z))
        rl_pts.append((ra_disp_x - c*rear_oh - s*W0, ra_disp_y - s*rear_oh + c*W0, z))
        # Mid-body points are the geometric midpoint of the front and rear body corners.
        mx = (fa_x + ra_disp_x) * 0.5; my = (fa_y + ra_disp_y) * 0.5
        mr_pts.append((mx + s*W0, my - c*W0, z))
        ml_pts.append((mx - s*W0, my + c*W0, z))

        body_h_prev = body_h

    return ar_pts, al_pts, cr_pts, cl_pts, rr_pts, rl_pts, mr_pts, ml_pts, unit_tracks, mta_events

# names (Point3D, clr, PolySeg, etc.) are injected from SCR_Imports.py via __dict__.update at load time

class SCROverlayBag:
    
    def getarrowlocations(l1, intervals):
        pts = []
        if l1 is None:
            return pts

        if l1.Normal.Horizon < math.pi / 2 - 1e-9:
            # UCS line: flatten, linearise, transform chord back, walk segments for position + direction
            orgNormal = l1.Normal.Clone()
            centerp   = l1.OriginOfUcs
            nv        = orgNormal.Clone()
            vx        = nv.Clone()
            vx.RotateAboutZ(math.pi / 2)
            vx.Horizon = 0
            rottozero          = Spinor3D.ComputeRotation(vx, nv, Vector3D(1, 0, 0), Vector3D(0, 0, 1))
            matrixtoflat       = Matrix4D.BuildTransformMatrix(Vector3D(centerp), Vector3D(0, 0, 0), rottozero, Vector3D(1, 1, 1))
            matrixbackfromflat = Matrix4D.Inverse(matrixtoflat)
            l1.Normal = Vector3D(0, 0, 1)
            try:
                polyseg   = l1.ComputePolySeg().Clone()
                polyseg_v = l1.ComputeVerticalPolySeg()
                chord     = polyseg.Linearize(0.001, 0.001, 50, polyseg_v, False)
                chord.Transform(matrixbackfromflat)
                world_pts = chord.ToPoint3DArray()
            finally:
                l1.Normal = orgNormal

            if world_pts is None or len(world_pts) < 2:
                return pts

            segs = []
            cum  = 0.0
            for i in range(len(world_pts) - 1):
                p0, p1 = world_pts[i], world_pts[i + 1]
                dx, dy  = p1.X - p0.X, p1.Y - p0.Y
                slen    = math.sqrt(dx * dx + dy * dy)
                if slen > 1e-9:
                    segs.append((cum, p0, p1, dx / slen, dy / slen, slen))
                    cum += slen

            if not segs or cum < 1e-9:
                return pts

            step = cum / intervals
            d    = step * 0.5
            si   = 0
            for _ in range(intervals):
                while si < len(segs) - 1 and segs[si][0] + segs[si][5] <= d:
                    si += 1
                seg_start, p0, p1, ux, uy, slen = segs[si]
                t  = d - seg_start
                mp = Point3D(p0.X + ux * t, p0.Y + uy * t, p0.Z + (p1.Z - p0.Z) * (t / slen))
                pts.Add([mp, math.atan2(uy, -ux)])   # atan2(uy,-ux) = right-perp azimuth; caller uses math.pi - this
                d += step

        else:
            # flat line: use FindPointFromStation directly
            polyseg = l1.ComputePolySeg().Clone()
            polyseg = polyseg.ToWorld()
            polyseg_v = l1.ComputeVerticalPolySeg()
            if polyseg_v:
                polyseg_v = polyseg_v.Clone()

            interval       = polyseg.ComputeStationing() / intervals
            computestation = interval * 0.5
            for _ in range(intervals):
                outSegment    = clr.StrongBox[Segment]()
                out_t         = clr.StrongBox[float]()
                outPointOnCL  = clr.StrongBox[Point3D]()
                perpVector3D  = clr.StrongBox[Vector3D]()
                outdeflection = clr.StrongBox[float]()

                polyseg.FindPointFromStation(computestation, outSegment, out_t, outPointOnCL, perpVector3D, outdeflection)
                p = outPointOnCL.Value
                if polyseg_v is not None:
                    p.Z = polyseg_v.ComputeVerticalSlopeAndGrade(computestation)[1]

                pts.Add([p, perpVector3D.Value.Azimuth])
                computestation += interval

        return pts

    def getpolypoints(l):
        if l is None:
            return Array[Point3D]([])

        orgNormal = None
        matrixbackfromflat = None
        if l.Normal.Horizon < math.pi / 2 - 1e-9:
            orgNormal = l.Normal.Clone()
            centerp = l.OriginOfUcs
            nv = orgNormal.Clone()
            vx = nv.Clone()
            vx.RotateAboutZ(math.pi / 2)
            vx.Horizon = 0
            rottozero = Spinor3D.ComputeRotation(vx, nv, Vector3D(1, 0, 0), Vector3D(0, 0, 1))
            matrixtoflat = Matrix4D.BuildTransformMatrix(Vector3D(centerp), Vector3D(0, 0, 0), rottozero, Vector3D(1, 1, 1))
            matrixbackfromflat = Matrix4D.Inverse(matrixtoflat)
            l.Normal = Vector3D(0, 0, 1)

        try:
            polyseg = l.ComputePolySeg().Clone()
            polyseg = polyseg.ToWorld()
            polyseg_v = l.ComputeVerticalPolySeg()
            if polyseg_v:
                polyseg_v = polyseg_v.Clone()
            elif not polyseg_v and not polyseg.AllPointsAre3D:
                polyseg_v = PolySeg.PolySeg()
                polyseg_v.Add(Point3D(polyseg.BeginStation, 0, 0))
                polyseg_v.Add(Point3D(polyseg.ComputeStationing(), 0, 0))
            chord = polyseg.Linearize(0.001, 0.001, 50, polyseg_v, False)
            if matrixbackfromflat is not None:
                chord.Transform(matrixbackfromflat)
        finally:
            if orgNormal is not None:
                l.Normal = orgNormal

        return chord.ToPoint3DArray()

    # need to Clone the Polyseg upon creation
    # otherwise something weird is going on
    # and the functions above (getpoly and arrow) may return clipped sections
    # as if the ClipStationRange somehow transpires through to the original Linestring
    # although the Linestring stays unchanged on screen
    # no clue what is going on
    def getclippedpolypoints(l, startstation, endstation):
        
        chord = None
        polyseg = None
        polyseg_v = None

        if l != None and not math.isnan(startstation) and not math.isnan(endstation):
            polyseg = l.ComputePolySeg().Clone()
            polyseg = polyseg.ToWorld()
            
            if polyseg:
                if math.isnan(endstation):
                    endstation = polyseg.ComputeStationing()
                polyseg.ClipStationRange(startstation, endstation, True)
                polyseg.Trim()
                
                polyseg_v = l.ComputeVerticalPolySeg()
                if polyseg_v:
                    polyseg_v = polyseg_v.Clone()
                elif not polyseg_v and not polyseg.AllPointsAre3D:
                    polyseg_v = PolySeg.PolySeg()
                    polyseg_v.Add(Point3D(polyseg.BeginStation,0,0))
                    polyseg_v.Add(Point3D(polyseg.ComputeStationing(), 0, 0))

                minmax = CommandHelper.GetMinMaxElevations(l)
                # MinMax is used on the original object; if it doesn't have elevations we'll get NaN back
                if math.isnan(minmax[1]): lmin = 0
                else:                     lmin = minmax[1]
                if math.isnan(minmax[2]): lmax = 0
                else:                     lmax = minmax[2]
                polyseg_v.Clip(Limits3D(Point3D(startstation, lmin-1, 0), Point3D(endstation, lmax+1, 0)), Side.Out)
                polyseg_v.Trim()

                chord = polyseg.Linearize(0.001, 0.001, 50, polyseg_v, False)

        if chord:
            return chord.ToPoint3DArray()
        else:
            return Array[Point3D]([])


class SCREntityPicker:

    @staticmethod
    def _find_inner(parent):
        for ctrl in parent.Controls:
            if type(ctrl).__name__ in ('CheckedListBox', 'ListBox', 'ListView'):
                return ctrl
            found = SCREntityPicker._find_inner(ctrl)
            if found is not None:
                return found
        return None

    @staticmethod
    def _on_selection_changed(sender, e):
        state = sender.Tag
        if not isinstance(state, dict) or not state.get('_scr'):
            return
        upd = state['upd']
        if upd[0]:
            return
        upd[0] = True
        try:
            from System.Windows.Forms import Control, Keys
            mods = Control.ModifierKeys
            ctrl_held = (int(mods) & int(Keys.Control)) != 0
            shift_held = (int(mods) & int(Keys.Shift)) != 0
            hl = state['hl']
            sel = state['sel']
            all_items = list(sender.Items)
            clicked = list(sender.SelectedItems)
            if not clicked:
                return

            if shift_held:
                target = sender.FocusedItem or clicked[-1]
                anchor_text = state.get('anchor')
                try:
                    anchor_idx = next(i for i, it in enumerate(all_items) if it.Text == anchor_text)
                except StopIteration:
                    anchor_idx = 0
                try:
                    target_idx = next(i for i, it in enumerate(all_items) if it is target)
                except StopIteration:
                    target_idx = len(all_items) - 1
                if not ctrl_held:
                    for item in all_items:
                        item.BackColor = sender.BackColor
                        item.ForeColor = Color.Black
                    state['sel'] = set()
                    sel = state['sel']
                lo, hi = min(anchor_idx, target_idx), max(anchor_idx, target_idx)
                for i in range(lo, hi + 1):
                    all_items[i].BackColor = hl
                    all_items[i].ForeColor = Color.White
                    sel.add(all_items[i].Text)

            elif ctrl_held:
                for item in clicked:
                    if item.Text in sel:
                        item.BackColor = sender.BackColor
                        item.ForeColor = Color.Black
                        sel.discard(item.Text)
                    else:
                        item.BackColor = hl
                        item.ForeColor = Color.White
                        sel.add(item.Text)
                state['anchor'] = clicked[-1].Text

            else:
                for item in all_items:
                    item.BackColor = sender.BackColor
                    item.ForeColor = Color.Black
                state['sel'] = set()
                sel = state['sel']
                for item in clicked:
                    item.BackColor = hl
                    item.ForeColor = Color.White
                    sel.add(item.Text)
                state['anchor'] = clicked[-1].Text

            for item in clicked:
                item.Selected = False
        finally:
            upd[0] = False

    @staticmethod
    def reapply_highlights(innerlist):
        if innerlist is None:
            return
        state = innerlist.Tag
        if not isinstance(state, dict) or not state.get('_scr'):
            return
        hl = state['hl']
        sel = state['sel']
        for item in innerlist.Items:
            if item.Text in sel:
                item.BackColor = hl
                item.ForeColor = Color.White
            else:
                item.BackColor = innerlist.BackColor
                item.ForeColor = Color.Black

    @staticmethod
    def get_selected_serials(picker, innerlist):
        if innerlist is None:
            return []
        state = innerlist.Tag
        if not isinstance(state, dict) or not state.get('_scr'):
            return []
        sel = state['sel']
        nameToSN = {kvp.Key.Description: kvp.Value for kvp in picker.EntitySerialNumbers}
        return [nameToSN[t] for t in sel if t in nameToSN]

    @staticmethod
    def _setup(host, search_container, entity_type, multi_select, backcolor, highlight_color):
        from Trimble.Vce.UI.Controls import CheckedListBoxEntityPicker as WFPicker
        picker = WFPicker()
        picker.ShowCheckBoxes = False
        host.Child = picker
        picker.SearchContainer = search_container
        picker.UseSelectionEngine = False
        picker.SetEntityType(entity_type, TrimbleOffice.TheOffice.CurrentProject)
        innerlist = SCREntityPicker._find_inner(picker)
        if innerlist is not None:
            if multi_select:
                innerlist.MultiSelect = True
            innerlist.HideSelection = False
            innerlist.BackColor = backcolor
            innerlist.Tag = {'_scr': True, 'hl': highlight_color, 'sel': set(), 'upd': [False], 'anchor': None}
            innerlist.SelectedIndexChanged += SCREntityPicker._on_selection_changed
        return picker, innerlist

    @staticmethod
    def create_layers(host, multi_select=True, backcolor=None, highlight_color=None):
        if backcolor is None: backcolor = Color.FromArgb(0xF3, 0xF3, 0xF7)
        if highlight_color is None: highlight_color = Color.FromArgb(0x28, 0x60, 0xC0)
        return SCREntityPicker._setup(host, Project.FixedSerial.LayerContainer,
                                      clr.GetClrType(Layer), multi_select, backcolor, highlight_color)

    @staticmethod
    def create_surfaces(host, multi_select=True, backcolor=None, highlight_color=None):
        if backcolor is None: backcolor = Color.FromArgb(0xF3, 0xF3, 0xF7)
        if highlight_color is None: highlight_color = Color.FromArgb(0x28, 0x60, 0xC0)
        types = Array[Type]([clr.GetClrType(ProjectedSurface)]) + Array[Type](SurfaceTypeLists.AllWithCutFillMap)
        return SCREntityPicker._setup(host, Project.FixedSerial.WorldView,
                                      types, multi_select, backcolor, highlight_color)

    @staticmethod
    def create(host, search_container, entity_type, multi_select=True, backcolor=None, highlight_color=None):
        if backcolor is None: backcolor = Color.FromArgb(0xF3, 0xF3, 0xF7)
        if highlight_color is None: highlight_color = Color.FromArgb(0x28, 0x60, 0xC0)
        return SCREntityPicker._setup(host, search_container, entity_type,
                                      multi_select, backcolor, highlight_color)


class SCROptions:

    @staticmethod
    def LoadMacroOptions(obj, prefix, defaults, project=None):
        wv = project[Project.FixedSerial.WorldView] if project else None
        for name, default in defaults.items():
            ctrl = getattr(obj, name, None)
            if ctrl is None:
                continue
            key = prefix + "." + name
            t = type(ctrl).__name__
            if t in ("CheckBox", "RadioButton"):
                try:
                    ctrl.IsChecked = bool(OptionsManager.GetBool(key, default))
                except:
                    ctrl.IsChecked = bool(default)
            elif t == "DistanceEdit":
                ctrl.Distance = OptionsManager.GetDouble(key, default)
            elif t == "NumericEdit":
                ctrl.Value = OptionsManager.GetDouble(key, default)
            elif t == "LayerPicker":
                sn = OptionsManager.GetUint(key, default)
                o = project.Concordance.Lookup(sn)
                if o is not None and isinstance(o.GetSite(), LayerCollection):
                    ctrl.SetSelectedSerialNumber(sn, InputMethod(3))
                else:
                    ctrl.SetSelectedSerialNumber(default, InputMethod(3))
            elif t == "ColorPicker":
                argb = OptionsManager.GetInt(key, 0)
                ctrl.SelectedColor = Color.FromArgb(argb) if argb != 0 else default
            elif t == "BearingEdit":
                ctrl.Direction = OptionsManager.GetDouble(key, default)
            elif t == "OffsetEdit":
                ctrl.Offset = OptionsManager.GetDouble(key, default)
            elif t == "AngularEdit":
                ctrl.Angle = OptionsManager.GetDouble(key, default)
            elif t == "StationEdit":
                ctrl.SetStation(OptionsManager.GetDouble(key, default), project)
            elif t == "CoordinateEdit":
                x = OptionsManager.GetDouble(key + "_x", 0.0)
                y = OptionsManager.GetDouble(key + "_y", 0.0)
                z = OptionsManager.GetDouble(key + "_z", 0.0)
                ctrl.SetCoordinate(Point3D(x, y, z), project, wv.CoordinateSystemDefinition)
            elif t in ("TextBox", "TextBlock"):
                ctrl.Text = OptionsManager.GetString(key, default if default is not None else "")
            elif t == "ComboBoxEntityPicker":
                try:
                    ctrl.SelectIndex(OptionsManager.GetInt(key, default if default is not None else 0))
                except:
                    ctrl.SelectIndex(0)
            elif t == "ElevationEdit":
                ctrl.Elevation = OptionsManager.GetDouble(key, default)
            elif t == "NameEdit":
                ctrl.SelectedName = OptionsManager.GetString(key, default if default is not None else "")
            elif t == "LineweightPicker":
                ctrl.Lineweight = OptionsManager.GetInt(key, default if default is not None else 0)
            elif t == "ComboBox":
                if isinstance(default, str):
                    ctrl.Text = OptionsManager.GetString(key, default)
                else:
                    ctrl.SelectedIndex = OptionsManager.GetInt(key, default if default is not None else 0)

    @staticmethod
    def SaveMacroOptions(obj, prefix, defaults):
        for name in defaults:
            ctrl = getattr(obj, name, None)
            if ctrl is None:
                continue
            key = prefix + "." + name
            t = type(ctrl).__name__
            if t in ("CheckBox", "RadioButton"):
                OptionsManager.SetValue(key, bool(ctrl.IsChecked) if ctrl.IsChecked is not None else False)
            elif t == "DistanceEdit":
                OptionsManager.SetValue(key, ctrl.Distance)
            elif t == "NumericEdit":
                OptionsManager.SetValue(key, ctrl.Value)
            elif t == "LayerPicker":
                OptionsManager.SetValue(key, ctrl.SelectedSerialNumber)
            elif t == "ColorPicker":
                OptionsManager.SetValue(key, ctrl.SelectedColor.ToArgb())
            elif t == "BearingEdit":
                OptionsManager.SetValue(key, ctrl.Direction)
            elif t == "OffsetEdit":
                OptionsManager.SetValue(key, ctrl.Offset)
            elif t == "AngularEdit":
                OptionsManager.SetValue(key, ctrl.Angle)
            elif t == "StationEdit":
                OptionsManager.SetValue(key, ctrl.Distance)
            elif t == "CoordinateEdit":
                c = ctrl.Coordinate
                if c is not None:
                    OptionsManager.SetValue(key + "_x", c.X)
                    OptionsManager.SetValue(key + "_y", c.Y)
                    OptionsManager.SetValue(key + "_z", c.Z)
            elif t in ("TextBox", "TextBlock"):
                OptionsManager.SetValue(key, ctrl.Text)
            elif t == "ComboBoxEntityPicker":
                OptionsManager.SetValue(key, ctrl.SelectedIndex)
            elif t == "ElevationEdit":
                OptionsManager.SetValue(key, ctrl.Elevation)
            elif t == "NameEdit":
                OptionsManager.SetValue(key, ctrl.SelectedName)
            elif t == "LineweightPicker":
                OptionsManager.SetValue(key, ctrl.Lineweight)
            elif t == "ComboBox":
                if isinstance(defaults[name], str):
                    OptionsManager.SetValue(key, ctrl.Text)
                else:
                    OptionsManager.SetValue(key, ctrl.SelectedIndex)

    @staticmethod
    def LoadProjectOptions(obj, prefix, defaults, project):
        settings = ConstructionCommandsSettings.ProvideObject(project)
        for name, default in defaults.items():
            ctrl = getattr(obj, name, None)
            if ctrl is None:
                continue
            key = prefix + "." + name
            t = type(ctrl).__name__
            if t in ("CheckBox", "RadioButton"):
                try:
                    ctrl.IsChecked = bool(settings.GetBoolean(key, default))
                except:
                    ctrl.IsChecked = bool(default)
            elif t in ("TextBox", "TextBlock"):
                ctrl.Text = settings.GetString(key, default if default is not None else "")
            elif t == "DistanceEdit":
                ctrl.Distance = settings.GetDouble(key, default)
            elif t == "NumericEdit":
                ctrl.Value = settings.GetDouble(key, default)
            elif t == "LayerPicker":
                sn = settings.GetUInt32(key, default)
                o = project.Concordance[sn]
                if o is not None and isinstance(o.GetSite(), LayerCollection):
                    ctrl.SetSelectedSerialNumber(sn, InputMethod(3))
                else:
                    ctrl.SetSelectedSerialNumber(default, InputMethod(3))
            elif t == "ColorPicker":
                argb = settings.GetInt32(key, 0)
                ctrl.SelectedColor = Color.FromArgb(argb) if argb != 0 else default
            elif t == "BearingEdit":
                ctrl.Direction = settings.GetDouble(key, default)
            elif t == "ComboBoxEntityPicker":
                try:
                    ctrl.SelectIndex(settings.GetInt32(key, default if default is not None else 0))
                except:
                    ctrl.SelectIndex(0)
            elif t == "OffsetEdit":
                ctrl.Offset = settings.GetDouble(key, default)
            elif t == "AngularEdit":
                ctrl.Angle = settings.GetDouble(key, default)
            elif t == "StationEdit":
                ctrl.SetStation(settings.GetDouble(key, default), project)
            elif t == "ElevationEdit":
                ctrl.Elevation = settings.GetDouble(key, default)
            elif t == "NameEdit":
                ctrl.SelectedName = settings.GetString(key, default if default is not None else "")
            elif t == "LineweightPicker":
                ctrl.Lineweight = settings.GetInt32(key, default if default is not None else 0)
            elif t == "ComboBox":
                if isinstance(default, str):
                    ctrl.Text = settings.GetString(key, default)
                else:
                    ctrl.SelectedIndex = settings.GetInt32(key, default if default is not None else 0)

    @staticmethod
    def SaveProjectOptions(obj, prefix, defaults, project):
        settings = ConstructionCommandsSettings.ProvideObject(project)
        for name in defaults:
            ctrl = getattr(obj, name, None)
            if ctrl is None:
                continue
            key = prefix + "." + name
            t = type(ctrl).__name__
            if t in ("CheckBox", "RadioButton"):
                settings.SetBoolean(key, bool(ctrl.IsChecked) if ctrl.IsChecked is not None else False)
            elif t in ("TextBox", "TextBlock"):
                settings.SetString(key, ctrl.Text)
            elif t == "DistanceEdit":
                settings.SetDouble(key, ctrl.Distance)
            elif t == "NumericEdit":
                settings.SetDouble(key, ctrl.Value)
            elif t == "LayerPicker":
                settings.SetUInt32(key, ctrl.SelectedSerialNumber)
            elif t == "ColorPicker":
                settings.SetInt32(key, ctrl.SelectedColor.ToArgb())
            elif t == "BearingEdit":
                settings.SetDouble(key, ctrl.Direction)
            elif t == "ComboBoxEntityPicker":
                settings.SetInt32(key, ctrl.SelectedIndex)
            elif t == "OffsetEdit":
                settings.SetDouble(key, ctrl.Offset)
            elif t == "AngularEdit":
                settings.SetDouble(key, ctrl.Angle)
            elif t == "StationEdit":
                settings.SetDouble(key, ctrl.Distance)
            elif t == "ElevationEdit":
                settings.SetDouble(key, ctrl.Elevation)
            elif t == "NameEdit":
                settings.SetString(key, ctrl.SelectedName)
            elif t == "LineweightPicker":
                settings.SetInt32(key, ctrl.Lineweight)
            elif t == "ComboBox":
                if isinstance(defaults[name], str):
                    settings.SetString(key, ctrl.Text)
                else:
                    settings.SetInt32(key, ctrl.SelectedIndex)

    @staticmethod
    def _IsWindowOnAnyScreen(obj):
        from System.Windows.Forms import Screen
        for screen in Screen.AllScreens:
            wa = screen.WorkingArea
            if wa.Left <= obj.Left < wa.Left + wa.Width and wa.Top <= obj.Top < wa.Top + wa.Height:
                return True
        return False

    @staticmethod
    def LoadWindowState(obj, prefix, default_left=100, default_top=100, default_width=600, default_height=400):
        from System.Windows import WindowState
        obj.Left        = OptionsManager.GetDouble(prefix + ".windowleft",      default_left)
        obj.Top         = OptionsManager.GetDouble(prefix + ".windowtop",       default_top)
        obj.Width       = OptionsManager.GetDouble(prefix + ".windowwidth",     default_width)
        obj.Height      = OptionsManager.GetDouble(prefix + ".windowheight",    default_height)
        obj.WindowState = WindowState(OptionsManager.GetInt(prefix + ".windowstateint", 0))
        if not SCROptions._IsWindowOnAnyScreen(obj):
            obj.Left        = default_left
            obj.Top         = default_top
            obj.WindowState = WindowState(0)

    @staticmethod
    def SaveWindowState(obj, prefix):
        OptionsManager.SetValue(prefix + ".windowleft",      obj.Left)
        OptionsManager.SetValue(prefix + ".windowtop",       obj.Top)
        OptionsManager.SetValue(prefix + ".windowwidth",     obj.Width)
        OptionsManager.SetValue(prefix + ".windowheight",    obj.Height)
        OptionsManager.SetValue(prefix + ".windowstateint",  int(obj.WindowState))


class SCROctree(object):
    """
    Octree for fast 3D spatial proximity queries.  Each entry stores arbitrary
    caller-supplied data alongside its (x, y, z) position, so you can attach
    serials, segment descriptors, feature codes, or anything else.

    Build:
        tree = SCROctree()
        tree.build([(x, y, z, data), ...])
        tree.build_from_points([(Point3D, data), ...])   # .X .Y .Z accessor

    3-D queries:
        hits       = tree.find_within(x, y, z, threshold)                         # list[data]
        hits       = tree.find_within_box(x0, x1, y0, y1, z0, z1)                 # list[data]
        data, dist = tree.find_closest(x, y, z)                                   # (data, dist) or (None, None)
        data, dist = tree.find_closest(x, y, z, max_dist=1.0)                     # bounded search

    2-D queries (Z ignored — useful when elevation is unreliable or absent):
        hits       = tree.find_within_2d(x, y, threshold)
        data, dist = tree.find_closest_2d(x, y)
    """

    _BUCKET    = 16
    _MAX_DEPTH = 14

    class _Node(object):
        __slots__ = ['cx', 'cy', 'cz', 'hs', 'children', 'items', 'leaf']
        def __init__(self, cx, cy, cz, hs):
            self.cx = cx; self.cy = cy; self.cz = cz; self.hs = hs
            self.children = None
            self.items = []     # list of (x, y, z, data)
            self.leaf  = True

    def __init__(self, bucket=16, max_depth=14):
        self._bucket    = bucket
        self._max_depth = max_depth
        self._root      = None

    # ------------------------------------------------------------------ build

    def build(self, items):
        """items: iterable of (x, y, z, data).  Call once; does not support incremental inserts."""
        pts = list(items)
        if not pts:
            return
        min_x = min(p[0] for p in pts); max_x = max(p[0] for p in pts)
        min_y = min(p[1] for p in pts); max_y = max(p[1] for p in pts)
        min_z = min(p[2] for p in pts); max_z = max(p[2] for p in pts)
        cx = (min_x + max_x) * 0.5
        cy = (min_y + max_y) * 0.5
        cz = (min_z + max_z) * 0.5
        # half-size slightly larger than the data extent so all points land strictly inside
        hs = max(max_x - min_x, max_y - min_y, max_z - min_z) * 0.5005 + 0.001
        self._root = SCROctree._Node(cx, cy, cz, hs)
        for p in pts:
            self._ins(self._root, p, 0)

    def build_from_points(self, point_data_pairs):
        """point_data_pairs: iterable of (Point3D, data) — Point3D must expose .X .Y .Z"""
        self.build((pt.X, pt.Y, pt.Z, d) for pt, d in point_data_pairs)

    # --------------------------------------------------------------- internal

    @staticmethod
    def _oct(node, x, y, z):
        """Return 0-7 octant index for (x, y, z) relative to node centre."""
        return (4 if x >= node.cx else 0) | (2 if y >= node.cy else 0) | (1 if z >= node.cz else 0)

    def _ins(self, node, item, depth):
        if node.leaf:
            node.items.append(item)
            if len(node.items) > self._bucket and depth < self._max_depth:
                self._split(node, depth)
        else:
            self._ins(node.children[SCROctree._oct(node, item[0], item[1], item[2])], item, depth + 1)

    def _split(self, node, depth):
        h = node.hs * 0.5
        offsets = [(-1,-1,-1),(-1,-1,1),(-1,1,-1),(-1,1,1),(1,-1,-1),(1,-1,1),(1,1,-1),(1,1,1)]
        node.children = [SCROctree._Node(node.cx + ox*h, node.cy + oy*h, node.cz + oz*h, h)
                         for ox, oy, oz in offsets]
        node.leaf = False
        old = node.items
        node.items = []
        for item in old:
            self._ins(node.children[SCROctree._oct(node, item[0], item[1], item[2])], item, depth + 1)

    # ------------------------------------------------------------- 3D queries

    def find_within(self, x, y, z, threshold):
        """Return list[data] for all stored points within threshold of (x, y, z)."""
        if self._root is None:
            return []
        result = []
        t2 = threshold * threshold
        stack = [self._root]
        while stack:
            node = stack.pop()
            dx = max(abs(x - node.cx) - node.hs, 0.0)
            dy = max(abs(y - node.cy) - node.hs, 0.0)
            dz = max(abs(z - node.cz) - node.hs, 0.0)
            if dx*dx + dy*dy + dz*dz > t2:
                continue
            if node.leaf:
                for ix, iy, iz, data in node.items:
                    ddx = ix - x; ddy = iy - y; ddz = iz - z
                    if ddx*ddx + ddy*ddy + ddz*ddz <= t2:
                        result.append(data)
            else:
                for child in node.children:
                    stack.append(child)
        return result

    def find_within_box(self, min_x, max_x, min_y, max_y, min_z, max_z):
        """Return list[data] for all stored points inside the axis-aligned box."""
        result = []
        if self._root is None:
            return result
        stack = [self._root]
        while stack:
            node = stack.pop()
            if (node.cx + node.hs < min_x or node.cx - node.hs > max_x or
                    node.cy + node.hs < min_y or node.cy - node.hs > max_y or
                    node.cz + node.hs < min_z or node.cz - node.hs > max_z):
                continue
            if node.leaf:
                for ix, iy, iz, data in node.items:
                    if min_x <= ix <= max_x and min_y <= iy <= max_y and min_z <= iz <= max_z:
                        result.append(data)
            else:
                stack.extend(node.children)
        return result

    def find_closest(self, x, y, z, max_dist=None):
        """Return (data, distance) of the nearest stored point, or (None, None).
        Optionally limit search radius with max_dist."""
        if self._root is None:
            return None, None
        best = [None, max_dist * max_dist if max_dist is not None else float('inf')]
        self._closest_3d(self._root, x, y, z, best)
        if best[0] is None:
            return None, None
        return best[0], best[1] ** 0.5

    def _closest_3d(self, node, x, y, z, best):
        dx = max(abs(x - node.cx) - node.hs, 0.0)
        dy = max(abs(y - node.cy) - node.hs, 0.0)
        dz = max(abs(z - node.cz) - node.hs, 0.0)
        if dx*dx + dy*dy + dz*dz >= best[1]:
            return
        if node.leaf:
            for ix, iy, iz, data in node.items:
                ddx = ix - x; ddy = iy - y; ddz = iz - z
                d2 = ddx*ddx + ddy*ddy + ddz*ddz
                if d2 < best[1]:
                    best[0] = data
                    best[1] = d2
        else:
            # visit the child whose octant contains the query point first so best[1]
            # tightens early and sibling nodes are pruned more aggressively
            first = SCROctree._oct(node, x, y, z)
            self._closest_3d(node.children[first], x, y, z, best)
            for i in range(8):
                if i != first:
                    self._closest_3d(node.children[i], x, y, z, best)

    # ------------------------------------------------------------- 2D queries

    def find_within_2d(self, x, y, threshold):
        """Return list[data] for all stored points within threshold of (x, y), ignoring Z."""
        if self._root is None:
            return []
        result = []
        t2 = threshold * threshold
        stack = [self._root]
        while stack:
            node = stack.pop()
            # AABB prune uses only X/Y; Z extent is always considered overlapping
            dx = max(abs(x - node.cx) - node.hs, 0.0)
            dy = max(abs(y - node.cy) - node.hs, 0.0)
            if dx*dx + dy*dy > t2:
                continue
            if node.leaf:
                for ix, iy, _, data in node.items:
                    ddx = ix - x; ddy = iy - y
                    if ddx*ddx + ddy*ddy <= t2:
                        result.append(data)
            else:
                for child in node.children:
                    stack.append(child)
        return result

    def find_closest_2d(self, x, y, max_dist=None):
        """Return (data, distance_2d) of the nearest stored point by XY only, or (None, None)."""
        if self._root is None:
            return None, None
        best = [None, max_dist * max_dist if max_dist is not None else float('inf')]
        self._closest_2d(self._root, x, y, best)
        if best[0] is None:
            return None, None
        return best[0], best[1] ** 0.5

    def _closest_2d(self, node, x, y, best):
        dx = max(abs(x - node.cx) - node.hs, 0.0)
        dy = max(abs(y - node.cy) - node.hs, 0.0)
        if dx*dx + dy*dy >= best[1]:
            return
        if node.leaf:
            for ix, iy, _, data in node.items:
                ddx = ix - x; ddy = iy - y
                d2 = ddx*ddx + ddy*ddy
                if d2 < best[1]:
                    best[0] = data
                    best[1] = d2
        else:
            # for a 2D query visit the two children in the matching XY quadrant first
            # (both Z-halves), then the remaining six
            xy_quad = (4 if x >= node.cx else 0) | (2 if y >= node.cy else 0)
            self._closest_2d(node.children[xy_quad],     x, y, best)
            self._closest_2d(node.children[xy_quad | 1], x, y, best)
            for i in range(8):
                if i != xy_quad and i != (xy_quad | 1):
                    self._closest_2d(node.children[i], x, y, best)


class SCRExpanders:
    """Wire Expander/RadioButton pairs so each direction keeps the other in sync.

    Clicking the expander chevron checks its RadioButton (which unchecks the
    others, collapsing their expanders).  Clicking a RadioButton expands its
    expander (and collapses the rest via their Unchecked handlers).

    Usage in OnLoad:
        SCRExpanders.wire_pairs([
            (self.expander_a, self.radiobutton_a),
            (self.expander_b, self.radiobutton_b),
        ])
    The XAML expanders must NOT use IsExpanded binding; set IsExpanded="True/False"
    directly to match the initial RadioButton state.
    """

    @staticmethod
    def wire_pairs(pairs):
        syncing = [False]

        def make_expanded(rb):
            def handler(*_):
                if not syncing[0]:
                    syncing[0] = True
                    rb.IsChecked = True
                    syncing[0] = False
            return handler

        def make_checked(ex):
            def handler(*_):
                if not syncing[0]:
                    ex.IsExpanded = True
            return handler

        def make_unchecked(ex):
            def handler(*_):
                ex.IsExpanded = False
            return handler

        for expander, radiobutton in pairs:
            expander.Expanded        += make_expanded(radiobutton)
            radiobutton.Checked      += make_checked(expander)
            radiobutton.Unchecked    += make_unchecked(expander)

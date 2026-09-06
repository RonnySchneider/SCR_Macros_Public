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

    def has_within(self, x, y, z, threshold):
        """Same query as find_within, but returns True as soon as one match is found instead of
        collecting every match in range into a list first - for a caller that only needs to know
        whether anything is nearby (filter_min_spacing below being the main one), this skips both the
        list allocation and, more importantly, exploring the rest of the tree once the answer is
        already known."""
        if self._root is None:
            return False
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
                        return True
            else:
                for child in node.children:
                    stack.append(child)
        return False

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

    # ------------------------------------------------------------- resampling

    def filter_min_spacing(self, points, spacing, progressCallback=None, bounds=None):
        """Greedily keep points at least `spacing` apart (radius-based, uniform minimum spacing -
        closer to PDAL's Poisson-disk-like "sample" filter than a density-target voxel grid).

        Builds this tree's bounds from the full `points` set up front, then walks them in order,
        testing each against what's already accepted before inserting it - first point in a cluster
        wins. Slower than a voxel grid (O(log n) tree traversal per point instead of a dict lookup,
        and every acceptance grows the tree it's checked against) but gives an actual minimum
        distance between kept points rather than one-per-cell.

        points: iterable of (x, y, z, ...) - extra trailing fields are preserved untouched.
        progressCallback, if given, is called periodically as progressCallback(index, count) - same
        time-gated (~1s) convention used elsewhere in this macro set. Return True to cancel early.
        bounds, if given, is (min_x, min_y, min_z, max_x, max_y, max_z) precomputed by the caller (e.g.
        SCRLasZip.read_points already tracks this for free while reading) - skips this method's own
        full pass over `points` computing the same thing, which is otherwise a real extra full scan at
        tens of millions of points before the point-by-point pass even starts.
        Returns a new list; does not mutate `points`."""
        points = list(points)
        if not points:
            return []
        count = len(points)
        if bounds is not None:
            min_x, min_y, min_z, max_x, max_y, max_z = bounds
        else:
            # one pass over all the points rather than six (min/max on each of x, y, z separately) - at
            # tens of millions of points that's the difference between one full scan and six before the
            # actual point-by-point pass (which has its own progress reporting) even starts
            first = points[0]
            min_x = max_x = first[0]
            min_y = max_y = first[1]
            min_z = max_z = first[2]
            for p in points:
                x, y, z = p[0], p[1], p[2]
                if x < min_x: min_x = x
                elif x > max_x: max_x = x
                if y < min_y: min_y = y
                elif y > max_y: max_y = y
                if z < min_z: min_z = z
                elif z > max_z: max_z = z
        cx = (min_x + max_x) * 0.5
        cy = (min_y + max_y) * 0.5
        cz = (min_z + max_z) * 0.5
        hs = max(max_x - min_x, max_y - min_y, max_z - min_z) * 0.5005 + 0.001
        self._root = SCROctree._Node(cx, cy, cz, hs)
        kept = []
        lastReport = timer()
        for i, p in enumerate(points):
            # progress is checked before the accept/reject branch below, not after it - checking only on
            # acceptance meant a long run of rejected points (e.g. a dense cluster all within `spacing`
            # of one already-kept point, which is exactly what a large spacing setting produces on a
            # dense cloud) skipped the check entirely for however long that run took, which is what made
            # the display look stuck on one message the whole time instead of just updating less often
            if progressCallback and (timer() - lastReport) > 1.0:
                lastReport = timer()
                if progressCallback(i, count):
                    break

            if self.has_within(p[0], p[1], p[2], spacing):
                continue
            self._ins(self._root, (p[0], p[1], p[2], p), 0)
            kept.append(p)
        return kept

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


class SCRLasZip:
    """Read raw X/Y/Z points out of a LAS/LAZ file via the LASzip C API (laszip3.dll, shipped
    by TBC in its install root), without importing the file into the project - no
    PointCloudDatabase object is ever created. Use this when only the coordinates are needed
    (e.g. to build a temporary surface for a LandXML export) rather than a real TBC point cloud.

    IronPython has no ctypes/ffi of its own, so the P/Invoke declarations are compiled on first
    use into an in-memory assembly via CSharpCodeProvider and cached on the class.

    Usage:
        for x, y, z in SCRLasZip.read_points(r"C:\data\survey.laz"):
            ...
    """

    _interop_type = None

    _INTEROP_SOURCE = r"""
using System;
using System.Runtime.InteropServices;

public static class SCR_LasZipInterop
{
    [DllImport("laszip3.dll", CallingConvention = CallingConvention.Cdecl)]
    public static extern int laszip_create(out IntPtr pointer);

    [DllImport("laszip3.dll", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
    public static extern int laszip_open_reader(IntPtr pointer, string file_name, out int is_compressed);

    [DllImport("laszip3.dll", CallingConvention = CallingConvention.Cdecl)]
    public static extern int laszip_get_point_count(IntPtr pointer, out long count);

    [DllImport("laszip3.dll", CallingConvention = CallingConvention.Cdecl)]
    public static extern int laszip_read_point(IntPtr pointer);

    [DllImport("laszip3.dll", CallingConvention = CallingConvention.Cdecl)]
    public static extern int laszip_get_coordinates(IntPtr pointer, [Out] double[] coordinates);

    [DllImport("laszip3.dll", CallingConvention = CallingConvention.Cdecl)]
    public static extern int laszip_close_reader(IntPtr pointer);

    [DllImport("laszip3.dll", CallingConvention = CallingConvention.Cdecl)]
    public static extern int laszip_destroy(IntPtr pointer);

    [DllImport("laszip3.dll", CallingConvention = CallingConvention.Cdecl)]
    public static extern int laszip_get_error(IntPtr pointer, out IntPtr error);

    public static string GetErrorMessage(IntPtr reader)
    {
        IntPtr msgPtr;
        laszip_get_error(reader, out msgPtr);
        return msgPtr == IntPtr.Zero ? "" : Marshal.PtrToStringAnsi(msgPtr);
    }
}
"""

    @classmethod
    def _ensure_interop(cls):
        if cls._interop_type is not None:
            return cls._interop_type

        import System.Diagnostics
        from Microsoft.CSharp import CSharpCodeProvider
        from System.CodeDom.Compiler import CompilerParameters

        # laszip3.dll sits directly next to TBC.exe, which is already on the default DllImport
        # search path (the calling exe's own folder) - no PATH tweak needed here, unlike
        # gdal_wrap.dll which lives in a "gdal\x64" subfolder (see SCR_CloudRevise.py).
        exePath = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName
        tbcDir = os.path.dirname(exePath)
        laszipPath = os.path.join(tbcDir, "laszip3.dll")
        if not os.path.isfile(laszipPath):
            raise Exception("laszip3.dll not found next to TBC.exe (" + tbcDir + ")")

        provider = CSharpCodeProvider()
        compilerParams = CompilerParameters()
        compilerParams.GenerateInMemory = True
        compilerParams.ReferencedAssemblies.Add("System.dll")
        results = provider.CompileAssemblyFromSource(compilerParams, cls._INTEROP_SOURCE)
        if results.Errors.HasErrors:
            errs = "; ".join([str(e) for e in results.Errors])
            raise Exception("Failed to compile LASzip interop shim: " + errs)

        # Assembly.GetType("SCR_LasZipInterop") would hand back a raw System.Type (RuntimeType) -
        # reflection metadata, not something IronPython lets you call static methods on by name
        # (that's what produced "'RuntimeType' object has no attribute 'laszip_create'"). Adding the
        # compiled assembly as a CLR reference and importing the class by name instead gives back
        # IronPython's own wrapper type, which does support normal attribute-style static calls.
        clr.AddReference(results.CompiledAssembly)
        import SCR_LasZipInterop
        cls._interop_type = SCR_LasZipInterop
        return cls._interop_type

    @classmethod
    def read_points(cls, path, progressCallback=None, relative=False, expectedCount=None):
        """Returns a list of (x, y, z) tuples read directly out of a LAS/LAZ file.

        progressCallback, if given, is called periodically as progressCallback(pointIndex,
        pointCount) - return True from it to cancel early (matches TBC_ProgressBar.SetProgress's
        cancel convention). Coordinates come back already scaled/offset to real-world units by
        LASzip - no further transform is needed.

        Note: LASzip's API streams one point at a time (laszip_read_point / laszip_get_coordinates
        per point) - there is no bulk "read the whole cloud" call like GDAL's Band.ReadRaster, so
        this is inherently point-by-point. Per-call P/Invoke overhead is still far smaller than the
        SWIG/COM marshaling GDAL goes through, so tens of millions of points remain workable, but
        it will not be as fast as a single bulk raster read.

        relative=True changes both the return shape AND the memory footprint: instead of a plain list
        of (x, y, z) tuples, returns (offsetX, offsetY, offsetZ, xs, ys, zs, bounds) where xs/ys/zs are
        array.array('f', ...) - flat, unboxed 32-bit-float arrays - holding each point's coordinate
        MINUS the first point's (the offset). Subtracting the offset first is what makes 32-bit safe:
        raw survey coordinates (easily in the hundreds of thousands to millions) would lose meters of
        precision at that width, but values near zero keep sub-millimeter precision. The bigger saving
        isn't actually the halved bit width - it's that a Python list of tuples means one boxed float
        object per coordinate plus one tuple object per point (three-plus separate heap allocations per
        point); array.array stores raw values packed contiguously with none of that, which is what
        actually matters at tens of millions of points. Add the offset back before writing real-world
        coordinates anywhere (e.g. into a LandXML file) - see laz_points_to_landxml_surface.

        bounds is (min_x, min_y, min_z, max_x, max_y, max_z) of the offset-relative coordinates, tracked
        for free alongside the same read loop that's already visiting every point - or None if no points
        were read. Hand this to SCROctree.filter_min_spacing's own bounds parameter to skip its separate
        full pass over the points computing the exact same thing.

        expectedCount, if given (some workflows name their LAZ files with the point count baked into
        the filename), pre-sizes the relative=True output arrays up front instead of growing them one
        append() at a time - array.array's automatic growth reallocates and copies the whole buffer
        periodically as it grows, which is real, avoidable overhead at tens of millions of points. It's
        used only as a hint, not trusted blindly: laszip_get_point_count() has already proven unreliable
        for some real files (see the comment below), so a filename-derived count gets exactly the same
        skepticism - reading falls back to normal appending for any points beyond it, and any unused
        pre-sized tail is trimmed off at the end if the file actually held fewer points.
        """
        interop = cls._ensure_interop()

        ret, reader = interop.laszip_create()
        if ret != 0:
            raise Exception("laszip_create failed (code %d)" % ret)

        try:
            ret, isCompressed = interop.laszip_open_reader(reader, path)
            if ret != 0:
                raise Exception("laszip_open_reader failed for " + path + ": " + interop.GetErrorMessage(reader))

            # laszip_get_point_count() (the "friendly" API) comes back 0 for at least some real-world
            # LAS 1.4 files even though the points read fine - a known quirk of this wrapper rather
            # than the file being empty. Hand-decoding number_of_point_records from a fixed byte
            # offset in the header was tried as a workaround and turned out worse: it read a
            # plausible-looking but wrong value (the header struct isn't laid out exactly like the
            # raw LAS spec header), which then made a real end-of-file look like a failure partway
            # through. So: use the count only as an optional hint for the progress bar, and let
            # laszip_read_point's own return value be the sole authority on when the file is done -
            # confirmed by testing against a real file that it fails cleanly and exactly at the true
            # last point, with a matching "point N of N" message from laszip itself.
            ret, countHint = interop.laszip_get_point_count(reader)
            if ret != 0:
                countHint = 0
            if expectedCount:
                countHint = expectedCount

            coords = Array.CreateInstance(Double, 3)

            if relative:
                import array
                preSized = bool(expectedCount) and expectedCount > 0
                if preSized:
                    # IronPython's array module (unlike CPython's) rejects a raw string/bytes initializer
                    # for a numeric typecode - "cannot use a str to initialize an array with typecode 'f'"
                    # - so pre-size via sequence repetition instead, which is standard array behavior
                    xs = array.array('f', [0.0]) * expectedCount
                    ys = array.array('f', [0.0]) * expectedCount
                    zs = array.array('f', [0.0]) * expectedCount
                else:
                    xs, ys, zs = array.array('f'), array.array('f'), array.array('f')
                offsetX = offsetY = offsetZ = None
                # tracked alongside the read loop below (already visiting every point once) rather than
                # left for a caller to compute afterward with its own separate full pass - SCROctree's
                # bounding-box scan is exactly that kind of redundant extra pass, see filter_min_spacing's
                # bounds parameter
                minX = minY = minZ = maxX = maxY = maxZ = None
            else:
                points = []

            i = 0
            lastReport = timer()
            while True:
                ret = interop.laszip_read_point(reader)
                if ret != 0:
                    break  # end of file (or an unrecoverable error - laszip gives no separate signal for the two)

                ret = interop.laszip_get_coordinates(reader, coords)
                if ret != 0:
                    raise Exception("laszip_get_coordinates failed at index %d: %s" % (i, interop.GetErrorMessage(reader)))

                if relative:
                    if offsetX is None:
                        offsetX, offsetY, offsetZ = coords[0], coords[1], coords[2]
                    x, y, z = coords[0] - offsetX, coords[1] - offsetY, coords[2] - offsetZ
                    if preSized and i < expectedCount:
                        xs[i] = x
                        ys[i] = y
                        zs[i] = z
                    else:
                        # either not pre-sized, or the hint undercounted and we've run past its capacity -
                        # either way, append for the rest same as the no-hint path always does
                        preSized = False
                        xs.append(x)
                        ys.append(y)
                        zs.append(z)

                    if minX is None:
                        minX = maxX = x
                        minY = maxY = y
                        minZ = maxZ = z
                    else:
                        if x < minX: minX = x
                        elif x > maxX: maxX = x
                        if y < minY: minY = y
                        elif y > maxY: maxY = y
                        if z < minZ: minZ = z
                        elif z > maxZ: maxZ = z
                else:
                    points.append((coords[0], coords[1], coords[2]))
                i += 1

                # time-gated rather than count-gated (a fixed point-count interval either updates far too
                # often on a fast machine/small file or barely at all on a slow one/huge file - the same
                # lesson learned the hard way with TBC_ProgressBar) - same >1s cadence used elsewhere in
                # this macro set (see laz_points_to_landxml_surface's vertex-adding loop)
                if progressCallback and (timer() - lastReport) > 1.0:
                    lastReport = timer()
                    if progressCallback(i, countHint if countHint > 0 else i):
                        break

            if relative:
                if preSized and i < expectedCount:
                    # the hint overcounted - trim the unused pre-sized (zero-filled) tail
                    xs, ys, zs = xs[:i], ys[:i], zs[:i]
                bounds = (minX, minY, minZ, maxX, maxY, maxZ) if minX is not None else None
                return offsetX or 0.0, offsetY or 0.0, offsetZ or 0.0, xs, ys, zs, bounds
            return points
        finally:
            interop.laszip_close_reader(reader)
            interop.laszip_destroy(reader)


def resample_points_voxel_grid(points, spacing, progressCallback=None):
    """Fast density-target point cloud resampling: buckets points into a `spacing`-sized voxel grid
    and keeps one point per occupied voxel (first point encountered) - same behavior as PDAL's voxel
    grid filter. O(1) per point via a dict lookup, no tree construction - much cheaper than
    SCROctree.filter_min_spacing for a plain density target, at the cost of not guaranteeing a true
    minimum distance between kept points (two points in adjacent voxels can end up closer together
    than `spacing`).

    points: iterable of (x, y, z, ...) - extra trailing fields are preserved untouched.
    progressCallback, if given, is called periodically as progressCallback(index, count) - same
    time-gated (~1s) convention used elsewhere in this macro set. Return True to cancel early. count
    requires materializing `points` into a list up front if it isn't already one.
    Returns a new list; does not mutate `points`."""
    seen = {}
    if progressCallback:
        points = points if hasattr(points, "__len__") else list(points)
        count = len(points)
        lastReport = timer()
        for i, p in enumerate(points):
            key = (int(p[0] // spacing), int(p[1] // spacing), int(p[2] // spacing))
            if key not in seen:
                seen[key] = p
            if (timer() - lastReport) > 1.0:
                lastReport = timer()
                if progressCallback(i, count):
                    break
    else:
        for p in points:
            key = (int(p[0] // spacing), int(p[1] // spacing), int(p[2] // spacing))
            if key not in seen:
                seen[key] = p
    return list(seen.values())


def resample_points_grid_2d(points, spacing, bounds=None, progressCallback=None):
    """Regular-grid ("DEM-style") resampling: lays a plain XY grid across the point cloud's extent at
    `spacing` intervals, and for every grid cell takes the nearest actual point by 2D distance (Z
    ignored for the search itself, but the matched point's real Z is what gets kept - never
    interpolated). Different in kind from voxel_grid/filter_min_spacing above, which both thin the
    cloud while following its own point distribution; this instead produces perfectly uniform output
    spacing, closer to how a raster DEM samples.

    Fast specifically for a dense, roughly uniform cloud at a spacing coarser than the native point
    spacing: grid cell count scales with (extent/spacing)^2 while point count scales with
    (extent/native_spacing)^2, so there are far fewer nearest-point searches than input points - e.g.
    ~0.05m native spacing thinned to 10m is a ~40,000x reduction in query count.

    Uses a simple spatial hash grid (dict of (cellx, celly) -> list of points, cell size = spacing) for
    the nearest-point search, rather than SCROctree: an octree built over tens of millions of points
    holds one tree node object plus a duplicated (x, y, z, data) tuple per point, which at real drone-LAZ
    point counts (tens of millions) is enough object overhead to push a 16GB machine into swapping. The
    hash grid holds exactly one list-entry reference per point (no per-point node objects, no duplicated
    coordinates), and building it is a single O(1)-per-point pass, same as resample_points_voxel_grid's
    own dict - so this method now shares that one's memory profile instead of SCROctree's.

    points: iterable of (x, y, z, ...) - extra trailing fields are preserved untouched on a match.
    bounds, if given, is (min_x, min_y, min_z, max_x, max_y, max_z) precomputed by the caller (e.g.
    SCRLasZip.read_points already tracks this for free while reading) - skips computing it again here.
    progressCallback, if given, is called periodically as progressCallback(index, count) over the grid
    cells (not the input points) - same time-gated (~1s) convention used elsewhere. Return True to
    cancel early - a cancelled run returns whatever was matched so far.
    Returns a new list; does not mutate `points`. A grid cell with no point within `spacing` (e.g. near
    an irregular site boundary) is simply skipped, not filled in with a distant/wrong match."""
    points = list(points)
    if not points:
        return []

    if bounds is not None:
        min_x, min_y, min_z, max_x, max_y, max_z = bounds
    else:
        first = points[0]
        min_x = max_x = first[0]
        min_y = max_y = first[1]
        min_z = max_z = first[2]
        for p in points:
            x, y, z = p[0], p[1], p[2]
            if x < min_x: min_x = x
            elif x > max_x: max_x = x
            if y < min_y: min_y = y
            elif y > max_y: max_y = y
            if z < min_z: min_z = z
            elif z > max_z: max_z = z

    # bucket size == spacing, so any point within `spacing` of a grid centre is guaranteed to fall in
    # the centre's own cell or one of its 8 immediate neighbors - a 3x3 neighborhood search is enough
    buckets = {}
    for p in points:
        key = (int((p[0] - min_x) // spacing), int((p[1] - min_y) // spacing))
        bucket = buckets.get(key)
        if bucket is None:
            buckets[key] = [p]
        else:
            bucket.append(p)

    countX = int((max_x - min_x) / spacing) + 1
    countY = int((max_y - min_y) / spacing) + 1
    count = countX * countY
    maxD2 = spacing * spacing

    kept = []
    seenIds = set()
    lastReport = timer()
    i = 0
    neighborOffsets = [(-1,-1),(-1,0),(-1,1),(0,-1),(0,0),(0,1),(1,-1),(1,0),(1,1)]
    for ix in range(countX):
        gx = min_x + ix * spacing
        cx = ix
        for iy in range(countY):
            gy = min_y + iy * spacing
            cy = iy

            bestP = None
            bestD2 = maxD2
            for ox, oy in neighborOffsets:
                bucket = buckets.get((cx + ox, cy + oy))
                if not bucket:
                    continue
                for p in bucket:
                    ddx = p[0] - gx; ddy = p[1] - gy
                    d2 = ddx*ddx + ddy*ddy
                    if d2 <= bestD2:
                        bestD2 = d2
                        bestP = p

            if bestP is not None:
                # dedupe by which native bucket bestP itself falls into (same cell key/size used to build
                # `buckets` above), NOT by point identity - two neighboring grid queries only share the
                # exact same nearest point often enough to thin properly when the native spacing happens
                # to be a clean multiple of the target spacing. In the (common) case where it isn't - e.g.
                # 1m native spacing being thinned to a 1.26m target - most adjacent queries instead snap to
                # DIFFERENT-but-still-distinct real points, so identity-based dedup lets nearly every
                # native point through with no effective thinning at all in that axis, while an axis whose
                # native spacing is already coarser than the target looks fine and masks the bug. Cell-key
                # dedup instead guarantees at most one output point per spacing-sized cell regardless of
                # query alignment, the same hard guarantee resample_points_voxel_grid already has.
                cellKey = (int((bestP[0] - min_x) // spacing), int((bestP[1] - min_y) // spacing))
                if cellKey not in seenIds:
                    seenIds.add(cellKey)
                    kept.append(bestP)

            i += 1
            if progressCallback and (timer() - lastReport) > 1.0:
                lastReport = timer()
                if progressCallback(i, count):
                    return kept

    return kept

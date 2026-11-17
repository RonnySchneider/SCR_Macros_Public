#   GNU GPLv3
#   <this is an add-on Script/Macro for the geospatial software "Trimble Business Center" aka TBC>
#   <you'll need at least the "Survey Advanced" licence of TBC in order to run this script>
#	<see the ToolTip section below for a brief explanation what the script does>
#	<see the Help-Files for more details>
#   Copyright (C) 2026 Ronny Schneider
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
    "stepwidth":   0.002,
    "surfacename": "IFC Underside",
    "gapvoids":    True,
}

# all tolerances in metres
VERTICALTOL = 1e-3     # |nz| of the unit normal below this = vertical face, ignored
SNAPTOL     = 1e-4     # vertices closer than this in plan are welded into one
ZTOL        = 2e-4     # undersides closer than this in height count as coincident
STEPTOL     = 1e-3     # height jumps above this at a shared edge are steps that get nudged
MINAREA     = 1e-10    # pieces smaller than this are dropped
MINWIDTH    = 1e-5     # pieces thinner than this are dropped (slivers)


def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_IFCUnderside"
    cmdData.CommandName = "SCR_IFCUnderside"
    cmdData.Caption = "_SCR_IFCUnderside"
    cmdData.UIForm = "SCR_IFCUnderside"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "IFC"
        cmdData.ShortCaption = "IFC Underside"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.000
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""

        cmdData.ToolTipTitle = "IFC Underside"
        cmdData.ToolTipTextFormatted = "creates a surface along the underside (lower envelope) of the selected IFC objects - vertical steps are replaced by steep slopes"

    except:
        pass

    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_IFCUnderside(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_IFCUnderside.xaml") as s:
            wpf.LoadComponent (self, s)
        self.currentProject = currentProject
        self.macroFileFolder = macroFileFolder

    def HelpClicked(self, cmd, e):
        webbrowser.open("C:\ProgramData\Trimble\MacroCommands3\SCR Macros\MacroHelp\MacroHelp.htm#" + type(self).__name__)

    def OnLoad(self, cmd, buttons, event):
        self.okBtn = buttons[0]
        buttons[2].Content = "Help"
        buttons[2].Visibility = Visibility.Visible
        buttons[2].Click += self.HelpClicked
        self.Caption = cmd.Command.Caption

        optionMenu = SelectionContextMenuHandler()
        optionMenu.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        self.objs.ButtonContextMenu = optionMenu
        self.objs.IsEntityValidCallback = self.IsValidBIM

        self.stepwidth.FormatProperty.AddSuffix = ControlBoolean(1)

        try:
            self.SetDefaultOptions()
        except:
            pass

    def SetDefaultOptions(self):
        SCROptions.LoadMacroOptions(self, "SCR_IFCUnderside", _OPTIONS, self.currentProject)

    def SaveOptions(self):
        SCROptions.SaveMacroOptions(self, "SCR_IFCUnderside", _OPTIONS)

    def IsValidBIM(self, serial):
        o = self.currentProject.Concordance[serial]
        return isinstance(o, BIMEntity)

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def Dispose(self, cmd, disposing):
        pass

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content = ''
        self.success.Content = ''

        stepwidth = self.stepwidth.Distance
        if math.isnan(stepwidth) or stepwidth <= SNAPTOL * 2:
            self.error.Content = 'enter a step width larger than ' + str(SNAPTOL * 2) + ' m'
            return

        surfacename = self.surfacename.Text.strip()
        if not surfacename:
            self.error.Content = 'enter a surface name'
            return

        gapvoids = self.gapvoids.IsChecked

        wv = self.currentProject [Project.FixedSerial.WorldView]

        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)

        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                tris, nmeshes = self.collecttriangles()
                if len(tris) == 0:
                    self.error.Content = 'no IFC objects with non-vertical faces selected'
                else:
                    pieces = lowerenvelope(tris, self.progress)
                    if pieces is None:
                        self.error.Content = 'cancelled'
                    elif len(pieces) == 0:
                        self.error.Content = 'no underside found'
                    else:
                        pieces, bands, nsteps = nudgesteps(pieces, stepwidth, self.progress)
                        if pieces is None:
                            self.error.Content = 'cancelled'
                        else:
                            vertices, breaklines = tinsetup(pieces, self.progress)
                            name = Model3D.GetUniqueName(surfacename, None, wv)
                            nvoid = self.createsurface(wv, name, vertices, breaklines, pieces, bands, gapvoids)

                            self.success.Content = 'surface "' + name + '" created' + \
                                '\n' + str(nmeshes) + ' mesh(es), ' + str(len(tris)) + ' non-vertical triangles' + \
                                '\n' + str(len(pieces)) + ' visible pieces, ' + str(nsteps) + ' step edges nudged'
                            if gapvoids:
                                self.success.Content += '\n' + str(nvoid) + ' gap triangles set to void'

                failGuard.Commit()
                self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
                UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())

        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            # EndMark MUST be set no matter what
            # otherwise TBC won't work anymore and needs to be restarted
            self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
            UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
            self.error.Content += '\nan Error occurred - Result probably incomplete\n' + str(exc_type) + '\n' + str(exc_obj) + '\nLine ' + str(exc_tb.tb_lineno)

        ProgressBar.TBC_ProgressBar.Title = ""
        self.SaveOptions()

    def progress(self, title, pct):
        # shared progress callback for the geometry functions, returns True on cancel
        if title is not None:
            ProgressBar.TBC_ProgressBar.Title = title
        return ProgressBar.TBC_ProgressBar.SetProgress(int(pct))

    def collecttriangles(self):
        # walk the selected BIM entities (and their children) and collect all mesh triangles in world coordinates
        # every ShellMeshInstance is only used once, even if a parent and its child are both selected
        meshes = []
        seen = set()

        def walk(entity):
            for o in entity:
                if isinstance(o, ShellMeshInstance):
                    if o.SerialNumber not in seen:
                        seen.add(o.SerialNumber)
                        meshes.Add(o)
                elif isinstance(o, BIMEntity):
                    walk(o)

        for o in self.objs.SelectedMembers(self.currentProject):
            if isinstance(o, BIMEntity):
                walk(o)

        tris = []
        ProgressBar.TBC_ProgressBar.Title = "reading IFC meshes"
        last_t = timer()
        for m, meshinstance in enumerate(meshes):
            if timer() - last_t > 0.5:
                ProgressBar.TBC_ProgressBar.SetProgress(m * 100 // len(meshes))
                last_t = timer()

            shellmeshdata = meshinstance.GetShellMeshData()
            vertexlist = shellmeshdata.GetVertexList(meshinstance.GlobalTransformation)
            # triangulationContext 0 - NonTriangulated; 1 - Triangulated; 2 - TriangulatedImplicit without the length entry 3
            indices = shellmeshdata.GetIndices(TriangulationContext(1))
            pts = [(p.X, p.Y, p.Z) for p in vertexlist]
            idx = list(indices)
            for i in range(0, len(idx) - 3, 4):
                t = maketriangle(pts[idx[i + 1]], pts[idx[i + 2]], pts[idx[i + 3]], len(tris))
                if t is not None:
                    tris.Add(t)

        return tris, len(meshes)

    def createsurface(self, wv, name, vertices, breaklines, pieces, bands, gapvoids):
        ProgressBar.TBC_ProgressBar.Title = "building surface"
        newSurface = wv.Add(clr.GetClrType(Model3D))
        newSurface.Name = name
        # SmoothShading | ElevationContours | ShowBackfaces | FilledTriangles | ElevationContours_2D | FilledTriangles_2D
        newSurface.Mode = DisplayMode(1 + 64 + 128 + 512 + 1024 + 65536)

        builder = newSurface.GetGemBatchBuilder()
        builderindex = []
        for x, y, z in vertices:
            builderindex.Add(builder.AddVertex(Point3D(x, y, z))[0])
        soft = Byte(DTMSharpness.eSoft)
        for v0, v1 in breaklines:
            builder.AddBreakline(soft, builderindex[v0], builderindex[v1])
        builder.Construction()

        nvoid = 0
        if gapvoids:
            # triangles that don't lie on an underside piece or in a step band span a gap between objects
            # or the outside of a concave footprint - give them the null material
            cover = Coverage(pieces, bands)
            gemmap = GemMaterialMap()
            for t in range(builder.NumberOfTriangles):
                if not builder.IsTrianglePresent(t):
                    continue
                cx = cy = 0.0
                for s in range(3):
                    p = builder.GetTriangleVertex(t, s)[-1]
                    cx += p.X
                    cy += p.Y
                if not cover.covered(cx / 3.0, cy / 3.0):
                    builder.AttachMaterial(0, gemmap, t)
                    nvoid += 1

        builder.Commit()
        return nvoid


# ---------------------------------------------------------------------------------------------------
# geometry engine - plain Python, no TBC types, so it can be tested outside of TBC
#
# a triangle/piece plane is z = a*x + b*y + c
# a polygon is a list of (x, y) tuples, convex and counter-clockwise in plan
# a half-plane (ha, hb, hc) is the region ha*x + hb*y + hc >= 0
# ---------------------------------------------------------------------------------------------------

class Tri(object):
    __slots__ = ('poly', 'a', 'b', 'c', 'zmin', 'zmax', 'bbox', 'idx')


def maketriangle(p0, p1, p2, idx):
    ux, uy, uz = p1[0] - p0[0], p1[1] - p0[1], p1[2] - p0[2]
    vx, vy, vz = p2[0] - p0[0], p2[1] - p0[1], p2[2] - p0[2]
    nx = uy * vz - uz * vy
    ny = uz * vx - ux * vz
    nz = ux * vy - uy * vx
    nl = math.sqrt(nx * nx + ny * ny + nz * nz)
    if nl == 0.0 or abs(nz) / nl < VERTICALTOL:
        return None   # degenerate or vertical, can't be seen from below
    if abs(nz) * 0.5 < MINAREA:
        return None
    t = Tri()
    t.a = -nx / nz
    t.b = -ny / nz
    t.c = p0[2] - t.a * p0[0] - t.b * p0[1]
    if nz > 0:
        t.poly = [(p0[0], p0[1]), (p1[0], p1[1]), (p2[0], p2[1])]
    else:
        t.poly = [(p0[0], p0[1]), (p2[0], p2[1]), (p1[0], p1[1])]
    t.zmin = min(p0[2], p1[2], p2[2])
    t.zmax = max(p0[2], p1[2], p2[2])
    t.bbox = polybbox(t.poly)
    t.idx = idx
    return t


def polybbox(poly):
    xs = [p[0] for p in poly]
    ys = [p[1] for p in poly]
    return (min(xs), min(ys), max(xs), max(ys))


def polyarea(poly):
    s = 0.0
    n = len(poly)
    for i in range(n):
        x0, y0 = poly[i]
        x1, y1 = poly[(i + 1) % n]
        s += x0 * y1 - x1 * y0
    return s * 0.5


def polyok(poly):
    # rejects empty, tiny and sliver polygons
    if len(poly) < 3:
        return False
    area = polyarea(poly)
    if area < MINAREA:
        return False
    longest = 0.0
    n = len(poly)
    for i in range(n):
        dx = poly[(i + 1) % n][0] - poly[i][0]
        dy = poly[(i + 1) % n][1] - poly[i][1]
        longest = max(longest, dx * dx + dy * dy)
    return 2.0 * area / math.sqrt(longest) > MINWIDTH


def edgehalfplanes(poly):
    # inward half-planes of a ccw convex polygon, normalised so the value is a distance
    hps = []
    n = len(poly)
    for i in range(n):
        x0, y0 = poly[i]
        x1, y1 = poly[(i + 1) % n]
        ha, hb = -(y1 - y0), x1 - x0
        l = math.sqrt(ha * ha + hb * hb)
        if l == 0.0:
            continue
        ha /= l
        hb /= l
        hps.Add((ha, hb, -(ha * x0 + hb * y0)))
    return hps


def splitpoly(poly, hp):
    # splits a convex polygon by a half-plane into (inside, outside)
    ha, hb, hc = hp
    vals = [ha * x + hb * y + hc for x, y in poly]
    eps = 1e-12
    if min(vals) >= -eps:
        return poly, None
    if max(vals) <= eps:
        return None, poly
    inside = []
    outside = []
    n = len(poly)
    for i in range(n):
        p, vp = poly[i], vals[i]
        q, vq = poly[(i + 1) % n], vals[(i + 1) % n]
        if vp >= 0:
            inside.Add(p)
        if vp <= 0:
            outside.Add(p)
        if (vp > 0 and vq < 0) or (vp < 0 and vq > 0):
            t = vp / (vp - vq)
            x = (p[0] + t * (q[0] - p[0]), p[1] + t * (q[1] - p[1]))
            inside.Add(x)
            outside.Add(x)
    return (inside if len(inside) >= 3 else None), (outside if len(outside) >= 3 else None)


def subtractconvex(pieces, hps):
    # removes the convex region given by the half-planes from a list of convex polygons
    # pieces that don't touch the region are returned unchanged, so they don't get fragmented
    result = []
    changed = False
    for poly in pieces:
        out = []
        cur = poly
        for hp in hps:
            cur, o = splitpoly(cur, hp)
            if o is not None:
                out.Add(o)
            if cur is None:
                break
        if cur is None or not polyok(cur):
            result.Add(poly)   # no real overlap
        else:
            changed = True
            for o in out:
                if polyok(o):
                    result.Add(o)
    return result, changed


class Grid(object):
    # bounding box grid index
    def __init__(self, bboxes, cell=None):
        if cell is None:
            if len(bboxes) == 0:
                cell = 1.0
            else:
                s = sum(max(b[2] - b[0], b[3] - b[1]) for b in bboxes) / len(bboxes)
                cell = max(s * 2.0, 1e-3)
        self.cell = cell
        self.cells = {}
        for i, b in enumerate(bboxes):
            self.insert(i, b)

    def insert(self, i, b):
        c = self.cell
        for gx in range(int(math.floor(b[0] / c)), int(math.floor(b[2] / c)) + 1):
            for gy in range(int(math.floor(b[1] / c)), int(math.floor(b[3] / c)) + 1):
                self.cells.setdefault((gx, gy), []).Add(i)

    def query(self, b):
        c = self.cell
        found = set()
        for gx in range(int(math.floor(b[0] / c)), int(math.floor(b[2] / c)) + 1):
            for gy in range(int(math.floor(b[1] / c)), int(math.floor(b[3] / c)) + 1):
                l = self.cells.get((gx, gy))
                if l:
                    found.update(l)
        return found


def bboxoverlap(b1, b2, tol=0.0):
    return b1[0] <= b2[2] + tol and b2[0] <= b1[2] + tol and b1[1] <= b2[3] + tol and b2[1] <= b1[3] + tol


class Piece(object):
    __slots__ = ('poly', 'a', 'b', 'c', 'src', 'bbox')

    def __init__(self, poly, a, b, c, src):
        self.poly = poly
        self.a = a
        self.b = b
        self.c = c
        self.src = src
        self.bbox = polybbox(poly)

    def z(self, x, y):
        return self.a * x + self.b * y + self.c


def lowerenvelope(tris, progress):
    # for every triangle T, remove the parts where another triangle S is lower
    # the region where S hides T is convex: inside S and z_T - z_S > threshold
    # so the visible part of T is T minus a number of convex regions - exact, no sampling
    # coincident faces: the one with the lower index wins, so exactly one of them survives
    grid = Grid([t.bbox for t in tris])
    order = sorted(range(len(tris)), key=lambda i: tris[i].zmin)
    pieces = []
    last_t = timer()
    progress("computing lower envelope", 0)
    for n, ti in enumerate(order):
        if timer() - last_t > 0.5:
            if progress(None, n * 100 // len(order)):
                return None
            last_t = timer()

        T = tris[ti]
        cands = [si for si in grid.query(T.bbox) if si != ti and tris[si].zmin <= T.zmax + ZTOL]
        cands.sort(key=lambda si: tris[si].zmin)
        polys = [T.poly]
        for si in cands:
            S = tris[si]
            if not bboxoverlap(T.bbox, S.bbox):
                continue
            thr = -ZTOL if S.idx < T.idx else ZTOL
            da, db, dc = T.a - S.a, T.b - S.b, T.c - S.c - thr
            l = math.sqrt(da * da + db * db)
            hps = edgehalfplanes(S.poly)
            if l < 1e-12:
                if dc <= 0.0:
                    continue   # parallel and S not below T
            else:
                hps.Add((da / l, db / l, dc / l))
            polys, changed = subtractconvex(polys, hps)
            if not polys:
                break
        for p in polys:
            pieces.Add(Piece(p, T.a, T.b, T.c, T.idx))
    return pieces


class Welder(object):
    # welds plan positions closer than SNAPTOL into one vertex id
    def __init__(self):
        self.pts = []
        self.cells = {}

    def id(self, x, y):
        gx = int(math.floor(x / SNAPTOL))
        gy = int(math.floor(y / SNAPTOL))
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for i in self.cells.get((gx + dx, gy + dy), ()):
                    px, py = self.pts[i]
                    if (px - x) ** 2 + (py - y) ** 2 <= SNAPTOL * SNAPTOL:
                        return i
        i = len(self.pts)
        self.pts.Add((x, y))
        self.cells.setdefault((gx, gy), []).Add(i)
        return i


def buildedges(pieces):
    # welds all piece corners and splits piece edges at vertices of neighbouring pieces (T-junctions)
    # returns the welded points, per piece the ring of vertex ids, and edge -> [(piece, from, to)]
    welder = Welder()
    rings = []
    for p in pieces:
        ring = []
        for x, y in p.poly:
            i = welder.id(x, y)
            if not ring or ring[-1] != i:
                ring.Add(i)
        if len(ring) > 1 and ring[0] == ring[-1]:
            ring.pop()
        rings.Add(ring)
    pts = welder.pts

    vgrid = Grid([(x, y, x, y) for x, y in pts], cell=max(1e-3, gridcell(pieces)))

    edges = {}
    for pi, ring in enumerate(rings):
        if len(ring) < 3:
            continue
        newring = []
        n = len(ring)
        for k in range(n):
            i, j = ring[k], ring[(k + 1) % n]
            newring.Add(i)
            x0, y0 = pts[i]
            x1, y1 = pts[j]
            dx, dy = x1 - x0, y1 - y0
            l2 = dx * dx + dy * dy
            if l2 == 0.0:
                continue
            inner = []
            for v in vgrid.query((min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1))):
                if v == i or v == j:
                    continue
                vx, vy = pts[v]
                t = ((vx - x0) * dx + (vy - y0) * dy) / l2
                if t <= 0.0 or t >= 1.0:
                    continue
                cx, cy = x0 + t * dx - vx, y0 + t * dy - vy
                if cx * cx + cy * cy <= SNAPTOL * SNAPTOL:
                    inner.Add((t, v))
            inner.sort()
            for t, v in inner:
                newring.Add(v)
        rings[pi] = newring
        n = len(newring)
        for k in range(n):
            i, j = newring[k], newring[(k + 1) % n]
            if i == j:
                continue
            key = (i, j) if i < j else (j, i)
            edges.setdefault(key, []).Add((pi, i, j))
    return pts, rings, edges


def gridcell(pieces):
    if not pieces:
        return 1.0
    s = sum(max(p.bbox[2] - p.bbox[0], p.bbox[3] - p.bbox[1]) for p in pieces) / len(pieces)
    return s * 2.0


def nudgesteps(pieces, eps, progress):
    # every edge with a height jump is a step; the upper side gets cut back by eps
    # the cut is a rectangle along the step edge, extended by eps at both ends so corners are covered too
    # it is removed from all pieces that are higher than the lower side of the step, the lower side stays untouched
    progress("finding steps", 0)
    pts, rings, edges = buildedges(pieces)

    bands = []   # (half-planes, polygon, lower piece, mid x, mid y, lower z at mid)
    for key, users in edges.items():
        if len(users) != 2:
            continue
        (p0, i0, j0), (p1, i1, j1) = users
        x0, y0 = pts[key[0]]
        x1, y1 = pts[key[1]]
        mx, my = (x0 + x1) * 0.5, (y0 + y1) * 0.5
        z0 = pieces[p0].z(mx, my)
        z1 = pieces[p1].z(mx, my)
        if abs(z0 - z1) <= STEPTOL:
            continue
        if z0 > z1:
            up, ui, uj, low, zlow = p0, i0, j0, p1, z1
        else:
            up, ui, uj, low, zlow = p1, i1, j1, p0, z0
        # the upper piece is ccw, so its inside is left of ui -> uj
        ax, ay = pts[ui]
        bx, by = pts[uj]
        l = math.sqrt((bx - ax) ** 2 + (by - ay) ** 2)
        tx, ty = (bx - ax) / l * eps, (by - ay) / l * eps
        nx, ny = -ty, tx
        rect = [(ax - tx, ay - ty), (bx + tx, by + ty), (bx + tx + nx, by + ty + ny), (ax - tx + nx, ay - ty + ny)]
        bands.Add((edgehalfplanes(rect), rect, low, mx, my, zlow))

    if not bands:
        return pieces, [], 0

    progress("nudging steps", 0)
    # per piece, the bands that cut it
    grid = Grid([p.bbox for p in pieces])
    cuts = {}
    for bi, (hps, rect, low, mx, my, zlow) in enumerate(bands):
        rb = polybbox(rect)
        for pi in grid.query(rb):
            if pi == low:
                continue
            p = pieces[pi]
            if not bboxoverlap(p.bbox, rb):
                continue
            if p.z(mx, my) <= zlow + STEPTOL:
                continue   # belongs to the lower side
            cuts.setdefault(pi, []).Add(bi)

    result = []
    last_t = timer()
    for pi, p in enumerate(pieces):
        if timer() - last_t > 0.5:
            if progress(None, pi * 100 // len(pieces)):
                return None, None, 0
            last_t = timer()
        if pi not in cuts:
            result.Add(p)
            continue
        polys = [p.poly]
        for bi in cuts[pi]:
            polys, changed = subtractconvex(polys, bands[bi][0])
            if not polys:
                break
        for poly in polys:
            result.Add(Piece(poly, p.a, p.b, p.c, p.src))

    return result, [b[1] for b in bands], len(bands)


def tinsetup(pieces, progress):
    # vertices and breaklines for the TIN
    # a vertex shared by several pieces gets the lowest of their heights, so the surface never lies above an object
    # edges between coplanar pieces (same source plane) are no breaklines, they would only clutter the surface
    progress("preparing surface", 0)
    pts, rings, edges = buildedges(pieces)

    zs = [None] * len(pts)
    for pi, ring in enumerate(rings):
        p = pieces[pi]
        for v in ring:
            z = p.z(pts[v][0], pts[v][1])
            if zs[v] is None or z < zs[v]:
                zs[v] = z

    used = {}
    vertices = []
    breaklines = []
    for key, users in edges.items():
        if len(users) == 2:
            pa, pb = pieces[users[0][0]], pieces[users[1][0]]
            if abs(pa.a - pb.a) < 1e-9 and abs(pa.b - pb.b) < 1e-9 and abs(pa.c - pb.c) < ZTOL:
                continue
        ids = []
        for v in key:
            if v not in used:
                used[v] = len(vertices)
                vertices.Add((pts[v][0], pts[v][1], zs[v]))
            ids.Add(used[v])
        breaklines.Add((ids[0], ids[1]))

    # vertices that only sit inside coplanar areas are still needed, e.g. T-junctions on a breakline
    for v in range(len(pts)):
        if v not in used and zs[v] is not None:
            used[v] = len(vertices)
            vertices.Add((pts[v][0], pts[v][1], zs[v]))

    return vertices, breaklines


class Coverage(object):
    # point-in-region test for the pieces and the step bands, with a small tolerance against slivers
    def __init__(self, pieces, bands):
        self.polys = [edgehalfplanes(p.poly) for p in pieces] + [edgehalfplanes(b) for b in bands]
        bboxes = [p.bbox for p in pieces] + [polybbox(b) for b in bands]
        self.bboxes = bboxes
        self.grid = Grid(bboxes)

    def covered(self, x, y):
        tol = SNAPTOL * 10
        for i in self.grid.query((x, y, x, y)):
            b = self.bboxes[i]
            if x < b[0] - tol or x > b[2] + tol or y < b[1] - tol or y > b[3] + tol:
                continue
            if all(ha * x + hb * y + hc >= -tol for ha, hb, hc in self.polys[i]):
                return True
        return False

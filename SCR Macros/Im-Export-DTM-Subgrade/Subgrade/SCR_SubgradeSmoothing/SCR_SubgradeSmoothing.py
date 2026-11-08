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
    "surfacepicker":  0,
    "interval":       1.0,
    "fitwindow":      5.0,
    "fixcl":          False,
    "crownfit":       True,
    "crowncl":        False,
    "crownnone":      False,
    "smoothcl":       10.0,
    "smoothgrade":    20.0,
    "smoothxfall":    20.0,
    "smoothcrown":    30.0,
    "smoothwidth":    10.0,
    "voffset":        0.2,
    "depthtransition": 20.0,
    "keepabove":      True,
    "limitmax":       False,
    "maxextra":       0.1,
    "extension":      0.0,
    "crossspacing":   1.5,
    "drawlines":      False,
    "layerpicker":    8,
    "writecsv":       True,
}

# must match the ComboBoxItems in the XAML section table
CROWNMODES = ["find Crown", "Crown on CL", "single Crossfall"]

def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_SubgradeSmoothing"
    cmdData.CommandName = "SCR_SubgradeSmoothing"
    cmdData.Caption = "_SCR_SubgradeSmoothing"
    cmdData.UIForm = "SCR_SubgradeSmoothing"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = ""
    cmdData.HelpTopic = ""

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "Subgrade"
        cmdData.ShortCaption = "Subgrade Smoothing"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.013
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""

        cmdData.ToolTipTitle = "Subgrade Smoothing"
        cmdData.ToolTipTextFormatted = "create a smooth surface with gradual crossfall and grade changes from a rough road surface, following a rough centreline"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


def whittaker(y, w, length, h):
    # Whittaker smoother with 2nd difference penalty
    # minimizes sum(w*(y-z)^2) + lam * sum((d2 z)^2)
    # length is roughly the wavelength [m] that gets damped to half amplitude
    # positions with weight 0 are interpolated from their neighbours
    n = len(y)
    if n < 3:
        return list(y)
    lam = (length / (2.0 * math.pi * h)) ** 4 if length > 0 else 0.0
    lam = max(lam, 1e-6)

    # pentadiagonal symmetric matrix W + lam * D'D
    d0 = [6.0] * n
    d0[0] = d0[n - 1] = 1.0
    d0[1] = d0[n - 2] = 5.0 if n > 3 else 2.0
    d1 = [-4.0] * (n - 1)
    d1[0] = d1[n - 2] = -2.0
    d2 = [1.0] * (n - 2)
    for i in range(n):
        d0[i] = w[i] + lam * d0[i] + 1e-12
    for i in range(n - 1):
        d1[i] *= lam
    for i in range(n - 2):
        d2[i] *= lam

    # banded Cholesky, L has the diagonal l0 and two sub diagonals l1, l2
    l0 = [0.0] * n
    l1 = [0.0] * n
    l2 = [0.0] * n
    for i in range(n):
        if i >= 2:
            l2[i] = d2[i - 2] / l0[i - 2]
        if i >= 1:
            l1[i] = (d1[i - 1] - l2[i] * l1[i - 1]) / l0[i - 1]
        l0[i] = math.sqrt(max(d0[i] - l1[i] * l1[i] - l2[i] * l2[i], 1e-15))

    b = [w[i] * y[i] for i in range(n)]
    t = [0.0] * n
    for i in range(n):
        v = b[i]
        if i >= 1: v -= l1[i] * t[i - 1]
        if i >= 2: v -= l2[i] * t[i - 2]
        t[i] = v / l0[i]
    z = [0.0] * n
    for i in range(n - 1, -1, -1):
        v = t[i]
        if i + 1 < n: v -= l1[i + 1] * z[i + 1]
        if i + 2 < n: v -= l2[i + 2] * z[i + 2]
        z[i] = v / l0[i]
    return z


def hingefit(o, z, w, c):
    # weighted least squares fit of  z = a + sL*min(o-c,0) + sR*max(o-c,0)
    # returns (a, sL, sR, sse) or None if a side doesn't have enough points
    swl = sol = sool = szl = sozl = 0.0
    swr = sor = soor = szr = sozr = 0.0
    szz = 0.0
    nl = nr = 0
    for i in range(len(o)):
        wi = w[i]; oi = o[i] - c; zi = z[i]
        szz += wi * zi * zi
        if oi < 0:
            nl += 1; swl += wi; sol += wi * oi; sool += wi * oi * oi; szl += wi * zi; sozl += wi * oi * zi
        else:
            nr += 1; swr += wi; sor += wi * oi; soor += wi * oi * oi; szr += wi * zi; sozr += wi * oi * zi
    return solvehinge(nl, swl, sol, sool, szl, sozl, nr, swr, sor, soor, szr, sozr, szz)


def solvehinge(nl, swl, sul, suul, szl, suzl, nr, swr, sur, suur, szr, suzr, szz):
    # sums are already relative to the hinge, u = o - c
    if nl < 3 or nr < 3 or suul < 1e-6 or suur < 1e-6:
        return None
    wt = swl + swr
    zt = szl + szr
    den = wt - sul * sul / suul - sur * sur / suur
    if abs(den) < 1e-12:
        return None
    a = (zt - sul * suzl / suul - sur * suzr / suur) / den
    sl = (suzl - a * sul) / suul
    sr = (suzr - a * sur) / suur
    sse = szz - (a * zt + sl * suzl + sr * suzr)
    return (a, sl, sr, sse)


def planefit(o, z, w, c):
    # z = a + s*(o-c), returned in hinge form (a, s, s, sse)
    sw = so = soo = sz = soz = szz = 0.0
    for i in range(len(o)):
        wi = w[i]; oi = o[i] - c; zi = z[i]
        sw += wi; so += wi * oi; soo += wi * oi * oi; sz += wi * zi; soz += wi * oi * zi; szz += wi * zi * zi
    den = sw * soo - so * so
    if len(o) < 3 or abs(den) < 1e-12:
        return None
    s = (sw * soz - so * sz) / den
    a = (sz - s * so) / sw
    sse = szz - (a * sz + s * soz)
    return (a, s, s, sse)


def bestcrown(o, z, w, step, cmin, cmax):
    # search the hinge position between cmin and cmax with the smallest residual
    # o must be sorted, prefix sums make each candidate O(1)
    m = len(o)
    if m < 6:
        return None
    pw = [0.0] * (m + 1); po = [0.0] * (m + 1); poo = [0.0] * (m + 1)
    pz = [0.0] * (m + 1); poz = [0.0] * (m + 1)
    szz = 0.0
    for i in range(m):
        wi = w[i]; oi = o[i]; zi = z[i]
        pw[i + 1] = pw[i] + wi; po[i + 1] = po[i] + wi * oi; poo[i + 1] = poo[i] + wi * oi * oi
        pz[i + 1] = pz[i] + wi * zi; poz[i + 1] = poz[i] + wi * oi * zi
        szz += wi * zi * zi

    best = None
    c = max(o[2], cmin) + step / 2.0
    p = 0
    while c < min(o[m - 3], cmax):
        while p < m and o[p] < c:
            p += 1
        # left part [0,p), right part [p,m) - shift all sums to u = o - c
        swl = pw[p]; sol = po[p]; sool = poo[p]; szl = pz[p]; sozl = poz[p]
        swr = pw[m] - swl; sor = po[m] - sol; soor = poo[m] - sool; szr = pz[m] - szl; sozr = poz[m] - sozl
        res = solvehinge(p, swl, sol - c * swl, sool - 2 * c * sol + c * c * swl, szl, sozl - c * szl,
                         m - p, swr, sor - c * swr, soor - 2 * c * sor + c * c * swr, szr, sozr - c * szr, szz)
        if res != None and (best == None or res[3] < best[1][3]):
            best = (c, res)
        c += step
    return best


def upperenvelope(y, length, h):
    # smooth curve that stays on or above y everywhere
    # the smoothing targets get pushed up wherever the curve is still below y
    n = len(y)
    ones = [1.0] * n
    target = list(y)
    for it in range(100):
        z = whittaker(target, ones, length, h)
        short = max([y[k] - z[k] for k in range(n)])
        if short <= 0.0002:
            break
        target = [target[k] + max(0.0, y[k] - z[k]) for k in range(n)]
    # close the last small gap with a constant shift
    short = max(0.0, max([y[k] - z[k] for k in range(n)]))
    return [z[k] + short for k in range(n)]


def bandsmooth(lo, hi, length, h, target=None):
    # smooth profile that stays between lo and hi where possible (asymmetric least squares)
    # without target everything pulls softly towards lo, with target towards target - only where lo exists
    # violated limits get a strong weight, repeated until nothing changes
    # positions without limits (lo = None) are interpolated from their neighbours
    n = len(lo)
    strong = 1e4
    state = [0] * n     # 0 inside, 1 below lo, 2 above hi
    z = [0.0] * n
    for it in range(50):
        y = [0.0] * n; w = [0.0] * n
        for k in range(n):
            if lo[k] == None: continue
            if state[k] == 1:
                y[k] = lo[k]; w[k] = strong
            elif state[k] == 2:
                y[k] = hi[k]; w[k] = strong
            else:
                y[k] = lo[k] if target == None else target[k]; w[k] = 1.0
        z = whittaker(y, w, length, h)
        changed = 0
        for k in range(n):
            s = 0
            if lo[k] != None:
                if z[k] < lo[k]: s = 1
                elif hi[k] != None and z[k] > hi[k]: s = 2
            if s != state[k]:
                state[k] = s; changed += 1
        if changed == 0:
            break
    return z


def rowfit(u, target, lo, hi, single):
    # lift across one cross section: one straight line (single crossfall) or two straight lines meeting at the crown
    # u are the offsets from the crown, fitted to target, kept between lo and hi where possible
    # returns (lift at crown, slope left, slope right)
    n = len(u)
    m = 2 if single else 3
    useall = all([v == None for v in lo])     # nothing below this cross section - just follow the target
    state = [0] * n
    x = [0.0, 0.0, 0.0]
    for it in range(30):
        A = [[0.0] * m for p in range(m)]
        b = [0.0] * m
        for p in range(1, m):
            A[p][p] += 1e-6     # keeps a side without any points solvable
        for j in range(n):
            if lo[j] == None and not useall: continue
            f = [1.0, u[j]] if single else [1.0, min(u[j], 0.0), max(u[j], 0.0)]
            if state[j] == 1:
                y = lo[j]; w = 1e4
            elif state[j] == 2:
                y = hi[j]; w = 1e4
            else:
                y = target[j]; w = 1.0
            for p in range(m):
                b[p] += w * f[p] * y
                for q in range(m):
                    A[p][q] += w * f[p] * f[q]
        # gaussian elimination, the system is tiny
        for p in range(m):
            for q in range(p + 1, m):
                fq = A[q][p] / A[p][p]
                for r in range(p, m):
                    A[q][r] -= fq * A[p][r]
                b[q] -= fq * b[p]
        sol = [0.0] * m
        for p in range(m - 1, -1, -1):
            sol[p] = (b[p] - sum([A[p][r] * sol[r] for r in range(p + 1, m)])) / A[p][p]
        x = [sol[0], sol[1], sol[1]] if single else sol
        changed = 0
        for j in range(n):
            s = 0
            if lo[j] != None:
                v = x[0] + (x[1] * u[j] if u[j] < 0 else x[2] * u[j])
                if v < lo[j]: s = 1
                elif hi[j] != None and v > hi[j]: s = 2
            if s != state[j]:
                state[j] = s; changed += 1
        if changed == 0:
            break
    return x


def boxfilter(y, half):
    # moving average over +- half stations, the ends are extended with their end values
    n = len(y)
    if half < 1:
        return list(y)
    prefix = [0.0]
    for i in range(-half, n + half):
        prefix.Add(prefix[-1] + y[min(max(i, 0), n - 1)])
    return [(prefix[k + 2 * half + 1] - prefix[k]) / (2 * half + 1) for k in range(n)]


def transition(y, length, h):
    # turns steps into S-shaped transitions of the given total length, centred on the step
    # two moving averages of half the length each, so the grade changes gradually at both ends
    half = int(round(length / 4.0 / h))
    return boxfilter(boxfilter(y, half), half)


def interp(values, h, s):
    # linear interpolation in an equally spaced list
    k = s / h
    i = int(math.floor(k))
    if i < 0: return values[0]
    if i >= len(values) - 1: return values[-1]
    f = k - i
    return values[i] * (1 - f) + values[i + 1] * f


class SCR_SubgradeSmoothing(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_SubgradeSmoothing.xaml") as s:
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

        types = Array [Type] (SurfaceTypeLists.AllWithCutFillMap)
        self.surfacepicker.FilterByEntityTypes = types
        self.surfacepicker.AllowNone = False

        self.lType = clr.GetClrType(IPolyseg)
        self.linepicker.IsEntityValidCallback = self.IsValid
        self.linepicker.ValueChanged += self.lineChanged

        # crown mode sections along the centreline
        self.sections = DataTable()
        self.sections.Columns.Add("from", System.Type.GetType("System.Double"))
        self.sections.Columns.Add("to",   System.Type.GetType("System.Double"))
        self.sections.Columns.Add("mode", System.Type.GetType("System.String"))
        self.sections.Columns.Add("depth", System.Type.GetType("System.String"))   # text, so it can stay empty
        self.sections.DefaultView.Sort = "from ASC"
        self.sectiongrid.DataContext = self.sections
        self.addsection.Click += self.addsectionClicked
        self.removesection.Click += self.removesectionClicked

        self.voffset.DistanceMin = -1000.0
        self.extension.DistanceMin = 0.0
        self.depthtransition.DistanceMin = 0.0
        for c in [self.interval, self.fitwindow, self.crossspacing, self.maxextra]:
            c.DistanceMin = 0.01
        for c in [self.smoothcl, self.smoothgrade, self.smoothxfall, self.smoothcrown, self.smoothwidth]:
            c.DistanceMin = 0.0

        self.fixcl.Checked += self.fixclChanged
        self.fixcl.Unchecked += self.fixclChanged

        try:
            self.SetDefaultOptions()
        except:
            pass
        self.fixclChanged(None, None)

    def fixclChanged(self, sender, e):
        # a fixed centreline is used as it is - no horizontal smoothing, no crown search
        free = not self.fixcl.IsChecked
        self.crownfit.IsEnabled = free     # single crossfall still makes sense on a fixed centreline
        self.smoothcl.IsEnabled = free
        self.smoothcrown.IsEnabled = free

    def SetDefaultOptions(self):
        # the sections are chainages of a specific road, so they are stored with the project
        settings = ConstructionCommandsSettings.ProvideObject(self.currentProject)
        for entry in settings.GetString("SCR_SubgradeSmoothing.sections", "").split("|"):
            values = entry.split(";")
            if len(values) >= 3 and values[2] in CROWNMODES:
                self.addsectionrow(float(values[0]), float(values[1]), values[2], values[3] if len(values) > 3 else "")
        SCROptions.LoadMacroOptions(self, "SCR_SubgradeSmoothing", _OPTIONS, self.currentProject)

    def SaveOptions(self):
        settings = ConstructionCommandsSettings.ProvideObject(self.currentProject)
        settings.SetString("SCR_SubgradeSmoothing.sections",
                           "|".join(["{!r};{!r};{};{}".format(r["from"], r["to"], r["mode"], self.sectiondepthtext(r)) for r in self.sections.Rows]))
        SCROptions.SaveMacroOptions(self, "SCR_SubgradeSmoothing", _OPTIONS)

    def IsValid(self, serial):
        o = self.currentProject.Concordance[serial]
        return isinstance(o, self.lType)

    def lineChanged(self, ctrl, e):
        l1 = self.linepicker.Entity
        if l1 != None:
            self.fromstation.StationProvider = l1
            self.tostation.StationProvider = l1

    def defaultmode(self):
        if self.crownnone.IsChecked: return CROWNMODES[2]
        if self.fixcl.IsChecked or self.crowncl.IsChecked: return CROWNMODES[1]
        return CROWNMODES[0]

    def addsectionrow(self, fromst, tost, mode, depth=""):
        if tost < fromst: fromst, tost = tost, fromst
        row = self.sections.NewRow()
        row["from"] = fromst
        row["to"] = tost
        row["mode"] = mode
        row["depth"] = depth
        self.sections.Rows.Add(row)

    def sectiondepthtext(self, r):
        # DBNull if the cell was never filled
        return "" if r.IsNull("depth") else str(r["depth"]).strip()

    def sectiondepth(self, r):
        # depth offset of a section, None if empty or not a number - then the global value applies
        try:
            return float(self.sectiondepthtext(r).replace(",", "."))
        except:
            return None

    def addsectionClicked(self, sender, e):
        self.addsectionrow(self.fromstation.Distance, self.tostation.Distance, self.defaultmode())

    def removesectionClicked(self, sender, e):
        rows = [item.Row for item in list(self.sectiongrid.SelectedItems)]
        for r in rows:
            self.sections.Rows.Remove(r)

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content = ''
        self.success.Content = ''

        wv = self.currentProject [Project.FixedSerial.WorldView]
        surface = wv.Lookup(self.surfacepicker.SelectedSerial)
        l1 = self.linepicker.Entity
        if surface == None:
            self.error.Content = 'no Surface selected'
            return
        if l1 == None:
            self.error.Content = 'no Centreline selected'
            return

        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        wv.PauseGraphicsCache(True)

        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                if self.compute(wv, surface, l1):
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
        wv.PauseGraphicsCache(False)

    def compute(self, wv, surface, l1):

        h = self.interval.Distance
        fitwindow = self.fitwindow.Distance
        fixcl = self.fixcl.IsChecked

        polyseg = l1.ComputePolySeg().ToWorld()
        beginstation = polyseg.BeginStation
        cllength = polyseg.Length()
        outpoint = clr.StrongBox[Point3D]()

        if fixcl:
            # ---- 1. fixed centreline - sample the line as it is at the station interval
            n = max(3, int(round(cllength / h)) + 1)
            h = cllength / (n - 1)
            px = []; py = []; chainage = []
            for k in range(n):
                st = beginstation + min(k * h, cllength)
                if not polyseg.FindPointFromStation(st, outpoint):
                    self.error.Content = 'could not compute the Centreline at Chainage {:.3f}'.format(st)
                    return False
                px.Add(outpoint.Value.X); py.Add(outpoint.Value.Y); chainage.Add(st)
        else:
            (n, h, px, py, chainage) = self.smoothcentreline(polyseg, beginstation, cllength, h, outpoint)
            if n == 0:
                self.error.Content = 'could not compute the Centreline'
                return False

        # crown mode per station - default first, then the sections from the table
        # same for the depth offset, an empty depth cell keeps the global value
        modes = [CROWNMODES.index(self.defaultmode())] * n
        depth = [self.voffset.Distance] * n
        for r in self.sections.DefaultView:
            m = CROWNMODES.index(r["mode"]) if r["mode"] in CROWNMODES else 0
            if fixcl and m == 0: m = 1     # no crown search on a fixed centreline
            d = self.sectiondepth(r.Row)
            for k in range(n):
                if r["from"] <= chainage[k] <= r["to"]:
                    modes[k] = m
                    if d != None: depth[k] = d
        depth = transition(depth, self.depthtransition.Distance, h)

        return self.computesurface(wv, surface, n, h, px, py, chainage, modes, depth, fitwindow, fixcl)

    def smoothcentreline(self, polyseg, beginstation, cllength, h, outpoint):
        # ---- 1. centreline - densify, smooth horizontally, resample at the station interval
        densestep = min(h, 1.0) / 2.0
        ndense = max(2, int(math.ceil(cllength / densestep)) + 1)
        cx = []; cy = []; cst = []
        for i in range(ndense):
            st = beginstation + cllength * i / (ndense - 1)
            if polyseg.FindPointFromStation(st, outpoint):
                cx.Add(outpoint.Value.X); cy.Add(outpoint.Value.Y); cst.Add(st)
        if cx.Count < 3:
            return (0, h, None, None, None)

        ones = [1.0] * cx.Count
        cx = whittaker(cx, ones, self.smoothcl.Distance, densestep)
        cy = whittaker(cy, ones, self.smoothcl.Distance, densestep)

        # cumulative length of the smoothed line, then resample equally spaced
        cum = [0.0]
        for i in range(1, cx.Count):
            cum.Add(cum[-1] + math.hypot(cx[i] - cx[i - 1], cy[i] - cy[i - 1]))
        total = cum[-1]
        n = max(3, int(round(total / h)) + 1)
        h = total / (n - 1)
        px = []; py = []
        chainage = []   # chainage on the selected line, the smoothed line is shorter in bends
        j = 0
        for k in range(n):
            s = k * h
            while j < cx.Count - 2 and cum[j + 1] < s:
                j += 1
            seglen = cum[j + 1] - cum[j]
            f = (s - cum[j]) / seglen if seglen > 0 else 0.0
            px.Add(cx[j] + (cx[j + 1] - cx[j]) * f)
            py.Add(cy[j] + (cy[j + 1] - cy[j]) * f)
            chainage.Add(cst[j] + (cst[j + 1] - cst[j]) * f)
        return (n, h, px, py, chainage)

    def computesurface(self, wv, surface, n, h, px, py, chainage, modes, depth, fitwindow, fixcl):
        # unit tangents and right hand normals
        tx = []; ty = []
        for k in range(n):
            k0 = max(k - 1, 0); k1 = min(k + 1, n - 1)
            dx = px[k1] - px[k0]; dy = py[k1] - py[k0]
            d = math.hypot(dx, dy)
            tx.Add(dx / d); ty.Add(dy / d)

        # ---- 2. source vertices -> station/offset
        ProgressBar.TBC_ProgressBar.Title = "collecting Surface Vertices"
        used = set()
        tris = []   # triangles with material, needed later for the source elevation under the new vertices
        for t in range(surface.NumberOfTriangles):
            if surface.IsTriangleMaterialPresent(t):
                tri = (surface.GetTriangleIVertex(t, 0), surface.GetTriangleIVertex(t, 1), surface.GetTriangleIVertex(t, 2))
                tris.Add(tri)
                for v in tri:
                    used.add(v)
        used = list(used)
        vpoints = [surface.GetVertexPoint(i) for i in used]
        vindex = dict(zip(used, range(used.Count)))
        tris = [(vindex[t[0]], vindex[t[1]], vindex[t[2]]) for t in tris]

        # bucket grid over the centreline segments for a fast nearest segment search
        cell = 5.0
        grid = {}
        for k in range(n - 1):
            ix0 = int(math.floor(min(px[k], px[k + 1]) / cell)); ix1 = int(math.floor(max(px[k], px[k + 1]) / cell))
            iy0 = int(math.floor(min(py[k], py[k + 1]) / cell)); iy1 = int(math.floor(max(py[k], py[k + 1]) / cell))
            for ix in range(ix0, ix1 + 1):
                for iy in range(iy0, iy1 + 1):
                    grid.setdefault((ix, iy), []).Add(k)
        gxmin = min([g[0] for g in grid]); gxmax = max([g[0] for g in grid])
        gymin = min([g[1] for g in grid]); gymax = max([g[1] for g in grid])

        ProgressBar.TBC_ProgressBar.Title = "projecting " + vpoints.Count.ToString("N0") + " Vertices onto Centreline"
        pts = []    # (station, offset, z)
        beyondends = 0
        last_t = timer()
        for i in range(vpoints.Count):
            if timer() - last_t > 0.5:
                if ProgressBar.TBC_ProgressBar.SetProgress(i * 100 // vpoints.Count):
                    return False
                last_t = timer()
            p = vpoints[i]
            if not p.Is3D: continue
            gx = int(math.floor(p.X / cell)); gy = int(math.floor(p.Y / cell))
            # search rings of cells around the point until no closer segment is possible
            rmax = max(abs(gx - gxmin), abs(gx - gxmax), abs(gy - gymin), abs(gy - gymax))
            bestd = None
            r = 0
            while r <= rmax:
                for ix in range(gx - r, gx + r + 1):
                    for iy in range(gy - r, gy + r + 1):
                        if max(abs(ix - gx), abs(iy - gy)) != r: continue
                        segs = grid.get((ix, iy))
                        if segs == None: continue
                        for k in segs:
                            ax = px[k]; ay = py[k]
                            dx = px[k + 1] - ax; dy = py[k + 1] - ay
                            f = ((p.X - ax) * dx + (p.Y - ay) * dy) / (dx * dx + dy * dy)
                            f = min(max(f, 0.0), 1.0)
                            ex = p.X - (ax + dx * f); ey = p.Y - (ay + dy * f)
                            d2 = ex * ex + ey * ey
                            if bestd == None or d2 < bestd:
                                bestd = d2; bestk = k; bestf = f
                                bestside = dx * (p.Y - ay) - dy * (p.X - ax)  # > 0 is left of the line
                # anything in further rings is at least r * cell away
                if bestd != None and bestd <= (r * cell) ** 2: break
                r += 1
            if bestd == None: continue
            # points before the start or after the end of the centreline would distort the widths
            if (bestk == 0 and bestf == 0.0) or (bestk == n - 2 and bestf == 1.0):
                beyondends += 1
                continue
            o = math.sqrt(bestd)
            if bestside > 0: o = -o
            pts.Add(((bestk + bestf) * h, o, p.Z))

        if pts.Count < 10:
            self.error.Content = 'not enough Surface Vertices along the Centreline'
            return False
        pts.sort()

        # ---- 3. per station fit
        sigma = fitwindow / 4.0
        half = fitwindow / 2.0
        widthhalf = fitwindow   # widths come from a window twice as long, edge shots are sparse

        rawa = [0.0] * n; rawsl = [0.0] * n; rawsr = [0.0] * n; rawc = [0.0] * n
        rawwl = [0.0] * n; rawwr = [0.0] * n
        wfit = [0.0] * n; wcrown = [0.0] * n; wwidth = [0.0] * n
        npts = [0] * n; rms = [0.0] * n

        ProgressBar.TBC_ProgressBar.Title = "fitting Cross Sections"
        lo = 0; hi = 0; wlo = 0; whi = 0
        last_t = timer()
        for k in range(n):
            if timer() - last_t > 0.5:
                if ProgressBar.TBC_ProgressBar.SetProgress(k * 100 // n):
                    return False
                last_t = timer()
            s = k * h
            while lo < pts.Count and pts[lo][0] < s - half: lo += 1
            while hi < pts.Count and pts[hi][0] <= s + half: hi += 1
            while wlo < pts.Count and pts[wlo][0] < s - widthhalf: wlo += 1
            while whi < pts.Count and pts[whi][0] <= s + widthhalf: whi += 1

            # widths - distance of the outermost points left and right of the centreline
            # negative if the centreline runs outside the road, i.e. on the inside of a smoothed tight bend
            wl = -1e9; wr = -1e9
            for i in range(wlo, whi):
                o = pts[i][1]
                if -o > wl: wl = -o
                if o > wr: wr = o
            if whi > wlo:
                rawwl[k] = wl; rawwr[k] = wr; wwidth[k] = 1.0

            window = sorted([(pts[i][1], pts[i][2], math.exp(-0.5 * ((pts[i][0] - s) / sigma) ** 2)) for i in range(lo, hi)])
            npts[k] = window.Count
            if window.Count < 3: continue
            o = [q[0] for q in window]; z = [q[1] for q in window]; w = [q[2] for q in window]

            res = None
            if modes[k] == 0:
                # only search the middle half of the section, otherwise the roll off at the edges is mistaken for a crown
                quarter = (o[-1] - o[0]) / 4.0
                best = bestcrown(o, z, w, 0.25, o[0] + quarter, o[-1] - quarter)
                if best != None:
                    rawc[k] = best[0]; res = best[1]
                    # a distinct crown gets a high weight, an ambiguous one (one way crossfall) almost none
                    wcrown[k] = min(abs(res[2] - res[1]) / 0.04, 1.0) ** 2
            elif modes[k] == 1:
                res = hingefit(o, z, w, 0.0)
                wcrown[k] = 1.0     # crown is fixed on the centreline
            if res == None:
                res = planefit(o, z, w, rawc[k])
            if res == None: continue

            rawa[k], rawsl[k], rawsr[k] = res[0], res[1], res[2]
            rms[k] = math.sqrt(max(res[3], 0.0) / sum(w))
            wfit[k] = 1.0

        if sum(wfit) < 2:
            self.error.Content = 'not enough Cross Sections could be fitted'
            return False

        # ---- 4. smoothing
        ProgressBar.TBC_ProgressBar.Title = "smoothing"
        # ambiguous crown positions and single crossfall sections are pulled softly towards the centreline
        # so the crown position moves smoothly between the different section types
        if fixcl:
            c = [0.0] * n   # crown stays on the given line
        else:
            base = 0.01
            ctarget = [(wcrown[k] * rawc[k]) / (wcrown[k] + base) for k in range(n)]
            c = whittaker(ctarget, [wcrown[k] + base for k in range(n)], self.smoothcrown.Distance, h)

        # refit with the smoothed crown position fixed
        lo = 0; hi = 0
        for k in range(n):
            s = k * h
            while lo < pts.Count and pts[lo][0] < s - half: lo += 1
            while hi < pts.Count and pts[hi][0] <= s + half: hi += 1
            if wfit[k] == 0: continue
            o = [pts[i][1] for i in range(lo, hi)]; z = [pts[i][2] for i in range(lo, hi)]
            w = [math.exp(-0.5 * ((pts[i][0] - s) / sigma) ** 2) for i in range(lo, hi)]
            res = hingefit(o, z, w, c[k]) if modes[k] != 2 else None
            if res == None:
                res = planefit(o, z, w, c[k])
            if res != None:
                rawa[k], rawsl[k], rawsr[k] = res[0], res[1], res[2]
                rms[k] = math.sqrt(max(res[3], 0.0) / sum(w))

        a = whittaker(rawa, wfit, self.smoothgrade.Distance, h)
        sl = whittaker(rawsl, wfit, self.smoothxfall.Distance, h)
        sr = whittaker(rawsr, wfit, self.smoothxfall.Distance, h)
        wl = whittaker(rawwl, wwidth, self.smoothwidth.Distance, h)
        wr = whittaker(rawwr, wwidth, self.smoothwidth.Distance, h)

        extension = self.extension.Distance
        for k in range(n):
            # keep a minimum width of 0.2 m
            grow = max(0.0, 0.2 - (wl[k] + wr[k])) / 2.0
            wl[k] += grow + extension
            wr[k] += grow + extension
            if fixcl:
                # the crown must stay on the line, so widen the road instead if the line runs along an edge
                wl[k] = max(wl[k], 0.05); wr[k] = max(wr[k], 0.05)
            else:
                c[k] = min(max(c[k], -wl[k] + 0.05), wr[k] - 0.05)

        # only build where we had data
        fitted = [k for k in range(n) if wfit[k] > 0]
        kfirst = fitted[0]; klast = fitted[-1]

        # ---- 4a. plan position of the output vertices, the elevations come later
        spacing = self.crossspacing.Distance
        nl = max(1, int(math.ceil(max([wl[k] + c[k] for k in range(kfirst, klast + 1)]) / spacing)))
        nr = max(1, int(math.ceil(max([wr[k] - c[k] for k in range(kfirst, klast + 1)]) / spacing)))
        ncol = nl + nr + 1

        gridxy = []     # rows of (x, y, offset) or None where the offset line folds back
        folded = 0
        lastvalid = [None] * ncol
        for k in range(kfirst, klast + 1):
            row = []
            nx = ty[k]; ny = -tx[k]   # right hand normal
            for j in range(ncol):
                if j <= nl:
                    o = -wl[k] + (c[k] + wl[k]) * j / nl
                else:
                    o = c[k] + (wr[k] - c[k]) * (j - nl) / nr
                x = px[k] + nx * o; y = py[k] + ny * o
                # on the inside of tight bends offset lines may run backwards - skip those points
                q = lastvalid[j]
                if q is not None and (x - q[0]) * tx[k] + (y - q[1]) * ty[k] <= 0.001:
                    row.Add(None)
                    folded += 1
                    continue
                lastvalid[j] = (x, y)
                row.Add((x, y, o))
            gridxy.Add(row)

        # ---- 5. QA - deviation of the source vertices from the smoothed model
        sumsq = 0.0; maxdev = 0.0; cnt = 0
        need = [0.0] * n    # how much each station has to come up so no source vertex sticks out above the model
        devs = []
        for (s, o, z) in pts:
            k = s / h
            if k < kfirst or k > klast: continue
            ck = interp(c, h, s)
            if o < -interp(wl, h, s) or o > interp(wr, h, s): continue
            u = o - ck
            zm = interp(a, h, s) + (interp(sl, h, s) * u if u < 0 else interp(sr, h, s) * u)
            d = z - zm
            sumsq += d * d; cnt += 1
            if abs(d) > abs(maxdev): maxdev = d
            devs.Add((s, o, d))
            if d > 0:
                # the model between two stations is interpolated, so both have to come up
                i = int(math.floor(k))
                need[i] = max(need[i], d)
                if i + 1 < n: need[i + 1] = max(need[i + 1], d)

        # between its vertices the source is flat triangles, the result is smooth - in sags and with long source
        # triangles the result can get closer than the vertical offset there, so every new vertex gets checked too
        gsamples = []   # (row, column, d) - source under every new vertex
        if self.keepabove.IsChecked:
            ProgressBar.TBC_ProgressBar.Title = "sampling Source at the new Vertices"
            # interpolated from the source triangles directly, PickSurface works in 3D on non-DTM surfaces
            tcell = 5.0
            tgrid = {}
            for t in range(tris.Count):
                p0 = vpoints[tris[t][0]]; p1 = vpoints[tris[t][1]]; p2 = vpoints[tris[t][2]]
                for ix in range(int(math.floor(min(p0.X, p1.X, p2.X) / tcell)), int(math.floor(max(p0.X, p1.X, p2.X) / tcell)) + 1):
                    for iy in range(int(math.floor(min(p0.Y, p1.Y, p2.Y) / tcell)), int(math.floor(max(p0.Y, p1.Y, p2.Y) / tcell)) + 1):
                        tgrid.setdefault((ix, iy), []).Add(t)

            def sourcez(x, y):
                for t in tgrid.get((int(math.floor(x / tcell)), int(math.floor(y / tcell))), []):
                    p0 = vpoints[tris[t][0]]; p1 = vpoints[tris[t][1]]; p2 = vpoints[tris[t][2]]
                    den = (p1.Y - p2.Y) * (p0.X - p2.X) + (p2.X - p1.X) * (p0.Y - p2.Y)
                    if abs(den) < 1e-12: continue
                    b0 = ((p1.Y - p2.Y) * (x - p2.X) + (p2.X - p1.X) * (y - p2.Y)) / den
                    b1 = ((p2.Y - p0.Y) * (x - p2.X) + (p0.X - p2.X) * (y - p2.Y)) / den
                    b2 = 1.0 - b0 - b1
                    if b0 >= -1e-9 and b1 >= -1e-9 and b2 >= -1e-9:
                        return b0 * p0.Z + b1 * p1.Z + b2 * p2.Z
                return None

            for r in range(gridxy.Count):
                k = kfirst + r
                for j in range(ncol):
                    q = gridxy[r][j]
                    if q is None: continue
                    zs = sourcez(q[0], q[1])
                    if zs == None: continue
                    u = q[2] - c[k]
                    d = zs - (a[k] + (sl[k] * u if u < 0 else sr[k] * u))
                    gsamples.Add((r, j, d))
                    if d > 0: need[k] = max(need[k], d)

        # ---- 5a. optional lift so the result never gets closer than the vertical offset
        # each new vertex has its own lift; with a max offset it gets smoothed along the road with the Grade length
        # and across with the Crossfall length, so the cross sections can bend where the source needs it
        nrows = gridxy.Count
        lift = [[0.0] * ncol for r in range(nrows)]
        maxlift = 0.0; minclear = None; maxclear = None; overmax = 0
        tol = self.maxextra.Distance if self.keepabove.IsChecked and self.limitmax.IsChecked else None

        def colpos(k, o):
            # fractional column of an offset in the row of station k
            if o < c[k]: return nl * (o + wl[k]) / max(c[k] + wl[k], 0.001)
            return nl + nr * (o - c[k]) / max(wr[k] - c[k], 0.001)

        def liftat(s, o):
            # lift at any position, interpolated like the triangles of the new surface
            rf = min(max(s / h - kfirst, 0.0), nrows - 1.0)
            r0 = min(int(rf), max(nrows - 2, 0)); fr = rf - r0
            v = 0.0
            for (r, wrow) in ((r0, 1 - fr), (min(r0 + 1, nrows - 1), fr)):
                jf = min(max(colpos(kfirst + r, o), 0.0), ncol - 1.0)
                j0 = min(int(jf), ncol - 2); fj = jf - j0
                v += wrow * ((1 - fj) * lift[r][j0] + fj * lift[r][j0 + 1])
            return v

        if self.keepabove.IsChecked:
            if tol != None:
                ProgressBar.TBC_ProgressBar.Title = "fitting between min and max Offset"
                lo = [[None] * ncol for r in range(nrows)]
                hi = [[None] * ncol for r in range(nrows)]
                for (r, j, d) in gsamples:
                    lo[r][j] = d if lo[r][j] == None else max(lo[r][j], d)
                    hi[r][j] = d + tol if hi[r][j] == None else min(hi[r][j], d + tol)
                # source vertices between the new vertices - their minimum goes onto the new vertices around them
                for (s, o, d) in devs:
                    rf = s / h - kfirst
                    r0 = int(math.floor(rf)); fr = rf - r0
                    for (r, wrow) in ((r0, 1 - fr), (r0 + 1, fr)):
                        if wrow < 0.25 or r < 0 or r >= nrows: continue
                        jf = colpos(kfirst + r, o)
                        j0 = int(math.floor(jf)); fj = jf - j0
                        for (j, wcol) in ((j0, 1 - fj), (j0 + 1, fj)):
                            if wcol < 0.25 or j < 0 or j >= ncol: continue
                            lo[r][j] = d if lo[r][j] == None else max(lo[r][j], d)
                # the minimum wins where both can't be met
                for r in range(nrows):
                    for j in range(ncol):
                        if hi[r][j] != None and hi[r][j] < lo[r][j]: hi[r][j] = lo[r][j]
                for j in range(ncol):
                    col = bandsmooth([lo[r][j] for r in range(nrows)], [hi[r][j] for r in range(nrows)], self.smoothgrade.Distance, h)
                    for r in range(nrows): lift[r][j] = col[r]
                # the cross sections keep their shape - one straight line for single crossfall, two meeting at the crown
                # fit that shape within the limits, smooth crown lift and slopes along the road, repeat
                us = []
                for r in range(nrows):
                    k = kfirst + r
                    us.Add([(-wl[k] + (c[k] + wl[k]) * j / nl if j <= nl else c[k] + (wr[k] - c[k]) * (j - nl) / nr) - c[k] for j in range(ncol)])
                ones = [1.0] * nrows
                target = lift
                for it in range(20):
                    fits = [rowfit(us[r], target[r], lo[r], hi[r], modes[kfirst + r] == 2) for r in range(nrows)]
                    lc = whittaker([f[0] for f in fits], ones, self.smoothgrade.Distance, h)
                    ql = whittaker([f[1] for f in fits], ones, self.smoothxfall.Distance, h)
                    qr = whittaker([f[2] for f in fits], ones, self.smoothxfall.Distance, h)
                    # no forcing ql = qr in single crossfall sections - the fits there already have one slope,
                    # the smoothing only makes them differ near a section change, i.e. a gradual transition
                    lift = [[lc[r] + (ql[r] * u if u < 0 else qr[r] * u) for u in us[r]] for r in range(nrows)]
                    # the smoothing moved some vertices out of their limits again - shift their targets by that much
                    worst = 0.0
                    for r in range(nrows):
                        for j in range(ncol):
                            if lo[r][j] == None: continue
                            v = lo[r][j] - lift[r][j] if lift[r][j] < lo[r][j] else (hi[r][j] - lift[r][j] if hi[r][j] != None and lift[r][j] > hi[r][j] else 0.0)
                            if v != 0.0:
                                target[r][j] += v; worst = max(worst, abs(v))
                    if worst < 0.002: break
            else:
                env = upperenvelope(need, self.smoothgrade.Distance, h)
                lift = [[env[kfirst + r]] * ncol for r in range(nrows)]

            # anything still short of the minimum gets a smooth local raise
            short = [0.0] * n
            for (r, j, d) in gsamples:
                v = d - lift[r][j]
                if v > 0: short[kfirst + r] = max(short[kfirst + r], v)
            for (s, o, d) in devs:
                v = d - liftat(s, o)
                if v > 0:
                    i = int(math.floor(s / h))
                    short[i] = max(short[i], v)
                    if i + 1 < n: short[i + 1] = max(short[i + 1], v)
            if max(short) > 0:
                bump = upperenvelope(short, self.smoothgrade.Distance, h)
                lift = [[v + bump[kfirst + r] for v in lift[r]] for r in range(nrows)]

            for (r, j, d) in gsamples:
                q = lift[r][j] - d; clear = q + depth[kfirst + r]
                if minclear == None or clear < minclear: minclear = clear
                if maxclear == None or clear > maxclear: maxclear = clear
                if tol != None and q > tol + 0.001: overmax += 1
            for (s, o, d) in devs:
                clear = liftat(s, o) - d + interp(depth, h, s)
                if minclear == None or clear < minclear: minclear = clear
                if maxclear == None or clear > maxclear: maxclear = clear
            maxlift = max([max(row) for row in lift])

        # ---- 6. output surface
        ProgressBar.TBC_ProgressBar.Title = "building Surface"
        grid3d = []     # rows of Point3D or None where the offset line folds back
        for r in range(nrows):
            k = kfirst + r
            row = []
            for j in range(ncol):
                q = gridxy[r][j]
                if q is None:
                    row.Add(None)
                    continue
                u = q[2] - c[k]
                row.Add(Point3D(q[0], q[1], a[k] + (sl[k] * u if u < 0 else sr[k] * u) + lift[r][j] + depth[k]))
            grid3d.Add(row)

        # crown and edge values for report and summary, the crossfalls from crown to edge
        liftc = [0.0] * n; liftl = [0.0] * n; liftr = [0.0] * n
        for r in range(nrows):
            k = kfirst + r
            liftc[k] = lift[r][nl]; liftl[k] = lift[r][0]; liftr[k] = lift[r][ncol - 1]
            sl[k] -= (liftl[k] - liftc[k]) / max(wl[k] + c[k], 0.001)
            sr[k] += (liftr[k] - liftc[k]) / max(wr[k] - c[k], 0.001)
            a[k] += liftc[k]

        newsurface = wv.Add(clr.GetClrType(Model3D))
        newsurface.Name = Model3D.GetUniqueName(surface.Name + " - smoothed", None, wv)
        builder = newsurface.GetGemBatchBuilder()
        vidx = []
        for row in grid3d:
            vidx.Add([builder.AddVertex(p)[0] if p is not None else None for p in row])

        soft = Byte(DTMSharpness.eSoft)
        bndy = Byte(DTMSharpness.eSharpAndTextureBndy)
        nrows = vidx.Count
        # longitudinal lines
        for j in range(ncol):
            prev = None
            for r in range(nrows):
                v = vidx[r][j]
                if v is None: continue
                if prev is not None:
                    builder.AddBreakline(bndy if j == 0 or j == ncol - 1 else soft, prev, v)
                prev = v
        # cross sections
        for r in range(nrows):
            prev = None
            for j in range(ncol):
                v = vidx[r][j]
                if v is None: continue
                if prev is not None:
                    builder.AddBreakline(bndy if r == 0 or r == nrows - 1 else soft, prev, v)
                prev = v

        builder.Construction()
        # null material on the convex hull triangles outside of our boundary, i.e. on the inside of bends
        gemmap = GemMaterialMap()
        for t in range(builder.NumberOfTriangles):
            for side in range(3):
                if not builder.GetTriangleOuterSide(t, side):
                    continue
                (bl, external, sharp) = builder.GetTriangleBreakline(t, side)
                if not bl:
                    builder.AttachMaterial(0, gemmap, t)
                    break
        builder.Commit()
        # keep the default display settings, but hide the breaklines in 2D and 3D
        mode = int(newsurface.Mode) & ~(int(DisplayMode.BreakLines) | int(DisplayMode.BreakLines_2D))
        newsurface.Mode = DisplayMode(mode)

        # ---- 7. optional crown and edge lines
        if self.drawlines.IsChecked:
            for j in sorted(set([0, nl, ncol - 1])):
                ls = wv.Add(clr.GetClrType(Linestring))
                ls.Layer = self.layerpicker.SelectedSerialNumber
                for row in grid3d:
                    if row[j] is None: continue
                    el = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                    el.Position = row[j]
                    ls.AppendElement(el)

        # ---- 8. optional CSV with raw and smoothed values
        csvname = os.path.expanduser('~/Downloads/SCR_SubgradeSmoothing.csv')
        if self.writecsv.IsChecked:
            with open(csvname, 'w') as f:
                f.write("Station,Mode,Points,FitRMS,CrownZ raw,CrownZ,Lift,Depth,Grade %,XfallL raw %,XfallL %,XfallR raw %,XfallR %,"
                        "Crown Offset raw,Crown Offset,Left Edge Offset raw,Left Edge Offset,Right Edge Offset raw,Right Edge Offset,"
                        "Lift Left Edge,Lift Right Edge\n")
                for k in range(n):
                    k0 = max(k - 1, 0); k1 = min(k + 1, n - 1)
                    grade = (a[k1] - a[k0]) / ((k1 - k0) * h) * 100
                    # crossfalls are given outwards from the crown, negative = falling away from the crown
                    f.write("{:.3f},{},{},{:.4f},{:.4f},{:.4f},{:.4f},{:.4f},{:.3f},{:.3f},{:.3f},{:.3f},{:.3f},{:.3f},{:.3f},{:.3f},{:.3f},{:.3f},{:.3f},{:.4f},{:.4f}\n".format(
                        chainage[k], CROWNMODES[modes[k]], npts[k], rms[k], rawa[k], a[k], liftc[k], depth[k], grade,
                        -rawsl[k] * 100, -sl[k] * 100, rawsr[k] * 100, sr[k] * 100,
                        rawc[k], c[k], -rawwl[k], -(wl[k] - extension), rawwr[k], wr[k] - extension, liftl[k], liftr[k]))

        # ---- summary
        maxdxfall = 0.0; maxdgrade = 0.0
        for k in range(kfirst + 1, klast):
            maxdxfall = max(maxdxfall, abs(sl[k + 1] - sl[k]) / h * 100, abs(sr[k + 1] - sr[k]) / h * 100)
            maxdgrade = max(maxdgrade, abs(a[k + 1] - 2 * a[k] + a[k - 1]) / (h * h) * 100)
        self.success.Content = 'created ' + newsurface.Name
        if cnt > 0:
            self.success.Content += '\nSource deviation RMS {:.3f} / max {:.3f}'.format(math.sqrt(sumsq / cnt), maxdev)
        if self.keepabove.IsChecked and minclear != None:
            self.success.Content += '\nraised by up to {:.3f}\nClearance min {:.3f} / max {:.3f}'.format(maxlift, minclear, maxclear)
            self.success.Content += '\n' + str(gsamples.Count) + ' new Vertices checked against the Source'
            if overmax > 0:
                self.success.Content += '\n' + str(overmax) + ' Points still above the max Offset'
        self.success.Content += '\nmax Crossfall change {:.2f} %/m'.format(maxdxfall)
        self.success.Content += '\nmax Grade change {:.3f} %/m'.format(maxdgrade)
        if folded > 0:
            self.success.Content += '\n' + str(folded) + ' points skipped in tight bends'
        if beyondends > 0:
            self.success.Content += '\n' + str(beyondends) + ' Vertices beyond the Centreline ends ignored'
        if self.writecsv.IsChecked:
            self.success.Content += '\nCSV written to\n' + csvname
        return True

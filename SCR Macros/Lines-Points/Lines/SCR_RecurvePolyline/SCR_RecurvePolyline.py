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
    "htol":          0.002,
    "minarcchords":  3.0,
    "maxradius":     10000.0,
    "roundradius":   0.0,
    "fitvertical":   True,
    "vtol":          0.002,
    "minvcchords":   3.0,
    "layerpicker":   8,
    "writecsv":      True,
}

def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_RecurvePolyline"
    cmdData.CommandName = "SCR_RecurvePolyline"
    cmdData.Caption = "_SCR_RecurvePolyline"
    cmdData.UIForm = "SCR_RecurvePolyline"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = ""
    cmdData.HelpTopic = ""

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "Lines"
        cmdData.ShortCaption = "Recurve Polyline"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.009
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""

        cmdData.ToolTipTitle = "Recurve Polyline"
        cmdData.ToolTipTextFormatted = "re-engineer chorded polylines into Straights/Arcs and Grades/Vertical Curves"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


# ---- geometry core - plain Python, no TBC types, can be tested outside of TBC
# all coordinates are local (shifted to the first vertex) to keep the circle math precise
# an element covers the vertex index range i..j, neighbouring elements share the boundary vertex

TWOPI = 2.0 * math.pi

def circlefrom3(ax, ay, bx, by, cx, cy):
    # circumcircle, returns (xc, yc, r) or None if the points are collinear
    bx -= ax; by -= ay; cx -= ax; cy -= ay
    d = 2.0 * (bx * cy - by * cx)
    if abs(d) < 1e-12:
        return None
    b2 = bx * bx + by * by
    c2 = cx * cx + cy * cy
    ux = (cy * b2 - by * c2) / d
    uy = (bx * c2 - cx * b2) / d
    return (ax + ux, ay + uy, math.hypot(ux, uy))


def reach(fits, i, jmin, n):
    # largest j >= jmin with fits(i, j), None if not even jmin fits
    # galloping search followed by bisection, assumes that a fitting range stays fitting when shortened
    if jmin > n - 1 or not fits(i, jmin):
        return None
    lo = jmin
    step = 1
    hi = None
    while lo < n - 1:
        j = min(lo + step, n - 1)
        if fits(i, j):
            lo = j
            step *= 2
        else:
            hi = j
            break
    if hi is None:
        return lo
    while hi - lo > 1:
        mid = (lo + hi) // 2
        if fits(i, mid):
            lo = mid
        else:
            hi = mid
    return lo


def mergeneighbours(elems, make):
    # neighbouring elements of the same type that fit as one (e.g. an arc split by rounding noise)
    changed = True
    while changed:
        changed = False
        for k in range(len(elems) - 1):
            e1 = elems[k]; e2 = elems[k + 1]
            if e1["t"] != e2["t"]:
                continue
            merged = make(e1["t"], e1["i"], e2["j"])
            if merged is not None:
                elems[k:k + 2] = [merged]
                changed = True
                break


def refineboundaries(elems, make, sse, window):
    # the greedy growth lets an element swallow the first vertices of the next one as long as they are within tolerance
    # slide every shared vertex a few positions and keep the split with the smallest squared residuals of both elements
    for passno in range(3):
        changed = False
        for k in range(len(elems) - 1):
            e1 = elems[k]; e2 = elems[k + 1]
            b = e1["j"]
            best = (sse(e1) + sse(e2), b, e1, e2)
            for nb in range(max(e1["i"] + 1, b - window), min(e2["j"] - 1, b + window) + 1):
                if nb == b:
                    continue
                n1 = make(e1["t"], e1["i"], nb)
                if n1 is None:
                    continue
                n2 = make(e2["t"], nb, e2["j"])
                if n2 is None:
                    continue
                v = sse(n1) + sse(n2)
                if v < best[0] - 1e-15:
                    best = (v, nb, n1, n2)
            if best[1] != b:
                elems[k] = best[2]; elems[k + 1] = best[3]
                changed = True
        if not changed:
            break


def hlinefits(x, y, i, j, tol):
    dx = x[j] - x[i]; dy = y[j] - y[i]
    length = math.hypot(dx, dy)
    if length < 1e-9:
        return False
    ux = dx / length; uy = dy / length
    tprev = 0.0
    for k in range(i + 1, j):
        px = x[k] - x[i]; py = y[k] - y[i]
        if abs(px * uy - py * ux) > tol:
            return False
        t = px * ux + py * uy
        # the vertices must progress along the line, no zig-zag
        if t < tprev - tol or t > length + tol:
            return False
        tprev = t
    return True


def arcsweep(x, y, i, j, xc, yc, side):
    a0 = math.atan2(y[i] - yc, x[i] - xc)
    a1 = math.atan2(y[j] - yc, x[j] - xc)
    return ((a1 - a0) * side) % TWOPI


def harcfits(x, y, i, j, tol, rmax):
    # circle through first and last vertex (they are shared with the neighbours), the centre slides along the
    # chord bisector and is fitted by least squares to all vertices in between
    # all vertices must be within tol and progress in one direction, returns the arc dict or None
    if j - i < 2:
        return None
    m = (i + j) // 2
    c = circlefrom3(x[i], y[i], x[m], y[m], x[j], y[j])
    if c is None or c[2] > rmax:
        return None
    side = 1 if (x[m] - x[i]) * (y[j] - y[m]) - (y[m] - y[i]) * (x[j] - x[m]) > 0 else -1   # 1 = left/ccw

    if j - i > 2:
        dx = x[j] - x[i]; dy = y[j] - y[i]
        chord = math.hypot(dx, dy)
        half = chord / 2.0
        mx = x[i] + dx / 2.0; my = y[i] + dy / 2.0
        nx = -dy / chord; ny = dx / chord
        t = (c[0] - mx) * nx + (c[1] - my) * ny
        for iteration in range(4):
            xc = mx + t * nx; yc = my + t * ny
            r = math.hypot(half, t)
            sjr = 0.0; sjj = 0.0
            for k in range(i + 1, j):
                px = x[k] - xc; py = y[k] - yc
                d = math.hypot(px, py)
                res = d - r
                jac = -(px * nx + py * ny) / d - t / r
                sjr += jac * res; sjj += jac * jac
            if sjj < 1e-30:
                break
            step = sjr / sjj
            t -= step
            if abs(step) < 1e-9:
                break
        c = (mx + t * nx, my + t * ny, math.hypot(half, t))
        if c[2] > rmax:
            return None

    xc, yc, r = c
    sweep = 0.0
    ax = x[i] - xc; ay = y[i] - yc
    for k in range(i + 1, j + 1):
        bx = x[k] - xc; by = y[k] - yc
        if abs(math.hypot(bx, by) - r) > tol:
            return None
        a = math.atan2(ax * by - ay * bx, ax * bx + ay * by)
        if a * side <= 0:
            return None
        sweep += abs(a)
        ax = bx; ay = by
    if sweep >= TWOPI - 1e-9:
        return None
    return {"t": "A", "i": i, "j": j, "xc": xc, "yc": yc, "r": r, "side": side, "sweep": sweep}


def hdirection(x, y, e, atend):
    # unit direction of travel at the start or end of a horizontal element
    if e["t"] == "L":
        dx = x[e["j"]] - x[e["i"]]; dy = y[e["j"]] - y[e["i"]]
        d = math.hypot(dx, dy)
        return (dx / d, dy / d)
    k = e["j"] if atend else e["i"]
    rx = x[k] - e["xc"]; ry = y[k] - e["yc"]
    d = math.hypot(rx, ry)
    return (-e["side"] * ry / d, e["side"] * rx / d)


def hkink(x, y, e1, e2):
    # deflection angle [rad] between the end of e1 and the start of e2, 0 = tangent
    d1 = hdirection(x, y, e1, True)
    d2 = hdirection(x, y, e2, False)
    return abs(math.atan2(d1[0] * d2[1] - d1[1] * d2[0], d1[0] * d2[0] + d1[1] * d2[1]))


def tangentjoint(x, y, e1, e2, arc, tol):
    # forcing tangency moves the arc by roughly kink * arc length / 4, accept it if that stays within tol
    return hkink(x, y, e1, e2) * arc["r"] * arc["sweep"] / 4.0 <= tol


def arctangency(x, y, elems, tol):
    # linestring element type per horizontal element
    # a run of arcs between two straights:  Begin, Begin, ..., TangentTangent
    #   (a single arc between straights is TangentTangent)
    # a run of arcs at the start or end of the line (no straight on one side) stays plain arcs with radius
    # forced[k] is True if a tangency is imposed although the kink in the data is bigger than the tolerance
    n = len(elems)
    types = ["Straight" if e["t"] == "L" else "Arc" for e in elems]
    forced = [False] * n
    k = 0
    while k < n:
        if elems[k]["t"] != "A":
            k += 1
            continue
        a = k
        while k < n and elems[k]["t"] == "A":
            k += 1
        b = k - 1
        if a > 0 and b + 1 < n:     # straights on both sides
            for m in range(a, b + 1):
                types[m] = "Begin"
                forced[m] = not tangentjoint(x, y, elems[m - 1], elems[m], elems[m], tol)
            types[b] = "TangentTangent"
            forced[b] = forced[b] or not tangentjoint(x, y, elems[b], elems[b + 1], elems[b], tol)
    return types, forced


def segmenthorizontal(x, y, tol, rmax, minchords):
    n = len(x)
    elems = []
    linefits = lambda a, b: hlinefits(x, y, a, b, tol)
    arcfits = lambda a, b: harcfits(x, y, a, b, tol, rmax) is not None
    i = 0
    while i < n - 1:
        js = reach(linefits, i, i + 1, n)
        if js is None:
            js = i + 1    # can only happen with duplicate vertices
        ja = reach(arcfits, i, i + 2, n)
        # the longer element wins, a straight wins a tie
        if ja is not None and ja > js:
            arc = harcfits(x, y, i, ja, tol, rmax)
            accept = (ja - i) >= minchords
            # short arcs (few chords) can't be told apart from angle points, unless they are tangent to the previous element
            if not accept and elems:
                accept = hkink(x, y, elems[-1], arc) * math.hypot(x[i + 1] - x[i], y[i + 1] - y[i]) <= tol
            if accept:
                elems.append(arc)
                i = ja
                continue
        elems.append({"t": "L", "i": i, "j": js})
        i = js

    def make(t, a, b):
        if t == "L":
            return {"t": "L", "i": a, "j": b} if hlinefits(x, y, a, b, tol) else None
        return harcfits(x, y, a, b, tol, rmax)

    mergeneighbours(elems, make)
    absorbsinglechords(elems, make)
    refineboundaries(elems, make, lambda e: hsse(x, y, e), 10)
    mergeneighbours(elems, make)
    return elems


def absorbsinglechords(elems, make):
    # a single chord left over next to an arc, e.g. the closing chord of a full circle (an arc can't do 360 deg)
    # if it lies on the arc, arc + chord are split into two arcs
    k = 0
    while k < len(elems):
        e = elems[k]
        if e["t"] == "L" and e["j"] - e["i"] == 1:
            for n in [k - 1, k + 1]:
                if 0 <= n < len(elems) and elems[n]["t"] == "A":
                    a = min(e["i"], elems[n]["i"]); b = max(e["j"], elems[n]["j"])
                    mid = (a + b) // 2
                    a1 = make("A", a, mid)
                    a2 = make("A", mid, b) if a1 is not None else None
                    if a2 is not None:
                        lo = min(k, n)
                        elems[lo:lo + 2] = [a1, a2]
                        break
        k += 1


def roundarc(x, y, e, step, tol):
    # snap the radius to a multiple of step, keeping BC and EC, if all vertices stay within tol
    if step <= 0:
        return False
    i = e["i"]; j = e["j"]
    rr = round(e["r"] / step) * step
    if rr <= 0 or abs(rr - e["r"]) < 1e-9:
        return False
    dx = x[j] - x[i]; dy = y[j] - y[i]
    chord = math.hypot(dx, dy)
    half = chord / 2.0
    if rr < half:
        return False
    mx = x[i] + dx / 2.0; my = y[i] + dy / 2.0
    nx = -dy / chord; ny = dx / chord
    h = math.sqrt(rr * rr - half * half)
    sgn = 1.0 if (e["xc"] - mx) * nx + (e["yc"] - my) * ny >= 0 else -1.0
    xc = mx + sgn * h * nx; yc = my + sgn * h * ny
    for k in range(i, j + 1):
        if abs(math.hypot(x[k] - xc, y[k] - yc) - rr) > tol:
            return False
    e["xc"] = xc; e["yc"] = yc; e["r"] = rr
    e["sweep"] = arcsweep(x, y, i, j, xc, yc, e["side"])
    return True


def hresiduals(x, y, e):
    i = e["i"]; j = e["j"]
    if e["t"] == "L":
        dx = x[j] - x[i]; dy = y[j] - y[i]
        d = math.hypot(dx, dy)
        return [((x[k] - x[i]) * dy - (y[k] - y[i]) * dx) / d for k in range(i, j + 1)]
    return [math.hypot(x[k] - e["xc"], y[k] - e["yc"]) - e["r"] for k in range(i, j + 1)]


def hdeviation(x, y, e):
    return max(abs(r) for r in hresiduals(x, y, e))


def hsse(x, y, e):
    return sum(r * r for r in hresiduals(x, y, e))


def stationing(x, y, elems):
    # arc length along the recovered horizontal for every vertex, also stores s0/length per element
    s = [0.0] * len(x)
    st = 0.0
    for e in elems:
        i = e["i"]; j = e["j"]
        if e["t"] == "L":
            dx = x[j] - x[i]; dy = y[j] - y[i]
            length = math.hypot(dx, dy)
            for k in range(i, j + 1):
                t = ((x[k] - x[i]) * dx + (y[k] - y[i]) * dy) / length
                s[k] = st + min(max(t, 0.0), length)
        else:
            e["sweep"] = arcsweep(x, y, i, j, e["xc"], e["yc"], e["side"])
            length = e["r"] * e["sweep"]
            s[i] = st
            for k in range(i + 1, j + 1):
                s[k] = st + min(e["r"] * arcsweep(x, y, i, k, e["xc"], e["yc"], e["side"]), length)
        e["s0"] = st
        e["length"] = length
        st += length
    return s


def vlinefits(s, z, i, j, tol):
    g = (z[j] - z[i]) / (s[j] - s[i])
    for k in range(i + 1, j):
        if abs(z[i] + g * (s[k] - s[i]) - z[k]) > tol:
            return False
    return True


def vparabolafits(s, z, i, j, tol):
    # parabola z = z_i + b*u + c*u^2 (u = s - s_i) through first and last vertex (shared with the neighbours)
    # the curvature c is fitted by least squares to all vertices in between
    if j - i < 2:
        return None
    uj = s[j] - s[i]
    g = (z[j] - z[i]) / uj
    swr = 0.0; sww = 0.0
    for k in range(i + 1, j):
        u = s[k] - s[i]
        w = u * (u - uj)
        swr += w * (z[k] - z[i] - g * u)
        sww += w * w
    c = swr / sww
    b = g - c * uj
    for k in range(i + 1, j):
        u = s[k] - s[i]
        if abs(z[i] + b * u + c * u * u - z[k]) > tol:
            return None
    return {"t": "P", "i": i, "j": j, "b": b, "c": c}


def vgrades(s, z, e):
    # grade at the start and at the end of a vertical element
    if e["t"] == "G":
        g = (z[e["j"]] - z[e["i"]]) / (s[e["j"]] - s[e["i"]])
        return (g, g)
    return (e["b"], e["b"] + 2.0 * e["c"] * (s[e["j"]] - s[e["i"]]))


def segmentvertical(s, z, tol, minchords):
    n = len(s)
    elems = []
    linefits = lambda a, b: vlinefits(s, z, a, b, tol)
    parfits = lambda a, b: vparabolafits(s, z, a, b, tol) is not None
    i = 0
    while i < n - 1:
        js = reach(linefits, i, i + 1, n)
        jp = reach(parfits, i, i + 2, n)
        if jp is not None and jp > js:
            par = vparabolafits(s, z, i, jp, tol)
            accept = (jp - i) >= minchords
            if not accept and elems:
                accept = abs(vgrades(s, z, elems[-1])[1] - par["b"]) * (s[i + 1] - s[i]) <= tol
            if accept:
                elems.append(par)
                i = jp
                continue
        elems.append({"t": "G", "i": i, "j": js})
        i = js

    def make(t, a, b):
        if t == "G":
            return {"t": "G", "i": a, "j": b} if vlinefits(s, z, a, b, tol) else None
        return vparabolafits(s, z, a, b, tol)

    mergeneighbours(elems, make)
    refineboundaries(elems, make, lambda e: vsse(s, z, e), 10)
    flattenparabolas(s, z, elems, tol)
    mergeneighbours(elems, make)
    return elems


def flattenparabolas(s, z, elems, tol):
    # a parabola that bows less than tol off its chord is a grade
    # (TBC derives the VPI from the in/out tangents, nearly parallel ones break the linestring vertical
    # and the profile ends at that curve, the rest of the line gets extrapolated)
    for k, e in enumerate(elems):
        if e["t"] == "P" and abs(e["c"]) * (s[e["j"]] - s[e["i"]]) ** 2 / 4.0 <= tol and vlinefits(s, z, e["i"], e["j"], tol):
            elems[k] = {"t": "G", "i": e["i"], "j": e["j"]}


def vresiduals(s, z, e):
    i = e["i"]; j = e["j"]
    if e["t"] == "G":
        g = (z[j] - z[i]) / (s[j] - s[i])
        return [z[i] + g * (s[k] - s[i]) - z[k] for k in range(i, j + 1)]
    return [z[i] + e["b"] * (s[k] - s[i]) + e["c"] * (s[k] - s[i]) ** 2 - z[k] for k in range(i, j + 1)]


def vdeviation(s, z, e):
    return max(abs(r) for r in vresiduals(s, z, e))


def vsse(s, z, e):
    return sum(r * r for r in vresiduals(s, z, e))

# ---- end of geometry core


class SCR_RecurvePolyline(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_RecurvePolyline.xaml") as s:
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

        self.lType = clr.GetClrType(IPolyseg)
        self.objs.IsEntityValidCallback = self.IsValid

        self.htol.DistanceMin = 0.00001
        self.vtol.DistanceMin = 0.00001
        self.maxradius.DistanceMin = 1.0
        self.roundradius.DistanceMin = 0.0
        for c in [self.minarcchords, self.minvcchords]:
            c.NumberOfDecimals = 0
            c.MinValue = 2
            c.MaxValue = 1000

        try:
            self.SetDefaultOptions()
        except:
            pass

    def SetDefaultOptions(self):
        SCROptions.LoadMacroOptions(self, "SCR_RecurvePolyline", _OPTIONS, self.currentProject)

    def SaveOptions(self):
        SCROptions.SaveMacroOptions(self, "SCR_RecurvePolyline", _OPTIONS)

    def IsValid(self, serial):
        o = self.currentProject.Concordance[serial]
        return isinstance(o, self.lType)

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content = ''
        self.success.Content = ''

        lines = [o for o in self.objs if isinstance(o, self.lType)]
        if not lines:
            self.error.Content = 'no Lines selected'
            return

        wv = self.currentProject [Project.FixedSerial.WorldView]

        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        wv.PauseGraphicsCache(True)

        report = []
        totals = [0, 0, 0, 0, 0.0, 0.0]   # lines, straights, arcs, vertical curves, max dev H, max dev V
        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                lasttime = 0
                for count, l in enumerate(lines):
                    if time.time() - lasttime > 1:
                        lasttime = time.time()
                        ProgressBar.TBC_ProgressBar.Title = 'recurving Line {} of {}'.format(count + 1, len(lines))
                        if ProgressBar.TBC_ProgressBar.SetProgress(int(100.0 * count / len(lines))):
                            break
                    self.recurve(wv, l, report, totals)

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
        wv.PauseGraphicsCache(False)

        self.success.Content += '{} Lines recurved\n{} Straights, {} Arcs, {} Vertical Curves\nmax Deviation H {:.4f} / V {:.4f}'.format(
            totals[0], totals[1], totals[2], totals[3], totals[4], totals[5])

        if self.writecsv.IsChecked and report:
            csvname = os.path.expanduser('~/Downloads/SCR_RecurvePolyline.csv')
            try:
                with open(csvname, 'w') as f:
                    f.write("Line,Element,Type,Linestring Element,Start Chainage,End Chainage,Length,Radius,Direction,Grade In %,Grade Out %,K-Value,"
                            "Start X,Start Y,End X,End Y,Kink to previous [deg],Vertices,max Deviation\n")
                    for row in report:
                        f.write(row + "\n")
                self.success.Content += '\nCSV written to\n' + csvname
            except Exception as e:
                self.error.Content += '\ncould not write the CSV - is it open in another program?'

        self.SaveOptions()

    def recurve(self, wv, l, report, totals):

        name = IName.Name.__get__(l)
        polyseg = l.ComputePolySeg()
        is3d = polyseg.AllPointsAre3D or (l.ComputeVerticalPolySeg() is not None)

        # vertices in world coordinates, any existing curves are chorded tightly
        pts = SCROverlayBag.getpolypoints(l)
        if pts.Length < 2:
            self.error.Content += '\nskipped "{}" - not enough Vertices'.format(name)
            return

        # local coordinates, consecutive duplicates (same XY) are dropped
        ox = pts[0].X; oy = pts[0].Y
        x = []; y = []; z = []
        for p in pts:
            px = p.X - ox; py = p.Y - oy
            if x and math.hypot(px - x[-1], py - y[-1]) < 1e-6:
                continue
            x.append(px); y.append(py); z.append(p.Z if is3d else 0.0)
        if len(x) < 2:
            self.error.Content += '\nskipped "{}" - not enough Vertices'.format(name)
            return

        htol = self.htol.Distance
        vtol = self.vtol.Distance

        # ---- horizontal
        helems = segmenthorizontal(x, y, htol, self.maxradius.Distance, int(self.minarcchords.Value))
        if self.roundradius.Distance > 0:
            for e in helems:
                if e["t"] == "A":
                    roundarc(x, y, e, self.roundradius.Distance, htol)
        s = stationing(x, y, helems)
        hlength = s[-1]

        # ---- horizontal as linestring elements, 2D - the elevations go into the vertical tab
        etypes, forced = arctangency(x, y, helems, htol)
        if any(forced):
            self.error.Content += '\n"{}" - {} Arc(s) made tangent although the kink exceeds the Tolerance - see CSV'.format(name, sum(forced))
        newline = wv.Add(clr.GetClrType(Linestring))
        newline.Name = name
        newline.Layer = self.layerpicker.SelectedSerialNumber

        el = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
        el.Position = Point3D(x[0] + ox, y[0] + oy)
        newline.AppendElement(el)
        for e, etype in zip(helems, etypes):
            if etype == "Straight":
                el = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
            elif etype == "Arc":
                el = ElementFactory.Create(clr.GetClrType(IArcSegment), clr.GetClrType(IXYZLocation))
                el.Radius = e["r"]
                # LeftRight is a Direction enum, take the type straight from the property instead of importing it
                directiontype = clr.GetClrType(IArcSegment).GetProperty("LeftRight").PropertyType
                el.LeftRight = System.Enum.Parse(directiontype, "Left" if e["side"] > 0 else "Right")
                el.LargeSmall = ArcLargeSmall.Large if e["sweep"] > math.pi else ArcLargeSmall.Small
            else:
                el = ElementFactory.Create(clr.GetClrType(ITangentArcSegment), clr.GetClrType(IXYZLocation))
                el.TangentType = ArcTangentType.Begin if etype == "Begin" else ArcTangentType.TangentTangent
            el.Position = Point3D(x[e["j"]] + ox, y[e["j"]] + oy)
            newline.AppendElement(el)

        # TBC computes the tangent arcs from the tangency, the result should only differ marginally from the fit
        polyseg = newline.ComputePolySeg()
        lengthdiff = polyseg.ComputeStationing() - polyseg.BeginStation - hlength
        if abs(lengthdiff) > 0.005:
            self.error.Content += '\n"{}" - Linestring length differs by {:.3f} from the fit'.format(name, lengthdiff)

        # ---- vertical, on the new horizontal
        # every original 3D vertex is projected onto the linestring TBC computed, giving its chainage
        # the vertical is then fitted to those chainages/elevations, so it covers the real line from start to end
        # (TBC computes the tangent arcs from the tangency, the length can differ from the horizontal fit)
        vs = []; vz = []; velems = []
        if is3d:
            vs, vz, dropped, maxoffset = self.projectvertices(polyseg, x, y, z, ox, oy)
            if dropped:
                self.error.Content += '\n"{}" - {} Vertices not used for the Vertical, Chainage not increasing'.format(name, dropped)
            if maxoffset > htol * 2:
                self.error.Content += '\n"{}" - Vertices up to {:.3f} off the new Linestring'.format(name, maxoffset)
            if len(vs) >= 2:
                if self.fitvertical.IsChecked:
                    velems = segmentvertical(vs, vz, vtol, int(self.minvcchords.Value))
                else:
                    # keep the original elevations - every vertex becomes a VPI
                    velems = [{"t": "G", "i": k, "j": k + 1} for k in range(len(vs) - 1)]

        if velems:
            vpoly = PolySeg.PolySeg()
            for e in velems:
                i = e["i"]; j = e["j"]
                if e["t"] == "G":
                    vpoly.Add(SegmentLine(Point3D(vs[i], vz[i], 0), Point3D(vs[j], vz[j], 0)))
                else:
                    half = (vs[j] - vs[i]) / 2.0
                    vpi = Point3D(vs[i] + half, vz[i] + e["b"] * half, 0)
                    vpoly.Add(ParabolaSegment(Point3D(vs[i], vz[i], 0), vpi, Point3D(vs[j], vz[j], 0)))
            Linestring.ConvertPolySegToLinestringVertical(newline, vpoly, 0)

        # ---- report
        totals[0] += 1
        for k, e in enumerate(helems):
            dev = hdeviation(x, y, e)
            totals[4] = max(totals[4], dev)
            kink = math.degrees(hkink(x, y, helems[k - 1], e)) if k > 0 else 0.0
            i = e["i"]; j = e["j"]
            if e["t"] == "L":
                totals[1] += 1
                typ = "Straight"; radius = ""; direction = ""
            else:
                totals[2] += 1
                typ = "Arc"; radius = "{:.4f}".format(e["r"]); direction = "Left" if e["side"] > 0 else "Right"
            report.append('"{}",H{},{},{},{:.4f},{:.4f},{:.4f},{},{},,,,{:.4f},{:.4f},{:.4f},{:.4f},{:.6f},{},{:.5f}'.format(
                name, k + 1, typ, etypes[k] + (" (forced)" if forced[k] else ""), e["s0"], e["s0"] + e["length"], e["length"], radius, direction,
                x[i] + ox, y[i] + oy, x[j] + ox, y[j] + oy, kink, j - i + 1, dev))

        for k, e in enumerate(velems if self.fitvertical.IsChecked else []):
            dev = vdeviation(vs, vz, e)
            totals[5] = max(totals[5], dev)
            i = e["i"]; j = e["j"]
            gin, gout = vgrades(vs, vz, e)
            if e["t"] == "G":
                typ = "Grade"; kvalue = ""
            else:
                totals[3] += 1
                typ = "Vertical Curve"
                kvalue = "{:.3f}".format((vs[j] - vs[i]) / abs((gout - gin) * 100.0)) if abs(gout - gin) > 1e-12 else ""
            report.append('"{}",V{},{},,{:.4f},{:.4f},{:.4f},,,{:.4f},{:.4f},{},{:.4f},{:.4f},{:.4f},{:.4f},,{},{:.5f}'.format(
                name, k + 1, typ, vs[i], vs[j], vs[j] - vs[i], gin * 100.0, gout * 100.0, kvalue,
                vs[i], vz[i], vs[j], vz[j], j - i + 1, dev))

    def projectvertices(self, polyseg, x, y, z, ox, oy):
        # chainage of every original vertex on the new horizontal, relative to its begin station
        # first/last vertex are the line ends, vertices whose chainage doesn't increase are dropped
        begin = polyseg.BeginStation
        total = polyseg.ComputeStationing() - begin
        onl = clr.StrongBox[Point3D]()
        sta = clr.StrongBox[float]()
        n = len(x)
        vs = [0.0]; vz = [z[0]]
        dropped = 0; maxoffset = 0.0
        for k in range(1, n - 1):
            px = x[k] + ox; py = y[k] + oy
            if not polyseg.FindPointFromPoint(Point3D(px, py, 0), onl, sta):
                dropped += 1
                continue
            c = sta.Value - begin
            if c <= vs[-1] + 1e-4 or c >= total - 1e-4:
                dropped += 1
                continue
            maxoffset = max(maxoffset, math.hypot(onl.Value.X - px, onl.Value.Y - py))
            vs.append(c); vz.append(z[k])
        vs.append(total); vz.append(z[-1])
        return vs, vz, dropped, maxoffset

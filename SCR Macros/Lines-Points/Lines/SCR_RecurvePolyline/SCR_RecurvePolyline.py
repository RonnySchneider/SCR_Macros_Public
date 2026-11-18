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
    "modedesign":    True,
    "modesurvey":    False,
    "htol":          0.002,
    "minarcchords":  3.0,
    "hscatter":      0.05,
    "hminlength":    20.0,
    "maxradius":     10000.0,
    "roundradius":   0.0,
    "fitvertical":   True,
    "vtol":          0.002,
    "minvcchords":   3.0,
    "vscatter":      0.02,
    "vminlength":    20.0,
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

        cmdData.Version = 1.011
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""

        cmdData.ToolTipTitle = "Recurve Polyline"
        cmdData.ToolTipTextFormatted = "re-engineer chorded or surveyed polylines into Straights/Arcs and Grades/Vertical Curves"

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


def tangenttypes(kinds):
    # linestring element type per element for a chain that is tangent everywhere ('L' straight, 'A' arc)
    # a run of arcs between two straights: Begin, ..., TangentTangent, arcs at the line ends stay plain arcs
    n = len(kinds)
    types = ["Straight" if t == 'L' else "Arc" for t in kinds]
    k = 0
    while k < n:
        if kinds[k] != 'A':
            k += 1
            continue
        a = k
        while k < n and kinds[k] == 'A':
            k += 1
        if a > 0 and k < n:
            for q in range(a, k):
                types[q] = "Begin"
            types[k - 1] = "TangentTangent"
    return types


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

# ---- survey mode
# the elements are fitted by least squares, they don't have to pass through any vertex
# 1. segmentation of the heading (grade) diagram by dynamic programming, constant pieces = straights (grades),
#    linear pieces = arcs (vertical curves)
# 2. adjustment of the whole chain by a robust Levenberg-Marquardt: start pose + length/curvature per element,
#    the chain is tangent continuous by construction
# 3. elements that shrink away are removed, the chain is refitted


def chorddiagram(s, v):
    # value per chord (heading or grade) at the chord middle, with weight
    # the noise of a chord value is ~ sqrt(2) * scatter / chord length, the weight makes the residuals comparable to the scatter
    ms = []; mv = []; mw = []
    for k in range(len(s) - 1):
        ds = s[k + 1] - s[k]
        if ds < 1e-6:
            continue
        ms.append((s[k] + s[k + 1]) / 2.0); mv.append(v[k]); mw.append(ds * ds / 2.0)
    return ms, mv, mw


def segmentdiagram(ms, mv, mw, sigma, minlen, kmin, maxchords=1500):
    # optimal partition of the diagram into constant and linear pieces
    # cost = weighted squared residuals / sigma^2 + penalty per parameter
    # two constant pieces in a row aren't allowed (that would be an angle point), linear pieces flatter than kmin aren't allowed
    # returns a list of (type 0 const / 1 linear, start, end, value at start, slope)
    n = len(ms)
    penalty = 2.0 * math.log(max(n, 2))
    cw = [0.0]; cs = [0.0]; css = [0.0]; cv = [0.0]; cvv = [0.0]; csv = [0.0]
    s0 = ms[0]
    for k in range(n):
        w = mw[k]; t = ms[k] - s0; v = mv[k]
        cw.append(cw[-1] + w); cs.append(cs[-1] + w * t); css.append(css[-1] + w * t * t)
        cv.append(cv[-1] + w * v); cvv.append(cvv[-1] + w * v * v); csv.append(csv[-1] + w * t * v)
    s2 = sigma * sigma

    def edge(k):
        if n == 1:
            return ms[0] if k == 0 else ms[0] + 1e-3
        if k == 0:
            return ms[0] - (ms[1] - ms[0]) / 2.0
        if k == n:
            return ms[-1] + (ms[-1] - ms[-2]) / 2.0
        return (ms[k - 1] + ms[k]) / 2.0

    edges = [edge(k) for k in range(n + 1)]
    INF = float("inf")
    best0 = [INF] * (n + 1); best1 = [INF] * (n + 1)     # cost of chords 0..b-1 ending with a constant/linear piece
    back0 = [None] * (n + 1); back1 = [None] * (n + 1)
    best0[0] = 0.0; best1[0] = 0.0
    for b in range(1, n + 1):
        sb = edges[b]
        for a in range(b - 1, max(-1, b - 1 - maxchords), -1):
            if sb - edges[a] < minlen and not (a == 0 and b == n):
                continue
            W = cw[b] - cw[a]; S = cs[b] - cs[a]; SS = css[b] - css[a]
            V = cv[b] - cv[a]; VV = cvv[b] - cvv[a]; SV = csv[b] - csv[a]
            mean = V / W
            ssec = max(VV - V * mean, 0.0)
            prevlin = best1[a] if a > 0 else 0.0
            c = prevlin + ssec / s2 + penalty
            if c < best0[b]:
                best0[b] = c; back0[b] = (a, 1, mean, 0.0, S / W)
            den = SS - S * S / W
            if b - a >= 3 and den > 1e-12:
                slope = (SV - S * V / W) / den
                if abs(slope) >= kmin:
                    ssel = max(ssec - slope * slope * den, 0.0)
                    prevany = min(best0[a], best1[a]) if a > 0 else 0.0
                    c = prevany + ssel / s2 + 2.0 * penalty
                    if c < best1[b]:
                        best1[b] = c
                        back1[b] = (a, (0 if best0[a] <= best1[a] else 1) if a > 0 else 0, mean, slope, S / W)
    out = []
    b = n; t = 0 if best0[n] <= best1[n] else 1
    while b > 0:
        a, pt, mean, slope, sm = back0[b] if t == 0 else back1[b]
        sa = edges[a]; sb = edges[b]
        out.append((t, sa, sb, mean + slope * (sa - s0 - sm), slope))
        b = a; t = pt
    out.reverse()
    return out


def jointvalues(pieces):
    # value (heading/grade) at every joint, a constant piece dictates it, between linear pieces the mean is used
    n = len(pieces)
    jv = []
    for k in range(n + 1):
        if k > 0 and pieces[k - 1][0] == 0:
            jv.append(pieces[k - 1][3])
        elif k < n and pieces[k][0] == 0:
            jv.append(pieces[k][3])
        else:
            vals = []
            if k > 0:
                t, sa, sb, v0, sl = pieces[k - 1]; vals.append(v0 + sl * (sb - sa))
            if k < n:
                vals.append(pieces[k][3])
            jv.append(sum(vals) / len(vals))
    return jv


def solvelinear(A, b):
    # Gaussian elimination with partial pivoting, A and b are modified
    n = len(b)
    for c in range(n):
        p = c
        for r in range(c + 1, n):
            if abs(A[r][c]) > abs(A[p][c]):
                p = r
        if abs(A[p][c]) < 1e-300:
            return None
        A[c], A[p] = A[p], A[c]; b[c], b[p] = b[p], b[c]
        Ac = A[c]
        for r in range(c + 1, n):
            f = A[r][c] / Ac[c]
            if f != 0.0:
                Ar = A[r]
                for q in range(c, n):
                    Ar[q] -= f * Ac[q]
                b[r] -= f * b[c]
    xs = [0.0] * n
    for c in range(n - 1, -1, -1):
        s = b[c]
        Ac = A[c]
        for q in range(c + 1, n):
            s -= Ac[q] * xs[q]
        xs[c] = s / Ac[c]
    return xs


def lmrobust(p, model, sigma, iterations):
    # robust (Huber) Levenberg-Marquardt
    # model.residuals(p, store) -> residual list or None, model.jacobian(p) -> list of (lo, column from lo) per parameter
    # (store=True keeps the point/element assignment and whatever the Jacobian needs)
    r = model.residuals(p, True)
    lam = 1e-3
    npar = len(p)
    n = len(r)
    for it in range(iterations):
        med = sorted(abs(v) for v in r)[n // 2]
        cth = 2.0 * max(sigma, 1.4826 * med)       # tightens towards 2 sigma as the fit converges
        w = [1.0 if abs(v) <= cth else cth / abs(v) for v in r]
        c0 = 0.0
        for wi, v in zip(w, r):
            c0 += wi * v * v
        cols = model.jacobian(p)
        JtJ = [[0.0] * npar for q in range(npar)]
        Jtr = [0.0] * npar
        wcols = []
        for q in range(npar):
            lo, col = cols[q]
            wc = [w[lo + i] * v for i, v in enumerate(col)]
            wcols.append(wc)
            s = 0.0
            for i, v in enumerate(wc):
                s += v * r[lo + i]
            Jtr[q] = s
        for q1 in range(npar):
            lo1, c1 = cols[q1]; wc1 = wcols[q1]
            for q2 in range(q1, npar):
                lo2, c2 = cols[q2]
                lo = max(lo1, lo2)
                s = 0.0
                o1 = lo - lo1; o2 = lo - lo2
                for i in range(min(len(c1) - o1, len(c2) - o2)):
                    s += wc1[o1 + i] * c2[o2 + i]
                JtJ[q1][q2] = s; JtJ[q2][q1] = s
        improved = False
        for attempt in range(10):
            A = [row[:] for row in JtJ]
            for q in range(npar):
                A[q][q] += lam * max(JtJ[q][q], 1e-12)
            d = solvelinear(A, [-v for v in Jtr])
            if d is not None:
                pn = [a + b for a, b in zip(p, d)]
                rn = model.residuals(pn, False)
                if rn is not None:
                    c1 = 0.0
                    for wi, v in zip(w, rn):
                        c1 += wi * v * v
                    if c1 < c0:
                        p = pn; lam = max(lam / 3.0, 1e-9); improved = True
                        break
            lam *= 4.0
        if not improved:
            break
        r = model.residuals(p, True)
        c1 = 0.0
        for wi, v in zip(w, r):
            c1 += wi * v * v
        if c0 - c1 < 1e-5 * c0:
            break
    return p, r


# -- horizontal chain: start pose (x, y, heading), elements with length and curvature (0 = straight)
# the first element extends backwards, the last one is open ended

def elementend(pose, L, k):
    # chord length 2 sin(kL/2) / k in the direction h + kL/2, stable for tiny curvatures
    x, y, h = pose
    a = k * L / 2.0
    chord = L * (math.sin(a) / a if abs(a) > 1e-6 else 1.0 - a * a / 6.0)
    return x + chord * math.cos(h + a), y + chord * math.sin(h + a), h + 2.0 * a


def chainposes(pose, lengths, curvs):
    out = [pose]
    for L, k in zip(lengths, curvs):
        out.append(elementend(out[-1], L, k))
    return out


def elementproject(px, py, pose, L, k, first, last):
    # station u along the element and signed offset (left positive), None if the foot is outside the element
    x, y, h = pose
    if abs(k) < 1e-9:
        dx = px - x; dy = py - y
        c = math.cos(h); sn = math.sin(h)
        u = dx * c + dy * sn
        off = dy * c - dx * sn
    else:
        r = 1.0 / k
        xc = x - math.sin(h) * r; yc = y + math.cos(h) * r
        sg = 1.0 if k > 0 else -1.0
        da = (math.atan2(py - yc, px - xc) - math.atan2(y - yc, x - xc)) * sg
        mid = abs(k) * L / 2.0
        da = (da - mid + math.pi) % TWOPI - math.pi + mid     # angle along the direction of travel, centred on the element middle
        u = da / abs(k)
        off = r - sg * math.hypot(px - xc, py - yc)
    if (u < 0 and not first) or (u > L and not last):
        return None
    return u, off


def chainfoot(px, py, poses, lengths, curvs, guess):
    # foot of a point on the chain, searching the elements around guess
    # returns (offset, element, station on the element, foot x, foot y, normal x, normal y)
    m = len(lengths)
    best = None
    for e in range(max(0, guess - 1), min(m, guess + 2)):
        r = elementproject(px, py, poses[e], lengths[e], curvs[e], e == 0, e == m - 1)
        if r is not None and (best is None or abs(r[1]) < abs(best[0])):
            best = (r[1], e, r[0])
    if best is None:
        for e in (guess - 2, guess + 2):
            if 0 <= e < m:
                r = elementproject(px, py, poses[e], lengths[e], curvs[e], e == 0, e == m - 1)
                if r is not None and (best is None or abs(r[1]) < abs(best[0])):
                    best = (r[1], e, r[0])
    if best is not None:
        off, e, u = best
        fx, fy, h = elementend(poses[e], u, curvs[e])
        return off, e, u, fx, fy, -math.sin(h), math.cos(h)
    # outside the corner between two elements, the nearest joint is the foot
    for e in range(max(1, guess - 2), min(m, guess + 3)):
        x, y, h = poses[e]
        d = math.hypot(px - x, py - y)
        if best is None or d < best[0]:
            best = (d, e, x, y, h)
    d, e, x, y, h = best
    if d < 1e-12:
        return 0.0, e, 0.0, x, y, -math.sin(h), math.cos(h)
    side = 1.0 if (px - x) * -math.sin(h) + (py - y) * math.cos(h) >= 0 else -1.0
    return side * d, e, 0.0, x, y, side * (px - x) / d, side * (py - y) / d


class HChain(object):
    # parameters: x0, y0, h0, log of every length but the last, curvature * 1000 of the free arcs
    def __init__(self, x, y, types, fixed):
        self.x = x; self.y = y
        self.types = types
        self.fixed = fixed            # {element: curvature} for rounded radii
        self.guess = [0] * len(x)
        self.feet = None

    def unpack(self, p):
        m = len(self.types)
        lengths = [math.exp(min(v, 12.0)) for v in p[3:3 + m - 1]] + [0.0]
        curvs = []
        q = 3 + m - 1
        for k, t in enumerate(self.types):
            if t == 'L':
                curvs.append(0.0)
            elif k in self.fixed:
                curvs.append(self.fixed[k])
            else:
                curvs.append(p[q] / 1000.0); q += 1
        lengths[-1] = 1e9 if curvs[-1] == 0.0 else (TWOPI - 0.1) / abs(curvs[-1])
        return (p[0], p[1], p[2]), lengths, curvs

    def pack(self, pose, lengths, curvs):
        p = list(pose) + [math.log(max(L, 1e-3)) for L in lengths[:-1]]
        for k, t in enumerate(self.types):
            if t == 'A' and k not in self.fixed:
                p.append(curvs[k] * 1000.0)
        return p

    def residuals(self, p, store):
        pose, lengths, curvs = self.unpack(p)
        poses = chainposes(pose, lengths, curvs)
        out = []
        feet = []
        for i in range(len(self.x)):
            f = chainfoot(self.x[i], self.y[i], poses, lengths, curvs, self.guess[i])
            out.append(f[0])
            feet.append(f)
        if store:
            self.feet = feet
            self.poses = poses
            for i in range(len(self.x)):
                self.guess[i] = feet[i][1]
        return out

    def jacobian(self, p):
        # a change of an element moves everything after it rigidly: translation d and rotation w about the element end E
        # d offset = -(d + w * perp(F - E)) . n   (F foot, n left normal at the foot)
        pose, lengths, curvs = self.unpack(p)
        poses = self.poses
        feet = self.feet
        npt = len(feet)
        m = len(self.types)

        def rigid(lo, dx, dy, w, ex, ey):
            col = []
            for i in range(lo, npt):
                f = feet[i]
                if f[1] < lo_el[0]:
                    col.append(0.0)
                    continue
                fx = f[3] - ex; fy = f[4] - ey
                col.append(-((dx - w * fy) * f[5] + (dy + w * fx) * f[6]))
            return col

        # first point index per element (points are roughly in order)
        firstpt = [npt] * (m + 1)
        for i in range(npt - 1, -1, -1):
            e = feet[i][1]
            for k in range(e, -1, -1):
                if firstpt[k] > i:
                    firstpt[k] = i
                else:
                    break
        cols = []
        lo_el = [0]
        cols.append((0, [-f[5] for f in feet]))
        cols.append((0, [-f[6] for f in feet]))
        cols.append((0, rigid(0, 0.0, 0.0, 1.0, pose[0], pose[1])))
        for k in range(m - 1):
            # log length: the end moves along the end tangent, an arc also turns the rest
            ex, ey, eh = poses[k + 1]
            L = lengths[k]
            lo_el[0] = k + 1
            lo = firstpt[k + 1] if k + 1 < m else npt
            lo = min(lo, firstpt[k])
            cols.append((lo, rigid(lo, L * math.cos(eh), L * math.sin(eh), curvs[k] * L, ex, ey)))
        for k, t in enumerate(self.types):
            if t != 'A' or k in self.fixed:
                continue
            # curvature: the points on the arc numerically, the rest rigidly
            eps = 1e-9
            L = lengths[k]
            ex, ey, eh = poses[k + 1] if k + 1 < len(poses) else elementend(poses[k], L, curvs[k])
            ex2, ey2, eh2 = elementend(poses[k], L, curvs[k] + eps)
            scale = 1e-3 / eps
            lo = firstpt[k]
            lo_el[0] = k + 1
            col = rigid(lo, (ex2 - ex) * 1e-3 / eps, (ey2 - ey) * 1e-3 / eps, (eh2 - eh) * 1e-3 / eps, ex, ey) if k < m - 1 else [0.0] * (npt - lo)
            for i in range(lo, npt):
                f = feet[i]
                if f[1] == k:
                    r = elementproject(self.x[i], self.y[i], poses[k], L, curvs[k] + eps, k == 0, k == m - 1)
                    col[i - lo] = (r[1] - f[0]) * scale if r is not None else 0.0
            cols.append((lo, col))
        return cols


def hfit(x, y, types, pose, lengths, curvs, sigma, fixed, iterations):
    ch = HChain(x, y, types, fixed)
    # initial assignment: walk forward, the points are roughly in order
    poses = chainposes(pose, lengths[:-1] + [1e9], curvs)
    e = 0
    for i in range(len(x)):
        while e + 1 < len(types):
            f1 = chainfoot(x[i], y[i], poses, lengths, curvs, e)
            if f1[1] > e:
                e = f1[1]
            else:
                break
        ch.guess[i] = e
    p = ch.pack(pose, lengths, curvs)
    p, r = lmrobust(p, ch, sigma, iterations)
    pose, lengths, curvs = ch.unpack(p)
    return pose, lengths, curvs, r, [f[1] for f in ch.feet], ch.feet


def prunechain(types, lengths, curvs, minlen, kmin, kinks=True):
    # removes elements shorter than minlen without disturbing the chain: the length goes to the neighbours,
    # the turn of a removed arc to the neighbouring arcs (a short arc between two straights is an angle point,
    # it's removed too if kinks, the refit has to absorb its turn)
    # arcs flatter than kmin become straights, neighbouring straights are merged
    # works for horizontal (curvature) and vertical (rate of grade change) alike, returns (types, lengths, curvs) or None
    els = [[t, L, c * L] for t, L, c in zip(types, lengths, curvs)]
    lastc = curvs[-1]
    els[-1][1] = 0.0; els[-1][2] = 0.0
    changed = False
    k = 0
    while k < len(els) - 1 and len(els) > 1:
        e = els[k]
        nb = [j for j in (k - 1, k + 1) if 0 <= j < len(els)]
        arcs = [j for j in nb if els[j][0] == 'A' and j < len(els) - 1]
        if e[1] < minlen and (e[0] == 'L' or arcs or kinks):
            for j in nb:
                els[j][1] += e[1] / len(nb)
            if e[0] == 'A':
                for j in arcs:
                    els[j][2] += e[2] / len(arcs)
            del els[k]
            changed = True
            k = max(k - 1, 0)
            continue
        k += 1
    for e in els[:-1]:
        if e[0] == 'A' and abs(e[2]) < kmin * e[1]:
            e[0] = 'L'; e[2] = 0.0; changed = True
    out = []
    for e in els:
        if out and out[-1][0] == 'L' and e[0] == 'L':
            out[-1][1] += e[1]
            changed = True
        else:
            out.append(e)
    if not changed:
        return None
    return ([e[0] for e in out], [e[1] for e in out],
            [(e[2] / e[1] if e[1] > 0 else lastc) if e[0] == 'A' else 0.0 for e in out])


def trimchain(pose, lengths, curvs, types, first, last):
    # cut the chain at the feet of the first and the last point, first/last = (element, station on it)
    e0, u0 = first; e1, u1 = last
    poses = chainposes(pose, lengths, curvs)
    newpose = elementend(poses[e0], u0, curvs[e0])
    lengths = list(lengths); curvs = list(curvs); types = list(types)
    if e1 == e0:
        lengths[e0] = u1 - u0
    else:
        lengths[e0] -= u0
        lengths[e1] = u1
    return newpose, types[e0:e1 + 1], lengths[e0:e1 + 1], curvs[e0:e1 + 1]


def hfitprune(x, y, types, pose, lengths, curvs, sigma, minlen, rmax, iterations):
    # fit, drop elements that shrink away, refit until the structure is stable
    for rnd in range(8):
        pose, lengths, curvs, r, el, feet = hfit(x, y, types, pose, lengths, curvs, sigma, {}, 30)
        new = prunechain(types, lengths, curvs, minlen / 2.0, 1.0 / rmax)
        if new is None:
            break
        types, lengths, curvs = new
    pose, lengths, curvs, r, el, feet = hfit(x, y, types, pose, lengths, curvs, sigma, {}, iterations)
    new = prunechain(types, lengths, curvs, minlen / 2.0, 1.0 / rmax)
    if new is not None:
        types, lengths, curvs = new
        pose, lengths, curvs, r, el, feet = hfit(x, y, types, pose, lengths, curvs, sigma, {}, iterations)
    return pose, types, lengths, curvs, r, feet


def worstrun(r, sigma, tried):
    # the heaviest run of at least 3 consecutive residuals beyond 2.5 sigma on the same side, (first, last) or None
    # runs whose middle is in tried are skipped
    best = None
    i = 0
    n = len(r)
    lim = 2.5 * sigma
    while i < n:
        if abs(r[i]) <= lim:
            i += 1
            continue
        j = i
        while j + 1 < n and abs(r[j + 1]) > lim and (r[j + 1] > 0) == (r[i] > 0):
            j += 1
        if j - i >= 2 and (i + j) // 2 not in tried:
            w = sum(abs(v) for v in r[i:j + 1])
            if best is None or w > best[0]:
                best = (w, i, j)
        i = j + 1
    return None if best is None else (best[1], best[2])


def hcost(r, sigma):
    # robust cost (Huber at 2 sigma)
    c = 2.0 * sigma
    return sum(v * v if abs(v) <= c else c * (2 * abs(v) - c) for v in r)


def surveyhorizontal(x, y, sigma, minlen, rmax, roundstep, iterations=150):
    # returns (start pose, types, lengths, curvatures, residuals, element per point)
    n = len(x)
    s = [0.0]
    for k in range(1, n):
        s.append(s[-1] + math.hypot(x[k] - x[k - 1], y[k] - y[k - 1]))
    hd = []
    for k in range(n - 1):
        a = math.atan2(y[k + 1] - y[k], x[k + 1] - x[k])
        if hd:
            a = hd[-1] + (a - hd[-1] + math.pi) % TWOPI - math.pi
        hd.append(a)
    ms, mv, mw = chorddiagram(s, hd)
    if len(ms) < 2:
        types = ['L']; lengths = [s[-1]]; curvs = [0.0]; pose = (x[0], y[0], hd[0] if hd else 0.0)
    else:
        pieces = segmentdiagram(ms, mv, mw, sigma, minlen, 1.0 / rmax)
        jh = jointvalues(pieces)
        types = ['L' if pc[0] == 0 else 'A' for pc in pieces]
        lengths = [pc[2] - pc[1] for pc in pieces]
        curvs = [0.0 if pc[0] == 0 else (jh[k + 1] - jh[k]) / (pc[2] - pc[1]) for k, pc in enumerate(pieces)]
        pose = (x[0], y[0], jh[0])
    fixed = {}
    pose, types, lengths, curvs, r, feet = hfitprune(x, y, types, pose, lengths, curvs, sigma, minlen, rmax, iterations)
    # systematic misfits the segmentation missed (it sees heading errors, not their sum): split the element there and refit
    tried = set()
    ins = 0
    for ins in range(20):
        run = worstrun(r, sigma, tried)
        if run is None:
            break
        i0, i1 = run
        tried.add((i0 + i1) // 2)
        e = feet[(i0 + i1) // 2][1]
        ua = max(feet[i0][2] if feet[i0][1] == e else 0.0, 0.0)
        ub = feet[i1][2] if feet[i1][1] == e else lengths[e]
        if e == len(types) - 1 and feet[i1][1] == e:
            ub = feet[i1][2]
        # curvature that bends the chain by the peak misfit over the run (positive misfit = points left of the chain)
        peak = max(r[i0:i1 + 1], key=abs)
        kins = 8.0 * peak / max(ub - ua, 1.0) ** 2
        kins = math.copysign(min(max(abs(kins), 2.0 / rmax), 0.01), peak)
        if types[e] == 'L':
            # straight -> straight, short arc, straight
            ua = max(ua, 1.0); ub = max(ub, ua + 2.0)
            nt = ['L', 'A', 'L']; nl = [ua, ub - ua, lengths[e] - ub]; nc = [0.0, kins, 0.0]
            if e == len(types) - 1:
                nl[2] = 1e9
        else:
            um = (ua + ub) / 2.0
            nt = ['A', 'A']; nl = [um, lengths[e] - um]; nc = [curvs[e], curvs[e]]
        if e == 0 and ua <= 1.0 and types[e] == 'L':
            nt = ['A', 'L']; nl = [ub, lengths[e] - ub]; nc = [kins, 0.0]
        t2 = types[:e] + nt + types[e + 1:]
        l2 = lengths[:e] + nl + lengths[e + 1:]
        c2 = curvs[:e] + nc + curvs[e + 1:]
        res = hfitprune(x, y, t2, pose, l2, c2, sigma, minlen, rmax, 40)
        if hcost(res[4], sigma) < hcost(r, sigma) * 0.98:
            pose, types, lengths, curvs, r, feet = res
            tried = set()
    if ins > 0:
        pose, types, lengths, curvs, r, feet = hfitprune(x, y, types, pose, lengths, curvs, sigma, minlen, rmax, iterations)
    if roundstep > 0:
        for k, t in enumerate(types[:-1] if types[-1] == 'L' else types):
            if t == 'A':
                rr = round(1.0 / abs(curvs[k]) / roundstep) * roundstep
                if rr > 0:
                    fixed[k] = math.copysign(1.0 / rr, curvs[k])
        if fixed:
            pose, lengths, curvs, r, el, feet = hfit(x, y, types, pose, lengths, curvs, sigma, fixed, iterations)
    pose, types, lengths, curvs = trimchain(pose, lengths, curvs, types, (feet[0][1], feet[0][2]), (feet[-1][1], feet[-1][2]))
    el = [f[1] - feet[0][1] for f in feet]
    return pose, types, lengths, curvs, r, el


# -- vertical chain: z0, g0 at s0, elements with length and rate of grade change (0 = grade)
# z = Z + G u + c u^2 / 2 along an element

class VChain(object):
    def __init__(self, s, z, types):
        self.s = s; self.z = z; self.types = types
        self.state = None

    def unpack(self, p):
        m = len(self.types)
        lengths = [math.exp(min(v, 12.0)) for v in p[2:2 + m - 1]] + [1e9]
        curvs = []
        q = 2 + m - 1
        for t in self.types:
            if t == 'L':
                curvs.append(0.0)
            else:
                curvs.append(p[q] / 1e4); q += 1
        return p[0], p[1], lengths, curvs

    def pack(self, z0, g0, lengths, curvs):
        p = [z0, g0] + [math.log(max(L, 1e-3)) for L in lengths[:-1]]
        for k, t in enumerate(self.types):
            if t == 'A':
                p.append(curvs[k] * 1e4)
        return p

    def evaluate(self, z0, g0, lengths, curvs):
        # element start station/height/grade, element and model height/grade per point
        S = [self.s[0]]; Z = [z0]; G = [g0]
        for L, c in zip(lengths[:-1], curvs[:-1]):
            Z.append(Z[-1] + G[-1] * L + c * L * L / 2.0); G.append(G[-1] + c * L); S.append(S[-1] + L)
        el = []; zm = []; gm = []
        e = 0
        m = len(lengths)
        for si in self.s:
            while e + 1 < m and si >= S[e + 1]:
                e += 1
            u = si - S[e]
            el.append(e); zm.append(Z[e] + G[e] * u + curvs[e] * u * u / 2.0); gm.append(G[e] + curvs[e] * u)
        return S, Z, G, el, zm, gm

    def residuals(self, p, store):
        z0, g0, lengths, curvs = self.unpack(p)
        st = self.evaluate(z0, g0, lengths, curvs)
        if store:
            self.state = (lengths, curvs) + st
        return [a - b for a, b in zip(st[4], self.z)]

    def jacobian(self, p):
        lengths, curvs, S, Z, G, el, zm, gm = self.state
        n = len(self.s); m = len(self.types)
        s = self.s
        firstpt = [n] * (m + 1)
        for i in range(n - 1, -1, -1):
            for k in range(el[i], -1, -1):
                if firstpt[k] > i:
                    firstpt[k] = i
                else:
                    break
        cols = [(0, [1.0] * n), (0, [si - S[0] for si in s])]
        for k in range(m - 1):
            # log length: everything after the element end shifts along, d/dL = -z'(s) + G_E + c_k (s - S_E)
            L = lengths[k]; se = S[k + 1]; ge = G[k + 1]; ck = curvs[k]
            lo = firstpt[k + 1]
            cols.append((lo, [L * (-gm[i] + ge + ck * (s[i] - se)) if el[i] > k else 0.0 for i in range(lo, n)]))
        for k, t in enumerate(self.types):
            if t != 'A':
                continue
            L = lengths[k]; se = S[k + 1] if k + 1 < m else 0.0; sk = S[k]
            lo = firstpt[k]
            col = []
            for i in range(lo, n):
                if el[i] == k:
                    u = s[i] - sk
                    col.append(u * u / 2.0 / 1e4)
                elif el[i] > k:
                    col.append((L * L / 2.0 + L * (s[i] - se)) / 1e4)
                else:
                    col.append(0.0)
            cols.append((lo, col))
        return cols


def vfit(s, z, types, z0, g0, lengths, curvs, sigma, iterations):
    ch = VChain(s, z, types)
    p = ch.pack(z0, g0, lengths, curvs)
    p, r = lmrobust(p, ch, sigma, iterations)
    z0, g0, lengths, curvs = ch.unpack(p)
    return z0, g0, lengths, curvs, r, ch.state[5], ch.state[2]


def vfitprune(s, z, types, z0, g0, lengths, curvs, sigma, minlen, iterations):
    # fit, drop elements that shrink away or end up beyond the last point, flatten curves TBC can't handle, refit
    for rnd in range(8):
        z0, g0, lengths, curvs, r, el, S = vfit(s, z, types, z0, g0, lengths, curvs, sigma, 30 if rnd < 7 else iterations)
        changed = False
        # elements starting beyond the last point
        while len(types) > 1 and S[len(types) - 1] >= s[-1]:
            types = types[:-1]; lengths = lengths[:-1]; curvs = curvs[:-1]
            changed = True
        new = prunechain(types, lengths, curvs, minlen / 2.0, 0.0)
        if new is not None:
            types, lengths, curvs = new
            changed = True
        # a curve that bows less than 1mm off its chord becomes a grade (TBC can't handle flat parabolas)
        types = list(types); curvs = list(curvs)
        for k in range(len(types) - 1):
            if types[k] == 'A' and abs(curvs[k]) * lengths[k] ** 2 / 8.0 < 0.001:
                types[k] = 'L'; curvs[k] = 0.0; changed = True
        merged = prunechain(types, lengths, curvs, 0.0, 0.0, False)
        if merged is not None:
            types, lengths, curvs = merged
        if not changed:
            break
    z0, g0, lengths, curvs, r, el, S = vfit(s, z, types, z0, g0, lengths, curvs, sigma, iterations)
    return z0, g0, list(types), list(lengths), list(curvs), r, el, S


def surveyvertical(s, z, sigma, minlen, iterations=100):
    # s must be increasing, returns (z0, g0, types, lengths, curvs, residuals, element per point)
    gr = []
    for k in range(len(s) - 1):
        ds = s[k + 1] - s[k]
        gr.append((z[k + 1] - z[k]) / ds if ds > 1e-6 else 0.0)
    ms, mv, mw = chorddiagram(s, gr)
    if len(ms) < 2:
        types = ['L']; lengths = [1e9]; curvs = [0.0]; z0 = z[0]; g0 = gr[0] if gr else 0.0
    else:
        pieces = segmentdiagram(ms, mv, mw, sigma, minlen, 0.0)
        jg = jointvalues(pieces)
        types = ['L' if pc[0] == 0 else 'A' for pc in pieces]
        lengths = [pc[2] - pc[1] for pc in pieces]
        curvs = [0.0 if pc[0] == 0 else (jg[k + 1] - jg[k]) / (pc[2] - pc[1]) for k, pc in enumerate(pieces)]
        z0 = z[0]; g0 = jg[0]
        lengths[0] -= s[0] - pieces[0][1]
    z0, g0, types, lengths, curvs, r, el, S = vfitprune(s, z, types, z0, g0, lengths, curvs, sigma, minlen, 40)
    # systematic misfits the segmentation missed: split the element there and refit
    tried = set()
    ins = 0
    for ins in range(20):
        run = worstrun(r, sigma, tried)
        if run is None:
            break
        i0, i1 = run
        tried.add((i0 + i1) // 2)
        e = el[(i0 + i1) // 2]
        ua = max(s[i0] - S[e], 0.0)
        ub = min(s[i1] - S[e], lengths[e] if e < len(types) - 1 else 1e9)
        peak = max(r[i0:i1 + 1], key=abs)
        # the model is above the points (positive residual) -> the curve has to bend down
        cins = -8.0 * peak / max(ub - ua, 1.0) ** 2
        if types[e] == 'L':
            ua = max(ua, 1.0); ub = max(ub, ua + 2.0)
            nt = ['L', 'A', 'L']; nl = [ua, ub - ua, lengths[e] - ub]; nc = [0.0, cins, 0.0]
            if e == 0 and ua <= 1.0:
                nt = ['A', 'L']; nl = [ub, lengths[e] - ub]; nc = [cins, 0.0]
        else:
            um = (ua + ub) / 2.0
            nt = ['A', 'A']; nl = [um, lengths[e] - um]; nc = [curvs[e], curvs[e]]
        res = vfitprune(s, z, types[:e] + nt + types[e + 1:], z0, g0, lengths[:e] + nl + lengths[e + 1:],
                        curvs[:e] + nc + curvs[e + 1:], sigma, minlen, 40)
        if hcost(res[5], sigma) < hcost(r, sigma) * 0.98:
            z0, g0, types, lengths, curvs, r, el, S = res
            tried = set()
    if ins > 0:
        z0, g0, types, lengths, curvs, r, el, S = vfitprune(s, z, types, z0, g0, lengths, curvs, sigma, minlen, iterations)
    lengths[-1] = s[-1] - S[-1]
    return z0, g0, types, lengths, curvs, r, el

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
        self.hscatter.DistanceMin = 0.0001
        self.vscatter.DistanceMin = 0.0001
        self.hminlength.DistanceMin = 0.1
        self.vminlength.DistanceMin = 0.1
        self.maxradius.DistanceMin = 1.0
        self.roundradius.DistanceMin = 0.0
        for c in [self.minarcchords, self.minvcchords]:
            c.NumberOfDecimals = 0
            c.MinValue = 2
            c.MaxValue = 1000

        self.modedesign.Checked += self.ModeChanged
        self.modesurvey.Checked += self.ModeChanged

        try:
            self.SetDefaultOptions()
        except:
            pass
        self.ModeChanged(None, None)

    def ModeChanged(self, sender, e):
        # Design mode: interpolating fit with a hard tolerance, Survey mode: least squares fit to the point scatter
        survey = self.modesurvey.IsChecked == True
        for c in [self.designh, self.designv]:
            c.Visibility = Visibility.Collapsed if survey else Visibility.Visible
        for c in [self.surveyh, self.surveyv, self.surveynote]:
            c.Visibility = Visibility.Visible if survey else Visibility.Collapsed

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
        # lines, straights, arcs, vertical curves, max dev H, max dev V, sum of squares H, count H, sum of squares V, count V
        totals = [0, 0, 0, 0, 0.0, 0.0, 0.0, 0, 0.0, 0]
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
        if totals[7] > 0:
            self.success.Content += '\nRMS H {:.4f}'.format(math.sqrt(totals[6] / totals[7]))
            if totals[9] > 0:
                self.success.Content += ' / V {:.4f}'.format(math.sqrt(totals[8] / totals[9]))

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

        survey = self.modesurvey.IsChecked == True
        totals[0] += 1

        # ---- horizontal, as a 2D linestring - the elevations go into the vertical tab
        if survey:
            newline, hlength = self.horizontalsurvey(wv, name, x, y, ox, oy, report, totals)
        else:
            newline, hlength = self.horizontaldesign(wv, name, x, y, ox, oy, report, totals)

        # TBC computes the tangent arcs from the tangency, the result should only differ marginally from the fit
        polyseg = newline.ComputePolySeg()
        lengthdiff = polyseg.ComputeStationing() - polyseg.BeginStation - hlength
        if abs(lengthdiff) > 0.005:
            self.error.Content += '\n"{}" - Linestring length differs by {:.3f} from the fit'.format(name, lengthdiff)

        # ---- vertical, on the new horizontal
        # every original 3D vertex is projected onto the linestring TBC computed, giving its chainage
        # the vertical is then fitted to those chainages/elevations, so it covers the real line from start to end
        # (TBC computes the tangent arcs from the tangency, the length can differ from the horizontal fit)
        if not is3d:
            return
        if survey:
            vs, vz, dropped, maxoffset = self.projectall(polyseg, x, y, z, ox, oy)
            if maxoffset > 6.0 * self.hscatter.Distance:
                self.error.Content += '\n"{}" - Vertices up to {:.3f} off the new Linestring'.format(name, maxoffset)
        else:
            vs, vz, dropped, maxoffset = self.projectvertices(polyseg, x, y, z, ox, oy)
            if maxoffset > self.htol.Distance * 2:
                self.error.Content += '\n"{}" - Vertices up to {:.3f} off the new Linestring'.format(name, maxoffset)
        if dropped:
            self.error.Content += '\n"{}" - {} Vertices not used for the Vertical, Chainage not increasing'.format(name, dropped)
        if len(vs) < 2:
            return

        # vertical elements: (type 'G' grade / 'P' parabola, start chainage, start height, end chainage, end height, grade at start)
        if not self.fitvertical.IsChecked:
            # keep the original elevations - every vertex becomes a VPI
            vitems = [('G', vs[k], vz[k], vs[k + 1], vz[k + 1], 0.0) for k in range(len(vs) - 1)]
        elif survey:
            vitems = self.verticalsurvey(name, vs, vz, report, totals)
        else:
            vitems = self.verticaldesign(name, vs, vz, report, totals)

        vpoly = PolySeg.PolySeg()
        for t, sa, za, sb, zb, ga in vitems:
            if t == 'G':
                vpoly.Add(SegmentLine(Point3D(sa, za, 0), Point3D(sb, zb, 0)))
            else:
                half = (sb - sa) / 2.0
                vpi = Point3D(sa + half, za + ga * half, 0)
                vpoly.Add(ParabolaSegment(Point3D(sa, za, 0), vpi, Point3D(sb, zb, 0)))
        Linestring.ConvertPolySegToLinestringVertical(newline, vpoly, 0)

    def buildlinestring(self, wv, name, startx, starty, items):
        # items: (linestring element type, end x, end y, radius, side, sweep) in world coordinates
        # element types: Straight, Arc (radius and direction), Begin / TangentTangent (tangent arcs, TBC computes the radius)
        newline = wv.Add(clr.GetClrType(Linestring))
        newline.Name = name
        newline.Layer = self.layerpicker.SelectedSerialNumber

        el = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
        el.Position = Point3D(startx, starty)
        newline.AppendElement(el)
        for etype, ex, ey, radius, side, sweep in items:
            if etype == "Straight":
                el = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
            elif etype == "Arc":
                el = ElementFactory.Create(clr.GetClrType(IArcSegment), clr.GetClrType(IXYZLocation))
                el.Radius = radius
                # LeftRight is a Direction enum, take the type straight from the property instead of importing it
                directiontype = clr.GetClrType(IArcSegment).GetProperty("LeftRight").PropertyType
                el.LeftRight = System.Enum.Parse(directiontype, "Left" if side > 0 else "Right")
                el.LargeSmall = ArcLargeSmall.Large if sweep > math.pi else ArcLargeSmall.Small
            else:
                el = ElementFactory.Create(clr.GetClrType(ITangentArcSegment), clr.GetClrType(IXYZLocation))
                el.TangentType = ArcTangentType.Begin if etype == "Begin" else ArcTangentType.TangentTangent
            el.Position = Point3D(ex, ey)
            newline.AppendElement(el)
        return newline

    def horizontaldesign(self, wv, name, x, y, ox, oy, report, totals):
        htol = self.htol.Distance
        helems = segmenthorizontal(x, y, htol, self.maxradius.Distance, int(self.minarcchords.Value))
        if self.roundradius.Distance > 0:
            for e in helems:
                if e["t"] == "A":
                    roundarc(x, y, e, self.roundradius.Distance, htol)
        s = stationing(x, y, helems)

        etypes, forced = arctangency(x, y, helems, htol)
        if any(forced):
            self.error.Content += '\n"{}" - {} Arc(s) made tangent although the kink exceeds the Tolerance - see CSV'.format(name, sum(forced))
        items = []
        for e, etype in zip(helems, etypes):
            items.append((etype, x[e["j"]] + ox, y[e["j"]] + oy, e.get("r", 0.0), e.get("side", 1), e.get("sweep", 0.0)))
        newline = self.buildlinestring(wv, name, x[0] + ox, y[0] + oy, items)

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
        return newline, s[-1]

    def horizontalsurvey(self, wv, name, x, y, ox, oy, report, totals):
        # least squares fit, the chain is tangent continuous by construction, so every joint is a tangent joint
        pose, types, lengths, curvs, r, el = surveyhorizontal(x, y, self.hscatter.Distance, self.hminlength.Distance,
                                                              self.maxradius.Distance, self.roundradius.Distance)
        types = ['A' if t == 'A' and c != 0.0 else 'L' for t, c in zip(types, curvs)]
        poses = chainposes(pose, lengths, curvs)
        etypes = tangenttypes(types)
        items = []
        for k in range(len(types)):
            ex, ey, eh = poses[k + 1]
            radius = 1.0 / abs(curvs[k]) if types[k] == 'A' else 0.0
            items.append((etypes[k], ex + ox, ey + oy, radius, 1 if curvs[k] > 0 else -1, abs(curvs[k]) * lengths[k]))
        newline = self.buildlinestring(wv, name, pose[0] + ox, pose[1] + oy, items)

        m = len(types)
        st = 0.0
        for k in range(m):
            res = [abs(r[i]) for i in range(len(r)) if min(max(el[i], 0), m - 1) == k]
            dev = max(res) if res else 0.0
            totals[4] = max(totals[4], dev)
            if types[k] == 'L':
                totals[1] += 1
                typ = "Straight"; radius = ""; direction = ""
            else:
                totals[2] += 1
                typ = "Arc"; radius = "{:.4f}".format(1.0 / abs(curvs[k])); direction = "Left" if curvs[k] > 0 else "Right"
            report.append('"{}",H{},{},{},{:.4f},{:.4f},{:.4f},{},{},,,,{:.4f},{:.4f},{:.4f},{:.4f},{:.6f},{},{:.5f}'.format(
                name, k + 1, typ, etypes[k], st, st + lengths[k], lengths[k], radius, direction,
                poses[k][0] + ox, poses[k][1] + oy, poses[k + 1][0] + ox, poses[k + 1][1] + oy, 0.0, len(res), dev))
            st += lengths[k]
        totals[6] += sum(v * v for v in r); totals[7] += len(r)
        return newline, st

    def verticaldesign(self, name, vs, vz, report, totals):
        velems = segmentvertical(vs, vz, self.vtol.Distance, int(self.minvcchords.Value))
        vitems = []
        for k, e in enumerate(velems):
            i = e["i"]; j = e["j"]
            gin, gout = vgrades(vs, vz, e)
            vitems.append(('G' if e["t"] == "G" else 'P', vs[i], vz[i], vs[j], vz[j], gin))
            dev = vdeviation(vs, vz, e)
            totals[5] = max(totals[5], dev)
            if e["t"] == "G":
                typ = "Grade"; kvalue = ""
            else:
                totals[3] += 1
                typ = "Vertical Curve"
                kvalue = "{:.3f}".format((vs[j] - vs[i]) / abs((gout - gin) * 100.0)) if abs(gout - gin) > 1e-12 else ""
            report.append('"{}",V{},{},,{:.4f},{:.4f},{:.4f},,,{:.4f},{:.4f},{},{:.4f},{:.4f},{:.4f},{:.4f},,{},{:.5f}'.format(
                name, k + 1, typ, vs[i], vs[j], vs[j] - vs[i], gin * 100.0, gout * 100.0, kvalue,
                vs[i], vz[i], vs[j], vz[j], j - i + 1, dev))
        return vitems

    def verticalsurvey(self, name, vs, vz, report, totals):
        z0, g0, types, lengths, curvs, r, el = surveyvertical(vs, vz, self.vscatter.Distance, self.vminlength.Distance)
        # the profile starts at chainage 0, the first vertex can sit a bit further along
        s0 = vs[0]
        za = z0 - g0 * s0 + curvs[0] * s0 * s0 / 2.0
        ga = g0 - curvs[0] * s0
        lengths = list(lengths)
        lengths[0] += s0
        sa = 0.0
        vitems = []
        m = len(types)
        for k in range(m):
            L = lengths[k]; c = curvs[k]
            sb = sa + L; zb = za + ga * L + c * L * L / 2.0; gb = ga + c * L
            vitems.append(('G' if types[k] == 'L' else 'P', sa, za, sb, zb, ga))
            res = [abs(r[i]) for i in range(len(r)) if el[i] == k]
            dev = max(res) if res else 0.0
            totals[5] = max(totals[5], dev)
            if types[k] == 'L':
                typ = "Grade"; kvalue = ""
            else:
                totals[3] += 1
                typ = "Vertical Curve"
                kvalue = "{:.3f}".format(L / abs((gb - ga) * 100.0)) if abs(gb - ga) > 1e-12 else ""
            report.append('"{}",V{},{},,{:.4f},{:.4f},{:.4f},,,{:.4f},{:.4f},{},{:.4f},{:.4f},{:.4f},{:.4f},,{},{:.5f}'.format(
                name, k + 1, typ, sa, sb, L, ga * 100.0, gb * 100.0, kvalue, sa, za, sb, zb, len(res), dev))
            sa = sb; za = zb; ga = gb
        totals[8] += sum(v * v for v in r); totals[9] += len(r)
        return vitems

    def projectall(self, polyseg, x, y, z, ox, oy):
        # chainage of every original vertex on the new horizontal, relative to its begin station, sorted
        # a survey pickup can step back a little, vertices with the same chainage as the previous one are dropped
        begin = polyseg.BeginStation
        total = polyseg.ComputeStationing() - begin
        onl = clr.StrongBox[Point3D]()
        sta = clr.StrongBox[float]()
        pairs = []
        dropped = 0; maxoffset = 0.0
        for k in range(len(x)):
            px = x[k] + ox; py = y[k] + oy
            if not polyseg.FindPointFromPoint(Point3D(px, py, 0), onl, sta):
                dropped += 1
                continue
            maxoffset = max(maxoffset, math.hypot(onl.Value.X - px, onl.Value.Y - py))
            pairs.append((min(max(sta.Value - begin, 0.0), total), z[k]))
        pairs.sort()
        vs = []; vz = []
        for c, h in pairs:
            if vs and c <= vs[-1] + 1e-4:
                dropped += 1
                continue
            vs.append(c); vz.append(h)
        return vs, vz, dropped, maxoffset

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

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
    "surface1droplist": 0,
    "surface2droplist": 0,
    "layerpicker":      0,
}

def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_IntersectSurfaceAnySurface"
    cmdData.CommandName = "SCR_IntersectSurfaceAnySurface"
    cmdData.Caption = "_SCR_IntersectSurfaceAnySurface"
    cmdData.UIForm = "SCR_IntersectSurfaceAnySurface"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "DTM"
        cmdData.ShortCaption = "Intersect Surfaces"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.06
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Intersect any Surfaces"
        cmdData.ToolTipTextFormatted = "Intersect any kind Surface"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_IntersectSurfaceAnySurface(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_IntersectSurfaceAnySurface.xaml") as s:
            wpf.LoadComponent (self, s)
        self.currentProject = currentProject
        self.macroFileFolder = macroFileFolder

    def SetDefaultOptions(self):
        SCROptions.LoadMacroOptions(self, "SCR_IntersectSurfaceAnySurface", _OPTIONS, self.currentProject)

    def SaveOptions(self):
        SCROptions.SaveMacroOptions(self, "SCR_IntersectSurfaceAnySurface", _OPTIONS)

    def HelpClicked(self, cmd, e):
        webbrowser.open("C:\ProgramData\Trimble\MacroCommands3\SCR Macros\MacroHelp\MacroHelp.htm#" + type(self).__name__)

    def OnLoad(self, cmd, buttons, event):
        self.okBtn = buttons[0]
        buttons[2].Content = "Help"
        buttons[2].Visibility = Visibility.Visible
        buttons[2].Click += self.HelpClicked
        self.Caption = cmd.Command.Caption
        types = Array [Type] ([clr.GetClrType (ProjectedSurface)]) + Array [Type] (SurfaceTypeLists.AllWithCutFillMap)    # we fill an array with TBC object types, we could combine different types

                                                                                                                                                                                                                                                                                                                           # +Array[Type](SurfaceTypeLists.AllWithCutFillMap)
        self.surface1droplist.FilterByEntityTypes = types    # we fill the dropdownlist by applying that types array as filter
        self.surface2droplist.FilterByEntityTypes = types

        self.surface1droplist.AllowNone = False              # our list shall not show an empty field
        self.surface2droplist.AllowNone = False              # our list shall not show an empty field
                                                             # I haven't found
 # a way yet to auto-select the top most value in the list, it says it's read only

        self.surface1droplist.ValueChanged += self.surfacedroplistChanged    # elevation field and the dropdown list the ability to react to changes
        self.surface2droplist.ValueChanged += self.surfacedroplistChanged    # elevation field and the dropdown list the ability to react to changes

        try:
            self.SetDefaultOptions()
        except:
            pass


    def surfacedroplistChanged(self, sender, e):        # in case we select a new surface from the list we update the min/max
                                                          # textfields
        wv = self.currentProject [Project.FixedSerial.WorldView]
        if self.surface1droplist.SelectedSerial != 0 and self.surface2droplist.SelectedSerial != 0:
            surface1 = wv.Lookup (self.surface1droplist.SelectedSerial) # we get our selected surface as object
            surface2 = wv.Lookup (self.surface2droplist.SelectedSerial) # we get our selected surface as object
            nTri1 = surface1.NumberOfTrianglesWithMaterial
            nTri2 = surface2.NumberOfTrianglesWithMaterial
            
            self.Label_nTri1.Content = 'Surface 1: ' + str (nTri1) # we get some surface values into some labels
            self.Label_nTri2.Content = 'Surface 2: ' + str (nTri2)
            #self.Label_iterations.Content = 'Iterations: ' + str (nTri1 * nTri2)

        
    def CancelClicked(self, thisCmd, args):
        thisCmd.CloseUICommand ()


    def OkClicked(self, thisCmd, e):
        self.label_benchmark.Content = ''
        self.error.Content = ''
        start_t = timer()
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)

        wv = self.currentProject[Project.FixedSerial.WorldView]

        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                surface1 = wv.Lookup(self.surface1droplist.SelectedSerial)
                surface2 = wv.Lookup(self.surface2droplist.SelectedSerial)
                nTri1 = surface1.NumberOfTriangles
                nTri2 = surface2.NumberOfTriangles

                # bail early if the two surface bounding boxes don't overlap at all
                s1_lim = surface1.GetLimits()
                s2_lim = surface2.GetLimits()
                if Limits3D(s1_lim[0], s1_lim[1]).LimitsInLimits(Limits3D(s2_lim[0], s2_lim[1]), True)[0] == Side.Out:
                    failGuard.Commit()
                    self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
                    UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
                    return

                is_proj1 = isinstance(surface1, ProjectedSurface)
                is_proj2 = isinstance(surface2, ProjectedSurface)
                null1 = surface1.NullMaterialIndex()
                null2 = surface2.NullMaterialIndex()

                # pull all API data into Python lists once so the hot loop never crosses .NET boundary
                s2_data = []
                octree_items = []   # (cx, cy, cz, idx) for the spatial index
                s2_max_hd = 0.0     # largest half-diagonal of any s2 bbox — used to expand octree queries

                ProgressBar.TBC_ProgressBar.Title = "Precomputing Surface 2 ..."
                last_t = timer()
                for i in range(nTri2):
                    if timer() - last_t > 0.5:
                        if ProgressBar.TBC_ProgressBar.SetProgress(i * 100 // nTri2):
                            break
                        last_t = timer()
                    if surface2.GetTriangleMaterial(i) == null2:
                        continue
                    if is_proj2:
                        v0 = surface2.TransformPointToWorldDelegate(surface2.GetVertexPoint(surface2.GetTriangleIVertex(i, 0)))
                        v1 = surface2.TransformPointToWorldDelegate(surface2.GetVertexPoint(surface2.GetTriangleIVertex(i, 1)))
                        v2 = surface2.TransformPointToWorldDelegate(surface2.GetVertexPoint(surface2.GetTriangleIVertex(i, 2)))
                    else:
                        v0 = surface2.GetVertexPoint(surface2.GetTriangleIVertex(i, 0))
                        v1 = surface2.GetVertexPoint(surface2.GetTriangleIVertex(i, 1))
                        v2 = surface2.GetVertexPoint(surface2.GetTriangleIVertex(i, 2))
                    bx0 = min(v0.X, v1.X, v2.X); bx1 = max(v0.X, v1.X, v2.X)
                    by0 = min(v0.Y, v1.Y, v2.Y); by1 = max(v0.Y, v1.Y, v2.Y)
                    bz0 = min(v0.Z, v1.Z, v2.Z); bz1 = max(v0.Z, v1.Z, v2.Z)
                    dx = bx1 - bx0; dy = by1 - by0; dz = bz1 - bz0
                    hd = math.sqrt(dx*dx + dy*dy + dz*dz) * 0.5
                    if hd > s2_max_hd:
                        s2_max_hd = hd
                    idx = len(s2_data)
                    s2_data.append((v0, v1, v2, Plane3D(v0, v1, v2)[0], bx0, bx1, by0, by1, bz0, bz1))
                    cx = (v0.X + v1.X + v2.X) / 3.0
                    cy = (v0.Y + v1.Y + v2.Y) / 3.0
                    cz = (v0.Z + v1.Z + v2.Z) / 3.0
                    octree_items.append((cx, cy, cz, idx))
                ProgressBar.TBC_ProgressBar.Title = ""

                s2_tree = SCROctree()
                if octree_items:
                    s2_tree.build(octree_items)

                s1_data = []
                ProgressBar.TBC_ProgressBar.Title = "Precomputing Surface 1 ..."
                last_t = timer()
                for i in range(nTri1):
                    if timer() - last_t > 0.5:
                        if ProgressBar.TBC_ProgressBar.SetProgress(i * 100 // nTri1):
                            break
                        last_t = timer()
                    if surface1.GetTriangleMaterial(i) == null1:
                        continue
                    if is_proj1:
                        v0 = surface1.TransformPointToWorldDelegate(surface1.GetVertexPoint(surface1.GetTriangleIVertex(i, 0)))
                        v1 = surface1.TransformPointToWorldDelegate(surface1.GetVertexPoint(surface1.GetTriangleIVertex(i, 1)))
                        v2 = surface1.TransformPointToWorldDelegate(surface1.GetVertexPoint(surface1.GetTriangleIVertex(i, 2)))
                    else:
                        v0 = surface1.GetVertexPoint(surface1.GetTriangleIVertex(i, 0))
                        v1 = surface1.GetVertexPoint(surface1.GetTriangleIVertex(i, 1))
                        v2 = surface1.GetVertexPoint(surface1.GetTriangleIVertex(i, 2))
                    bx0 = min(v0.X, v1.X, v2.X); bx1 = max(v0.X, v1.X, v2.X)
                    by0 = min(v0.Y, v1.Y, v2.Y); by1 = max(v0.Y, v1.Y, v2.Y)
                    bz0 = min(v0.Z, v1.Z, v2.Z); bz1 = max(v0.Z, v1.Z, v2.Z)
                    s1_data.append((v0, v1, v2, Plane3D(v0, v1, v2)[0], bx0, bx1, by0, by1, bz0, bz1))
                ProgressBar.TBC_ProgressBar.Title = ""

                # for each s1 triangle, query the octree with the s1 bbox expanded by s2_max_hd —
                # guarantees any s2 triangle whose bbox overlaps s1's bbox has its centroid in range.
                D = s2_max_hd
                polysegs = List[PolySeg.PolySeg]()
                n_s1 = len(s1_data)

                ProgressBar.TBC_ProgressBar.Title = "Intersecting surfaces ..."
                last_t = timer()
                for ti, (s1_v1, s1_v2, s1_v3, plane1, t1x0, t1x1, t1y0, t1y1, t1z0, t1z1) in enumerate(s1_data):
                    if timer() - last_t > 0.5:
                        if ProgressBar.TBC_ProgressBar.SetProgress(ti * 100 // n_s1):
                            break
                        last_t = timer()

                    candidates = s2_tree.find_within_box(
                        t1x0 - D, t1x1 + D,
                        t1y0 - D, t1y1 + D,
                        t1z0 - D, t1z1 + D)

                    for idx in candidates:
                        v0, v1, v2, plane2, bx0, bx1, by0, by1, bz0, bz1 = s2_data[idx]
                        if bx0 > t1x1 or bx1 < t1x0 or by0 > t1y1 or by1 < t1y0 or bz0 > t1z1 or bz1 < t1z0:
                            continue

                        ip0 = plane2.IntersectLine(s1_v1, s1_v2)[5].point
                        ip1 = plane2.IntersectLine(s1_v1, s1_v3)[5].point
                        ip2 = plane2.IntersectLine(s1_v2, s1_v3)[5].point
                        ip3 = plane1.IntersectLine(v0, v1)[5].point
                        ip4 = plane1.IntersectLine(v0, v2)[5].point
                        ip5 = plane1.IntersectLine(v1, v2)[5].point

                        ptcount = 0
                        lp0 = lp1 = None
                        for ip in (ip0, ip1, ip2, ip3, ip4, ip5):
                            if PointInsideTriangle3D(s1_v1, s1_v2, s1_v3, ip) and \
                               PointInsideTriangle3D(v0, v1, v2, ip):
                                if ptcount == 0:
                                    lp0 = ip
                                    ptcount = 1
                                elif ip != lp0:
                                    lp1 = ip
                                    ptcount = 2
                                    break

                        if ptcount == 2:
                            ps = PolySeg.PolySeg()
                            ps.Add(lp0)
                            ps.Add(lp1)
                            polysegs.Add(ps.Clone())

                ProgressBar.TBC_ProgressBar.Title = ""

                # join touching segments into continuous linestrings and write to worldview
                self.connect_linetuples(polysegs)

                failGuard.Commit()
                self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
                UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())

        except Exception as ex:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
            UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
            self.error.Content += '\nan Error occurred - Result probably incomplete\n' + str(exc_type) + '\n' + str(exc_obj) + '\nLine ' + str(exc_tb.tb_lineno)

        end_t = timer()
        self.label_benchmark.Content = 'elapsed time: ' + str(timedelta(seconds=end_t - start_t))
        self.SaveOptions()
    
    def connect_linetuples(self, polysegs):
        wv = self.currentProject[Project.FixedSerial.WorldView]
        PolySeg.PolySeg.JoinTouchingPolysegs(polysegs)
        for p in polysegs:
            if p and p.NumberOfNodes > 1:
                l = wv.Add(Linestring)
                l.Layer = self.layerpicker.SelectedSerialNumber
                l.Append(p, None, False, False)
             


def PointInsideTriangle3D(p1, p2, p3, p4):

    ax = p2.X - p1.X; ay = p2.Y - p1.Y; az = p2.Z - p1.Z
    bx = p3.X - p1.X; by = p3.Y - p1.Y; bz = p3.Z - p1.Z
    wx = p4.X - p1.X; wy = p4.Y - p1.Y; wz = p4.Z - p1.Z
    aa = ax*ax + ay*ay + az*az
    ab = ax*bx + ay*by + az*bz
    bb = bx*bx + by*by + bz*bz
    wa = wx*ax + wy*ay + wz*az
    wb = wx*bx + wy*by + wz*bz
    d = ab*ab - aa*bb
    # rounding gives better results when the point is very close to a triangle edge
    s = round((ab*wb - bb*wa) / d, 5)
    if s < 0 or s > 1:
        return False
    t = round((ab*wa - aa*wb) / d, 5)
    return not (t < 0 or (s + t) > 1)

                     
                
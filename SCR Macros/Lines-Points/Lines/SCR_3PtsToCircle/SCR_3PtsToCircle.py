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
    "layerpicker":      8,
    "drawcircle":       True,     "drawcenter":      True,
    "threepoint":       True,     "bestfit":         False,
    "givenradius":      True,     "bestfitradius":   False,   "radiusedit":       1.0,
    "avgplane":         True,     "bestfitplane":    False,   "userdefplane":     False,
    "samplingdistance": 0.001,    "samplingcount":   15000.0,
    "standardmode":     True,     "mode3d":          False,   "defineplane":      False,
    "selectpoints":     True,     "pickpoints":      False,
    "coordCtl1": None, "coordCtl2": None, "coordCtl3": None,
    "coordCtl7": None, "coordCtl8": None, "coordCtl9": None,
}

def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_3PtsToCircle"
    cmdData.CommandName = "SCR_3PtsToCircle"
    cmdData.Caption = "_SCR_3PtsToCircle"
    cmdData.UIForm = "SCR_3PtsToCircle"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "Lines"
        cmdData.ShortCaption = "Circle/Center from 3 Points"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.10
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Circle/Center from 3 Points"
        cmdData.ToolTipTextFormatted = "Circle/Center from 3 Points"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_3PtsToCircle(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_3PtsToCircle.xaml") as s:
            wpf.LoadComponent (self, s)
        self.currentProject = currentProject
        self.macroFileFolder = macroFileFolder

        self.ViewOverlay = Guid.NewGuid()
        self.overlayBag = OverlayBag(self.ViewOverlay)

    def HelpClicked(self, cmd, e):
        webbrowser.open("C:\ProgramData\Trimble\MacroCommands3\SCR Macros\MacroHelp\MacroHelp.htm#" + type(self).__name__)


    def OnLoad(self, cmd, buttons, event):
        self.okBtn = buttons[0]
        buttons[2].Content = "Help"
        buttons[2].Visibility = Visibility.Visible
        buttons[2].Click += self.HelpClicked
        self.Caption = cmd.Command.Caption

        self.ignoreGotFocus = False
        self.selectionControls = [self.objs1, self.objs2]
        self.pointcloudType = clr.GetClrType(PointCloudRegion)

        for objs in self.selectionControls:
            optionMenu = SelectionContextMenuHandler()
            optionMenu.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
            objs.ButtonContextMenu = optionMenu

        self.objs1.IsEntityValidCallback = self.IsValidPoints
        self.objs2.IsEntityValidCallback = self.IsValidBestFit

        self.samplingcount.NumberOfDecimals = 0
        self.samplingcount.MinValue = 1
        self.samplingcount.MaxValue = 15000

        self.radiusedit.DistanceMin = 0.00000001

        self.coordCtl1.ValueChanged += self.Coord1Changed
        self.coordCtl2.ValueChanged += self.Coord2Changed
        self.coordCtl3.ValueChanged += self.Coord3Changed
        self.coordCtl4.ValueChanged += self.Coord4Changed
        self.coordCtl5.ValueChanged += self.Coord5Changed
        self.coordCtl6.ValueChanged += self.Coord6Changed

        self.coordCtl7.ValueChanged += self.Coord7Changed
        self.coordCtl8.ValueChanged += self.Coord8Changed
        self.coordCtl9.ValueChanged += self.Coord9Changed

        SCRExpanders.wire_pairs([
            (self.expander_threepoint, self.threepoint),
            (self.expander_mode3d, self.mode3d),
            (self.expander_defineplane, self.defineplane),
            (self.expander_bestfit, self.bestfit),
            (self.expander_userdefplane, self.userdefplane),
        ])
		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def Selection_PreviewGotFocus(self, sender, e):
        self.ignoreGotFocus = True
        for ctrl in self.selectionControls:
            active = (sender == ctrl)
            ctrl.ProcessGlobalSelectionChanges = active
            ctrl.UpdateTextOnSelectionChange = active

    def Selection_ValueChanged(self, sender, e):
        if self.ignoreGotFocus:
            self.ignoreGotFocus = False
            return

    def SetDefaultOptions(self):
        SCROptions.LoadMacroOptions(self, "SCR_3PtsToCircle", _OPTIONS, self.currentProject)

    def SaveOptions(self):
        SCROptions.SaveMacroOptions(self, "SCR_3PtsToCircle", _OPTIONS)

    def IsValidPoints(self, serial):
        o = self.currentProject.Concordance[serial]
        if isinstance(o, CadPoint):
            return True
        if isinstance(o, CoordPoint):
            return True
        return False

    def IsValidBestFit(self, serial):
        o = self.currentProject.Concordance[serial]
        if isinstance(o, CadPoint):
            return True
        if isinstance(o, CoordPoint):
            return True
        if isinstance(o, self.pointcloudType):
            return True
        return False

    def _coord(self, ctl):
        c = ctl.Coordinate
        if c.Is2D: c.Z = 0
        return c

    def Coord1Changed(self, ctrl, e):
        self.coordCtl2.CursorStyle = CursorStyle.CrossHair | CursorStyle.RubberLine
        if self.coordCtl1.ResultCoordinateSystem:
            self.coordCtl2.AnchorPoint = MousePosition(self.coordCtl1.ClickWindow, self.coordCtl1.Coordinate, self.coordCtl1.ResultCoordinateSystem)
        else:
            self.coordCtl2.AnchorPoint = None

        coord1 = self._coord(self.coordCtl1)
        if not coord1.Is3D:
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

        coord2 = self._coord(self.coordCtl2)
        if not coord2.Is3D:
            self.coordCtl2.StatusMessage = "No valid coordinate defined, must be 3D"
        else:
            self.coordCtl2.StatusMessage = ""

        self.drawoverlay()

    def Coord3Changed(self, ctrl, e):
        self.coordCtl4.CursorStyle = CursorStyle.CrossHair | CursorStyle.RubberLine
        if self.coordCtl3.ResultCoordinateSystem:
            self.coordCtl4.AnchorPoint = MousePosition(self.coordCtl3.ClickWindow, self.coordCtl3.Coordinate, self.coordCtl3.ResultCoordinateSystem)
        else:
            self.coordCtl4.AnchorPoint = None

        coord3 = self._coord(self.coordCtl3)
        if not coord3.Is3D:
            self.coordCtl3.StatusMessage = "No valid coordinate defined, must be 3D"
        else:
            self.coordCtl3.StatusMessage = ""

        self.drawoverlay()

    def Coord4Changed(self, ctrl, e):
        # set keyboard focus if change was due to mouse pick
        self.coordCtl5.CursorStyle = CursorStyle.CrossHair | CursorStyle.RubberLine
        if self.coordCtl4.ResultCoordinateSystem:
            self.coordCtl5.AnchorPoint = MousePosition(self.coordCtl4.ClickWindow, self.coordCtl4.Coordinate, self.coordCtl4.ResultCoordinateSystem)
        else:
            self.coordCtl5.AnchorPoint = None

        coord4 = self._coord(self.coordCtl4)
        if self.mode3d.IsChecked and not coord4.Is3D:
            self.coordCtl4.StatusMessage = "No valid coordinate defined, must be 3D"
        elif self.standardmode.IsChecked and not coord4.Is3D and not coord4.Is2D:
            self.coordCtl4.StatusMessage = "No valid coordinate defined"
        else:
            self.coordCtl4.StatusMessage = ""

    def Coord5Changed(self, ctrl, e):
        # set keyboard focus if change was due to mouse pick
        self.coordCtl6.CursorStyle = CursorStyle.CrossHair | CursorStyle.RubberLine
        if self.coordCtl5.ResultCoordinateSystem:
            self.coordCtl6.AnchorPoint = MousePosition(self.coordCtl5.ClickWindow, self.coordCtl5.Coordinate, self.coordCtl5.ResultCoordinateSystem)
        else:
            self.coordCtl6.AnchorPoint = None

        coord5 = self._coord(self.coordCtl5)
        if self.mode3d.IsChecked and not coord5.Is3D:
            self.coordCtl5.StatusMessage = "No valid coordinate defined, must be 3D"
        elif self.standardmode.IsChecked and not coord5.Is3D and not coord5.Is2D:
            self.coordCtl5.StatusMessage = "No valid coordinate defined"
        else:
            self.coordCtl5.StatusMessage = ""

    def Coord6Changed(self, ctrl, e):
        # set keyboard focus if change was due to mouse pick
        self.coordCtl4.CursorStyle = CursorStyle.CrossHair | CursorStyle.RubberLine
        if self.coordCtl6.ResultCoordinateSystem:
            self.coordCtl4.AnchorPoint = MousePosition(self.coordCtl6.ClickWindow, self.coordCtl6.Coordinate, self.coordCtl6.ResultCoordinateSystem)
        else:
            self.coordCtl4.AnchorPoint = None

        coord6 = self._coord(self.coordCtl6)
        if self.mode3d.IsChecked and not coord6.Is3D:
            self.coordCtl6.StatusMessage = "No valid coordinate defined, must be 3D"
        elif self.standardmode.IsChecked and not coord6.Is3D and not coord6.Is2D:
            self.coordCtl6.StatusMessage = "No valid coordinate defined"
        else:
            self.coordCtl6.StatusMessage = ""

        self.OkClicked(None, None)

    def Coord7Changed(self, ctrl, e):
        self.coordCtl8.CursorStyle = CursorStyle.CrossHair | CursorStyle.RubberLine
        if self.coordCtl7.ResultCoordinateSystem:
            self.coordCtl8.AnchorPoint = MousePosition(self.coordCtl7.ClickWindow, self.coordCtl7.Coordinate, self.coordCtl7.ResultCoordinateSystem)
        else:
            self.coordCtl8.AnchorPoint = None

        coord7 = self._coord(self.coordCtl7)
        if not coord7.Is3D:
            self.coordCtl7.StatusMessage = "No valid coordinate defined, must be 3D"
        else:
            self.coordCtl7.StatusMessage = ""

        self.drawoverlay()

    def Coord8Changed(self, ctrl, e):
        self.coordCtl9.CursorStyle = CursorStyle.CrossHair | CursorStyle.RubberLine
        if self.coordCtl8.ResultCoordinateSystem:
            self.coordCtl9.AnchorPoint = MousePosition(self.coordCtl8.ClickWindow, self.coordCtl8.Coordinate, self.coordCtl8.ResultCoordinateSystem)
        else:
            self.coordCtl9.AnchorPoint = None

        coord8 = self._coord(self.coordCtl8)
        if not coord8.Is3D:
            self.coordCtl8.StatusMessage = "No valid coordinate defined, must be 3D"
        else:
            self.coordCtl8.StatusMessage = ""

        self.drawoverlay()

    def Coord9Changed(self, ctrl, e):
        coord9 = self._coord(self.coordCtl9)
        if not coord9.Is3D:
            self.coordCtl9.StatusMessage = "No valid coordinate defined, must be 3D"
        else:
            self.coordCtl9.StatusMessage = ""

        self.drawoverlay()

    def DefinePlaneClicked(self, sender, e):
        self.drawoverlay()

    def DrawOverlayClicked(self, sender, e):
        self.drawoverlay()

    def drawoverlay(self):
        wv = self.currentProject[Project.FixedSerial.WorldView]
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)
        self.overlayBag = OverlayBag(self.ViewOverlay)

        if self.threepoint.IsChecked and self.defineplane.IsChecked:
            c1 = self._coord(self.coordCtl1)
            c2 = self._coord(self.coordCtl2)
            c3 = self._coord(self.coordCtl3)
        elif self.bestfit.IsChecked and self.userdefplane.IsChecked:
            c1 = self._coord(self.coordCtl7)
            c2 = self._coord(self.coordCtl8)
            c3 = self._coord(self.coordCtl9)
        else:
            return

        if c1.Is3D and c2.Is3D and c3.Is3D:
            self.overlayBag.AddPolyline(Array[Point3D]([c1, c2, c3, c1]), Color.Blue.ToArgb(), 5)
            self.overlayBag.AddMarker(c1, GraphicMarkerTypes.HollowCircle_IndependentColor, Color.Magenta.ToArgb(), "   V1", 0, 0, 2.0)
            self.overlayBag.AddMarker(c2, GraphicMarkerTypes.HollowCircle_IndependentColor, Color.Magenta.ToArgb(), "   V2", 0, 0, 2.0)
            self.overlayBag.AddMarker(c3, GraphicMarkerTypes.HollowCircle_IndependentColor, Color.Magenta.ToArgb(), "   V3", 0, 0, 2.0)
            array = Array[Guid]([DisplayWindow.Hoops3DViewGUID, DisplayWindow.HoopsPlanViewGUID])
            TrimbleOffice.TheOffice.MainWindow.AppViewManager.AddOverlayGeometry(array, self.overlayBag)

    def CancelClicked(self, thisCmd, args):
        thisCmd.CloseUICommand()

    def Dispose(self, cmd, disposing):
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)


    def RunThreePoint(self, wv):
        inputok = True
        plist = []
        if self.selectpoints.IsChecked:
            for p in self.objs1:
                if isinstance(p, CoordPoint) or isinstance(p, CadPoint):
                    plist.Add(clr.StrongBox[Point3D](p.Position))

        elif self.pickpoints.IsChecked:
            for ctl in [self.coordCtl4, self.coordCtl5, self.coordCtl6]:
                c = self._coord(ctl)
                if self.mode3d.IsChecked:
                    if c.Is3D:
                        plist.Add(clr.StrongBox[Point3D](c))
                else:
                    if c.Is3D or c.Is2D:
                        plist.Add(clr.StrongBox[Point3D](c))

        if plist.Count != 3:
            self.success.Content = '\nSelect 3 points'
            inputok = False

        if inputok:
            outcenter = clr.StrongBox[Point3D]()
            outradius = clr.StrongBox[float]()
            outdir = clr.StrongBox[bool]()
            outsmall = clr.StrongBox[bool]()
            outangle = clr.StrongBox[float]()

            if self.standardmode.IsChecked:
                Arc.GetThreePointArcProperties(plist[0], plist[1], plist[2], outcenter, outradius, outdir, outsmall, outangle)

                circlecenter = outcenter.Value
                circlecenter.Z = (plist[0].Value.Z + plist[1].Value.Z + plist[2].Value.Z) / 3

            elif self.mode3d.IsChecked:
                c1 = plist[0].Value
                c2 = plist[1].Value
                c3 = plist[2].Value

                if self.defineplane.IsChecked:
                    p1 = self._coord(self.coordCtl1)
                    p2 = self._coord(self.coordCtl2)
                    p3 = self._coord(self.coordCtl3)
                else:
                    p1 = c1
                    p2 = c2
                    p3 = c3

                p = Plane3D(p1, p2, p3)[0]
                nv = p.normal

                rottozero = Spinor3D.ComputeRotation(Vector3D(p1, p2), nv, Vector3D(1,0,0), Vector3D(0,0,1))
                matrixtozero = Matrix4D.BuildTransformMatrix(Vector3D(p1), Vector3D(p1, Point3D(0, 0, 0)), rottozero, Vector3D(1,1,1))
                matrixback = Matrix4D.Inverse(matrixtozero)

                c10 = matrixtozero.TransformPoint(c1)
                c20 = matrixtozero.TransformPoint(c2)
                c30 = matrixtozero.TransformPoint(c3)

                clist = []
                clist.Add(clr.StrongBox[Point3D](c10))
                clist.Add(clr.StrongBox[Point3D](c20))
                clist.Add(clr.StrongBox[Point3D](c30))

                Arc.GetThreePointArcProperties(clist[0], clist[1], clist[2], outcenter, outradius, outdir, outsmall, outangle)

                circlecenter = outcenter.Value
                circlecenter.Z = 0
                circlecenter = matrixback.TransformPoint(circlecenter)

            if circlecenter:
                if self.drawcenter.IsChecked:
                    cadPoint = wv.Add(clr.GetClrType(CadPoint))
                    cadPoint.Layer = self.layerpicker.SelectedSerialNumber
                    cadPoint.Point0 = circlecenter

                if self.drawcircle.IsChecked:
                    newcircle = wv.Add(clr.GetClrType(CadCircle))
                    newcircle.CenterPoint = circlecenter
                    newcircle.Radius = outradius.Value
                    newcircle.Layer = self.layerpicker.SelectedSerialNumber
                    newcircle.Name = 'Radius: ' + str(outradius.Value)
                    if self.mode3d.IsChecked:
                        newcircle.Normal = nv

    def BestFitCircle2D(self, plist):
        # Kasa algebraic least-squares: solves A*x + B*y - C = x^2 + y^2
        # cx = A/2, cy = B/2, r = sqrt(C + cx^2 + cy^2)
        # Points are centred first to avoid float64 overflow with large UTM-scale coordinates.
        n = plist.Count
        mx = sum(plist[i].X for i in range(n)) / n
        my = sum(plist[i].Y for i in range(n)) / n
        sum_x = sum_y = sum_x2 = sum_y2 = sum_xy = 0.0
        sum_xb = sum_yb = sum_b = 0.0
        for i in range(n):
            xi = plist[i].X - mx
            yi = plist[i].Y - my
            bi = xi*xi + yi*yi
            sum_x  += xi;  sum_y  += yi
            sum_x2 += xi*xi; sum_y2 += yi*yi; sum_xy += xi*yi
            sum_xb += xi*bi; sum_yb += yi*bi; sum_b  += bi

        def det3(a,b,c, d,e,f, g,h,i):
            return a*(e*i - f*h) - b*(d*i - f*g) + c*(d*h - e*g)

        D  = det3(sum_x2, sum_xy, sum_x,  sum_xy, sum_y2, sum_y,  sum_x,  sum_y,  n)
        if abs(D) < 1e-10: return None
        DA = det3(sum_xb, sum_xy, sum_x,  sum_yb, sum_y2, sum_y,  sum_b,  sum_y,  n)
        DB = det3(sum_x2, sum_xb, sum_x,  sum_xy, sum_yb, sum_y,  sum_x,  sum_b,  n)
        DC = det3(sum_x2, sum_xy, sum_xb, sum_xy, sum_y2, sum_yb, sum_x,  sum_y,  sum_b)

        cx = (DA / D) / 2.0
        cy = (DB / D) / 2.0
        r_sq = (DC / D) + cx*cx + cy*cy
        if r_sq < 0: return None
        return cx + mx, cy + my, math.sqrt(r_sq)

    def MiniMaxCircle2D(self, plist):
        # Minimises (max di - min di) — equal deviation above and below radius.
        # Subgradient: at each step move toward the farthest point and away from the closest.
        n = plist.Count
        result = self.BestFitCircle2D(plist)
        if result:
            cx, cy, _ = result
        else:
            cx = sum(plist[i].X for i in range(n)) / n
            cy = sum(plist[i].Y for i in range(n)) / n

        avg_d = sum(math.sqrt((plist[i].X-cx)**2 + (plist[i].Y-cy)**2) for i in range(n)) / n
        step = avg_d * 0.5

        for _ in range(2000):
            dists = [math.sqrt((plist[i].X-cx)**2 + (plist[i].Y-cy)**2) for i in range(n)]
            i_far   = max(range(n), key=lambda k: dists[k])
            i_close = min(range(n), key=lambda k: dists[k])
            spread = dists[i_far] - dists[i_close]
            if spread < 1e-10: break

            dx = ((plist[i_far].X - cx)   / dists[i_far] +
                  (cx - plist[i_close].X) / dists[i_close])
            dy = ((plist[i_far].Y - cy)   / dists[i_far] +
                  (cy - plist[i_close].Y) / dists[i_close])
            dnorm = math.sqrt(dx*dx + dy*dy)
            if dnorm < 1e-10: break

            cx_new = cx + step * dx / dnorm
            cy_new = cy + step * dy / dnorm

            dists_new = [math.sqrt((plist[i].X-cx_new)**2 + (plist[i].Y-cy_new)**2) for i in range(n)]
            if max(dists_new) - min(dists_new) < spread:
                cx, cy = cx_new, cy_new
                step = min(step * 1.1, spread * 0.5)
            else:
                step *= 0.5
                if step < 1e-12: break

        dists = [math.sqrt((plist[i].X-cx)**2 + (plist[i].Y-cy)**2) for i in range(n)]
        r = (max(dists) + min(dists)) / 2.0
        return cx, cy, r

    def MiniMaxCenter2D(self, plist, radius):
        # Minimises max_i(|di - radius|) — equal deviation inside and outside.
        # At each step: if outer deviation dominates move toward farthest point,
        # if inner deviation dominates move away from closest point.
        n = plist.Count
        result = self.BestFitCircle2D(plist)
        if result:
            cx, cy, _ = result
        else:
            cx = sum(plist[i].X for i in range(n)) / n
            cy = sum(plist[i].Y for i in range(n)) / n

        step = radius * 0.5

        def obj(cx, cy):
            return max(abs(math.sqrt((plist[i].X-cx)**2 + (plist[i].Y-cy)**2) - radius) for i in range(n))

        for _ in range(2000):
            dists = [math.sqrt((plist[i].X-cx)**2 + (plist[i].Y-cy)**2) for i in range(n)]
            errs  = [d - radius for d in dists]
            e_pos = max(errs)
            e_neg = min(errs)
            if abs(e_pos) < 1e-10 and abs(e_neg) < 1e-10: break

            if e_pos >= -e_neg:
                i = errs.index(e_pos)
                dx = (plist[i].X - cx) / dists[i]
                dy = (plist[i].Y - cy) / dists[i]
            else:
                i = errs.index(e_neg)
                dx = (cx - plist[i].X) / dists[i]
                dy = (cy - plist[i].Y) / dists[i]

            cx_new = cx + step * dx
            cy_new = cy + step * dy

            if obj(cx_new, cy_new) < obj(cx, cy):
                cx, cy = cx_new, cy_new
                step = min(step * 1.1, radius * 0.5)
            else:
                step *= 0.5
                if step < 1e-12: break

        return cx, cy


    def _buildPlist(self):
        plist = []
        cloudselectionids = []
        cloudintegration = None
        for p in self.objs2:
            if isinstance(p, CoordPoint) or isinstance(p, CadPoint):
                pos = p.Position
                if pos.Is2D or math.isnan(pos.Z): pos.Z = 0
                plist.Add(pos)
            elif isinstance(p, self.pointcloudType):
                cloudselectionids.Add(p.Integration.GetSelectedCloudId())
                cloudintegration = p.Integration.PointCloudDatabase.Integration
        if cloudselectionids.Count > 0 and cloudintegration is not None:
            cps = cloudintegration.GetSelectedPoints(cloudselectionids)
            if cps.Count <= int(self.samplingcount.Value):
                for cp in cps:
                    plist.Add(Point3D(cp.X, cp.Y, cp.Z))
            else:
                cpssampleid = cloudintegration.CreateSpatiallySampledCloud(cloudselectionids, self.samplingdistance.Distance, int(self.samplingcount.Value))
                cps2 = cloudintegration.GetPoints(cpssampleid)
                for cp in cps2:
                    plist.Add(Point3D(cp.X, cp.Y, cp.Z))
        return plist

    def EstimateClicked(self, cmd, e):
        plist = self._buildPlist()
        if plist.Count < 3:
            return

        if self.avgplane.IsChecked:
            result = self.MiniMaxCircle2D(plist)
            if result:
                _, _, r = result
                self.radiusedit.Distance = r

        elif self.bestfitplane.IsChecked:
            if plist.Count > 15000:
                self.error.Content = 'Too many points for SVD plane fit (max 15000)'
                return
            rwcloudpoints = []
            for i in range(plist.Count):
                rwcloudpoints.Add(RwPoint3D(plist[i].X, plist[i].Y, plist[i].Z))
            try:
                rwplane = RwPlane3D.FitPlaneTo3DPoints(rwcloudpoints)
            except Exception as fit_ex:
                self.error.Content = 'Could not fit a plane to the selected points: ' + str(fit_ex)
                return
            centerp = Point3D(rwplane.Point.X, rwplane.Point.Y, rwplane.Point.Z)
            plane_normal = Vector3D(rwplane.NormalVector.X, rwplane.NormalVector.Y, rwplane.NormalVector.Z)
            vx = plane_normal.Clone()
            vx.RotateAboutZ(math.pi / 2)
            vx.Horizon = 0
            rottozero = Spinor3D.ComputeRotation(Vector3D(centerp, centerp + vx), plane_normal, Vector3D(1, 0, 0), Vector3D(0, 0, 1))
            matrixtozero = Matrix4D.BuildTransformMatrix(Vector3D(centerp), Vector3D(centerp, Point3D(0, 0, 0)), rottozero, Vector3D(1, 1, 1))
            plist2d = []
            for i in range(plist.Count):
                plist2d.Add(matrixtozero.TransformPoint(plist[i]))
            result = self.MiniMaxCircle2D(plist2d)
            if result:
                _, _, r = result
                self.radiusedit.Distance = r

        elif self.userdefplane.IsChecked:
            p1 = self._coord(self.coordCtl7)
            p2 = self._coord(self.coordCtl8)
            p3 = self._coord(self.coordCtl9)
            if not (p1.Is3D and p2.Is3D and p3.Is3D):
                self.error.Content = 'Define all 3 plane vertices'
                return
            plane_normal = Plane3D(p1, p2, p3)[0].normal
            rottozero = Spinor3D.ComputeRotation(Vector3D(p1, p2), plane_normal, Vector3D(1, 0, 0), Vector3D(0, 0, 1))
            matrixtozero = Matrix4D.BuildTransformMatrix(Vector3D(p1), Vector3D(p1, Point3D(0, 0, 0)), rottozero, Vector3D(1, 1, 1))
            plist2d = []
            for i in range(plist.Count):
                plist2d.Add(matrixtozero.TransformPoint(plist[i]))
            result = self.MiniMaxCircle2D(plist2d)
            if result:
                _, _, r = result
                self.radiusedit.Distance = r

    def RunBestFit(self, wv):
        plist = self._buildPlist()

        if plist.Count < 3:
            self.success.Content = '\nSelect at least 3 points'
            return

        avg_z = sum(plist[i].Z for i in range(plist.Count)) / plist.Count

        circlecenter = None
        outradius = None
        plane_normal = None

        if self.avgplane.IsChecked:
            if self.bestfitradius.IsChecked:
                result = self.MiniMaxCircle2D(plist)
                if result:
                    cx, cy, outradius = result
                    circlecenter = Point3D(cx, cy, avg_z)

            elif self.givenradius.IsChecked:
                outradius = self.radiusedit.Distance
                cx, cy = self.MiniMaxCenter2D(plist, outradius)
                circlecenter = Point3D(cx, cy, avg_z)

        elif self.bestfitplane.IsChecked:
            if plist.Count > 15000:
                self.error.Content = 'Too many points for SVD plane fit (max 15000)'
                return

            rwcloudpoints = []
            for i in range(plist.Count):
                rwcloudpoints.Add(RwPoint3D(plist[i].X, plist[i].Y, plist[i].Z))

            try:
                rwplane = RwPlane3D.FitPlaneTo3DPoints(rwcloudpoints)
            except Exception as fit_ex:
                self.error.Content = 'Could not fit a plane to the selected points: ' + str(fit_ex)
                return
            centerp = Point3D(rwplane.Point.X, rwplane.Point.Y, rwplane.Point.Z)
            plane_normal = Vector3D(rwplane.NormalVector.X, rwplane.NormalVector.Y, rwplane.NormalVector.Z)

            vx = plane_normal.Clone()
            vx.RotateAboutZ(math.pi / 2)
            vx.Horizon = 0

            rottozero = Spinor3D.ComputeRotation(Vector3D(centerp, centerp + vx), plane_normal, Vector3D(1, 0, 0), Vector3D(0, 0, 1))
            matrixtozero = Matrix4D.BuildTransformMatrix(Vector3D(centerp), Vector3D(centerp, Point3D(0, 0, 0)), rottozero, Vector3D(1, 1, 1))
            matrixback = Matrix4D.Inverse(matrixtozero)

            plist2d = []
            for i in range(plist.Count):
                plist2d.Add(matrixtozero.TransformPoint(plist[i]))

            if self.bestfitradius.IsChecked:
                result = self.MiniMaxCircle2D(plist2d)
                if result:
                    cx, cy, outradius = result
                    circlecenter = matrixback.TransformPoint(Point3D(cx, cy, 0))

            elif self.givenradius.IsChecked:
                outradius = self.radiusedit.Distance
                cx, cy = self.MiniMaxCenter2D(plist2d, outradius)
                circlecenter = matrixback.TransformPoint(Point3D(cx, cy, 0))

        elif self.userdefplane.IsChecked:
            p1 = self._coord(self.coordCtl7)
            p2 = self._coord(self.coordCtl8)
            p3 = self._coord(self.coordCtl9)
            if not (p1.Is3D and p2.Is3D and p3.Is3D):
                self.error.Content = 'Define all 3 plane vertices'
                return

            plane_normal = Plane3D(p1, p2, p3)[0].normal
            rottozero = Spinor3D.ComputeRotation(Vector3D(p1, p2), plane_normal, Vector3D(1, 0, 0), Vector3D(0, 0, 1))
            matrixtozero = Matrix4D.BuildTransformMatrix(Vector3D(p1), Vector3D(p1, Point3D(0, 0, 0)), rottozero, Vector3D(1, 1, 1))
            matrixback = Matrix4D.Inverse(matrixtozero)

            plist2d = []
            for i in range(plist.Count):
                plist2d.Add(matrixtozero.TransformPoint(plist[i]))

            if self.bestfitradius.IsChecked:
                result = self.MiniMaxCircle2D(plist2d)
                if result:
                    cx, cy, outradius = result
                    circlecenter = matrixback.TransformPoint(Point3D(cx, cy, 0))

            elif self.givenradius.IsChecked:
                outradius = self.radiusedit.Distance
                cx, cy = self.MiniMaxCenter2D(plist2d, outradius)
                circlecenter = matrixback.TransformPoint(Point3D(cx, cy, 0))

        if circlecenter:
            if self.drawcenter.IsChecked:
                cadPoint = wv.Add(clr.GetClrType(CadPoint))
                cadPoint.Layer = self.layerpicker.SelectedSerialNumber
                cadPoint.Point0 = circlecenter

            if self.drawcircle.IsChecked:
                newcircle = wv.Add(clr.GetClrType(CadCircle))
                newcircle.CenterPoint = circlecenter
                newcircle.Radius = outradius
                newcircle.Layer = self.layerpicker.SelectedSerialNumber
                newcircle.Name = 'Radius: ' + str(outradius)
                if plane_normal is not None:
                    newcircle.Normal = plane_normal

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content = ''

        if not self.threepoint.IsChecked and not self.bestfit.IsChecked:
            self.error.Content = 'select an operation first'
            return

        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)

        wv = self.currentProject[Project.FixedSerial.WorldView]

        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                if self.threepoint.IsChecked:
                    self.RunThreePoint(wv)
                elif self.bestfit.IsChecked:
                    self.RunBestFit(wv)

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

        if self.threepoint.IsChecked:
            if self.selectpoints.IsChecked:
                Keyboard.Focus(self.objs1)
            elif self.pickpoints.IsChecked:
                Keyboard.Focus(self.coordCtl4)
        elif self.bestfit.IsChecked:
            Keyboard.Focus(self.objs2)

        self.SaveOptions()

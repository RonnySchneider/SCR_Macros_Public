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
    "layerpicker": 8, "hztol": 0.001, "vztol": 0.001,
    "highlightmode": True, "deletemode": False,
    "usesourceselected": True, "usesourcelayer": False, "sourcelayerpicker": 8,
}

def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_Check4DoubleSegments"
    cmdData.CommandName = "SCR_Check4DoubleSegments"
    cmdData.Caption = "_SCR_Check4DoubleSegments"
    cmdData.UIForm = "SCR_Check4DoubleSegments"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "Lines"
        cmdData.ShortCaption = "Check4DoubleSegments"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.13
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "check for duplicate Segments"
        cmdData.ToolTipTextFormatted = "check for duplicate Segments"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_Check4DoubleSegments(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_Check4DoubleSegments.xaml") as s:
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


        #self.lType = clr.GetClrType(IPolyseg)
        self.lType = clr.GetClrType(Linestring)
        self.poly3dType = clr.GetClrType(CadPoly3D)

        self.objs.IsEntityValidCallback = self.IsValid

        self.hztol.DistanceMin = 0.00000001
        self.hztol.DistanceMax = 0.9999999
        self.vztol.DistanceMin = 0.00000001
        self.vztol.DistanceMax = 0.9999999
        
        self.vztol.DistanceType = DistanceType.Z


        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        #self.lfp = self.lunits.Properties.Copy()
        linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation
        #self.lfp.AddSuffix = False
        self.hztollabel.Content = "horizontal tolerance [" + linearsuffix + "] <="
        self.vztollabel.Content = "vertical tolerance [" + linearsuffix + "] <="

        SCRExpanders.wire_pairs([
            (self.expander_highlightmode, self.highlightmode),
            (self.expander_deletemode, self.deletemode),
        ])
        self.SetDefaultOptions()

    def IsValid(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        if isinstance(o, self.poly3dType):
            return True
        return False

    def GetSourceObjects(self):
        if self.usesourcelayer.IsChecked:
            layer_sn = self.sourcelayerpicker.SelectedSerialNumber
            wlc = self.currentProject[Project.FixedSerial.LayerContainer]
            wl = wlc[layer_sn]
            members = wl.Members

            result = []
            count = members.Count
            ProgressBar.TBC_ProgressBar.Title = "gathering line information"
            time1 = datetime.now()
            for i, sn in enumerate(members):
                if (datetime.now() - time1).seconds > 1:
                    if ProgressBar.TBC_ProgressBar.SetProgress(i * 100 // count):
                        break
                    time1 = datetime.now()
                result.append(self.currentProject.Concordance[sn])
            ProgressBar.TBC_ProgressBar.Title = ""
            return result
        else:
            return list(self.objs)

    def ComputeElementMidpoint(self, o, B, E):
        # a straight-line average of B/E sits off the actual curve for arc/bezier/etc.
        # elements, so locate the matching chord(s) in the full chorded PolySeg between
        # B and E and interpolate at half their combined length instead
        tol = 1e-4
        polyseg = o.ComputePolySeg().ToWorld()
        s = polyseg.FirstSegment
        interval = []
        collecting = False
        while s is not None:
            if not collecting and abs(s.BeginPoint.X - B.X) <= tol and abs(s.BeginPoint.Y - B.Y) <= tol:
                collecting = True
            if collecting:
                interval.append(s)
                if abs(s.EndPoint.X - E.X) <= tol and abs(s.EndPoint.Y - E.Y) <= tol:
                    break
            s = polyseg.Next(s)

        if not interval:
            return (B.X + E.X) / 2.0, (B.Y + E.Y) / 2.0

        total_len = sum(seg.Length2D for seg in interval)
        if total_len <= 0:
            return (B.X + E.X) / 2.0, (B.Y + E.Y) / 2.0

        half = total_len / 2.0
        acc = 0.0
        for n, seg in enumerate(interval):
            L = seg.Length2D
            if acc + L >= half or n == len(interval) - 1:
                t = (half - acc) / L if L > 0 else 0.5
                p = seg.ComputePoint(t)[1]
                return p.X, p.Y
            acc += L
        return (B.X + E.X) / 2.0, (B.Y + E.Y) / 2.0

    def SetDefaultOptions(self):
        SCROptions.LoadMacroOptions(self, "SCR_Check4DoubleSegments", _OPTIONS, self.currentProject)

    def SaveOptions(self):
        SCROptions.SaveMacroOptions(self, "SCR_Check4DoubleSegments", _OPTIONS)

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def OkClicked(self, cmd, e):

        Keyboard.Focus(self.okBtn)
        self.error.Content=''
        self.success.Content = ''
        self.error3Dduplicate.Content = ''
        self.error2Dduplicate.Content = ''

        if not self.deletemode.IsChecked and not self.highlightmode.IsChecked:
            self.error.Content = 'select an operation first'
            return

        wv = self.currentProject [Project.FixedSerial.WorldView]
        wv.PauseGraphicsCache(True)

        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                if self.deletemode.IsChecked:

                    sourceobjs = self.GetSourceObjects()
                    lineserials = []
                    time1 = datetime.now()
                    i1 = 0
                    ProgressBar.TBC_ProgressBar.Title = "preparing Lookup-Tables"
                    for o in sourceobjs:

                        i1 += 1
                        if (datetime.now() - time1).seconds > 1:
                            if ProgressBar.TBC_ProgressBar.SetProgress(i1 * 100 // len(sourceobjs)):
                                break
                            time1 = datetime.now()

                        if isinstance(o, self.lType) and o.ElementCount == 2:

                            p1 = o.GetElementInWCS(0).Position
                            p2 = o.GetElementInWCS(1).Position

                            if not (p1.Is3D and p2.Is3D):
                                polyseg1_v = o.ComputeVerticalPolySeg()
                                p1.Z = polyseg1_v.FirstSegment.BeginPoint.Y
                                p2.Z = polyseg1_v.FirstSegment.EndPoint.Y

                            if p1.Is3D and p2.Is3D:
                                lineserials.Add([o.SerialNumber, p1, p2, False])

                        elif isinstance(o, self.poly3dType) and o.VertexCount == 2:

                            p1 = o[0]
                            p2 = o[1]
                            lineserials.Add([o.SerialNumber, p1, p2, False])

                    # cache tolerances as Python floats — avoids repeated .NET property access in hot loops
                    hz = self.hztol.Distance
                    vz = self.vztol.Distance

                    # sqrt(hz² + vz²) is a safe bounding radius: any pair satisfying both per-axis
                    # tolerances independently is guaranteed to lie within this sphere
                    conservative_radius = math.sqrt(hz**2 + vz**2)
                    seg_items = []
                    for i, entry in enumerate(lineserials):
                        seg_items.append((entry[1].X, entry[1].Y, entry[1].Z, i))
                        seg_items.append((entry[2].X, entry[2].Y, entry[2].Z, i))
                    seg_tree = SCROctree()
                    seg_tree.build(seg_items)

                    ProgressBar.TBC_ProgressBar.Title = "checking for duplicates"
                    time1 = datetime.now()
                    for i1 in range(0, lineserials.Count - 1):

                        if lineserials[i1][3] == True:
                            continue

                        if (datetime.now() - time1).seconds > 1:
                            if ProgressBar.TBC_ProgressBar.SetProgress(i1 * 100 // lineserials.Count):
                                break
                            time1 = datetime.now()

                        # extract i1 coordinates once as Python floats to avoid repeated .NET access inside inner loop
                        x1 = lineserials[i1][1].X; y1 = lineserials[i1][1].Y; z1 = lineserials[i1][1].Z
                        x2 = lineserials[i1][2].X; y2 = lineserials[i1][2].Y; z2 = lineserials[i1][2].Z

                        # candidates must have one endpoint near p1 AND one near p2
                        # (covers both same-direction and flipped duplicates; per-axis check below filters precisely)
                        near_p1 = set(seg_tree.find_within(x1, y1, z1, conservative_radius))
                        near_p2 = set(seg_tree.find_within(x2, y2, z2, conservative_radius))
                        closebyindex = near_p1 & near_p2
                        closebyindex.discard(i1)

                        for i2 in closebyindex:

                                if i2 <= i1 or lineserials[i2][3] == True:
                                    continue

                                # extract i2 coordinates as Python floats — avoids .NET Vector3D allocations
                                ax = lineserials[i2][1].X; ay = lineserials[i2][1].Y; az = lineserials[i2][1].Z
                                bx = lineserials[i2][2].X; by = lineserials[i2][2].Y; bz = lineserials[i2][2].Z

                                # check 2D match in original (p1↔a, p2↔b) and flipped (p1↔b, p2↔a) orientation
                                orig_2d = (math.sqrt((x1-ax)**2 + (y1-ay)**2) <= hz and
                                           math.sqrt((x2-bx)**2 + (y2-by)**2) <= hz)
                                flip_2d = (math.sqrt((x1-bx)**2 + (y1-by)**2) <= hz and
                                           math.sqrt((x2-ax)**2 + (y2-ay)**2) <= hz)

                                if orig_2d or flip_2d:
                                    if (abs(z1-az) <= vz and abs(z2-bz) <= vz) or \
                                       (abs(z1-bz) <= vz and abs(z2-az) <= vz):
                                        lineserials[i2][3] = True

                    ProgressBar.TBC_ProgressBar.Title = "deleting duplicates"
                    
                    GlobalSelection.Clear() # need to clear here - otherwise it's too slow
                    deletecount = 0
                    time1 = datetime.now()
                    for i in range(0, lineserials.Count):
                        #if (i * 100 // lineserials.Count) % 2 == 0:
                        if (datetime.now() - time1).seconds > 1:
                            if ProgressBar.TBC_ProgressBar.SetProgress(i * 100 // lineserials.Count):
                                break
                            time1 = datetime.now()

                        if lineserials[i][3] == True:
                            
                            o = self.currentProject.Concordance[lineserials[i][0]]
                            o.GetSite().Remove(o.SerialNumber)
                            deletecount += 1

                    if not self.usesourcelayer.IsChecked:
                        # only reinstate the selection when we started from the user's own
                        # selection — in layer mode there was no prior selection to restore
                        ProgressBar.TBC_ProgressBar.Title = "reinstating selection"
                        finalselection = []
                        for i in range(0, lineserials.Count):
                            if lineserials[i][3] == False: # not deleted
                                finalselection.Add(lineserials[i][0])
                        GlobalSelection.Items(self.currentProject).Set(finalselection)

                    self.success.Content += '\ndeleted ' + str(deletecount) + ' lines'

                elif self.highlightmode.IsChecked:

                    # pre-compute all segments into a flat list with coordinates as Python floats.
                    # segments come from the linestring's own elements (its actual vertices), not
                    # the chorded PolySeg — a smoothed/curved line gets subdivided into thousands of
                    # near-but-not-quite-identical PolySeg chords, which made two "duplicate" smoothed
                    # lines fail to match segment-for-segment. Comparing the real elements is stable
                    # regardless of how finely a curve happens to get chorded.
                    # all_segs entries: (line_sn, elem_type, bx, by, bz, ex, ey, ez, is3d)
                    all_segs = []
                    seg_items = []  # (x, y, z, idx) for the octree
                    for o in self.GetSourceObjects():
                        if isinstance(o, self.lType) and o.ElementCount >= 2:
                            ec = o.ElementCount
                            polyseg_v = o.ComputeVerticalPolySeg()
                            has_vertical = polyseg_v is not None  # check once per object, not per element

                            elempts = []  # (Point3D, element .NET type)
                            station = 0.0
                            prevxy = None
                            for i in range(ec):
                                elem = o.GetElementInWCS(i)
                                p = elem.Position
                                if prevxy is not None:
                                    station += math.sqrt((p.X - prevxy[0])**2 + (p.Y - prevxy[1])**2)
                                prevxy = (p.X, p.Y)
                                if not p.Is3D and has_vertical:
                                    p.Z = polyseg_v.ComputeVerticalSlopeAndGrade(station)[1]
                                elempts.append((p, type(elem)))

                            for i in range(ec - 1):
                                B = elempts[i][0]
                                E, etype = elempts[i + 1]
                                idx = len(all_segs)
                                bx, by = B.X, B.Y
                                ex, ey = E.X, E.Y
                                # undefined elevation comes back as NaN, which poisons every distance
                                # comparison it touches (NaN <= anything is always False) — sanitize to
                                # 0.0 so the octree and Z-tolerance checks below stay well-defined;
                                # is3d (from the real, un-sanitized flag) gates whether Z is trusted at all
                                is3d = B.Is3D and E.Is3D
                                bz = B.Z if B.Is3D else 0.0
                                ez = E.Z if E.Is3D else 0.0
                                all_segs.append((o.SerialNumber, etype, bx, by, bz, ex, ey, ez, is3d, o, B, E))
                                seg_items.append((bx, by, bz, idx))
                                seg_items.append((ex, ey, ez, idx))

                    # cache tolerances as Python floats — avoids repeated .NET property access in hot loops
                    hz = self.hztol.Distance
                    vz = self.vztol.Distance
                    seg_tree = SCROctree()
                    seg_tree.build(seg_items)

                    time1 = datetime.now()
                    for i1 in range(len(all_segs)):

                        if (datetime.now() - time1).seconds > 1:
                            ProgressBar.TBC_ProgressBar.Title = "checking segment " + str(i1) + "/" + str(len(all_segs))
                            if ProgressBar.TBC_ProgressBar.SetProgress(i1 * 100 // len(all_segs)):
                                break
                            time1 = datetime.now()

                        sn1, etype1, s1_bx, s1_by, s1_bz, s1_ex, s1_ey, s1_ez, is3d1, o1, B1, E1 = all_segs[i1]

                        # 2D-only candidate search: a 2D duplicate can have wildly different (or
                        # undefined/sanitized-to-0) elevation on either side, so filtering candidates
                        # by a 3D radius would wrongly exclude those before the dup2d/dup3d checks below
                        # ever get a chance to run
                        near_B = set(seg_tree.find_within_2d(s1_bx, s1_by, hz))
                        near_E = set(seg_tree.find_within_2d(s1_ex, s1_ey, hz))
                        candidates = (near_B & near_E) - {i1}

                        for i2 in candidates:
                            if i2 <= i1:
                                continue
                            sn2, etype2, s2_bx, s2_by, s2_bz, s2_ex, s2_ey, s2_ez, is3d2, o2, B2, E2 = all_segs[i2]
                            if sn1 == sn2 or etype1 != etype2:
                                continue

                            # inline 2D distance — avoids creating 4 .NET Vector3D objects per candidate pair
                            orig_2d = (math.sqrt((s1_bx-s2_bx)**2 + (s1_by-s2_by)**2) <= hz and
                                       math.sqrt((s1_ex-s2_ex)**2 + (s1_ey-s2_ey)**2) <= hz)
                            flip_2d = (math.sqrt((s1_bx-s2_ex)**2 + (s1_by-s2_ey)**2) <= hz and
                                       math.sqrt((s1_ex-s2_bx)**2 + (s1_ey-s2_by)**2) <= hz)

                            dup2d = orig_2d or flip_2d
                            dup3d = False

                            if dup2d and is3d1 and is3d2:
                                # Z check only meaningful when both segments actually carry an elevation;
                                # otherwise leave the match as a 2D duplicate rather than falsely promoting
                                # it to 3D by comparing undefined (default 0) elevations
                                orig_z = (abs(s1_bz-s2_bz) <= vz and abs(s1_ez-s2_ez) <= vz)
                                flip_z = (abs(s1_bz-s2_ez) <= vz and abs(s1_ez-s2_bz) <= vz)
                                if orig_z or flip_z:
                                    dup3d = True
                                    dup2d = False

                            if dup3d or dup2d:
                                mx, my = self.ComputeElementMidpoint(o1, B1, E1)
                                circle_radius = math.sqrt((s1_ex-s1_bx)**2 + (s1_ey-s1_by)**2) / 4.0
                                newcircle = wv.Add(clr.GetClrType(CadCircle))
                                newcircle.Layer = self.layerpicker.SelectedSerialNumber
                                newcircle.CenterPoint = Point3D(mx, my, (s1_bz+s1_ez)/2)
                                newcircle.Radius = circle_radius
                                if dup3d:
                                    newcircle.Color = Color.Red
                                    self.error3Dduplicate.Content = '3D duplicate'
                                else:
                                    newcircle.Color = Color.Orange
                                    self.error2Dduplicate.Content = '2D duplicate'

                failGuard.Commit()
                self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
                UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
        
        except Exception as e:
            tt = sys.exc_info()
            exc_type, exc_obj, exc_tb = sys.exc_info()
            # EndMark MUST be set no matter what
            # otherwise TBC won't work anymore and needs to be restarted
            self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
            UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
            self.error.Content += '\nan Error occurred - Result probably incomplete\n' + str(exc_type) + '\n' + str(exc_obj) + '\nLine ' + str(exc_tb.tb_lineno)

        self.SaveOptions()           
        Keyboard.Focus(self.okBtn)        
        ProgressBar.TBC_ProgressBar.Title = ""
        
        wv.PauseGraphicsCache(False)
    

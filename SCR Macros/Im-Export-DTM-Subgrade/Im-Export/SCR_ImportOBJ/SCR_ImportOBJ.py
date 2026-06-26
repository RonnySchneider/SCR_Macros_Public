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
    "openfilename":  os.path.expanduser('~\\Downloads'),
    "layerpicker":   8,
    "drawlines":     False,
    "drawifc":       False,
    "createsurface": False,
}

def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_ImportOBJ"
    cmdData.CommandName = "SCR_ImportOBJ"
    cmdData.Caption = "_SCR_ImportOBJ"
    cmdData.UIForm = "SCR_ImportOBJ"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "Import"
        cmdData.ShortCaption = "Import OBJ"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.09
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""

        cmdData.ToolTipTitle = "Import OBJ file"
        cmdData.ToolTipTextFormatted = "Import OBJ file"

    except:
        pass

    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_ImportOBJ(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_ImportOBJ.xaml") as s:
            wpf.LoadComponent (self, s)
        self.currentProject = currentProject
        self.macroFileFolder = macroFileFolder
        if self.openfilename.Text=='': self.openfilename.Text = macroFileFolder

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

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass


    def SetDefaultOptions(self):
        SCROptions.LoadMacroOptions(self, "SCR_ImportOBJ", _OPTIONS, self.currentProject)

    def SaveOptions(self):
        SCROptions.SaveMacroOptions(self, "SCR_ImportOBJ", _OPTIONS)


    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def Dispose(self, cmd, disposing):
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)

    def openbutton_Click(self, sender, e):
        dialog = OpenFileDialog()
        dialog.InitialDirectory = self.openfilename.Text
        dialog.Filter=("OBJ|*.obj")

        tt=dialog.ShowDialog()
        if tt==DialogResult.OK:
            self.openfilename.Text = dialog.FileName

    def drawoverlay(self):

        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)
        self.overlayBag = OverlayBag(self.ViewOverlay) # create a new/empty overlaybag

        layerobject = self.currentProject.Concordance.Lookup(8) # get the layer as object

        for sn in layerobject.Members:
            l = self.currentProject.Concordance.Lookup(sn)
            polyseg = l.ComputePolySeg()
            self.overlayBag.AddPolyline(polyseg.ToPoint3DArray(), Color.Green.ToArgb(), 1)

        # if it only needs to be visible in all Planview then remove the Hoops3DViewGUID
        array = Array[Guid]([DisplayWindow.Hoops3DViewGUID, DisplayWindow.HoopsPlanViewGUID])
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.AddOverlayGeometry(array, self.overlayBag)

        return

    def parse_mtl_file(self, mtl_path):
        """Parse a .mtl file and return dict: material_name -> ARGB int color."""
        materials = {}
        current = None
        r, g, b, opacity = 1.0, 1.0, 1.0, 1.0

        if not File.Exists(mtl_path):
            return materials

        with open(mtl_path, 'r') as f:
            for line in f:
                parts = line.strip().split()
                if not parts or parts[0].startswith('#'):
                    continue
                if parts[0] == 'newmtl':
                    if current is not None:
                        col = Color.FromArgb(int(opacity * 255), int(r * 255), int(g * 255), int(b * 255))
                        materials[current] = col.ToArgb()
                    current = ' '.join(parts[1:])
                    r, g, b, opacity = 1.0, 1.0, 1.0, 1.0
                elif parts[0] == 'Kd' and len(parts) >= 4:
                    r, g, b = float(parts[1]), float(parts[2]), float(parts[3])
                elif parts[0] == 'd' and len(parts) >= 2:
                    opacity = float(parts[1])
                elif parts[0] == 'Tr' and len(parts) >= 2:
                    # Tr is inverse of d
                    opacity = 1.0 - float(parts[1])

        if current is not None:
            col = Color.FromArgb(int(opacity * 255), int(r * 255), int(g * 255), int(b * 255))
            materials[current] = col.ToArgb()

        return materials

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''

        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)

        wv = self.currentProject [Project.FixedSerial.WorldView]
        wv.PauseGraphicsCache(True)

        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                if self.developoverlay.IsChecked:

                    self.drawoverlay()

                else:

                    ProgressBar.TBC_ProgressBar.Title = "reading the file"

                    # set up the single IFC container for the whole file before parsing
                    ifcprojectname = os.path.basename(self.openfilename.Text)
                    layer_name = os.path.splitext(ifcprojectname)[0]
                    output_layer = Layer.FindOrCreateLayer(self.currentProject, layer_name)
                    bimprojectEntity = None
                    shellMeshDataColl = None
                    if self.drawifc.IsChecked:
                        bimEntityColl = BIMEntityCollection.ProvideEntityCollection(self.currentProject, True)
                        shellMeshDataColl = ShellMeshDataCollection.ProvideShellMeshDataCollection(self.currentProject, True)
                        for e in bimEntityColl:
                            if e.Description == ifcprojectname:
                                bimprojectEntity = e
                                break
                        if bimprojectEntity is None:
                            bimprojectEntity = bimEntityColl.Add(clr.GetClrType(BIMEntity))
                            bimprojectEntity.EntityType = "IFCPROJECT"
                            bimprojectEntity.Description = ifcprojectname
                            bimprojectEntity.BIMGuid = Guid.NewGuid()
                            bimprojectEntity.Mode = DisplayMode(1 + 2 + 64 + 128 + 512 + 4096)
                            bimprojectEntity.Layer = output_layer.SerialNumber

                    materials_colors = {}       # material_name -> ARGB int
                    objectname = ''
                    current_material = 'default'
                    vertices = []               # global list of Point3D (OBJ indices are file-global)
                    faces_by_material = {'default': []}  # material_name -> list of (v1, v2, v3)
                    linetuples = []             # list of (v1, v2) edge pairs

                    with open(self.openfilename.Text, 'r') as csvfile:
                        reader = csv.reader(csvfile, delimiter=' ', quotechar='|')
                        ProgressBar.TBC_ProgressBar.Title = "preparing data tables"

                        for row in reader:
                            if not row or row[0] == '' or row[0].startswith('#'):
                                continue

                            if row[0] == 'mtllib' and len(row) > 1:
                                mtl_name = ' '.join(row[1:])
                                mtl_path = os.path.join(os.path.dirname(self.openfilename.Text), mtl_name)
                                materials_colors = self.parse_mtl_file(mtl_path)

                            elif row[0] in ('g', 'o'):
                                # flush current group before starting the next one
                                if vertices and any(faces_by_material.values()):
                                    self.startdrawing(vertices, faces_by_material, linetuples, objectname, materials_colors, bimprojectEntity, shellMeshDataColl, output_layer.SerialNumber)
                                objectname = ' '.join(row[1:]) if len(row) > 1 else ''
                                # vertices are NOT reset — OBJ indices are file-global
                                faces_by_material = {'default': []}
                                current_material = 'default'
                                linetuples = []

                            elif row[0] == 'usemtl':
                                current_material = ' '.join(row[1:]) if len(row) > 1 else 'default'
                                if current_material not in faces_by_material:
                                    faces_by_material[current_material] = []

                            elif row[0] == 'v':
                                # OBJ is Y-up; TBC is Z-up: X=x, Y=-z, Z=y
                                vertices.append(Point3D(float(row[1]), -float(row[3]), float(row[2])))

                            elif row[0] == 'f':
                                # f v1[/vt1[/vn1]] v2[/vt2[/vn2]] ... — only vertex index matters
                                # OBJ indices are 1-based and file-global
                                v_indices = []
                                for i in range(1, len(row)):
                                    s = row[i].split('/')[0] if '/' in row[i] else row[i]
                                    if s:
                                        v_indices.append(int(s) - 1)

                                # fan-triangulate; handles triangles and n-gons
                                for i in range(1, len(v_indices) - 1):
                                    tri = (v_indices[0], v_indices[i], v_indices[i + 1])
                                    faces_by_material[current_material].append(tri)
                                    linetuples.append((v_indices[0], v_indices[i]))
                                    linetuples.append((v_indices[i], v_indices[i + 1]))
                                    linetuples.append((v_indices[i + 1], v_indices[0]))

                        # flush the final group (no trailing empty line needed)
                        if vertices and any(faces_by_material.values()):
                            self.startdrawing(vertices, faces_by_material, linetuples, objectname, materials_colors, bimprojectEntity, shellMeshDataColl, output_layer.SerialNumber)

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

        ProgressBar.TBC_ProgressBar.Title = ""

        wv.PauseGraphicsCache(False)
        self.SaveOptions()

    def startdrawing(self, vertices, faces_by_material, linetuples, objectname, materials_colors, bimprojectEntity, shellMeshDataColl, output_layer_sn):

        wv = self.currentProject [Project.FixedSerial.WorldView]
        all_faces = [f for faces in faces_by_material.values() for f in faces]

        if self.drawlines.IsChecked or self.createsurface.IsChecked:
            ProgressBar.TBC_ProgressBar.Title = "removing duplicate lines"
            if linetuples:
                # build octree over both endpoints of every edge so that
                # edges identical in 3D coords (even with different vertex indices)
                # are removed — same pattern as SCR_Check4DoubleSegments
                edge_items = []
                for idx, (ei, ej) in enumerate(linetuples):
                    p1, p2 = vertices[ei], vertices[ej]
                    edge_items.append((p1.X, p1.Y, p1.Z, idx))
                    edge_items.append((p2.X, p2.Y, p2.Z, idx))
                edge_tree = SCROctree()
                edge_tree.build(edge_items)
                tol = 1e-9
                seen = set()
                linetuplesclean = []
                for idx, (ei, ej) in enumerate(linetuples):
                    if idx in seen:
                        continue
                    p1, p2 = vertices[ei], vertices[ej]
                    near_p1 = set(edge_tree.find_within(p1.X, p1.Y, p1.Z, tol))
                    near_p2 = set(edge_tree.find_within(p2.X, p2.Y, p2.Z, tol))
                    for dup in (near_p1 & near_p2) - {idx}:
                        if dup > idx:
                            seen.add(dup)
                    linetuplesclean.append((ei, ej))
            else:
                linetuplesclean = []

        if self.drawlines.IsChecked:
            timerresults = []

            # that's just for the progressbar
            barvalueinc = 0.05
            barvalueinccount = len(linetuplesclean) * barvalueinc
            barvalueinci = 0

            for i in range(len(linetuplesclean)):
                barvalueinci += 1
                if barvalueinci >= barvalueinccount:
                    barvalueinci = 0
                    ProgressBar.TBC_ProgressBar.Title = "drawing line: " + str(i) + "/" + str(len(linetuplesclean))
                    if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(i * 100 / len(linetuplesclean))):
                        break   # function returns true if user pressed cancel

                drawtuple = linetuplesclean[i]

                start_t = timer()

                polyseg = PolySeg.PolySeg()
                polyseg.Add(vertices[drawtuple[0]])
                polyseg.Add(vertices[drawtuple[1]])
                l = wv.Add(clr.GetClrType(Linestring))
                l.Layer = output_layer_sn
                l.Append(polyseg, None, False, False)
                l.Name = objectname

                end_t=timer()
                timerresults.append(timedelta(seconds=end_t-start_t))

            filename = os.path.expanduser('~/Downloads/timerresult.csv')
            if File.Exists(filename):
                File.Delete(filename)
            with open(filename, 'w') as f:
                for i in range(len(timerresults)):
                    outputline = str(i + 1) + "," + str(timerresults[i]) + "\n"
                    f.write(outputline)
                f.close()


        if self.drawifc.IsChecked:
            ProgressBar.TBC_ProgressBar.Title = "creating the IFC objects"

            ifcvertices = Array[Point3D]([Point3D()] * len(vertices))
            for i in range(len(vertices)):
                ifcvertices[i] = vertices[i]

            ifcnormals = Array[Point3D]([Point3D()] * 0)

            known_colors = list(materials_colors.values())
            if known_colors:
                avg_r = sum(int(Color.FromArgb(c).R) for c in known_colors) // len(known_colors)
                avg_g = sum(int(Color.FromArgb(c).G) for c in known_colors) // len(known_colors)
                avg_b = sum(int(Color.FromArgb(c).B) for c in known_colors) // len(known_colors)
                fallback_color = Color.FromArgb(255, avg_r, avg_g, avg_b).ToArgb()
            else:
                fallback_color = Color.White.ToArgb()

            for mat_name, faces in faces_by_material.items():
                if not faces:
                    continue
                color = materials_colors.get(mat_name, fallback_color)

                bimobjectEntity = bimprojectEntity.Add(clr.GetClrType(BIMEntity))
                bimobjectEntity.EntityType = "IFCBUILDINGELEMENTPROXY"
                bimobjectEntity.Description = mat_name if mat_name != 'default' else objectname
                bimobjectEntity.BIMGuid = Guid.NewGuid()
                bimobjectEntity.Mode = DisplayMode(1 + 2 + 64 + 128 + 512 + 4096)
                bimobjectEntity.Layer = bimprojectEntity.Layer

                ifcfacelist = Array[Int32]([Int32()] * (len(faces) * 4))
                for idx, tri in enumerate(faces):
                    ifcfacelist[idx * 4 + 0] = 3
                    ifcfacelist[idx * 4 + 1] = tri[0]
                    ifcfacelist[idx * 4 + 2] = tri[1]
                    ifcfacelist[idx * 4 + 3] = tri[2]
                shellmeshdata = shellMeshDataColl.AddShellMeshData(self.currentProject)
                shellmeshdata.CreateShellMeshData(ifcvertices, ifcfacelist, ifcnormals)
                shellmeshdata.SetVolumeCalculationShell(ifcvertices, ifcfacelist)
                meshInstance = bimobjectEntity.Add(clr.GetClrType(ShellMeshInstance))
                meshInstance.CreateShell(0, shellmeshdata.SerialNumber, color, Matrix4D())


        if self.createsurface.IsChecked:
            ProgressBar.TBC_ProgressBar.Title = "creating new surface"
            newSurface = wv.Add(clr.GetClrType(Model3D))
            newSurface.Name = Model3D.GetUniqueName(os.path.basename(self.openfilename.Text) + ' - ' + objectname, None, wv) #make sure name is unique
            # SmoothShading | ElevationContours | ShowBackfaces | FilledTriangles | ElevationContours_2D | FilledTriangles_2D
            newSurface.Mode = DisplayMode(1 + 64 + 128 + 512 + 1024 + 65536)
            builder = newSurface.GetGemBatchBuilder()
            for i in range(len(vertices)):
                builder.AddVertex(vertices[i])

            for drawtuple in linetuplesclean:
                builder.AddBreakline(Byte(DTMSharpness.eSoft), drawtuple[0], drawtuple[1])

            builder.Construction()

            # now flag all edge triangles that don't have breakline edges
            nTri = builder.NumberOfTriangles
            gemmap = GemMaterialMap()
            for t in range(nTri):
                for side in range(3):
                    isOuter = builder.GetTriangleOuterSide(t, side)
                    if not isOuter:
                        continue
                    (bl, external, sharp) = builder.GetTriangleBreakline(t, side)
                    if bl:
                        continue
                    # we have edge triangle where edge is not breakline
                    # put null material on triangle
                    builder.AttachMaterial(0, gemmap, t)

            builder.Commit()


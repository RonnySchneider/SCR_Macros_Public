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
    "radius":          1.0,
    "verticaloffset":  0.0,
    "segments":        24,
    "layerpicker":     8,
    "createifc":       True,
    "ifccolorpicker":  Color.Orange,
    "createstockpile": False,
    "createbasin":     False,
}

IFCPROJECTNAME = "SCR_SphereClashCheck"
UNNAMEDBASE = "Sphere 1"

def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_SphereClashCheck"
    cmdData.CommandName = "SCR_SphereClashCheck"
    cmdData.Caption = "_SCR_SphereClashCheck"
    cmdData.UIForm = "SCR_SphereClashCheck"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "IFC"
        cmdData.ShortCaption = "Sphere Clash Check"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.005
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""

        cmdData.ToolTipTitle = "Sphere Clash Check"
        cmdData.ToolTipTextFormatted = "creates a sphere around each selected point - as IFC object and/or as 2 half-sphere DTMs (stockpile on top, basin below)"

    except:
        pass

    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_SphereClashCheck(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_SphereClashCheck.xaml") as s:
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

        self.cadpointType = clr.GetClrType(CadPoint)
        self.coordpointType = clr.GetClrType(CoordPoint)
        self.objs.IsEntityValidCallback = self.IsValidPoint

        self.radius.FormatProperty.AddSuffix = ControlBoolean(1)
        self.verticaloffset.FormatProperty.AddSuffix = ControlBoolean(1)

        try:
            self.SetDefaultOptions()
        except:
            pass

    def SetDefaultOptions(self):
        SCROptions.LoadMacroOptions(self, "SCR_SphereClashCheck", _OPTIONS, self.currentProject)

    def SaveOptions(self):
        SCROptions.SaveMacroOptions(self, "SCR_SphereClashCheck", _OPTIONS)

    def IsValidPoint(self, serial):
        o = self.currentProject.Concordance[serial]
        if isinstance(o, self.cadpointType):
            return True
        if isinstance(o, self.coordpointType):
            return True
        return False

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def Dispose(self, cmd, disposing):
        pass

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content = ''
        self.success.Content = ''

        createifc = self.createifc.IsChecked
        createstockpile = self.createstockpile.IsChecked
        createbasin = self.createbasin.IsChecked

        if not createifc and not createstockpile and not createbasin:
            self.error.Content = 'select at least one output'
            return

        radius = self.radius.Distance
        if math.isnan(radius) or radius <= 0.0:
            self.error.Content = 'enter a valid radius'
            return

        verticaloffset = self.verticaloffset.Distance
        if math.isnan(verticaloffset):
            verticaloffset = 0.0

        # segments around the equator - must be a multiple of 4 so that the
        # latitude rings step evenly from the equator up to the poles
        segments = int(round(self.segments.Value / 4.0)) * 4
        if segments < 8:
            segments = 8

        # collect centre points and names first
        centres = []
        skipped2d = 0
        for o in self.objs.SelectedMembers(self.currentProject):
            if isinstance(o, self.coordpointType):
                p = o.Position
                name = o.Name
            elif isinstance(o, self.cadpointType):
                p = o.Point0
                name = ''
            else:
                continue
            if not p.Is3D:
                skipped2d += 1
                continue
            centres.Add([Point3D(p.X, p.Y, p.Z + verticaloffset), name])

        if centres.Count == 0:
            self.error.Content = 'no 3D points selected'
            return

        wv = self.currentProject [Project.FixedSerial.WorldView]
        layer = self.layerpicker.SelectedSerialNumber
        ifccolor = self.ifccolorpicker.SelectedColor.ToArgb()

        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)

        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                # all spheres are identical, so we build the mesh only once around 0,0,0
                vertices, upperfaces, lowerfaces = self.buildsphere(Point3D(0, 0, 0), radius, segments)

                if createifc:
                    bimprojectEntity, shellMeshDataColl = self.getifcproject(layer)
                    ifcnames = set(e.Description for e in bimprojectEntity if isinstance(e, BIMEntity))
                    # one ShellMeshData shared by all spheres, each ShellMeshInstance only moves it to its point
                    shellmeshdata = self.createshellmeshdata(shellMeshDataColl, vertices, upperfaces + lowerfaces)
                else:
                    ifcnames = set()

                ProgressBar.TBC_ProgressBar.Title = "creating spheres"
                last_t = timer()
                done = 0
                for i, (centre, name) in enumerate(centres):
                    if timer() - last_t > 0.5:
                        if ProgressBar.TBC_ProgressBar.SetProgress(i * 100 // centres.Count):
                            break   # user pressed cancel
                        last_t = timer()

                    if not name:
                        name = self.unusedname(wv, ifcnames)

                    if createifc:
                        self.createifcobject(bimprojectEntity, shellmeshdata, name, centre, ifccolor, layer)
                        ifcnames.add(name)

                    if createstockpile or createbasin:
                        # DTMs need real coordinates
                        worldvertices = [Point3D(v.X + centre.X, v.Y + centre.Y, v.Z + centre.Z) for v in vertices]

                    if createstockpile:
                        self.createdtm(wv, name + ' - stockpile', worldvertices, upperfaces)

                    if createbasin:
                        self.createdtm(wv, name + ' - basin', worldvertices, lowerfaces)

                    done += 1

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

        if skipped2d > 0:
            self.error.Content += '\n' + str(skipped2d) + ' point(s) without elevation skipped'
        self.success.Content += '\n' + str(done) + ' sphere(s) created'

        self.SaveOptions()

    def buildsphere(self, centre, radius, segments):
        # latitude/longitude sphere
        # the equator is a full vertex ring shared by both halves, so stockpile and basin meet exactly
        # vertex order: equator ring, upper rings, top pole, lower rings, bottom pole
        # returns vertices plus upper/lower triangle lists, all oriented outwards (CCW seen from outside)
        latsteps = segments // 4
        latangle = (math.pi / 2.0) / latsteps
        cosa = [math.cos(2.0 * math.pi * j / segments) for j in range(segments)]
        sina = [math.sin(2.0 * math.pi * j / segments) for j in range(segments)]
        cx, cy, cz = centre.X, centre.Y, centre.Z

        vertices = []

        def addring(lat):
            start = len(vertices)
            ringradius = radius * math.cos(lat)
            z = cz + radius * math.sin(lat)
            for j in range(segments):
                vertices.Add(Point3D(cx + ringradius * cosa[j], cy + ringradius * sina[j], z))
            return start

        equator = addring(0.0)
        upperrings = [equator] + [addring(k * latangle) for k in range(1, latsteps)]
        toppole = len(vertices)
        vertices.Add(Point3D(cx, cy, cz + radius))
        lowerrings = [equator] + [addring(-k * latangle) for k in range(1, latsteps)]
        bottompole = len(vertices)
        vertices.Add(Point3D(cx, cy, cz - radius))

        upperfaces = []
        lowerfaces = []
        for j in range(segments):
            j1 = (j + 1) % segments
            for k in range(latsteps - 1):
                # a = ring closer to the equator, b = ring closer to the pole
                a, b = upperrings[k], upperrings[k + 1]
                upperfaces.Add((a + j, a + j1, b + j1))
                upperfaces.Add((a + j, b + j1, b + j))
                a, b = lowerrings[k], lowerrings[k + 1]
                lowerfaces.Add((a + j, b + j, b + j1))
                lowerfaces.Add((a + j, b + j1, a + j1))
            a = upperrings[-1]
            upperfaces.Add((a + j, a + j1, toppole))
            a = lowerrings[-1]
            lowerfaces.Add((a + j, bottompole, a + j1))

        return vertices, upperfaces, lowerfaces

    def unusedname(self, wv, ifcnames):
        # GetUniqueName returns the name unchanged if no surface uses it yet
        name = UNNAMEDBASE
        while name in ifcnames or \
              Model3D.GetUniqueName(name + ' - stockpile', None, wv) != name + ' - stockpile' or \
              Model3D.GetUniqueName(name + ' - basin', None, wv) != name + ' - basin':
            name = PointHelper.Helpers.Increment(name, None, True)
        return name

    def getifcproject(self, layer):
        # all spheres go as separate objects into one IFC project
        bimEntityColl = BIMEntityCollection.ProvideEntityCollection(self.currentProject, True)
        shellMeshDataColl = ShellMeshDataCollection.ProvideShellMeshDataCollection(self.currentProject, True)
        for e in bimEntityColl:
            if e.Description == IFCPROJECTNAME:
                return e, shellMeshDataColl

        bimprojectEntity = bimEntityColl.Add(clr.GetClrType(BIMEntity))
        bimprojectEntity.EntityType = "IFCPROJECT"
        bimprojectEntity.Description = IFCPROJECTNAME # must be set, otherwise export fails
        bimprojectEntity.BIMGuid = Guid.NewGuid()
        bimprojectEntity.Mode = DisplayMode(1 + 2 + 64 + 128 + 512 + 4096)
        bimprojectEntity.Layer = layer
        return bimprojectEntity, shellMeshDataColl

    def createshellmeshdata(self, shellMeshDataColl, vertices, faces):
        ifcvertices = Array[Point3D](vertices)
        ifcfacelist = Array[Int32]([Int32()] * (len(faces) * 4))
        for idx, tri in enumerate(faces):
            ifcfacelist[idx * 4 + 0] = 3
            ifcfacelist[idx * 4 + 1] = tri[0]
            ifcfacelist[idx * 4 + 2] = tri[1]
            ifcfacelist[idx * 4 + 3] = tri[2]
        ifcnormals = Array[Point3D]([Point3D()] * 0)

        shellmeshdata = shellMeshDataColl.AddShellMeshData(self.currentProject)
        shellmeshdata.CreateShellMeshData(ifcvertices, ifcfacelist, ifcnormals)
        shellmeshdata.SetVolumeCalculationShell(ifcvertices, ifcfacelist)
        return shellmeshdata

    def createifcobject(self, bimprojectEntity, shellmeshdata, name, centre, color, layer):
        bimobjectEntity = bimprojectEntity.Add(clr.GetClrType(BIMEntity))
        bimobjectEntity.EntityType = "IFCBUILDINGELEMENTPROXY"
        bimobjectEntity.Description = name
        bimobjectEntity.BIMGuid = Guid.NewGuid()
        bimobjectEntity.Mode = DisplayMode(1 + 2 + 64 + 128 + 512 + 4096)
        # MUST set layer, otherwise the properties manager will falsly show layer "0"
        # but the object is invisible until manually changing the layer
        bimobjectEntity.Layer = layer

        transform = Matrix4D.BuildTransformMatrix(Point3D(0, 0, 0), centre, 0, 1.0, 1.0, 1.0)
        meshInstance = bimobjectEntity.Add(clr.GetClrType(ShellMeshInstance))
        meshInstance.CreateShell(0, shellmeshdata.SerialNumber, color, transform)

    def createdtm(self, wv, name, vertices, faces):
        # a half sphere has no overhangs, so in plan view the mesh is a valid triangulation
        # every mesh edge is added as breakline, that forces TBC to triangulate exactly like our mesh
        localindex = {}
        for tri in faces:
            for v in tri:
                if v not in localindex:
                    localindex[v] = len(localindex)

        edges = set()
        for tri in faces:
            for s in range(3):
                v0, v1 = localindex[tri[s]], localindex[tri[(s + 1) % 3]]
                edges.add((v0, v1) if v0 < v1 else (v1, v0))

        newSurface = wv.Add(clr.GetClrType(Model3D))
        newSurface.Name = Model3D.GetUniqueName(name, None, wv) #make sure name is unique
        # SmoothShading | ElevationContours | ShowBackfaces | FilledTriangles | ElevationContours_2D | FilledTriangles_2D
        newSurface.Mode = DisplayMode(1 + 64 + 128 + 512 + 1024 + 65536)

        builder = newSurface.GetGemBatchBuilder()
        for v, i in sorted(localindex.items(), key=lambda item: item[1]):
            builder.AddVertex(vertices[v])
        for v0, v1 in edges:
            builder.AddBreakline(Byte(DTMSharpness.eSoft), v0, v1)
        builder.Construction()
        builder.Commit()

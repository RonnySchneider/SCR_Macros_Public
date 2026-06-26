from System.Collections.Generic import List, IEnumerable
exec(open(r"C:\ProgramData\Trimble\MacroCommands3\SCR Macros\SCR_Imports.py").read())

_OPTIONS = {
    "slopeCB": False,
    "LineLayers": 8,
    "createlinestring": True,
    "createcadpoly3d": False,
}

def Setup(cmdData, macroFileFolder):
    cmdData.Key = "ExplodeSurfaceToLines"
    cmdData.CommandName = "ExplodeSurfaceToLines"
    cmdData.Caption = "Explode Surface To Lines"
    cmdData.UIForm = "ExplodeSurfaceToLines"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    #cmdData.HelpTopic = "25744"

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "DTM"
        cmdData.ShortCaption = "Explode DTM 2 Lines"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.07
        cmdData.MacroAuthor = "Trimble"

        cmdData.ToolTipTitle = "Explode Surface into lines"
        cmdData.ToolTipTextFormatted = "Explode a Surface into lines only"
    except:
        pass
    try:
        b = Bitmap(macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass

class ExplodeSurfaceToLines(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader(macroFileFolder + r"\ExplodeSurfaceToLines.xaml") as s:
            wpf.LoadComponent(self, s)
        self.currentProject = currentProject

    def HelpClicked(self, cmd, e):
        webbrowser.open("C:\ProgramData\Trimble\MacroCommands3\SCR Macros\MacroHelp\MacroHelp.htm#" + type(self).__name__)

    def OnLoad(self, cmd, buttons, event):
        self.okBtn = buttons[0]
        buttons[2].Content = "Help"
        buttons[2].Visibility = Visibility.Visible
        buttons[2].Click += self.HelpClicked
        self.Caption = cmd.Command.Caption
        types = Array[Type](SurfaceTypeLists.AllWithCutFillMap)
        self.surface.FilterByEntityTypes = types
        self.surface.AllowNone = False
        try:
            self.SetDefaultOptions(None, None)
        except:
            pass

    def SetDefaultOptions(self, sender, e):
        SCROptions.LoadMacroOptions(self, "ExplodeSurfaceToLines", _OPTIONS, self.currentProject)
        self.slopeLimitCtl.Angle = OptionsManager.GetDouble("ExplodeSurfaceToLines.slopelimit", 0.0)
        self.surface.SelectedSerial = int(OptionsManager.GetDouble("ExplodeSurfaceToLines.surface", 0))

    def SaveOptions(self):
        SCROptions.SaveMacroOptions(self, "ExplodeSurfaceToLines", _OPTIONS)
        OptionsManager.SetValue("ExplodeSurfaceToLines.slopelimit", self.slopeLimitCtl.Angle)
        OptionsManager.SetValue("ExplodeSurfaceToLines.surface", float(self.surface.SelectedSerial))

    def CancelClicked(self, sender, args):
        sender.CloseUICommand()

    def DoTriangleEdge(self, triangle, side, pnts, surface, wv, layerSerial):
        # check if triangle edge is outside of surface or next to hole
        isOuter = surface.GetTriangleOuterSide(triangle, side)
        if not isOuter:
            (b, tri, ic) = surface.Gem.GetTriangleAdjacent(triangle, side)
            if not surface.IsTriangleMaterialPresent(tri):
                isOuter = True
        iVertexA = surface.GetTriangleIVertex(triangle, side)
        nextSide = side + 1
        if nextSide == 3:
            nextSide = 0
        iVertexB = surface.GetTriangleIVertex(triangle, nextSide)
        if isOuter:
            self.CreateLine(wv, layerSerial, pnts[iVertexA], pnts[iVertexB])
        elif iVertexA < iVertexB:
            self.CreateLine(wv, layerSerial, pnts[iVertexA], pnts[iVertexB])

    def TriangleSlope(self, surface, iTri):
        (v,p0) = surface.GetTriangleVertex(iTri, 0)
        (v,p1) = surface.GetTriangleVertex(iTri, 1)
        (v,p2) = surface.GetTriangleVertex(iTri, 2)
        plane = Plane3D(p0, p1, p2)
        v = plane[0].normal
        v.Normalize()
        return v

    def DoTriangleSlopeEdge(self, triangle, triSlope, side, surface, wv, layerSerial, slopeBreak):
        # check if triangle edge is outside of surface or next to hole
        isOuter = surface.GetTriangleOuterSide(triangle, side)
        if isOuter:
            return
        (b, tri, ic) = surface.Gem.GetTriangleAdjacent(triangle, side)
        iVertexA = surface.GetTriangleIVertex(triangle, side)
        nextSide = side + 1
        if nextSide == 3:
            nextSide = 0
        iVertexB = surface.GetTriangleIVertex(triangle, nextSide)
        if iVertexA > iVertexB:
            return
        triAdjSlope = self.TriangleSlope(surface, tri)
        (dot, v1, v2) = Vector3D.DotProduct(triSlope, triAdjSlope)
        # check for round-off
        if dot > 1.0:
            dot = 1.0
        if dot < -1.0:
            dot = -1.0
        angle = Math.Acos(dot)
        if angle > slopeBreak:
            self.CreateLine(wv, layerSerial, surface.GetVertexPoint(iVertexA), surface.GetVertexPoint(iVertexB))

    def CreateLineWithPoints(self, wv, layerSerial, p1, p2):
        if self.createcadpoly3d.IsChecked:
            pt1 = self.currentProject.Concordance[p1].Point
            pt2 = self.currentProject.Concordance[p2].Point
            l = wv.Add(clr.GetClrType(CadPoly3D))
            l.Append(pt1)
            l.Append(pt2)
            l.Layer = layerSerial
        else:
            l = wv.Add(clr.GetClrType(Linestring))
            e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IPointIdLocation))
            e.LocationSerialNumber = p1
            l.AppendElement(e)
            e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IPointIdLocation))
            e.LocationSerialNumber = p2
            l.AppendElement(e)
            l.Layer = layerSerial

    def CreateLine(self, wv, layerSerial, p1, p2):
        if self.createcadpoly3d.IsChecked:
            l = wv.Add(clr.GetClrType(CadPoly3D))
            l.Append(p1)
            l.Append(p2)
            l.Layer = layerSerial
        else:
            l = wv.Add(clr.GetClrType(Linestring))
            e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
            e.Position = p1
            l.AppendElement(e)
            e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
            e.Position = p2
            l.AppendElement(e)
            l.Layer = layerSerial

    def OkClicked(self, sender, e):
        if self.surface.SelectedSerial == 0:
            self.surface.StatusMessage = "Select reference surface"
            return False

        if self.slopeCB.IsChecked:
            slope = self.slopeLimitCtl.Angle
            if Double.IsNaN(slope):
                self.slopeLimitCtl.StatusMessage = "Enter valid slope"
                return False
            
        wv = self.currentProject[Project.FixedSerial.WorldView]
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        # pause the graphics so every change doesn't draw
        wv.PauseGraphicsCache(True)
        try:
            layerSerial = self.LineLayers.SelectedSerialNumber
            self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
            surface = wv.Lookup(self.surface.SelectedSerial)
            # the "with" statement will unroll any changes if something go wrong
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:             
                nVertices = surface.NumberOfVertices
                nTri = surface.NumberOfTriangles
                if self.slopeCB.IsChecked:
                    for i in range(nTri):  
                        if not surface.IsTriangleMaterialPresent(i): # skip "holes"
                            continue
                        triSlope = self.TriangleSlope(surface, i)
                        for side in range(3): # do 3 sides of triangle
                            self.DoTriangleSlopeEdge(i, triSlope, side, surface, wv, layerSerial, slope)
                else:
                    pnts={}
                    ProgressBar.TBC_ProgressBar.Title = "reading vertices: 0/" + str(nVertices)
                    time1 = datetime.now()
                    for i in range(nVertices):
                        if (datetime.now() - time1).seconds > 0.5:
                            ProgressBar.TBC_ProgressBar.Title = "reading vertices: " + str(i) + "/" + str(nVertices)
                            if ProgressBar.TBC_ProgressBar.SetProgress(i * 100 // nVertices):
                                break
                            time1 = datetime.now()
                        p = surface.GetVertexPoint(i)
                        #cadPoint = wv.Add(clr.GetClrType(CadPoint))
                        #cadPoint.Point0 = p
                        #if layerSerial:
                        #    cadPoint.Layer = layerSerial
                        ## map surface index to serial number
                        #pnts.Add(i, cadPoint.SerialNumber)
                        pnts.Add(i, p)
                    ProgressBar.TBC_ProgressBar.Title = "drawing lines: 0/" + str(nTri)
                    time1 = datetime.now()
                    for i in range(nTri):
                        if (datetime.now() - time1).seconds > 0.5:
                            ProgressBar.TBC_ProgressBar.Title = "drawing lines: " + str(i) + "/" + str(nTri)
                            if ProgressBar.TBC_ProgressBar.SetProgress(i * 100 // nTri):
                                break
                            time1 = datetime.now()
                        if not surface.IsTriangleMaterialPresent(i): # skip "holes"
                            continue
                        for side in range(3): # do 3 sides of triangle
                            self.DoTriangleEdge(i, side, pnts, surface, wv, layerSerial)
                    ProgressBar.TBC_ProgressBar.Title = ""
                    
                    ProgressBar.TBC_ProgressBar.Title = "updating Worldview"

                failGuard.Commit()

        finally:
            self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
            UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
            wv.PauseGraphicsCache(False)
            ProgressBar.TBC_ProgressBar.Title = ""
        self.SaveOptions()
        #sender.CloseUICommand()


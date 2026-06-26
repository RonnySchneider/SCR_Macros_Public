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
    "create3dfaces": False,
    "createlinestring": True,
    "createcadpoly3d": False,
    "anglefilter": True,
    "setcolor": True,
    "outcolorpicker": TrimbleColor.ByLayer,
    "picklayer": False,
    "layerpicker": 8,
    "addlayersuffix": True,
    "addtolayername": " - Lines",
}

def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_ExplodeIFC"
    cmdData.CommandName = "SCR_ExplodeIFC"
    cmdData.Caption = "_SCR_ExplodeIFC"
    cmdData.UIForm = "SCR_ExplodeIFC" # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!

    try:
        cmdData.DefaultTabKey = "SCR Expld-SNR-Relay-Prop"
        cmdData.DefaultTabGroupKey = "Explode"
        cmdData.ShortCaption = "Explode IFC"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3
     
        cmdData.Version = 1.28
        cmdData.MacroAuthor = "SCR"
        cmdData.ToolTipTitle = "explode IFC"
        cmdData.ToolTipTextFormatted = "explode IFC solid objects with angle filter"
    except:
        pass
    try:
        b = Bitmap(macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass

# the name of this class must match name from cmdData.UIForm (above)
class SCR_ExplodeIFC(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader(macroFileFolder + r"\SCR_ExplodeIFC.xaml") as s:
            wpf.LoadComponent(self, s)
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
        
        #self.ifcType = clr.GetClrType(IFCMesh)
        self.ifcType = clr.GetClrType(BIMEntity)

        self.ifcs.IsEntityValidCallback = self.IsValidIFC


        # populate combobox
        # item = ComboBoxItem()
        # item.Content = "Up"
        # item.FontSize = 12
        # self.facedirection.Items.Add(item)
        # item = ComboBoxItem()
        # item.Content = "Down"
        # item.FontSize = 12
        # self.facedirection.Items.Add(item)
        # item = ComboBoxItem()
        # item.Content = "Vertical"
        # item.FontSize = 12
        # self.facedirection.Items.Add(item)



        SCRExpanders.wire_pairs([
            (self.expander_anglefilter, self.anglefilter),
        ])
		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions(None, None)
        except:
            pass

        #self.toldecimals.NumberOfDecimals = 0
        #self.toldecimals.MinValue = 0.0
        #self.Loaded += self.SetDefaultOptions
    

    def IsValidIFC(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.ifcType):
            return True
        if isinstance(o, clr.GetClrType(IFCMesh)):
            return True
        return False
        
    
    def SetDefaultOptions(self, sender, e):
        SCROptions.LoadMacroOptions(self, "SCR_ExplodeIFC", _OPTIONS, self.currentProject)
        self.maindirectionui.Angle = OptionsManager.GetDouble("SCR_ExplodeIFC.maindirection", 0.0)
        self.anglevariationupui.Angle = OptionsManager.GetDouble("SCR_ExplodeIFC.anglevariationup", 0.0)
        self.anglevariationdownui.Angle = OptionsManager.GetDouble("SCR_ExplodeIFC.anglevariationdown", 0.0)

    def SaveOptions(self):
        SCROptions.SaveMacroOptions(self, "SCR_ExplodeIFC", _OPTIONS)
        OptionsManager.SetValue("SCR_ExplodeIFC.maindirection", self.maindirectionui.Angle)
        OptionsManager.SetValue("SCR_ExplodeIFC.anglevariationup", self.anglevariationupui.Angle)
        OptionsManager.SetValue("SCR_ExplodeIFC.anglevariationdown", self.anglevariationdownui.Angle)

    def CancelClicked(self, thisCmd, args):
        thisCmd.CloseUICommand()

    def OkClicked(self, thisCmd, e):
        Keyboard.Focus(self.okBtn)  # a trick to evaluate all input fields before execution, otherwise you'd have to click in another field first

        wv = self.currentProject[Project.FixedSerial.WorldView]
        wv.PauseGraphicsCache(True)

        inputok=True
        try:
            #anglevariation = abs(math.radians(self.anglevariation.Value))
            self.anglevariationup = abs(self.anglevariationupui.Angle)
            self.anglevariationupui.Angle = self.anglevariationup
            self.anglevariationdown= abs(self.anglevariationdownui.Angle)
            self.anglevariationdownui.Angle = self.anglevariationdown
        except:
            inputok = False

        if inputok:
            UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
            self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
            try:
                with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
 
                    self.testlimit = None
                    activeForm = TrimbleOffice.TheOffice.MainWindow.AppViewManager.ActiveView
                    if activeForm.View.LimitBoxActive():
                        limitbox = self.currentProject.Concordance.Lookup(activeForm.View.LimitBoxSerial)
                        
                        # unfortunately the Limitbox doesn't provide it's Matrix4D, so we need to recreate one ourself
                        # create some helper points
                        p0 = Point3D(0,0,0)
                        px = Point3D(1000,0,0)
                        py = Point3D(0,1000,0)
                        pz = Point3D(0,0,1000)
                        
                        # rotate everything around x-axis
                        bivector = BiVector3D(Vector3D(p0, px), limitbox.RotateX)
                        rot = Matrix4D.BuildTransformMatrix(Vector3D(p0), Vector3D(0,0,0), Spinor3D(bivector), Vector3D(1,1,1))
                        pxrot = rot.TransformPoint(px)
                        pyrot = rot.TransformPoint(py)
                        pzrot = rot.TransformPoint(pz)
                        
                        # rotate everything around y-axis
                        bivector = BiVector3D(Vector3D(p0, pyrot), limitbox.RotateY)
                        rot = Matrix4D.BuildTransformMatrix(Vector3D(p0), Vector3D(0,0,0), Spinor3D(bivector), Vector3D(1,1,1))
                        pxrot = rot.TransformPoint(pxrot)
                        pyrot = rot.TransformPoint(pyrot)
                        pzrot = rot.TransformPoint(pzrot)

                        # rotate everything around z-axis
                        #!!! UNTIL NOW THIS ROTATION IS AROUND WORLD Z AND NOT THE LIMITBOX Z !!!!
                        #!!! CAN'T UNDERSTAND WHY THEY DID THAT, IT'S TOTALLY ILOGIC !!!
                        bivector = BiVector3D(Vector3D(p0, pz), limitbox.RotateZ)
                        rot = Matrix4D.BuildTransformMatrix(Vector3D(p0), Vector3D(0,0,0), Spinor3D(bivector), Vector3D(1,1,1))
                        pxrot = rot.TransformPoint(pxrot)
                        pyrot = rot.TransformPoint(pyrot)
                        pzrot = rot.TransformPoint(pzrot)

                        # rotation from World to Limitbox
                        newspinor = Spinor3D.ComputeRotation(Vector3D(p0, px), Vector3D(p0, py), Vector3D(p0, pxrot), Vector3D(p0, pyrot))
                        
                        # world to Limitbox
                        limitboxMatrix4D = Matrix4D.BuildTransformMatrix(Vector3D(limitbox.Origin), Vector3D(0,0,0), newspinor, Vector3D(1,1,1))
                        self.reverselimitboxMatrix4D = Matrix4D.Inverse(limitboxMatrix4D)
                        
                        l1 = Point3D(-limitbox.Length/2, -limitbox.Width/2, -limitbox.Height/2) + limitbox.Origin
                        l2 = Point3D(limitbox.Length/2, limitbox.Width/2, limitbox.Height/2) + limitbox.Origin

                        self.testlimit = Limits3D(l1, l2)

                    # we use IFCs
                    # ifclayer  = self.currentProject.Concordance.Lookup(self.ifclayerpicker.SelectedSerialNumber)    # we get just the ifc layer as an object
                    # ifcmembers = ifclayer.Members  # we get serial number list of all the elements on that layer
                    ifccount = 0
                    for o in self.ifcs.SelectedMembers(self.currentProject):
                        if isinstance(o, self.ifcType) or isinstance(o, IFCMesh): # in case it is an IFC Mesh we get us the coordinates
                            ifccount += 1   

                    j = 0
                    for o in self.ifcs.SelectedMembers(self.currentProject):
                        # o = self.currentProject.Concordance.Lookup(sn)
                        
                        if isinstance(o, self.ifcType): # in case it is an IFC Mesh we get us the coordinates
                            # need overall lists in case we have submeshes which need to add to those
                            self.verticiesGlobal = []
                            self.vertexIndices = []

                            self.setoutputlayer(o)

                            for shellMeshInstance in o.GetGeometry():

                                shellMeshData = shellMeshInstance.GetShellMeshData()

                                # DEPENDING ON THE TYPE OF IFC THE DIFFERENT METHODS RETURN EMPTY LISTS
                                try:
                                    self.vertexIndicesTemp = shellMeshData.GetTriangulatedFaceList() # this works for the bridge IFC
                                    if self.vertexIndicesTemp.Count == 0:
                                        self.vertexIndicesTemp = shellMeshData.GetFaces()    # this works for the geotech, but not the bridges
                                    for vi in self.vertexIndicesTemp:
                                        self.vertexIndices.Add(vi)
                                    verticiesLocal = shellMeshData.GetVertex() # verticies as Point3Ds
                                    for v in range(0, verticiesLocal.Count):
                                        self.verticiesGlobal.Add(shellMeshInstance.GlobalTransformation.TransformPoint(verticiesLocal[v]))

                                except: # new for 2024.00
                                    self.verticiesGlobal = []
                                    self.vertexIndices = []
                                    self.verticiesGlobalTemp = shellMeshData.GetVertexList(shellMeshInstance.GlobalTransformation)
                                    for v in self.verticiesGlobalTemp:
                                        self.verticiesGlobal.Add(v)
                                    self.vertexIndicesTemp = shellMeshData.GetIndices(TriangulationContext(1)) # triangulationContext 0 - NonTriangulated; 1 - Triangulated; 2 - TriangulatedImplicit without the length entry 3
                                    for vi in self.vertexIndicesTemp:
                                        self.vertexIndices.Add(vi)
                                
                                self.preparevertexmap()

                                self.checkanddrawlines()
                                self.vertexIndices = None
                                self.verticiesGlobal = None

                        elif isinstance(o, IFCMesh):

                            self.setoutputlayer(o)

                            self.vertexIndices = o. GetIndices(TriangulationContext(1))
                            self.verticiesGlobal = o.GetVertices(TransformationContext.Transformed_Global)

                            self.preparevertexmap()

                            self.checkanddrawlines()
                            self.vertexIndices = None
                            self.verticiesGlobal = None
                    
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
        wv.PauseGraphicsCache(False)

    def preparevertexmap(self):
        ProgressBar.TBC_ProgressBar.Title = "preparing vertex map"
        tolerance = 0.0001
        n = self.verticiesGlobal.Count
        items = []
        for i in range(n):
            v = self.verticiesGlobal[i]
            items.append((v.X, v.Y, v.Z, i))
        vtree = SCROctree()
        vtree.build(items)
        self.vertex_map = {}
        for i in range(n):
            v = self.verticiesGlobal[i]
            hits = vtree.find_within(v.X, v.Y, v.Z, tolerance)
            low = min(hits) if hits else i
            # resolve through any already-mapped index so transitive chains collapse to the true canonical
            canonical = self.vertex_map.get(low, low)
            if canonical != i:
                self.vertex_map[i] = canonical

    def _create_linestring(self, wv, p1, p2):
        if self.createcadpoly3d.IsChecked:
            l = wv.Add(clr.GetClrType(CadPoly3D))
            l.Append(p1)
            l.Append(p2)
            l.Layer = self.outputlayer.SerialNumber
            if self.setcolor.IsChecked:
                l.ColorOfObject = self.outcolorpicker.SelectedColor
        else:
            l = wv.Add(clr.GetClrType(Linestring))
            e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
            e.Position = p1
            l.AppendElement(e)
            e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
            e.Position = p2
            l.AppendElement(e)
            l.Layer = self.outputlayer.SerialNumber
            if self.setcolor.IsChecked:
                l.Color = self.outcolorpicker.SelectedColor

    def setoutputlayer(self, o):

        if self.picklayer.IsChecked:
            self.outputlayer = self.currentProject.Concordance.Lookup(self.layerpicker.SelectedSerialNumber)

        else:

            if o.Layer != 0:
                inputlayer = self.currentProject.Concordance.Lookup(o.Layer)
            else:
                inputlayer = self.currentProject.Concordance.Lookup(8)
            self.outputlayer = Layer.FindOrCreateLayer(self.currentProject, inputlayer.Name + self.addtolayername.Text)
            if self.outputlayer.LayerGroupSerial == 0: # only set layer group if it is a new layer, otherwise leave at its location
                self.outputlayer.LayerGroupSerial = inputlayer.LayerGroupSerial

        return

    def checkanddrawlines(self):

        wv = self.currentProject[Project.FixedSerial.WorldView]
        activeForm = TrimbleOffice.TheOffice.MainWindow.AppViewManager.ActiveView

        if self.vertexIndices and self.vertexIndices.Count > 0 and self.verticiesGlobal and self.verticiesGlobal.Count > 0:
            ProgressBar.TBC_ProgressBar.Title = "prepare Lines"
            linetuples = []
            wholetriangles = []
    
            linestodraw = 0
            j2 = 0
            time1 = datetime.now()
            for t in range(0, self.vertexIndices.Count, 4): # step through all the faces
                
                j2 += 1
                if (datetime.now() - time1).seconds > 0.5:
                    ProgressBar.TBC_ProgressBar.Title = "analyzing triangles: " + str(j2) + "/" + str(self.vertexIndices.Count)
                    if ProgressBar.TBC_ProgressBar.SetProgress(j2 * 100 // (self.vertexIndices.Count / 4)):
                        break   # function returns true if user pressed cancel
                    time1 = datetime.now()
                # it can be that the IFC contains "triangles" where all three points are on one line, we don't want those
                # that would lead to a division by zero in the algorithm that checks if the perpendicular solution is within the triangle
                # checking it here is doubling up some computations, but it could also mess up the plane and normal vector creation
                vertex1 = self.verticiesGlobal[self.vertexIndices[t+1]]
                vertex2 = self.verticiesGlobal[self.vertexIndices[t+2]]
                vertex3 = self.verticiesGlobal[self.vertexIndices[t+3]]
    
                a = Vector3D(vertex1, vertex2)
                b = Vector3D(vertex1, vertex3)
                
                aa = Vector3D.DotProduct(a,a)[0]
                ab = Vector3D.DotProduct(a,b)[0]
                bb = Vector3D.DotProduct(b,b)[0]
                
                d = round(ab * ab - aa * bb, 14) # otherwise it could still happen that "triangles" slip through
    
                # if vertex1.X == vertex2.X and vertex1.X == vertex3.X:
                #     tt = 0
    
                if d != 0:
                    # find out if we want that triangle
                    plane = Plane3D(vertex1, vertex2, vertex3)[0]
                    nv = plane.normal
                    nv.Length = -0.1 # the normal vectors seem to face to the inside of the solid object, so we turn it around
                    #nv.Length = 0.1 # the normal vectors now seem to face outwards
                    
                    # # test show the normal vectors
                    # l = wv.Add(clr.GetClrType(Linestring))
                    # e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                    # e.Position = vertex1 
                    # l.AppendElement(e)
                    # e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                    # e.Position = (vertex1 + nv)
                    # l.AppendElement(e)
                    # l.Layer = self.outputlayer.SerialNumber
                    # l.Color = Color.Blue
    
                    # Vector3D.Horizon is positive above the horizon and negative below
                    # VerticalAngleEdit starts with 0 in Zenith 
                    self.maindirection = (-1 * (self.maindirectionui.Angle - math.pi/2)) + math.pi/2
                    
                    addtriangle = False
                    # check against the slope filter
                    if self.anglefilter.IsChecked:
                        if  nv.Horizon >= (self.maindirection - self.anglevariationdown) and \
                            nv.Horizon <= (self.maindirection + self.anglevariationup):
                    
                            addtriangle = True
                    else:
                            addtriangle = True
    
                    # check against the limitbox
                    if activeForm.View.LimitBoxActive() and addtriangle == True:
                    
                        add3dface = 0
    
                        if self.testlimit.IsPointInside(self.reverselimitboxMatrix4D.TransformPoint(vertex1), True)[0] and \
                           self.testlimit.IsPointInside(self.reverselimitboxMatrix4D.TransformPoint(vertex2), True)[0]:
                            linetuples.Add([self.vertexIndices[t+1], self.vertexIndices[t+2]])
                            add3dface += 1
    
                        if self.testlimit.IsPointInside(self.reverselimitboxMatrix4D.TransformPoint(vertex2), True)[0] and \
                           self.testlimit.IsPointInside(self.reverselimitboxMatrix4D.TransformPoint(vertex3), True)[0]:
                            linetuples.Add([self.vertexIndices[t+2], self.vertexIndices[t+3]])
                            add3dface += 1
    
                        if self.testlimit.IsPointInside(self.reverselimitboxMatrix4D.TransformPoint(vertex3), True)[0] and \
                           self.testlimit.IsPointInside(self.reverselimitboxMatrix4D.TransformPoint(vertex1), True)[0]:
                            linetuples.Add([self.vertexIndices[t+3], self.vertexIndices[t+1]])
                            add3dface += 1
    
                        if add3dface == 3:
                            wholetriangles.Add([self.vertexIndices[t+1], self.vertexIndices[t+2], self.vertexIndices[t+3]])
    
                    else:
                        if addtriangle == True:
                            if self.create3dfaces.IsChecked:
                                wholetriangles.Add([self.vertexIndices[t+1], self.vertexIndices[t+2], self.vertexIndices[t+3]])
                            else:
                                linetuples.Add([self.vertexIndices[t+1], self.vertexIndices[t+2]])
                                linetuples.Add([self.vertexIndices[t+2], self.vertexIndices[t+3]])
                                linetuples.Add([self.vertexIndices[t+3], self.vertexIndices[t+1]])
    
            ## cleanup test
            #filename = os.path.expanduser('~/Downloads/test.csv')
            #if File.Exists(filename):
            #    File.Delete(filename)
            #with open(filename, 'w') as f:       
            #
            #    for t in linetuplesclean:
            #        outputline = str(t[0]) + ',' + str(t[1])
            #        f.write(outputline + "\n")
            #
            #    f.close()

            if self.create3dfaces.IsChecked:
                j2 = 0
                time1 = datetime.now()
                for triangle in wholetriangles:
                    
                    j2 += 1
                    if (datetime.now() - time1).seconds > 0.5:
                        ProgressBar.TBC_ProgressBar.Title = "drawing 3DFace: " + str(j2) + "/" + str(wholetriangles.Count)
                        if ProgressBar.TBC_ProgressBar.SetProgress(j2 * 100 // wholetriangles.Count):
                            break   # function returns true if user pressed cancel
                        time1 = datetime.now()
    
                    f = wv.Add(clr.GetClrType(Face3D))
                    
                    f.Point0 = self.verticiesGlobal[triangle[0]]
                    f.Point1 = self.verticiesGlobal[triangle[2]]
                    f.Point2 = self.verticiesGlobal[triangle[1]]
                    f.Point3 = self.verticiesGlobal[triangle[0]]
    
                    f.Transparency = 100
                    f.Layer = self.outputlayer.SerialNumber
                    if self.setcolor.IsChecked:
                        f.Color = self.outcolorpicker.SelectedColor
    
            else: # lines only

                # remap duplicate vertex IDs to their canonical counterpart
                for i in range(linetuples.Count):
                    linetuples[i][0] = self.vertex_map.get(linetuples[i][0], linetuples[i][0])
                    linetuples[i][1] = self.vertex_map.get(linetuples[i][1], linetuples[i][1])

                ProgressBar.TBC_ProgressBar.Title = "removing duplicate Lines"
                linetuplesclean = list(set(map(lambda t: tuple(sorted(t)), linetuples)))

                j2 = 0
                total = len(linetuplesclean)
                ProgressBar.TBC_ProgressBar.Title = "drawing Lines: 0/" + str(total)
                time1 = datetime.now()
                for drawtuple in linetuplesclean:
                    j2 += 1
                    if (datetime.now() - time1).seconds > 0.5:
                        ProgressBar.TBC_ProgressBar.Title = "drawing Lines: " + str(j2) + "/" + str(total)
                        if ProgressBar.TBC_ProgressBar.SetProgress(j2 * 100 // total):
                            break
                        time1 = datetime.now()
                    p1 = self.verticiesGlobal[drawtuple[0]]
                    p2 = self.verticiesGlobal[drawtuple[1]]
                    if Vector3D(p1, p2).Length != 0:
                        self._create_linestring(wv, p1, p2)
    

            linetuples = None
            linetuplesclean = None
            ProgressBar.TBC_ProgressBar.Title = "updating Worldview"    
            #j += 1
            #ProgressBar.TBC_ProgressBar.Title = "exploded IFC: " + str(j) + "/" + str(ifccount)
            #if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(j * 100 / ifccount)):
            #    break   # function returns true if user pressed cancel
    
        return


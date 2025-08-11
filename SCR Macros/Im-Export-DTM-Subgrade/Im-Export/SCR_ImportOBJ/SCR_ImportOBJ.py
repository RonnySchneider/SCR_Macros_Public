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

from System.Collections.Generic import List, IEnumerable # import here, otherwise there is a weird issue with Count and Add for lists
import os
exec(open("C:\ProgramData\Trimble\MacroCommands3\SCR Macros\SCR_Imports.py").read())

def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_ImportOBJ"
    cmdData.CommandName = "SCR_ImportOBJ"
    cmdData.Caption = "_SCR_ImportOBJ"
    cmdData.UIForm = "SCR_ImportOBJ"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "Import/Export"
        cmdData.ShortCaption = "Import OBJ"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.03
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
        self.openfilename.Text = OptionsManager.GetString("SCR_ImportOBJ.openfilename", os.path.expanduser('~\\Downloads'))

        self.drawlines.IsChecked = OptionsManager.GetBool("SCR_ImportOBJ.drawlines", False)
        self.drawifc.IsChecked = OptionsManager.GetBool("SCR_ImportOBJ.drawifc", False)
        self.createsurface.IsChecked = OptionsManager.GetBool("SCR_ImportOBJ.createsurface", False)

    def SaveOptions(self):
        OptionsManager.SetValue("SCR_ImportOBJ.openfilename", self.openfilename.Text)

        OptionsManager.SetValue("SCR_ImportOBJ.drawlines", self.drawlines.IsChecked)
        OptionsManager.SetValue("SCR_ImportOBJ.drawifc", self.drawifc.IsChecked)
        OptionsManager.SetValue("SCR_ImportOBJ.createsurface", self.createsurface.IsChecked)


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

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''

        #if File.Exists(self.openfilename.Text) == False:
        #    self.error.Content += 'no file selected\n'
        
        
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())

        wv = self.currentProject [Project.FixedSerial.WorldView]
        wv.PauseGraphicsCache(True)
            
        verticies = []
        faces = []
        linetuples = []

        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                if self.developoverlay.IsChecked:

                    self.drawoverlay()

                else:

                    ProgressBar.TBC_ProgressBar.Title = "reading the file"
                    with open(self.openfilename.Text,'r') as csvfile: 
                        reader = csv.reader(csvfile, delimiter=' ', quotechar='|') 
                        ProgressBar.TBC_ProgressBar.Title = "preparing data tables"
                        
                        objectname = ''
                        #polygongroup = ''
                        overallvertexcount = 0
                        currentobjectvertexcount = 0
                        prevrow = ['']

                        for row in reader:

                            if (prevrow[0] != row[0] and prevrow[0] == 'f') or len(row) == 0:
                                if verticies.Count > 0 and faces.Count > 0:
                                    self.startdrawing(verticies, faces, linetuples, objectname)
                            
                                # cleanup for next sub object
                                currentobjectvertexcount = 0
                                verticies = []
                                faces = []
                                linetuples = []
                            
                            if row[0] == "g":
                                objectname = str(row[1])
                            #elif row[0] == "o":
                            #    polygongroup = str(row[1])
                            elif row[0] == "v":
                                overallvertexcount += 1
                                currentobjectvertexcount += 1
                                verticies.Add([Point3D(float(row[1]), float(row[2]), float(row[3]))])
                            elif row[0] == "f":
                                # if the faces have a texture applied the face line could potentially look like this
                                # f v1/vt1/vn1 v2/vt2/vn2 v3/vt3/vn3
                                # v is the vertex, vt the texture coordinate and vn a vertex normal
                                # currently we only care about the vertex
                                verttemp = []
                                # the row is already split by spaces - i.e. f ; 1//1 ; 2//1 ; 3//1 
                                # 
                                for i in range (1, len(row)):
                                    if "/" in row[i]:
                                        temp = row[i].split("/")
                                        verttemp.Add(temp[0])
                                    else:
                                        verttemp.Add(row[i])

                                vindexoffset = overallvertexcount - currentobjectvertexcount
                                # int(verttemp[0]) is the overall vertex index
                                v1i = int(verttemp[0]) - vindexoffset - 1
                                v2i = int(verttemp[1]) - vindexoffset - 1
                                v3i = int(verttemp[2]) - vindexoffset - 1
                                # the obj vertice count is not zero based, hence the -1
                                
                                faces.Add([v1i, v2i, v3i])
                                
                                linetuples.Add([v1i, v2i])
                                linetuples.Add([v2i, v3i])
                                linetuples.Add([v3i, v1i])


                            prevrow = row
                        
                        # trigger another drawing for the end of the file
                        # in case there wasn't an empty line at the end  
                        if verticies.Count > 0 and faces.Count > 0:
                            self.startdrawing(verticies, faces, linetuples, objectname)

                  
                
                failGuard.Commit()
                UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
                self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
        
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

    def startdrawing(self, verticies, faces, linetuples, objectname):

        wv = self.currentProject [Project.FixedSerial.WorldView]
        
        if self.drawlines.IsChecked or self.createsurface.IsChecked:
            ProgressBar.TBC_ProgressBar.Title = "removing duplicate lines"
            linetuplesclean = list(set(map(lambda i: tuple(sorted(i)), linetuples)))
                    
        if self.drawlines.IsChecked:
            timerresults = []

            # that's just for the progressbar
            barvalueinc = 0.05
            barvalueinccount = linetuplesclean.Count * barvalueinc
            barvalueinci = 0

            for i in range(linetuplesclean.Count):
                barvalueinci += 1
                if barvalueinci >= barvalueinccount:
                    barvalueinci = 0
                    ProgressBar.TBC_ProgressBar.Title = "drawing line: " + str(i) + "/" + str(linetuplesclean.Count)
                    if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(i * 100 / linetuplesclean.Count)):
                        break   # function returns true if user pressed cancel

                drawtuple = linetuplesclean[i]
                
                start_t = timer()

                #l = wv.Add(clr.GetClrType(Linestring))
                #e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                #e.Position = verticies[drawtuple[0]]
                #l.AppendElement(e)
                #e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                #e.Position = verticies[drawtuple[1]]
                #l.AppendElement(e)

                polyseg = PolySeg.PolySeg()
                polyseg.Add(verticies[drawtuple[0]][0])
                polyseg.Add(verticies[drawtuple[1]][0])
                l = wv.Add(clr.GetClrType(Linestring))
                l.Append(polyseg, None, False, False)
                l.Name = objectname

                end_t=timer()
                timerresults.Add(timedelta(seconds=end_t-start_t))

            filename = os.path.expanduser('~/Downloads/timerresult.csv')
            if File.Exists(filename):
                File.Delete(filename)
            with open(filename, 'w') as f:         

                for i in range(timerresults.Count):
                    outputline = str(i + 1) + "," + str(timerresults[i]) + "\n"
                    f.write(outputline)

                f.close()


        if self.drawifc.IsChecked:
            ProgressBar.TBC_ProgressBar.Title = "creating the IFC objects"

            tempfaces = []
            tempvertices = []

            #for v in range(0, faces.Count):
            #    # go through the list and reindex the vertices until we may find another mesh name
            #    
            #    
            #    if faces[3] != prevname:
            #        # start create a new IFC


            ifcvertices = Array[Point3D]([Point3D()] * verticies.Count)
            for i in range(0, verticies.Count):
                ifcvertices[i] = verticies[i][0]
                
            ifcfacelist = Array[Int32]([Int32()] * faces.Count * 4) # define size of facelist array
            ifcfacelistindex = 0                                                                                    

            for f in faces:
                ifcfacelist[ifcfacelistindex + 0] = 3
                ifcfacelist[ifcfacelistindex + 1] = f[0]
                ifcfacelist[ifcfacelistindex + 2] = f[1]
                ifcfacelist[ifcfacelistindex + 3] = f[2]
                ifcfacelistindex += 4

            ifcnormals = Array[Point3D]([Point3D()]*0)

            #def createifcobject(ifcprojectname, ifcshellgroupname, ifcshellname, ifcvertices, ifcfacelist, ifcnormals, color, transform, layer)
            self.createifcobject(os.path.basename(self.openfilename.Text), objectname, ifcvertices, ifcfacelist, ifcnormals, Color.Red.ToArgb(), Matrix4D(), 8)


        if self.createsurface.IsChecked:
            ProgressBar.TBC_ProgressBar.Title = "creating new surface"
            newSurface = wv.Add(clr.GetClrType(Model3D))
            newSurface.Name = Model3D.GetUniqueName(os.path.basename(self.openfilename.Text) + ' - ' + objectname, None, wv) #make sure name is unique
            # SmoothShading | ElevationContours | ShowBackfaces | FilledTriangles | ElevationContours_2D | FilledTriangles_2D
            newSurface.Mode = DisplayMode(1 + 64 + 128 + 512 + 1024 + 65536)
            builder = newSurface.GetGemBatchBuilder()
            mapVertices = {}
            for i in range(0, verticies.Count):
                v = builder.AddVertex(verticies[i][0])
            
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

    def createifcobject(self, ifcprojectname, ifcshellgroupname, ifcvertices, ifcfacelist, ifcnormals, color, transform, layer):
  
        # get the BIM entity collection from the world view - if it doesn't exist yet it is created
        bimEntityColl = BIMEntityCollection.ProvideEntityCollection(self.currentProject, True)
        shellMeshDataColl = ShellMeshDataCollection.ProvideShellMeshDataCollection(self.currentProject, True)
        # theoretical IFC hierarchy in TBC
        # IFCPROJECT
        #   IFCSITE
        #       IFCBUILDING
        #           IFCBUILDINGSTOREY
        #               IFCBUILDINGELEMENTPROXY
        
        # test if provided project name exists, if yes then add to it
        found = False
    
        for e in bimEntityColl:
            if e.Name == ifcprojectname:
                bimprojectEntity = e
                found = True
                continue
        if not found:
            bimprojectEntity = bimEntityColl.Add(clr.GetClrType(BIMEntity))
            bimprojectEntity.EntityType = "IFCPROJECT"
            bimprojectEntity.Description = ifcprojectname # must be set, otherwise export fails
            bimprojectEntity.BIMGuid = Guid.NewGuid()

        
        # find unique group name
        while found == True:
            found = False
            for e in bimprojectEntity:
                if e.Name == ifcshellgroupname:
                    found = True
            if found:
                ifcshellgroupname = PointHelper.Helpers.Increment(ifcshellgroupname, None, True)
        
        # if we create at least IFCPROJECT and IFCBUILDINGELEMENTPROXY (which includes the IFCMesh)
        # we can export it properly into an IFC file but keep the meshes separate
        # upon import TBC will structure it as follows
        # IFCPROJECT --> IFCPROJECT
        # everything will end up under "Default Site - IFCSITE"
        # the IFCBUILDINGELEMENTPROXY with the subsequent IFCMesh get combined into separate BIM Objects

        bimobjectEntity = bimprojectEntity.Add(clr.GetClrType(BIMEntity))
        bimobjectEntity.EntityType = "IFCBUILDINGELEMENTPROXY"
        bimobjectEntity.Description = ifcshellgroupname
        bimobjectEntity.BIMGuid = Guid.NewGuid()
        bimobjectEntity.Mode = DisplayMode(1 + 2 + 64 + 128 + 512 + 4096)
        # MUST set layer, otherwise the properties manager will falsly show layer "0"
        # but the object is invisible until manually changing the layer
        bimobjectEntity.Layer = layer 

        shellmeshdata = shellMeshDataColl.AddShellMeshData(self.currentProject)
        shellmeshdata.CreateShellMeshData(ifcvertices, ifcfacelist, ifcnormals)
        shellmeshdata.SetVolumeCalculationShell(ifcvertices, ifcfacelist)
        #shellmeshdata.Description = ifcshellgroupname
        #shellmeshdata = shellMeshDataColl.TryCreateFromShell3D(None, 0)

        meshInstance = bimobjectEntity.Add(clr.GetClrType(ShellMeshInstance))
        meshInstance.CreateShell(0, shellmeshdata.SerialNumber, color, transform)
        tt = 1
        
        #meshInstance.SetVolumeCalculationShell(ifcvertices, ifcfacelist)

        # it looks as if the use specifc enumerators from surface models here
        # which names don't really make sense in this context
        # SmoothShading + ElevationContours + ShowBackfaces + FilledTriangles + DrapeLines_2D
        #ifcMesh.Mode = DisplayMode(1 + 64 + 128 + 512 + 4096)
        #mesh._TriangulatedFaceList = ifcfacelist # otherwise GetTriangulatedFaceList() in the perpDist macro will not get the necessary values

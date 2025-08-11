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
    cmdData.Key = "SCR_SplitCombineIFCMeshes"
    cmdData.CommandName = "SCR_SplitCombineIFCMeshes"
    cmdData.Caption = "_SCR_SplitCombineIFCMeshes"
    cmdData.UIForm = "SCR_SplitCombineIFCMeshes"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "IFC"
        cmdData.ShortCaption = "Split Combine IFCs"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.03
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "split or combine IFC meshes"
        cmdData.ToolTipTextFormatted = "a IFC object can be composed of multiple sub-meshes, either split or combine them"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png") # we have to include a icon revision, otherwise TBC might not show the new one
        cmdData.ImageSmall = b
    except:
        pass


class SCR_SplitCombineIFCMeshes(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_SplitCombineIFCMeshes.xaml") as s:
            wpf.LoadComponent (self, s)
        self.currentProject = currentProject

        tvi = TreeViewItem()
        tvi.Tag = BIMEntityCollection.ProvideEntityCollection(self.currentProject, True)
        tvi.Header = 'BIM Data'
        self.Tree.Items.Add(tvi)
        tvi.IsExpanded = True

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

        optionMenu = SelectionContextMenuHandler()
        # remove options that don't apply here
        optionMenu.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        self.objs.ButtonContextMenu = optionMenu

        # Description is the actual property name inside the item in ticklist.Content.Items
        sd = System.ComponentModel.SortDescription("Description", System.ComponentModel.ListSortDirection.Ascending)

        self.helperbox.SearchContainer = BIMEntityCollection.ProvideEntityCollection(self.currentProject, True).SerialNumber
        self.helperbox.UseSelectionEngine = False
        self.helperbox.SetEntityType(clr.GetClrType(BIMEntity), self.currentProject)
        self.helperbox.Content.Items.SortDescriptions.Add(sd)

        self.helperbox.Content.ItemContainerGenerator.StatusChanged += self.reloadtree

        self.pointcloudType = clr.GetClrType(PointCloudRegion)
        self.lType = clr.GetClrType(IPolyseg)

        self.meshpicker.ValueChanged += self.meshpickerChanged
        self.meshpicker.AutoTab = False
        #self.Loaded += self.reloadtree

        self.SetDefaultOptions()

    def Dispose(self, thisCmd, disposing):
        #self.Loaded -= self.reloadtree
        self.helperbox.Content.ItemContainerGenerator.StatusChanged -= self.reloadtree
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)

    def reloadtree(self, sender, e):
        # generator status - 0-not started; 1-generating; 2-generated/done; 3-error
        if sender == None or int(sender.Status) == 2:
        
            tvi = self.Tree.Items[0]
            tvi.Items.Clear()

            self.AddTreeNodes(tvi, BIMEntityCollection.ProvideEntityCollection(self.currentProject, True))
            tvi.IsExpanded = True

    def AddTreeNodes(self, containerTreeNode, container):
        include = False

        for o in container:
            tvi = TreeViewItem() 
            tvi.Tag = o
            tt = tvi.Margin
            tt.Left = -10
            tvi.Margin = tt
            if isinstance(o, BIMEntity):# and not isinstance(o, ShellMeshInstance):

                text = o.Name + ' - ' + o.EntityType
                include = True
                tvi.Header = text
                containerTreeNode.Items.Add(tvi)
                tvi.IsExpanded = True

                self.AddTreeNodes(tvi, o)

        return

    def radiochange(self, send, e):

        self.objs.IsEntityValidCallback = self.IsValidObject
        
        i = -1
        # may look a bit cluttered but may make it simpler to see what is enabled when
        if self.split.IsChecked: i = 0
        elif self.combine.IsChecked: i = 1
        elif self.move.IsChecked: i = 2
        elif self.addentry.IsChecked: i = 3
        elif self.createfrom3DFaces.IsChecked: i = 4
        elif self.createfrompoints.IsChecked: i = 5
        elif self.reversetriangulation.IsChecked: i = 6
        
        if i > -1:
            isenabled = []
                             #0 multiobjectpick #1 meshpicker #2 namefield   # 3 Tree    # 4 invisible   # 5 colorpicker     # 6 BIM Type    # 7 layer   # 8 del face
            isenabled.Add(  [1                  ,0            ,0              ,0          ,0              ,0                  ,0              ,0          , 0         ]) #split
            isenabled.Add(  [1                  ,0            ,1              ,1          ,0              ,0                  ,1              ,1          , 0         ]) #combine
            isenabled.Add(  [1                  ,0            ,0              ,1          ,1              ,0                  ,1              ,0          , 0         ]) #move
            isenabled.Add(  [1                  ,0            ,1              ,1          ,0              ,0                  ,1              ,1          , 0         ]) #addentry
            isenabled.Add(  [1                  ,0            ,1              ,1          ,0              ,1                  ,1              ,1          , 1         ]) #create from 3dfaces
            isenabled.Add(  [1                  ,0            ,1              ,1          ,0              ,1                  ,1              ,1          , 0         ]) #create from points
            isenabled.Add(  [0                  ,1            ,0              ,0          ,0              ,0                  ,0              ,0          , 0         ]) #reverse triangulation
            
            # Visibility Hidden = 1 --- Visible = 0
            self.objs.IsEnabled = bool(isenabled[i][0])
            self.objs.Visibility = Visibility(not bool(isenabled[i][0]))
            self.meshpicker.IsEnabled = bool(isenabled[i][1])
            self.meshpicker.Visibility = Visibility(not bool(isenabled[i][1]))
            self.combinedname.IsEnabled = bool(isenabled[i][2])
            self.combinedname.Visibility = Visibility(not bool(isenabled[i][2]))
            self.Tree.IsEnabled = bool(isenabled[i][3])
            self.Tree.Visibility = Visibility(not bool(isenabled[i][3]))
            self.triggerinvisible.IsEnabled = bool(isenabled[i][4])
            self.triggerinvisible.Visibility = Visibility(not bool(isenabled[i][4]))
            self.meshcolorpicker.IsEnabled = bool(isenabled[i][5])
            self.meshcolorpicker.Visibility = Visibility(not bool(isenabled[i][5]))
            self.bimtype.IsEnabled = bool(isenabled[i][6])
            self.bimtype.Visibility = Visibility(not bool(isenabled[i][6]))
            self.layerpicker.IsEnabled = bool(isenabled[i][7])
            self.layerpicker.Visibility = Visibility(not bool(isenabled[i][7]))
            self.deletefaces.IsEnabled = bool(isenabled[i][8])
            self.deletefaces.Visibility = Visibility(not bool(isenabled[i][8]))


    def IsValidObject(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        
        if self.createfrom3DFaces.IsChecked:
            if isinstance(o, Face3D):
                return True
        elif self.createfrompoints.IsChecked:
            if isinstance(o, CadPoint) or isinstance(o, CoordPoint) or isinstance(o, self.pointcloudType) or isinstance(o, self.lType):
                return True
        #elif not self.createfrom3DFaces.IsChecked:
        else:
            if isinstance(o, BIMEntity):
                return True

        return False

    def meshpickerChanged(self, send, e):

        self.overlaynormals = []
        o = self.meshpicker.Entity
        if o != None:
            GlobalSelection.Clear()
            GlobalSelection.Items(self.currentProject).Set([o.SerialNumber])


            if isinstance(o, BIMEntity):
                # only do it if we have single meshinstances
                if o.Count == 1:
                    for orgmeshinstance in o:
                    
                        orgshellmeshdata = orgmeshinstance.GetShellMeshData()
                        verticiesGlobalTemp = orgshellmeshdata.GetVertexList(orgmeshinstance.GlobalTransformation)
                        vertexIndicesTemp = orgshellmeshdata.GetIndices(TriangulationContext(1)) # triangulationContext 0 - NonTriangulated; 1 - Triangulated; 2 - TriangulatedImplicit without the length entry 3

                        for i in range(0, vertexIndicesTemp.Count, 4):

                            p = Plane3D(verticiesGlobalTemp[vertexIndicesTemp[i + 1]], verticiesGlobalTemp[vertexIndicesTemp[i + 2]], verticiesGlobalTemp[vertexIndicesTemp[i + 3]])[0]
                            nv = p.normal
                            nv.Length = 0.2

                            self.overlaynormals.Add([verticiesGlobalTemp[vertexIndicesTemp[i + 1]], verticiesGlobalTemp[vertexIndicesTemp[i + 1]] + nv])

        else:
            GlobalSelection.Clear()

        self.drawoverlay()

    def drawoverlay(self):

        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)
        self.overlayBag = OverlayBag(self.ViewOverlay) # create a new/empty overlaybag

        for pointpair in self.overlaynormals:
            self.overlayBag.AddPolyline(Array[Point3D]([pointpair[0], pointpair[1]]), Color.Blue.ToArgb(), 4)

        # if it only needs to be visible in all Planview then remove the Hoops3DViewGUID
        array = Array[Guid]([DisplayWindow.Hoops3DViewGUID, DisplayWindow.HoopsPlanViewGUID])
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.AddOverlayGeometry(array, self.overlayBag)

        return


    def SetDefaultOptions(self):
        
        self.split.IsChecked = OptionsManager.GetBool("SCR_SplitCombineIFCMeshes.split", True)
        self.combine.IsChecked = OptionsManager.GetBool("SCR_SplitCombineIFCMeshes.combine", False)
        self.move.IsChecked = OptionsManager.GetBool("SCR_SplitCombineIFCMeshes.move", False)
        self.addentry.IsChecked = OptionsManager.GetBool("SCR_SplitCombineIFCMeshes.addentry", False)
        self.createfrom3DFaces.IsChecked = OptionsManager.GetBool("SCR_SplitCombineIFCMeshes.createfrom3DFaces", False)
        self.createfrompoints.IsChecked = OptionsManager.GetBool("SCR_SplitCombineIFCMeshes.createfrompoints", False)
        self.reversetriangulation.IsChecked = OptionsManager.GetBool("SCR_SplitCombineIFCMeshes.reversetriangulation", False)
        self.deletefaces.IsChecked = OptionsManager.GetBool("SCR_SplitCombineIFCMeshes.deletefaces", False)
        self.triggerinvisible.IsChecked = OptionsManager.GetBool("SCR_SplitCombineIFCMeshes.triggerinvisible", False)
        self.combinedname.Text = OptionsManager.GetString("SCR_SplitCombineIFCMeshes.combinedname", "")

        self.bimtype.SelectedIndex = OptionsManager.GetInt("SCR_SplitCombineIFCMeshes.bimtype", 0)

        try:    self.meshcolorpicker.SelectedColor = Color.FromArgb(OptionsManager.GetInt("SCR_SplitCombineIFCMeshes.meshcolorpicker"))
        except: self.meshcolorpicker.SelectedColor = TrimbleColor.ByLayer
        
        lserial = OptionsManager.GetUint("SCR_SplitCombineIFCMeshes.layerpicker", 8)
        o = self.currentProject.Concordance.Lookup(lserial) # get the object withj the saved serial number
        if o != None:   # check if we actually got an object back
            if isinstance(o.GetSite(), LayerCollection):    # check if the object is a layer
                self.layerpicker.SetSelectedSerialNumber(lserial, InputMethod(3))
            else:  # if not then set Layer zero                       
                self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:   # if not then set Layer zero               
            self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))


        self.radiochange(None, None)

    def SaveOptions(self):

        OptionsManager.SetValue("SCR_SplitCombineIFCMeshes.split", self.split.IsChecked)
        OptionsManager.SetValue("SCR_SplitCombineIFCMeshes.combine", self.combine.IsChecked)
        OptionsManager.SetValue("SCR_SplitCombineIFCMeshes.move", self.move.IsChecked)
        OptionsManager.SetValue("SCR_SplitCombineIFCMeshes.addentry", self.addentry.IsChecked)
        OptionsManager.SetValue("SCR_SplitCombineIFCMeshes.createfrom3DFaces", self.createfrom3DFaces.IsChecked)
        OptionsManager.SetValue("SCR_SplitCombineIFCMeshes.createfrompoints", self.createfrompoints.IsChecked)
        OptionsManager.SetValue("SCR_SplitCombineIFCMeshes.reversetriangulation", self.reversetriangulation.IsChecked)
        OptionsManager.SetValue("SCR_SplitCombineIFCMeshes.deletefaces", self.deletefaces.IsChecked)
        OptionsManager.SetValue("SCR_SplitCombineIFCMeshes.triggerinvisible", self.triggerinvisible.IsChecked)
        OptionsManager.SetValue("SCR_SplitCombineIFCMeshes.combinedname", self.combinedname.Text)

        OptionsManager.SetValue("SCR_SplitCombineIFCMeshes.bimtype", self.bimtype.SelectedIndex)
        OptionsManager.SetValue("SCR_SplitCombineIFCMeshes.meshcolorpicker", self.meshcolorpicker.SelectedColor.ToArgb())
        OptionsManager.SetValue("SCR_SplitCombineIFCMeshes.layerpicker", self.layerpicker.SelectedSerialNumber)


    def OkClicked(self, cmd, e):
        
        self.success.Content = ""
        self.error.Content = ""

        wv = self.currentProject [Project.FixedSerial.WorldView]


        # make sure we work with BIMEntities only
        bimlist = []
        for o in self.objs:
            if self.createfrom3DFaces.IsChecked:
                if isinstance(o, Face3D):
                    bimlist.Add(o)
            elif not self.createfrom3DFaces.IsChecked:
                if isinstance(o, BIMEntity):
                    bimlist.Add(o)
        
        savedserials = []
        for o in GlobalSelection.SelectedMembers(self.currentProject):
            savedserials.Add(o.SerialNumber)

        GlobalSelection.Clear()
        
        if not self.move.IsChecked:
            if not self.split.IsChecked:
                if not self.reversetriangulation.IsChecked:
                    try:
                        newbimsite = self.Tree.SelectedItem.Tag
                    except:
                        self.error.Content += '\nSelect a target in the Tree'
                        return

            UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
            self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)

            try:

                with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                    
                        
                    # now either split or combine
                    

                    if self.split.IsChecked:
                        savedserials = []
                        for oldbimentity in bimlist: # the user may have selected multiple bimentities which we need to loop through

                            if oldbimentity.Count > 0: # only split if there is more than one mesh inside this bimentity

                                for oldmeshinstance in oldbimentity:

                                    # it could be that the user selected the current BIMEntity as container for the new one
                                    # and thus we may have a mix of types when looping through oldbimentity
                                    if isinstance(oldmeshinstance, ShellMeshInstance):
                                        oldbimsite = oldbimentity.GetSite()
                                        newbimentity = oldbimsite.Add(clr.GetClrType(BIMEntity))
                                        if not isinstance(oldbimsite, BIMEntityCollection):
                                            newbimentity.EntityType = oldbimsite.EntityType #"IFCBUILDINGELEMENTPROXY"
                                        else:
                                            newbimentity.EntityType = "IFCPROJECT"
                                        newbimentity.Name = oldbimentity.Name + ' - split'
                                        newbimentity.BIMGuid = Guid.NewGuid()
                                        newbimentity.Layer = oldmeshinstance.Layer
                                        newbimentity.MeshColor = oldmeshinstance.Color

                                        newmeshInstance = newbimentity.Add(clr.GetClrType(ShellMeshInstance))
                                        newmeshInstance.CreateShell(0, oldmeshinstance.GetShellMeshData().SerialNumber, oldmeshinstance.Color, oldmeshinstance.GlobalTransformation)

                                        savedserials.Add(newbimentity.SerialNumber)
                                # remove the whole old BIMEntity
                                try:
                                    oldbimsite.Remove(oldbimentity.SerialNumber)
                                except:
                                    pass
                        #self.reloadtree(None, None)

                    elif self.combine.IsChecked: # combine

                        #newbimsite = bimlist[0].GetSite() # set the new position in the bim tree to where the first bimentity currently sits

                        # create a new BIMEntity in the tree which will contain all submeshes
                        newbimentity = newbimsite.Add(clr.GetClrType(BIMEntity))
                        newbimentity.EntityType = self.bimtype.SelectedItem.Content #"IFCBUILDINGELEMENTPROXY"
                        newbimentity.Name = self.combinedname.Text
                        newbimentity.BIMGuid = Guid.NewGuid()
                        newbimentity.Layer = self.layerpicker.SelectedSerialNumber
                        tt = TrimbleColor()
                        newbimentity.MeshColor = TrimbleColorCodeValues.AsImported.ToArgb() #bimlist[0].MeshColor

                        for oldbimentity in bimlist: # the user may have selected multiple bimentities which we need to loop through

                            for oldmeshinstance in oldbimentity: # reference the old meshes (meshinstances) to the new combined BIMEntity

                                # it could be that the user selected the current BIMEntity as container for the new one
                                # and thus we may have a mix of types when looping through oldbimentity
                                if isinstance(oldmeshinstance, ShellMeshInstance):

                                    newmeshInstance = newbimentity.Add(clr.GetClrType(ShellMeshInstance))
                                    newmeshInstance.CreateShell(0, oldmeshinstance.GetShellMeshData().SerialNumber, oldmeshinstance.Color, oldmeshinstance.GlobalTransformation)

                                    oldbimentity.Remove(oldmeshinstance.SerialNumber)
                            # remove the old BIMEntity from its container
                            if not oldbimentity == newbimsite:
                                oldbimentity.GetSite().Remove(oldbimentity.SerialNumber)

                        savedserials = []
                        #self.reloadtree(None, None)

                    elif self.addentry.IsChecked:

                        # create a new BIMEntity in the tree which will contain all submeshes
                        newbimentity = newbimsite.Add(clr.GetClrType(BIMEntity))
                        newbimentity.EntityType = self.bimtype.SelectedItem.Content
                        newbimentity.Name = self.combinedname.Text
                        newbimentity.BIMGuid = Guid.NewGuid()
                        newbimentity.Layer = self.layerpicker.SelectedSerialNumber
                        newbimentity.MeshColor = Color.Green.ToArgb()

                        #self.reloadtree(None, None)

                    elif self.createfrom3DFaces.IsChecked:

                        tempverticetuples = []
                        tempfacelist = []

                        ProgressBar.TBC_ProgressBar.Title = "condensing face nodes and preparing triangulation"
                        c = 0
                        for face in bimlist:
                            c += 1
                            if ProgressBar.TBC_ProgressBar.SetProgress(c * 100 // bimlist.Count):
                                break                            
                            temptuples = []
                            temptuples.Add((face.Point0.X, face.Point0.Y, face.Point0.Z))
                            temptuples.Add((face.Point1.X, face.Point1.Y, face.Point1.Z))
                            temptuples.Add((face.Point2.X, face.Point2.Y, face.Point2.Z))

                            tempfacelist.Add(3)

                            for i in range(0, temptuples.Count):
                                
                                try:
                                    tempfacelist.Add(tempverticetuples.index(temptuples[i])) # find the index of the coord tuple in the list and use its index
                                except:
                                    tempverticetuples.Add(temptuples[i])        # otherwise create a new vertice and add this new index
                                    tempfacelist.Add(tempverticetuples.Count - 1)


                        # convert the lists to the appropriate arrays
                        ifcvertices = Array[Point3D]([Point3D()] * tempverticetuples.Count)
                        ifcfacelist = Array[Int32]([Int32()] * bimlist.Count * 4) # define size of facelist array
                        ifcnormals = Array[Point3D]([Point3D()]*0)
                        
                        for i in range(0, tempverticetuples.Count):
                            t = tempverticetuples[i]
                            ifcvertices[i] = Point3D(t[0], t[1], t[2])
                        for i in range(0, tempfacelist.Count):
                            ifcfacelist[i] = tempfacelist[i]

                        self.success.Content += '\n' + str(bimlist.Count * 3) + ' - Faces Nodes condensed to'
                        self.success.Content += '\n' + str(ifcvertices.Count) + ' - Faces Nodes'

                        # get the BIM entity collection if it doesn't exist yet it is created
                        bimEntityColl = BIMEntityCollection.ProvideEntityCollection(self.currentProject, True)
                        shellMeshDataColl = ShellMeshDataCollection.ProvideShellMeshDataCollection(self.currentProject, True)

                        bimobjectEntity = newbimsite.Add(clr.GetClrType(BIMEntity))
                        bimobjectEntity.EntityType = self.bimtype.SelectedItem.Content
                        bimobjectEntity.Name = self.combinedname.Text
                        bimobjectEntity.BIMGuid = Guid.NewGuid()
                        bimobjectEntity.Mode = DisplayMode(1 + 2 + 64 + 128 + 512 + 4096) 
                        bimobjectEntity.MeshColor = self.meshcolorpicker.SelectedColor.ToArgb()
                        # MUST set layer, otherwise the properties manager will falsly show layer "0"
                        # but the object is invisible until manually changing the layer
                        bimobjectEntity.Layer = self.layerpicker.SelectedSerialNumber


                        shellmeshdata = shellMeshDataColl.AddShellMeshData(self.currentProject)
                        shellmeshdata.CreateShellMeshData(ifcvertices, ifcfacelist, ifcnormals)
                        shellmeshdata.SetVolumeCalculationShell(ifcvertices, ifcfacelist)
                        #shellmeshdata.Description = ifcshellgroupname
                        #shellmeshdata = shellMeshDataColl.TryCreateFromShell3D(None, 0)

                        meshInstance = bimobjectEntity.Add(clr.GetClrType(ShellMeshInstance))
                        meshInstance.CreateShell(0, shellmeshdata.SerialNumber, self.meshcolorpicker.SelectedColor.ToArgb(), Matrix4D())

                        if self.deletefaces.IsChecked:
                            ProgressBar.TBC_ProgressBar.Title = "deleting 3DFaces"
                            #GlobalSelection.Clear()
                            for face in bimlist:
                                face.GetSite().Remove(face.SerialNumber)

                            savedserials = []

                    elif self.createfrompoints.IsChecked:

                        rwcloudpoints = []
                        temppoints = []
                        cloudselectionids = []
                        selectedpolysegs = []
                        for sn in savedserials:   
                            o = self.currentProject.Concordance[sn]
                            if isinstance(o, CoordPoint) or isinstance(o, CadPoint):
                                temppoints.Add(o.Position)
                            # compile a list with the ID's of the selected cloud portions
                            if isinstance(o, self.pointcloudType):
                                cloudselectionids.Add(o.Integration.GetSelectedCloudId())
                                cloudintegration = o.Integration.PointCloudDatabase.Integration
                            if isinstance(o, self.lType):

                                polyseg = o.ComputePolySeg()
                                polyseg = polyseg.ToWorld()
                                
                                selectedpolysegs.Add(polyseg.Clone())

                                for p in polyseg.Point3Ds():
                                    if p.Is3D:
                                        temppoints.Add(p)

                        # cleanup point list
                        for p1 in temppoints:
                            duplicate = False
                            for p2 in rwcloudpoints:
                                if Vector3D(p1, Point3D(p2.X, p2.Y, p2.Z)).Length < 0.0001:
                                    duplicate = True
                                    break
                            if not duplicate:
                                rwcloudpoints.Add(RwPoint3D(p1.X, p1.Y, p1.Z))


                        if cloudselectionids.Count > 0:
                            ProgressBar.TBC_ProgressBar.Title = "retrieve selected Cloud-Points"
                            #cps = o.Integration.PointCloudDatabase.Integration.GetPoints(o.Integration.GetSelectedCloudId())
                            # getselectedpoints accepts a list of CloudID's, so we just give it the ID's of the selected portions
                            cps = cloudintegration.GetSelectedPoints(cloudselectionids)
                            self.error.Content += '\nselected Cloud Point: ' + str(cps.Count)
                            ProgressBar.TBC_ProgressBar.Title = "downsampling Cloud-Points"
                            
                            if cps.Count < 15000:
                                for p in cps:
                                    rwcloudpoints.Add(RwPoint3D(p.X, p.Y, p.Z))

                            else:
                                # throw that list of all the selected portions onto the sampling algorithm
                                # it returns another ID
                                cpssampleid = cloudintegration.CreateSpatiallySampledCloud(cloudselectionids, 0.025, 15000)
                                # now get the points from that specific subcloud
                                cps2 = cloudintegration.GetPoints(cpssampleid)

                                for p in cps2:
                                    rwcloudpoints.Add(RwPoint3D(p.X, p.Y, p.Z))

                            self.error.Content += '\nsampled to #: ' + str(rwcloudpoints.Count)

                        # create duplicate of rwpoint list as normal Point3D
                        cloudpoints = []
                        for p in rwcloudpoints:
                            cloudpoints.Add(Point3D(p.X, p.Y, p.Z))

                        if rwcloudpoints.Count > 2:

                            # compute triangulation with normal surface builder
                            # selected points could from a vertical wall
                            # need to transform points to x/y plane first
                            # compute a best fit plane into the selected points -> transform to x/y -> triangulate -> use triangulation for IFC object

                            ProgressBar.TBC_ProgressBar.Title = "computing best fit plane from " + str(rwcloudpoints.Count) + " points"
                            rwplane = RwPlane3D.FitPlaneTo3DPoints(rwcloudpoints)
                            centerp = Point3D(rwplane.Point.X, rwplane.Point.Y, rwplane.Point.Z)
                            nv = Vector3D(rwplane.NormalVector.X, rwplane.NormalVector.Y, rwplane.NormalVector.Z)
                            vx = nv.Clone()
                            vx.RotateAboutZ(math.pi/2)
                            vx.Horizon = 0
                            vy = vx.Clone()
                            vy.Rotate(BiVector3D(nv, math.pi/2))

                            # rotate coordinate list to x/y plane
                            rottozero = Spinor3D.ComputeRotation(vx, nv, Vector3D(1,0,0), Vector3D(0,0,1))
                            matrixtozero = Matrix4D.BuildTransformMatrix(Vector3D(cloudpoints[0]), Vector3D(cloudpoints[0], Point3D(0, 0, 0)), rottozero, Vector3D(1,1,1))
                            matrixback = Matrix4D.Inverse(matrixtozero)

                            # create a helper surface with full triangulation
                            
                            boundarypolysegs = List[PolySeg.PolySeg]()
                            for poly in selectedpolysegs:
                                tt = poly.Clone()
                                tt.Transform(matrixtozero)
                                boundarypolysegs.Add(tt.Clone())

                            # use build in function to combine the segments as much as possible, spares us to do a manual Project-Cleanup
                            ttcount = PolySeg.PolySeg.JoinTouchingPolysegs(boundarypolysegs)
                            

                            # draw the profile linestring
                            l = wv.Add(clr.GetClrType(Linestring))
                            l.Append(boundarypolysegs[0], None, False, False)

                            tempsurface = wv.Add(clr.GetClrType(Model3D))
                            tempsurface.Name = Model3D.GetUniqueName("Best-Fit Intermediate", None, wv) #make sure name is unique
                            tempsurface.MaxEdgeLength = 20
                            builder = tempsurface.GetGemBatchBuilder()
                            for p in cloudpoints:
                                builder.AddVertex(Matrix4D.TransformPoint(matrixtozero, p))
                            
                            tempsurface.AddInfluences([l])

                            builder.Construction()
                            builder.Commit()

                            boundaryCollection = ModelBoundaries.GetModelBoundaries(self.currentProject, tempsurface.SerialNumber, True)
                            boundaryCollection.Clear()
                            boundaryCollection.Add(l.SerialNumber)
                            tempsurface.InvalidateSurface()
                            tempsurface.ConstructSurface()

                            # now that we have the triangulation we use to information to create a MeshData and MeshInstance
                            ## convert the lists to the appropriate arrays
                            ifcvertices = Array[Point3D]([Point3D()] * tempsurface.NumberOfVertices)
                            ifcfacelist = Array[Int32]([Int32()] * tempsurface.NumberOfTriangles * 4) # define size of facelist array
                            ifcnormals = Array[Point3D]([Point3D()]*0)
                            #
                            for i in range(0, tempsurface.NumberOfVertices):
                                ifcvertices[i] = Matrix4D.TransformPoint(matrixback, tempsurface.GetVertexPoint(i))
                            for i in range(0, tempsurface.NumberOfTriangles):
                                ifcfacelist[(i * 4) + 0] = 3
                                ifcfacelist[(i * 4) + 1] = tempsurface.GetTriangleIVertex(i, 0) # Gets the index of the vertex at the specified triangle corner
                                ifcfacelist[(i * 4) + 2] = tempsurface.GetTriangleIVertex(i, 1) # Gets the index of the vertex at the specified triangle corner
                                ifcfacelist[(i * 4) + 3] = tempsurface.GetTriangleIVertex(i, 2) # Gets the index of the vertex at the specified triangle corner

                            
                            # get the BIM entity collection if it doesn't exist yet it is created
                            bimEntityColl = BIMEntityCollection.ProvideEntityCollection(self.currentProject, True)
                            shellMeshDataColl = ShellMeshDataCollection.ProvideShellMeshDataCollection(self.currentProject, True)
                            
                            bimobjectEntity = newbimsite.Add(clr.GetClrType(BIMEntity))
                            bimobjectEntity.EntityType = self.bimtype.SelectedItem.Content
                            bimobjectEntity.Name = self.combinedname.Text
                            bimobjectEntity.BIMGuid = Guid.NewGuid()
                            bimobjectEntity.Mode = DisplayMode(1 + 2 + 64 + 128 + 512 + 4096) 
                            bimobjectEntity.MeshColor = self.meshcolorpicker.SelectedColor.ToArgb()
                            # MUST set layer, otherwise the properties manager will falsly show layer "0"
                            # but the object is invisible until manually changing the layer
                            bimobjectEntity.Layer = self.layerpicker.SelectedSerialNumber
                            
                            
                            shellmeshdata = shellMeshDataColl.AddShellMeshData(self.currentProject)
                            shellmeshdata.CreateShellMeshData(ifcvertices, ifcfacelist, ifcnormals)
                            shellmeshdata.SetVolumeCalculationShell(ifcvertices, ifcfacelist)
                            #shellmeshdata.Description = ifcshellgroupname
                            #shellmeshdata = shellMeshDataColl.TryCreateFromShell3D(None, 0)
                            
                            meshInstance = bimobjectEntity.Add(clr.GetClrType(ShellMeshInstance))
                            meshInstance.CreateShell(0, shellmeshdata.SerialNumber, self.meshcolorpicker.SelectedColor.ToArgb(), Matrix4D())


                        osite = wv.Remove(tempsurface.SerialNumber)
                        osite = wv.Remove(l.SerialNumber)

                    elif self.reversetriangulation.IsChecked:

                        o = self.meshpicker.Entity

                        if isinstance(o, BIMEntity):
                            # only do it if we have single meshinstances
                            if o.Count == 1:
                                for orgmeshinstance in o:
                                
                                    orgshellmeshdata = orgmeshinstance.GetShellMeshData()
                                    ifcvertices = orgshellmeshdata.GetVertexList(orgmeshinstance.Transform)
                                    ifcfacelist = orgshellmeshdata.GetIndices(TriangulationContext(1)) # triangulationContext 0 - NonTriangulated; 1 - Triangulated; 2 - TriangulatedImplicit without the length entry 3

                                    newifcvertices = Array[Point3D]([Point3D()] * ifcvertices.Count)
                                    newifcfacelist = Array[Int32]([Int32()] * ifcfacelist.Count) # define size of facelist array
                                    newifcnormals = Array[Point3D]([Point3D()]*0)

                                    for i in range(0, ifcfacelist.Count, 4):
                                        newifcfacelist[i + 0] = ifcfacelist[i + 0]
                                        newifcfacelist[i + 1] = ifcfacelist[i + 1]
                                        newifcfacelist[i + 2] = ifcfacelist[i + 3]
                                        newifcfacelist[i + 3] = ifcfacelist[i + 2]
                                    for i in range(0, ifcvertices.Count):
                                        newifcvertices[i] = ifcvertices[i]

                                    orgshellmeshdata.CreateShellMeshData(newifcvertices, newifcfacelist, newifcnormals)
                                    orgshellmeshdata.SetVolumeCalculationShell(newifcvertices, newifcfacelist)

                        self.meshpickerChanged(None, None)
                        Keyboard.Focus(self.meshpicker)                        

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


        # need to have the move operation here and outside of the TransactMethodcall since it's something that can't be undone
        else: # move in BIM tree

            try:
                for oldbimentity in bimlist: # the user may have selected multiple bimentities which we need to loop through

                    newbimsite = self.Tree.SelectedItem.Tag
                    oldbimentity.Layer = newbimsite.Layer
                    oldbimsite = oldbimentity.GetSite()

                    inputok = True

                    if oldbimsite == newbimsite:
                        self.error.Content += "\nsource and target are the same"
                        inputok = False
                    if self.containsrecursive(oldbimentity, newbimsite.SerialNumber):
                        self.error.Content += "\nnot possible - can't move to"
                        self.error.Content += "\nlocation within the source itself"
                        inputok = False

                    if inputok:


                        oldbimentity.SetSite(newbimsite)
                        oldbimentity.EntityType = self.bimtype.SelectedItem.Content #newbimsite.EntityType
                        self.currentProject.Concordance.AddObserver(oldbimentity.SerialNumber, newbimsite.SerialNumber)
                        #for obs in self.currentProject.Concordance.GetObserversOf(f.SerialNumber):
                        self.currentProject.Concordance.RemoveObserver(oldbimentity.SerialNumber, oldbimsite.SerialNumber)#RemoveObserver(uint target, uint source) source is the observer

                        if self.triggerinvisible.IsChecked:
                            oldbimentity.Visible = False

                # now update the project explorer manually - haven't found a better way yet
                te = ExplorerData.TheExplorer
                
                ## find the bimnodecollection - node
                for i in te.Items.AllItems:
                    if isinstance(i, BimEntityCollectionNode):
                        bimcn = i

                bimcn.Expanded = True # must expand it, otherwise the allitems list won't be populated
                
                tt = [] # need temp list since we may run into an error while removing data from a live enumerator
                for i in bimcn.Items.AllItems:
                    if isinstance(i, BimEntityNode):
                        tt.Add(i)
                
                # now call removeitem on all bimnodes - that basically empties the bim collection node in the project explorer
                for i in tt:
                    bimcn.RemoveItem(i)
                
                # trigger repopulate of all child items from bim collection node level
                # this will recreate the items list based on the project tree database we altered with addobserver/removeobserver
                bimcn.Populate()

                # re-expand the project explorer tree
                for i in bimcn.Items.AllItems:
                    if isinstance(i, BimEntityNode) or isinstance(i, BimGroupNode):
                        i.Expanded = True

                # redraw the project explorer ui
                te.UpdateData()

                self.currentProject.TransactionManager.ResetUndoRedoStack()

                self.reloadtree(None, None)

            except Exception as e:
                tt = sys.exc_info()
                exc_type, exc_obj, exc_tb = sys.exc_info()
                # EndMark MUST be set no matter what
                # otherwise TBC won't work anymore and needs to be restarted
                #self.currentProject.TransactionManager.AddEndMark(CommandGranularity.SubStep)
                #UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
                self.error.Content += '\nan Error occurred - Result probably incomplete\n' + str(exc_type) + '\n' + str(exc_obj) + '\nLine ' + str(exc_tb.tb_lineno)

        ProgressBar.TBC_ProgressBar.Title = ""
        
        GlobalSelection.Clear()
        GlobalSelection.Items(self.currentProject).Set(savedserials)
        
        self.SaveOptions()

    def containsrecursive(self, start, sntofind):

        if isinstance(start, BIMEntity):
            if start.Contains(sntofind):
                return True
            else:
                for j in start:
                    if self.containsrecursive(j, sntofind):
                        return True

        return False

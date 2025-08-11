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
    cmdData.Key = "SCR_DTMMerge"
    cmdData.CommandName = "SCR_DTMMerge"
    cmdData.Caption = "_SCR_DTMMerge"
    cmdData.UIForm = "SCR_DTMMerge"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "DTM"
        cmdData.ShortCaption = "Merge DTM"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.09
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Merge DTM"
        cmdData.ToolTipTextFormatted = "Merge DTM"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png") 
        cmdData.ImageSmall = b
    except:
        pass

class SCR_DTMMerge(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_DTMMerge.xaml") as s:
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

        types = Array [Type] (SurfaceTypeLists.AllWithCutFillMap)    # we fill an array with TBC object types, we could combine different types
        self.surfacepicker1.FilterByEntityTypes = types    # we fill the dropdownlist by applying that types array as filter
        self.surfacepicker1.AllowNone = False              # our list shall not show an empty field
        self.surfacepicker1.SelectOnAdd = False              # our list shall not show an empty field
        self.surfacepicker1.ValueChanged += self.surfacepickerChanged
        
        self.surfacepicker2.FilterByEntityTypes = types    # we fill the dropdownlist by applying that types array as filter
        self.surfacepicker2.AllowNone = False              # our list shall not show an empty field
        self.surfacepicker2.SelectOnAdd = False              # our list shall not show an empty field
        self.surfacepicker2.ValueChanged += self.surfacepickerChanged

        self.surfacepicker3.FilterByEntityTypes = types    # we fill the dropdownlist by applying that types array as filter
        self.surfacepicker3.AllowNone = False              # our list shall not show an empty field
        self.surfacepicker3.SelectOnAdd = False              # our list shall not show an empty field

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        #self.lfp = self.lunits.Properties.Copy()
        linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation
        #self.lfp.AddSuffix = False
        self.stepwidthlabel.Content = "Step-Width [" + linearsuffix + "]"
        self.stepwidth.DistanceMin = 0.001

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

        self.surfacepickerChanged(None, None)

    def SetDefaultOptions(self):
        try:    self.surfacepicker1.SelectIndex(OptionsManager.GetInt("SCR_DTMMerge.surfacepicker1", 0))
        except: self.surfacepicker1.SelectIndex(0)
        try:    self.surfacepicker2.SelectIndex(OptionsManager.GetInt("SCR_DTMMerge.surfacepicker2", 0))
        except: self.surfacepicker2.SelectIndex(0)
        try:    self.surfacepicker3.SelectIndex(OptionsManager.GetInt("SCR_DTMMerge.surfacepicker3", 0))
        except: self.surfacepicker3.SelectIndex(0)
        
        self.stepwidth.Distance = OptionsManager.GetDouble("SCR_DTMMerge.stepwidth", 0.001)

        self.usemanualname.IsChecked = OptionsManager.GetBool("SCR_DTMMerge.usemanualname", False)
        self.manualname.Text = OptionsManager.GetString("SCR_DTMMerge.manualname", "Temp1")
        self.overwritesurface.IsChecked = OptionsManager.GetBool("SCR_DTMMerge.overwritesurface", False)

    def SaveOptions(self):
        try:    # if nothing is selected it would throw an error
            OptionsManager.SetValue("SCR_DTMMerge.surfacepicker1", self.surfacepicker1.SelectedIndex)
            OptionsManager.SetValue("SCR_DTMMerge.surfacepicker2", self.surfacepicker2.SelectedIndex)
            OptionsManager.SetValue("SCR_DTMMerge.surfacepicker3", self.surfacepicker3.SelectedIndex)
        except:
            pass
        OptionsManager.SetValue("SCR_DTMMerge.stepwidth", self.stepwidth.Distance)

        OptionsManager.SetValue("SCR_DTMMerge.usemanualname", self.usemanualname.IsChecked)
        OptionsManager.SetValue("SCR_DTMMerge.manualname", self.manualname.Text)
        OptionsManager.SetValue("SCR_DTMMerge.overwritesurface", self.overwritesurface.IsChecked)

    def surfacepickerChanged(self, sender, e):        # in case we select a new surface from the list we update the min/max textfields
        exlist = []
        exlist.Add(self.surfacepicker1.SelectedSerial)
        exlist.Add(self.surfacepicker2.SelectedSerial)
        self.surfacepicker3.SetExcludedEntities(exlist)

    def CheckBoxChanged(self, sender, e):
        if sender.Name == "usemanualname":
            if self.usemanualname.IsChecked:
                self.overwritesurface.IsChecked = False
                return
        if sender.Name == "overwritesurface":
            if self.overwritesurface.IsChecked:
                self.usemanualname.IsChecked = False
                return


    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''
        self.success.Content = ""

        if self.stepwidth.Distance >= 0.001:
            #    self.error.Content += '\nStep-Width must not be Zero'
            #    return

            self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
            UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())

            wv = self.currentProject [Project.FixedSerial.WorldView]

            try:
                # the "with" statement will unroll any changes if something go wrong
                with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                    # get the surface as object
                    surface1 = wv.Lookup(self.surfacepicker1.SelectedSerial)
                    surface2 = wv.Lookup(self.surfacepicker2.SelectedSerial)
                    if self.overwritesurface.IsChecked:
                        surface3 = wv.Lookup(self.surfacepicker3.SelectedSerial)

                    # get the boundaries
                    poly1 = surface1.Boundary.Clone()
                    poly2 = surface2.Boundary.Clone()
                    # offset the boundary of surface 2 outwards
                    if poly2.IsClockWise(): # double check if the poly runs anti-clockwise, needed for expand
                        poly2.Reverse()
                    # expand ignores the sign and always offsets to the right
                    # .Expand seems broken in 5.90, need to use Offset instead

                    poly2 = poly2.Offset(Side.Right, abs(self.stepwidth.Distance))[1]
                    poly2.Close(True)
                    # double check if poly2 is completely inside poly1, in that case we need to do it differently
                    ints = Intersections()
                    poly1.Intersect(poly2, True, ints)
                    
                    # check if all poly2 nodes are inside poly1
                    poly2insidepoly1 = False
                    for p in poly2.ToPoint3DArray():
                        if poly1.PointInPolyseg(p) == Side.In:
                            poly2insidepoly1 = True

                    #l = wv.Add(clr.GetClrType(Linestring))
                    #l.Append(poly2, None, False, False)

                    # check if all poly1 nodes are inside poly2
                    poly1insidepoly2 = False
                    for p in poly1.ToPoint3DArray():
                        if poly2.PointInPolyseg(p) == Side.In:
                            poly1insidepoly2 = True

                    if ints.Count == 0 and not poly2.Overlaps(poly1)[1]:
                        polyscross = False
                    else:
                        polyscross = True

                    #l = wv.Add(clr.GetClrType(Linestring))
                    #l.Append(poly2, None, False, False)

                    # in order to clip the data of surface 1 we need a regionbuilder
                    # shown i.e. in the offset surface macro
                    rb = RegionBuilder()
                    replaceby2 = False
                    if not polyscross: 
                        if poly2insidepoly1:
                            # if surface 2 is completely inside surface 1 than remove the whole area 2 from surface 1
                            # need to punch a hole into surface 1 - needs two boundaries
                            rb.Add(poly1)
                            rb.Add(poly2)
                        elif poly1insidepoly2:
                            # surface 2 is sourrounding and hence replacing all of surface 1
                            # need to fix this further down
                            replaceby2 = True
                        elif not poly2insidepoly1 and not poly1insidepoly2:
                            # both surfaces stand apart from each other
                            # always combine them
                            rb.Add(poly1)

                    else:
                        polysegs = List[PolySeg.PolySeg]()
                        
                        tt = poly1.Hide(poly2, Side.In)
                        s = poly1.FirstSegment
                        while s is not None:
                            if s.Visible:
                                polysegs.Add(PolySeg.PolySeg(s.Clone()))
                            s = poly1.Next(s)
                        
                        tt = poly2.Hide(poly1, Side.Out)
                        s = poly2.FirstSegment
                        while s is not None:
                            if s.Visible:
                                polysegs.Add(PolySeg.PolySeg(s.Clone()))
                            s = poly2.Next(s)
                        
                        joincount = PolySeg.PolySeg.JoinTouchingPolysegs(polysegs)

                        for p in polysegs:
                            p.Close(True)
                            rb.Add(p)
                            #l = wv.Add(clr.GetClrType(Linestring))
                            #l.Append(p, None, False, False)

                        ###if ints.Count == 2:     # cut a chunk out of surface 1
                        ###    # remove part of boundary 1 that is inside the bigger boundary 2
                        ###    poly1.Clip(poly2, Side.In)
                        ###    # remove part of boundary 2 that is inside the original boundary 1
                        ###    poly2.Clip(surface1.Boundary, Side.Out)
                        ###
                        ###    # combine the two leftover pieces
                        ###    polysegs = List[PolySeg.PolySeg]()
                        ###    polysegs.Add(poly1)
                        ###    polysegs.Add(poly2)
                        ###    # use build in function to combine the segments as much as possible, spares us to do a manual Project-Cleanup
                        ###    joincount = PolySeg.PolySeg.JoinTouchingPolysegs(polysegs)
                        ###    bndy = polysegs[0]
                        ###
                        ###if ints.Count == 4:     # in this case we split surface 1 in two regions - need two boundaries
                        ###    # remove part of boundary 1 that is inside the bigger boundary 2
                        ###    tt = poly1.Clone()

                    nRegions = rb.Build()
                    
                    if not self.overwritesurface.IsChecked:
                        
                        newSurface = wv.Add(clr.GetClrType(Model3D))
                        
                        if self.usemanualname.IsChecked:
                            newname = Model3D.GetUniqueName(self.manualname.Text, None, wv)
                            self.manualname.Text = newname
                        else:
                            newname = Model3D.GetUniqueName(surface2.Name + ' merged onto ' + surface1.Name, None, wv) #make sure name is unique

                        newSurface.Name = newname

                    else: # overwrite existing surface data

                        newSurface = surface3

                    builder = newSurface.GetGemBatchBuilder()

                    if not replaceby2:
                        # use a copy of the original surface 1 data to create a new surface
                        newgem = surface1.GemCopy
                        newgem.External = False # make all data internal
                        newgem.IsLimited = False # recompute the min/max limits 

                        newSurface.Gem = newgem

                        # clip that new surface using the regionbuilder
                        ModelBoundaries.ClipModelByRegions(newSurface, rb, False)
                        self.TrimSurface(newSurface.Gem) # from trimble offset surface sample macro
                        
                        # the builder will create a new surface from scratch
                        # that's why the data of both surface geometries needs to be added
                        
                        builder = self.addGem(builder, newSurface.GemCopy)
                    
                    # add geometry of surface 2 - from the combine surface macro
                    builder = self.addGem(builder, surface2.GemCopy)

                    builder.Construction()
                    
                    # the following is also from the combine sample macro
                    # but I commented it out, otherwise the step will not be triangulated
                    # and instead would be a void
                    
                    ## now flag all edge triangles that don't have breakline edges
                    ##nTri = builder.NumberOfTriangles
                    ##map = GemMaterialMap()
                    ##for t in range(nTri):
                    ##    for side in range(3):
                    ##        isOuter = builder.GetTriangleOuterSide(t, side)
                    ##        if not isOuter:
                    ##            continue
                    ##        (bl, external, sharp) = builder.GetTriangleBreakline(t, side)
                    ##        if bl:
                    ##            continue
                    ##        # we have edge triangle where edge is not breakline
                    ##        # put null material on triangle
                    ##        builder.AttachMaterial(0, map, t)

                    builder.Commit() 

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
        self.SaveOptions()  

    def addGem(self, builder, surfacegem):
        
        mapVertices = {}
        nVertices = surfacegem.NumberOfVertices
        nTri = surfacegem.NumberOfTriangles
        for i in range(nVertices):
            if not surfacegem.IsVertexPresent(i):
                continue
            if not surfacegem.IsVertexTriangulated(i):
                continue
            p = surfacegem.GetVertexPoint(i)
            v = builder.AddVertex(p)
            mapVertices.Add(i, v[0])
        for t in range(nTri):
            if not surfacegem.IsTriangleMaterialPresent(t):
                continue
            for side in range(3):
                isOuter = surfacegem.GetTriangleOuter(t, side)
                if not isOuter:
                    (ok, tAdj, sideAdj) = surfacegem.GetTriangleAdjacent(t, side)
                    if not surfacegem.IsTriangleMaterialPresent(tAdj):
                        isOuter = True # treat edges next to null as valid
                iVertexA = surfacegem.GetTriangleVertex(t, side)
                nextSide = side + 1
                if nextSide == 3:
                    nextSide = 0
                iVertexB = surfacegem.GetTriangleVertex(t, nextSide)
                if isOuter or iVertexA < iVertexB:
                    b = DTMSharpness.eSoft
                    if isOuter:
                        b = DTMSharpness.eSharpAndTextureBndy
                    builder.AddBreakline(Byte(b), mapVertices[iVertexA], mapVertices[iVertexB])
        
        return builder

    def TrimSurface(self, surfaceGem): # from trimble offset surface sample macro
        v = 0
        triangles = surfaceGem.GetTriangleList()
        while v < surfaceGem.NumberOfVertices:
            iTriangleLast,iCornerLast= surfaceGem.GetVertexLastTriangle(v)
            if iTriangleLast == -1:
                surfaceGem.DeleteVertex(v, True)
                v += 1
                continue
            delV = True
            iTriangle = iTriangleLast
            iCorner = iCornerLast
            while True:
                if surfaceGem.IsTriangleMaterialPresent(iTriangle):
                    delV = False
                    break
                iTriangle, iCorner = triangles.LeftTriangle(iTriangle, iCorner)
                if iTriangle == iTriangleLast:
                    break
            if delV:
                surfaceGem.DeleteVertex(v, True)
            v += 1
        # deleting a vertex will cause surface to be marked as non-constructed.
        # we set this true so surface not rebuilt.
        surfaceGem.IsConstructed = True
        surfaceGem.Compact()
        # now set all data to "internal"
        v = 0
        external = False
        while v < surfaceGem.NumberOfVertices:
            surfaceGem.SetVertex(v, surfaceGem.GetVertexType(v), external, surfaceGem.GetVertexPoint(v))
            v += 1
        # now set the triangles
        triangles.SetAllExternalFlags(external)

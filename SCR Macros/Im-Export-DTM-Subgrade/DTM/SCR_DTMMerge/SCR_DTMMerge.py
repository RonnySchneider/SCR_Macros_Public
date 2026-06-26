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
    "surfacepicker1":    0,
    "surfacepicker2":    0,
    "surfacepicker3":    0,
    "filter1":           "",
    "filter2":           "",
    "filter3":           "",
    "compositemode":     False,
    "stepwidth":         0.001,
    "usemanualname":     False,
    "manualname":        "Temp1",
    "overwritesurface":  False,
}

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

        cmdData.Version = 1.14
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

        self.compositepicker.FilterByEntityTypes = Array [Type] ([clr.GetClrType(CompositeSurface)])    # we fill the dropdownlist by applying that types array as filter
        self.compositepicker.AllowNone = False              # our list shall not show an empty field
        self.compositepicker.SelectOnAdd = False              # our list shall not show an empty field

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

        self.filter1.TextChanged += self.filterchange
        self.filter2.TextChanged += self.filterchange
        self.filter3.TextChanged += self.filterchange

        SCRExpanders.wire_pairs([
            (self.expander_manualmode, self.manualmode),
            (self.expander_compositemode, self.compositemode),
        ])
		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

        self.surfacepickerChanged(None, None)

    def CancelClicked(self, thisCmd, args):
        thisCmd.CloseUICommand()

    def Dispose(self, cmd, disposing):
        self.filter1.TextChanged -= self.filterchange
        self.filter2.TextChanged -= self.filterchange
        self.filter3.TextChanged -= self.filterchange

    def filterchange(self, sender, e):

        # compiles and applies exclude list to listbox

        if sender.Name == "filter1":
            exclude = []
            self.surfacepicker1.SetExcludedEntities(exclude)
            
            tt = self.filter1.Text.lower()
            ticklistfilter = tt.split()
            
            for i in self.surfacepicker1.EntitySerialNumbers:
                for f in ticklistfilter:
                   if not f in i.Key.Description.lower():
                        exclude.Add(i.Value)
            
            self.surfacepicker1.SetExcludedEntities(exclude)

        elif sender.Name == "filter2":
            exclude = []
            self.surfacepicker2.SetExcludedEntities(exclude)
            
            tt = self.filter2.Text.lower()
            ticklistfilter = tt.split()
            
            for i in self.surfacepicker2.EntitySerialNumbers:
                for f in ticklistfilter:
                   if not f in i.Key.Description.lower():
                        exclude.Add(i.Value)
            
            self.surfacepicker2.SetExcludedEntities(exclude)

        elif sender.Name == "filter3":
            exclude = []
            self.surfacepicker3.SetExcludedEntities(exclude)
            
            tt = self.filter3.Text.lower()
            ticklistfilter = tt.split()
            
            for i in self.surfacepicker3.EntitySerialNumbers:
                for f in ticklistfilter:
                   if not f in i.Key.Description.lower():
                        exclude.Add(i.Value)
            
            self.surfacepicker3.SetExcludedEntities(exclude)

    def swapab_Click(self, sender, e):

        try: 
            oldsurfacea = self.surfacepicker1.SelectedSerial
            oldsurfaceb = self.surfacepicker2.SelectedSerial
            self.filter1.Text, self.filter2.Text = self.filter2.Text, self.filter1.Text
            self.surfacepicker1.SelectBySerialNumber(oldsurfaceb)
            self.surfacepicker2.SelectBySerialNumber(oldsurfacea)
        except:
            pass

    def SetDefaultOptions(self):
        SCROptions.LoadMacroOptions(self, "SCR_DTMMerge", _OPTIONS)

    def SaveOptions(self):
        SCROptions.SaveMacroOptions(self, "SCR_DTMMerge", _OPTIONS)

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


    def debugmerge(self, surface1, surface2, rb1, rb2, rbouter):

        wv = self.currentProject [Project.FixedSerial.WorldView]

        tmp1 = wv.Add(clr.GetClrType(Model3D))
        #tmp1 = Model3D()
        tmp1.Name = Model3D.GetUniqueName('Debug1', None, wv)

        tmp1builder = tmp1.GetGemBatchBuilder()

        # use a copy of the original surface 1 data to create a new surface
        tmp1gem = surface1.GemCopy
        tmp1gem.External = False # make all data internal
        tmp1gem.IsLimited = False # recompute the min/max limits 
        tmp1.Gem = tmp1gem

        # clip that new surface using the regionbuilder
        ModelBoundaries.ClipModelByRegions(tmp1, rb1, True)
        self.TrimSurface(tmp1.Gem) # from trimble offset surface sample macro
        
        # the builder will create a new surface from scratch
        tmp1builder = self.addModel3D(tmp1builder, tmp1)
        tmp1builder.Construction()
        #tmp1builder.Commit()
        
        
        # add geometry of surface 2 - from the combine surface macro
        #tmp2 = wv.Add(clr.GetClrType(Model3D))
        tmp2 = wv.Add(clr.GetClrType(Model3D))
        tmp2.Name = Model3D.GetUniqueName('Debug2', None, wv)
        tmp2builder = tmp2.GetGemBatchBuilder()
        
        tmp2gem = surface2.GemCopy
        tmp2gem.External = False # make all data internal
        tmp2gem.IsLimited = False # recompute the min/max limits 
        tmp2.Gem = tmp2gem
        
        # clip that new surface using the regionbuilder
        ModelBoundaries.ClipModelByRegions(tmp2, rb2, False)
        self.TrimSurface(tmp2.Gem)
        
        # the builder is generally empty - doesn't contain data of the 'source' surface
        # fill the builder with the geometry of 1 and 2
        tmp2builder = self.addModel3D(tmp2builder, tmp1)
        tmp2builder = self.addModel3D(tmp2builder, tmp2)
        tmp2builder.Construction()
        tmp2builder.Commit() # retriangulates and commit empties builder and writes triangulation into surface

        ModelBoundaries.ClipModelByRegions(tmp2, rbouter, False) # doesn't need retriangulation
        self.TrimSurface(tmp2.Gem)

        tmp1.GetSite().Remove(tmp1.SerialNumber)

        return tmp2

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''
        self.success.Content = ""
        self.debugging = False

        if not self.manualmode.IsChecked and not self.compositemode.IsChecked:
            self.error.Content = 'select an operation first'
            return

        if self.stepwidth.Distance >= 0.001:
            #    self.error.Content += '\nStep-Width must not be Zero'
            #    return

            UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
            self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)

            wv = self.currentProject [Project.FixedSerial.WorldView]
            surfaceserials = []

            try:
                # the "with" statement will unroll any changes if something go wrong
                with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                    
                    if self.manualmode.IsChecked:
                        surfaceserials.Add(self.surfacepicker1.SelectedSerial)
                        surfaceserials.Add(self.surfacepicker2.SelectedSerial)
                    else:
                        #comp = self.currentProject.Concordance[self.compositepicker.SelectedSerial]
                        surfaceserials += self.currentProject.Concordance[self.compositepicker.SelectedSerial].SourceSurfacesSerials


                    # get the 1st surface as object
                    surface1 = self.currentProject.Concordance[surfaceserials[0]]
                    
                    # keep adding surfaces until we're at the end of the surface list
                    tt = range(1, surfaceserials.Count)
                    for i in range(1, surfaceserials.Count):

                        surface2 = self.currentProject.Concordance[surfaceserials[i]]

                        if not self.overwritesurface.IsChecked:
                            
                            newSurface = wv.Add(clr.GetClrType(Model3D))
                            
                            if self.usemanualname.IsChecked:
                                newname = Model3D.GetUniqueName(self.manualname.Text, None, wv)
                                self.manualname.Text = newname
                            else:
                                newname = Model3D.GetUniqueName(surface2.Name + ' merged onto ' + surface1.Name, None, wv) #make sure name is unique
                
                            newSurface.Name = newname
                
                        else: # overwrite existing surface data
                
                            newSurface = self.currentProject.Concordance[self.surfacepicker3.SelectedSerial]

                        if self.debugging:
                            rb1, rb2, rbouter = self.prepareclippingregions(surface1, surface2)
                            surface1 = self.debugmerge(surface1, surface2, rb1, rb2, rbouter)

                        else:
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
                                
                                builder = self.addModel3D(builder, newSurface)
                        
                            # add geometry of surface 2 - from the combine surface macro
                            builder = self.addModel3D(builder, surface2)

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
        self.SaveOptions()  

    def clippolylists(self, list1: List[PolySeg.PolySeg](), side: Side, list2: List[PolySeg.PolySeg](), cliplist: List[PolySeg.PolySeg]()):
        
        # clip all polysegs with each other
        # i.e. list1 and Side.In will "delete" what's inside polysegs from list2

        # first just clip them, if side is set to remove things from inside we can just use this result
        # if side is set to keep inside things, we need to check if clipped segments don't lie within one of the
        # other boundaries and would get visible flag there

        tmp = []
        for p1 in list1:
            tmp1 = p1.Clone()
            tt = 1
            for p2 in list2:
                tt = tmp1.Hide(p2, side) 
            
            if side == Side.In:
                for p in self.getvisiblepolysegs2d(tmp1):
                    cliplist.Add(p)
            else:
                tmp.Add(tmp1.Clone())
                tt = 1

        # now check if we've falsely cut away segments that are visible inside one of the other boundaries
        if side == Side.Out:
            for p in tmp:
                 s = p.FirstSegment
                 while s is not None:
                     s2d = self.makesegment2d(s)

                     if s2d.Visible:
                         cliplist.Add(PolySeg.PolySeg(s2d.Clone()))

                     else:
                        for p2 in list2:
                            b = p2.PointInPolyseg(s2d.BeginPoint)
                            e = p2.PointInPolyseg(s2d.EndPoint)
                            if not b == Side.Out and not e == Side.Out:
                                s2d.Visible = True

                        if s2d.Visible:
                            cliplist.Add(PolySeg.PolySeg(s2d.Clone()))

                     s = p.Next(s)
                

        return cliplist


    def prepareclippingregions(self, surface1, surface2):

        outer1, inner1 = self.surfaceboundaries(surface1)
        if self.debugging: self.drawdebugline(outer1, 'Outer')
        if self.debugging: self.drawdebugline(inner1, 'Inner')
        outer2, inner2 = self.surfaceboundaries(surface2)
        if self.debugging: self.drawdebugline(outer2, 'Outer')
        if self.debugging: self.drawdebugline(inner2, 'Inner')
                        
        rb1 = RegionBuilder()
        rb2 = RegionBuilder()
        rbouter = RegionBuilder()

        clip1 = List[PolySeg.PolySeg]()
        clip2 = List[PolySeg.PolySeg]()
        clipmerge = List[PolySeg.PolySeg]()

        offset1 = abs(0) # cut from 2
        offset2 = abs(self.stepwidth.Distance) # cut from 1
        #offset2 = abs(self.stepwidth.Distance / 2)

        outer1off = []
        outer2off = []
        outer2offneg = []

        # compute the offset polysegs only once
        for o1 in outer1:
            poly1 = o1.Clone()
            poly1 = poly1.Offset(Side.Right, offset1)[1]
            outer1off.Add(poly1)
        #self.drawdebugline(outer1off, 'Offset1')
        # compute the offset polysegs only once
        for o2 in outer2:
            poly2 = o2.Clone()
            poly2 = poly2.Offset(Side.Right, offset2)[1]
            outer2off.Add(poly2)
        #self.drawdebugline(outer2off, 'Offset2')
        # compute the offset polysegs only once
        for o2 in outer2:
            poly2 = o2.Clone()
            poly2 = poly2.Offset(Side.Left, offset2)[1]
            outer2offneg.Add(poly2)
        #self.drawdebugline(outer2offneg, 'Offset2neg1')

        # prepare the merge clipping - doesn't need to be offset
        clipmerge = self.clippolylists(outer1, Side.In, outer2, clipmerge)
        clipmerge = self.clippolylists(outer2, Side.In, outer1, clipmerge)

        # keep from outline 1 what isn't replaced by offset 2 - clip what's inside offset 2
        clip1 = self.clippolylists(outer1, Side.In, outer2off, clip1) 
        # keep from offset 2 what's inside 1 - clip what is outside 1
        clip1 = self.clippolylists(outer2off, Side.Out, outer1, clip1) 

        # keep from neg offset 2 what is inside outer 1 - clip the rest
        clip2 = self.clippolylists(outer2offneg, Side.Out, outer1, clip2) 
        # keep from outer 2 the outline, that isnt inside the outer1 
        clip2 = self.clippolylists(outer2, Side.In, outer1, clip2) 



#                # double check if poly2 is completely inside poly1, in that case we need to do it differently
#                ints = Intersections()
#                poly1.Intersect(poly2, True, ints)
#                
#                # check if all poly2 nodes are inside poly1
#                poly2insidepoly1 = True
#                for p in poly2.ToPoint3DArray():
#                    if poly1.PointInPolyseg(p) == Side.Out:
#                        poly2insidepoly1 = False
#                
#                # check if all poly1 nodes are inside poly2
#                poly1insidepoly2 = True
#                for p in poly1.ToPoint3DArray():
#                    if poly2.PointInPolyseg(p) == Side.Out:
#                        poly1insidepoly2 = False
#                
#                if ints.Count == 0 and not poly2.Overlaps(poly1)[1]:
#                    polyscross = False
#                else:
#                    polyscross = True
#            
#                # in order to clip the data of surface 1 we need a regionbuilder
#                # after surface 1 has been clipped we can add surface 2 and retriangulate
#                # shown i.e. in the offset surface macro
#                # Attempts to build as many closed regions as possible from the current set of polysegs defined in this builder
#
#                replaceby2 = False
#    
#                if not polyscross: 
#                    if poly2insidepoly1:
#                        # if surface 2 is completely inside surface 1 than remove the whole area 2 from surface 1
#                        # need to punch a hole into surface 1 - needs two boundaries
#                        rb1.Add(poly1)
#                        rb1.Add(poly2)
#                    elif poly1insidepoly2:
#                        # surface 2 is sourrounding and hence replacing all of surface 1
#                        # need to fix this further down
#                        replaceby2 = True
#                    elif not poly2insidepoly1 and not poly1insidepoly2:
#                        # both surfaces stand apart from each other
#                        # always combine them
#                        rb1.Add(poly1)
#                        rb2.Add(poly1)
#                        rb2.Add(poly2)
#                
#                else: # polys cross each other
  
    
                ### new clipping for surface 2
                ##    tt = tmp3.Hide(poly1, Side.In) # keep from 2 what isn't replaced by 1
                ##    s = tmp3.FirstSegment
                ##    while s is not None:
                ##        if s.Visible:
                ##            s2d = self.makesegment2d(s)
                ##            #clipmerge.Add(PolySeg.PolySeg(s2d.Clone()))
                ##        s = tmp3.Next(s)
                ##
                ##    #joincount = PolySeg.PolySeg.JoinTouchingPolysegs(clip1)
                ##    #joincount = PolySeg.PolySeg.JoinTouchingPolysegs(clip2)

        #for p in clip1:
        #    
        #    b = p.FirstSegment.BeginPoint
        #    e = p.FirstSegment.EndPoint
        #
        #    isInside = False
        #    for o2 in outer2:
        #        if o2.PointInPolyseg(p.FirstSegment.BeginPoint) == Side.In or o2.PointInPolyseg(p.FirstSegment.EndPoint) == Side.In:
        #            isInside = True
        #
        #    if not isInside:
        #        finalclip1.Add(p)

        #for p in clip2:
        #    
        #    b = p.FirstSegment.BeginPoint
        #    e = p.FirstSegment.EndPoint
        #
        #    isInside = False
        #    for o1 in outer1:
        #        if o1.PointInPolyseg(p.FirstSegment.BeginPoint) == Side.In or o1.PointInPolyseg(p.FirstSegment.EndPoint) == Side.In:
        #            isInside = True
        #
        #    if not isInside:
        #        finalclip2.Add(p)
                
        joincount = PolySeg.PolySeg.JoinTouchingPolysegs(clip1)
        if self.debugging: self.drawdebugline(clip1, 'Clip 1')
        joincount = PolySeg.PolySeg.JoinTouchingPolysegs(inner1)

        joincount = PolySeg.PolySeg.JoinTouchingPolysegs(clip2)
        if self.debugging: self.drawdebugline(clip2, 'Clip 2')
        joincount = PolySeg.PolySeg.JoinTouchingPolysegs(inner2)

        joincount = PolySeg.PolySeg.JoinTouchingPolysegs(clipmerge)
        if self.debugging: self.drawdebugline(clipmerge, 'Clip Merge')
            
        for p in clip1:
            rb1.Add(p)
        for p in inner1:
            rb1.Add(p)
        nRegions1 = rb1.Build()

        for p in clip2:
            rb2.Add(p)
        for p in inner2:
            rb2.Add(p)
        nRegions2 = rb2.Build()

        for p in clipmerge:
            rbouter.Add(p)
        nRegionsouter = rbouter.Build()

#        for p in polysegs2:
#            rb2.Add(p)
#        nRegions2 = rb2.Build()

        return rb1, rb2, rbouter

    def getvisiblepolysegs2d(self, poly: PolySeg.PolySeg):            
        
        vis = List[PolySeg.PolySeg]()
        
        s = poly.FirstSegment
        while s is not None:
            if s.Visible:
                s2d = self.makesegment2d(s)
                vis.Add(PolySeg.PolySeg(s2d.Clone()))
            s = poly.Next(s)

        return vis

    def surfaceboundaries(self, surface: Model3D):

        outerbounds = List[PolySeg.PolySeg]()
        innerbounds = List[PolySeg.PolySeg]()
        
        nTri = surface.NumberOfTriangles
        for t in range(nTri):
            
            if not surface.GetTriangleMaterial(t) == surface.NullMaterialIndex():
            
                for side in range(3):
                    
                    isOuter = surface.GetTriangleOuterSide(t, side)

                    ok, adjTri, adjC = surface.GetTriangleAdjacent(t, side)
                    # if adjacent trianlge is invisible we have a hole and add this one to our own boundaries as well
                    if ok and surface.GetTriangleMaterial(adjTri) == surface.NullMaterialIndex():
                        isInner = True
                    else:
                        isInner = False

                    if isOuter or isInner:
                        bi, b = surface.GetTriangleVertex(t, side)
                        if side + 1 <= 2:
                            ei, e = surface.GetTriangleVertex(t, side + 1)
                        elif side + 1 == 3:
                            ei, e = surface.GetTriangleVertex(t, 0)
                        
                        p = PolySeg.PolySeg()
                        if isOuter:
                            outerbounds.Add(p.Add(PolySeg.Segment.Line(b, e)))
                        elif isInner:
                            innerbounds.Add(p.Add(PolySeg.Segment.Line(b, e)))

        joincount = PolySeg.PolySeg.JoinTouchingPolysegs(innerbounds)
        # we could have polysegs in the innerbounds list which belong to invisible triangles between islands
        # those won't be closed and belong to the outerbounds list
        finalinnerbounds = List[PolySeg.PolySeg]()
        for p in innerbounds:
            if p.IsClosed:
                if p.IsClockWise(): # double check if the poly runs counter-clockwise, needed for expand
                    p.Reverse()
                finalinnerbounds.Add(p)
            else:
                outerbounds.Add(p)

        joincount = PolySeg.PolySeg.JoinTouchingPolysegs(outerbounds)
        finalouterbounds = List[PolySeg.PolySeg]()
        for p in outerbounds:
            if p.IsClosed:
                if p.IsClockWise(): # double check if the poly runs counter-clockwise, needed for expand
                    p.Reverse()
                finalouterbounds.Add(p)
         
        return finalouterbounds, finalinnerbounds

    def makesegment2d(self, s: PolySeg.Segment):

        b = s.BeginPoint
        b.To2D()
        e = s.EndPoint
        e.To2D()

        news = s.GetType()(b, e)
        news.Visible = s.Visible
        tt = 2

        return news

    # testing intellisense declaration
    def drawdebugline(self, polyseg: List[PolySeg.PolySeg], layer: str):

        for p in polyseg:
            if isinstance(p, PolySeg.PolySeg):
                wv = self.currentProject [Project.FixedSerial.WorldView]
                l = wv.Add(clr.GetClrType(Linestring))
                l.Append(p, None, False, False)
                l.Layer = Layer.FindOrCreateLayer(self.currentProject, layer).SerialNumber



    def addModel3D(self, builder, surface: Model3D):
        
        mapVertices = {}
        nVertices = surface.NumberOfVertices
        nTri = surface.NumberOfTriangles
        for i in range(nVertices):
            if not surface.IsVertexPresent(i):
                continue
            if not surface.Gem.IsVertexTriangulated(i):
                continue
            p = surface.GetVertexPoint(i)
            v = builder.AddVertex(p)
            mapVertices.Add(i, v[0])
        for t in range(nTri):
            if not surface.IsTriangleMaterialPresent(t):
                continue
            for side in range(3):
                isOuter = surface.GetTriangleOuterSide(t, side)
                if not isOuter:
                    (ok, tAdj, sideAdj) = surface.GetTriangleAdjacent(t, side)
                    if not surface.IsTriangleMaterialPresent(tAdj):
                        isOuter = True # treat edges next to null as valid
                iVertexA = surface.GetTriangleIVertex(t, side)
                nextSide = side + 1
                if nextSide == 3:
                    nextSide = 0
                iVertexB = surface.GetTriangleIVertex(t, nextSide)
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

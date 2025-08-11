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
    cmdData.Key = "SCR_DTMShot"
    cmdData.CommandName = "SCR_DTMShot"
    cmdData.Caption = "_SCR_DTMShot"
    cmdData.UIForm = "SCR_DTMShot"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "DTM"
        cmdData.ShortCaption = "DTM Shot"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.05
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "intersect a DTM with a 2-Point Vector"
        cmdData.ToolTipTextFormatted = "intersect a DTM with a 2-Point Vector"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_DTMShot(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_DTMShot.xaml") as s:
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
        wv = self.currentProject [Project.FixedSerial.WorldView]

        self.ifcs.IsEntityValidCallback = self.IsValidIFC

        types = Array[Type](SurfaceTypeLists.AllWithCutFillMap)+Array[Type]([clr.GetClrType(ProjectedSurface)])
        #types.extend (Array[Type]([clr.GetClrType(ProjectedSurface)]))
        self.surfacepicker.FilterByEntityTypes = types
        self.surfacepicker.AllowNone = False

        self.coordpick1.ShowElevationIf3D = True
        self.coordpick2.ShowElevationIf3D = True
        self.coordpick2.ValueChanged += self.CoordPick2Changed
        self.coordpick2.AutoTab = False

        self.lType = clr.GetClrType(Linestring) # don't allow 2D CadLines
        self.linepicker1.ValueChanged += self.lineChanged
        self.linepicker1.AutoTab = False
        self.linepicker1.IsEntityValidCallback = self.IsValidLine

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def CoordPick2Changed(self, ctrl, e):
        if e.Cause == InputMethod.Mouse:     
            self.OkClicked(None, None)

    def IsValidLine(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        return False

    def IsValidIFC(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, BIMEntity):
            return True
        if isinstance(o, Shell3D):
            return True
        return False

    def lineChanged(self, ctrl, e):
        self.OkClicked(None, None)

    def SetDefaultOptions(self):
        self.shootifc.IsChecked = OptionsManager.GetBool("SCR_DTMShot.shootifc", True)
        self.shootdtm.IsChecked = OptionsManager.GetBool("SCR_DTMShot.shootdtm", False)

        lserial = OptionsManager.GetUint("SCR_DTMShot.layerpicker", 8)
        o = self.currentProject.Concordance.Lookup(lserial)
        if o != None:
            if isinstance(o.GetSite(), LayerCollection):    
                self.layerpicker.SetSelectedSerialNumber(lserial, InputMethod(3))
            else:                       
                self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:                       
            self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))

        # Select surface
        try:    self.surfacepicker.SelectIndex(OptionsManager.GetInt("SCR_DTMShot.surfacepicker", 0))
        except: self.surfacepicker.SelectIndex(0)

        self.usepoints.IsChecked = OptionsManager.GetBool("SCR_DTMShot.usepoints", True)
        self.extendline.IsChecked = OptionsManager.GetBool("SCR_DTMShot.extendline", False)
        self.breakline.IsChecked = OptionsManager.GetBool("SCR_DTMShot.breakline", False)
        self.drawlinepoint.IsChecked = OptionsManager.GetBool("SCR_DTMShot.drawlinepoint", False)

    def SaveOptions(self):
        OptionsManager.SetValue("SCR_DTMShot.shootifc", self.shootifc.IsChecked)      
        OptionsManager.SetValue("SCR_DTMShot.shootdtm", self.shootdtm.IsChecked)      
        OptionsManager.SetValue("SCR_DTMShot.layerpicker", self.layerpicker.SelectedSerialNumber)
        OptionsManager.SetValue("SCR_DTMShot.surfacepicker", self.surfacepicker.SelectedIndex)
        OptionsManager.SetValue("SCR_DTMShot.usepoints", self.usepoints.IsChecked)      
        OptionsManager.SetValue("SCR_DTMShot.extendline", self.extendline.IsChecked)      
        OptionsManager.SetValue("SCR_DTMShot.breakline", self.breakline.IsChecked)      
        OptionsManager.SetValue("SCR_DTMShot.drawlinepoint", self.drawlinepoint.IsChecked)      

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''

        self.success.Content = ''

        wv = self.currentProject [Project.FixedSerial.WorldView]
        wv.PauseGraphicsCache(True)

        inputok=True

        if inputok:
            self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
            UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
            try:
                with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                    
                    dtmintersection = None
                    vertexlist = self.createvertexlist()

                    if self.usepoints.IsChecked:
                        p1 = self.coordpick1.Coordinate
                        p2 = self.coordpick2.Coordinate
                        shot = Vector3D(p1, p2)
                        
                        #if isinstance(surface, ProjectedSurface):

                        dtmintersection = self.intersectdtm(vertexlist, p1, p2, False)

                        #else: # standard surface is easy
                        #    tiepoint = clr.StrongBox[Point3D]()     # we create us the variable the ComputeTie wants for the output
                        #    if surface.ComputeTie(p1, shot, math.pi/2 - (shot.Horizon), 10000, tiepoint): # we compute the surface intersection
                        #        dtmintersection = tiepoint.Value

                        if dtmintersection:
                            cadPoint = wv.Add(clr.GetClrType(CadPoint))
                            cadPoint.Layer = self.layerpicker.SelectedSerialNumber
                            cadPoint.Point0 = dtmintersection

                    else: # extend line
                        l1 = self.linepicker1.Entity
                        if l1:
                            polyseg = l1.ComputePolySeg()
                            polyseg = polyseg.ToWorld()
                            polyseg_v = l1.ComputeVerticalPolySeg()
                            polyseg = polyseg.Linearize(0.0001, 0.0001, 1000, polyseg_v, False)

                            if self.extendline.IsChecked and not self.breakline.IsChecked:
                                found, pout, pstation = polyseg.FindPointFromPoint(self.linepicker1.PickPointProjected)
                                if pstation < (polyseg.ComputeStationing() - pstation):
                                    pickedstartofline = True
                                    p1 = polyseg.FirstSegment.EndPoint
                                    p2 = polyseg.FirstSegment.BeginPoint
                                    #shot = Vector3D(p1, p2)
                                else:
                                    pickedstartofline = False
                                    #done, el, v = polyseg_v.ComputeVerticalSlopeAndGrade(polyseg.ComputeStationing())
                                    p1 = polyseg.LastSegment.BeginPoint
                                    p2 = polyseg.LastSegment.EndPoint
                                    #shot = Vector3D(p1, p2)

                                #if isinstance(surface, ProjectedSurface):

                                dtmintersection = self.intersectdtm(vertexlist, p1, p2, False)

                                #else: # standard surface is easy
                                #    tiepoint = clr.StrongBox[Point3D]()     # we create us the variable the ComputeTie wants for the output
                                #    # slope in Computetie is zenith angle with upwards=0
                                #    # Vector3D.Horizon is positive above the horizon and negative below
                                #    if surface.ComputeTie(p1, shot, math.pi/2 - (shot.Horizon), 10000, tiepoint): # we compute the surface intersection
                                #        dtmintersection = tiepoint.Value
        
                                if dtmintersection:

                                    if pickedstartofline:
                                        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                        e.Position = dtmintersection
                                        # if the old segment is longer than the distance to the new intersection
                                        if Vector3D(p1, p2).Length > Vector3D(p1, dtmintersection).Length:
                                            # replace as new start
                                            l1.ReplaceElementAt(e, 0)
                                        else:
                                            # add as new start
                                            l1.InsertElementAt(e, 0)
                                    else:
                                        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                        e.Position = dtmintersection
                                        # if the old segment is longer than the distance to the new intersection
                                        if Vector3D(p1, p2).Length > Vector3D(p1, dtmintersection).Length:
                                            # need to replace last element
                                            l1.ReplaceElementAt(e, l1.ElementCount - 1)
                                        else:
                                            # add extra element
                                            l1.AppendElement(e)

                                    if self.drawlinepoint.IsChecked:
                                        cadPoint = wv.Add(clr.GetClrType(CadPoint))
                                        cadPoint.Layer = self.layerpicker.SelectedSerialNumber
                                        cadPoint.Point0 = dtmintersection

                            if self.extendline.IsChecked and self.breakline.IsChecked:
                                s = polyseg.FirstSegment
                                while s is not None:
                                    dtmintersection = None
                                    if s.Visible:
                                        dtmintersection = self.intersectdtm(vertexlist, s.BeginPoint, s.EndPoint, True)
                                    
                                           
                                    if dtmintersection:
                                        pch = polyseg.FindPointFromPoint(dtmintersection)
                                        if not pch[2] == 0:
                                            if self.drawlinepoint.IsChecked:
                                                cadPoint = wv.Add(clr.GetClrType(CadPoint))
                                                cadPoint.Layer = self.layerpicker.SelectedSerialNumber
                                                cadPoint.Point0 = dtmintersection
                                            
                                            l1 = l1.BreakAtStation(pch[2]) # result is a new linestring after the break point                            
                                            # compute a new polyseg for the rest of the original linestring
                                            # and repeat the search for segments which intersect the surface
                                            polyseg = l1.ComputePolySeg()
                                            polyseg = polyseg.ToWorld()
                                            polyseg_v = l1.ComputeVerticalPolySeg()
                                            polyseg = polyseg.Linearize(0.0001, 0.0001, 1, polyseg_v, False)
                                            s = polyseg.FirstSegment
                                        else:
                                            s = polyseg.Next(s)
                                    s = polyseg.Next(s)

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

        self.success.Content += '\nDone'
        self.SaveOptions()           
        
        if self.usepoints.IsChecked:
            Keyboard.Focus(self.coordpick1)
        else:
            Keyboard.Focus(self.linepicker1)
        
        wv.PauseGraphicsCache(False)
    
    def PointInsideTriangle3D(self, vertex1, p2, p3, p4):
        # see also https://blackpawn.com/texts/pointinpoly/
    
        a = Vector3D(p2.X-vertex1.X, p2.Y-vertex1.Y, p2.Z-vertex1.Z)     # v1
        b = Vector3D(p3.X-vertex1.X, p3.Y-vertex1.Y, p3.Z-vertex1.Z)     # v0
        w = Vector3D(p4.X-vertex1.X, p4.Y-vertex1.Y, p4.Z-vertex1.Z)     # v2
    
        aa = Vector3D.DotProduct(a,a)[0]  # dot11
        ab = Vector3D.DotProduct(a,b)[0]  # dot10
        bb = Vector3D.DotProduct(b,b)[0]  # dot00
        wa = Vector3D.DotProduct(w,a)[0]  # dot21
        wb = Vector3D.DotProduct(w,b)[0]  # dot20
        
        # dot10 * dot10 - dot11 * dot00
        d = ab * ab - aa * bb
        if d == 0:
            # in case the three triangle vertices are in one line we can't compute it and ignore that "triangle"
            inside = False
        else:
            inside = True
                    # v = dot10 * dot20 - dot00 * dot21
            s = round((ab * wb - bb * wa) / d, 6) # rounding that value a bit down gives us better results when very close to the triangle side, what we are
            if s < 0 or s > 1:                    # otherwise we might miss a value
                inside = False
                    # u = dot10 * dot21 - dot11 * dot20
            t = round((ab * wa - aa * wb) / d, 6)
            if t < 0 or (s + t) > 1:
               inside = False
    
        return inside

    def intersectdtm(self, vertexlist, p1, p2, isectonsegmentonly):
        # isectonsegmentonly - is a bool value; for breaking lines we only want the locations when the intersection is between t=0 and t=1
        
        shot = Vector3D(p1, p2)

        # setup the line
        seg1 = SegmentLine(p1, p2)

        # prepare variables
        out_t = clr.StrongBox[float]()
        outPointOnCL = clr.StrongBox[Point3D]()
        testside = clr.StrongBox[Side]()

        shortest_intersection = None

        for i in range(0, vertexlist.Count, 3):
        
            p = Plane3D(vertexlist[i], vertexlist[i+1], vertexlist[i+2])[0] # the plane is returned as first element
        
            if not p.IsValid:
                continue
        
            pnew = Plane3D.IntersectWithRay(p, p1, shot)
        
            # we only want results were the intersecting point is within the tested triangle
            # otherwise we get hundreds of false results which don't lie on the DTM
        
            #if Triangle2D.IsPointInside(v1,v2,v3,pnew)[0] == True:
            ### if PointInsideTriangle3D(v1,v2,v3,pnew):
            if pnew != Point3D.Undefined and self.PointInsideTriangle3D(vertexlist[i], vertexlist[i+1], vertexlist[i+2], pnew):
                if not Point3D.IsDuplicate(pnew, p2, 0.000001)[0]:
                    # project the point - only if it's in front of us we want it
                    if seg1.ProjectPoint(pnew, out_t, outPointOnCL, testside):
                        if not isectonsegmentonly and out_t.Value > 0.0:
                            if not shortest_intersection:
                                shortest_intersection = pnew
                            else:
                                if Vector3D(p1, pnew).Length < Vector3D(p1, shortest_intersection).Length:
                                    shortest_intersection = pnew

                        if isectonsegmentonly and out_t.Value > 0.0 and out_t.Value <= 1.0:
                            if not shortest_intersection:
                                shortest_intersection = pnew
                            else:
                                if Vector3D(p1, pnew).Length < Vector3D(p1, shortest_intersection).Length:
                                    shortest_intersection = pnew

                    
        #tt1 = Vector3D(p1, shortest_intersection).Length
        #tt2 = Vector3D(p2, shortest_intersection).Length
                                    
        if not isectonsegmentonly:
            if shortest_intersection and Vector3D(p1, shortest_intersection).Length > 0 and Vector3D(p2, shortest_intersection).Length > 0:
                return shortest_intersection
            else:
                return None
        else:
            if shortest_intersection and Vector3D(p1, shortest_intersection).Length > 0 and Vector3D(p1, shortest_intersection).Length <= Vector3D(p1, p2).Length:
                return shortest_intersection
            else:
                return None
        
    def createvertexlist(self):
        # create a list of triangle vertices
        vertexlist = []
        if self.shootdtm.IsChecked:

            surface = self.currentProject.Concordance.Lookup(self.surfacepicker.SelectedSerial)
            nTri = surface.NumberOfTriangles

            if isinstance(surface,ProjectedSurface):
                projected=True
            else:
                projected=False
                
            for i in range(nTri):
                if surface.GetTriangleMaterial(i) == surface.NullMaterialIndex(): continue
                if projected==True:
                    vertexlist.Add(surface.TransformPointToWorldDelegate(surface.GetVertexPoint(surface.GetTriangleIVertex(i,0))))
                    vertexlist.Add(surface.TransformPointToWorldDelegate(surface.GetVertexPoint(surface.GetTriangleIVertex(i,1))))
                    vertexlist.Add(surface.TransformPointToWorldDelegate(surface.GetVertexPoint(surface.GetTriangleIVertex(i,2))))
                else:
                    vertexlist.Add(surface.GetVertexPoint(surface.GetTriangleIVertex(i,0)))
                    vertexlist.Add(surface.GetVertexPoint(surface.GetTriangleIVertex(i,1)))
                    vertexlist.Add(surface.GetVertexPoint(surface.GetTriangleIVertex(i,2)))
        else: # if we use IFCs

            for o in self.ifcs.SelectedMembers(self.currentProject):
                # o = self.currentProject.Concordance.Lookup(sn)
                verticesGlobal = []
                
                # create Point3D List of vertices, not in any order yet
                if  isinstance(o, Shell3D): # in case it is an IFC Mesh we get us the coordinates
                    try: #2023.11
                        vertexIndices = o.GetTriangulatedFaceList() # this works with self created 3DShells - i.e. SweepShape, or Linebundle
                        verticesLocal = o.GetVertex() # vertices as Point3Ds
                        for i in range(0, vertexIndices.Count, 4):
                            verticesGlobal.Add(o.GlobalTransformation.TransformPoint(verticesLocal[vertexIndices[i + 1]]))
                            verticesGlobal.Add(o.GlobalTransformation.TransformPoint(verticesLocal[vertexIndices[i + 2]]))
                            verticesGlobal.Add(o.GlobalTransformation.TransformPoint(verticesLocal[vertexIndices[i + 3]]))

                    except: # 2024.00
                        tt = o.GetTrianglesForInspection()
                        for t in tt:
                            verticesGlobal.Add(t.pointA)
                            verticesGlobal.Add(t.pointB)
                            verticesGlobal.Add(t.pointC)

                elif isinstance(o, BIMEntity):
                    verticesGlobal = []
                    for shellMeshInstance in o.GetGeometry():
                        shellMeshData = shellMeshInstance.GetShellMeshData()
                        
                        try: #2023.11
                            # DEPENDING ON THE TYPE OF IFC THE DIFFERENT METHODS RETURN EMPTY LISTS
                            vertexIndices = shellMeshData.GetTriangulatedFaceList() # this works for the bridge IFC
                            if vertexIndices.Count == 0:
                                vertexIndices = shellMeshData.GetFaces()    # this works for the geotech, but not the bridges
                                verticesLocal = shellMeshData.GetVertex() # vertices as Point3Ds

                            for i in range(0, vertexIndices.Count, 4):
                                verticesGlobal.Add(o.GlobalTransformation.TransformPoint(verticesLocal[vertexIndices[i + 1]]))
                                verticesGlobal.Add(o.GlobalTransformation.TransformPoint(verticesLocal[vertexIndices[i + 2]]))
                                verticesGlobal.Add(o.GlobalTransformation.TransformPoint(verticesLocal[vertexIndices[i + 3]]))
                        
                        except: # 2024.00
                            tt = shellMeshData.GetTrianglesForInspectionInternal(shellMeshInstance.GlobalTransformation)
                            for t in tt:
                                verticesGlobal.Add(t.pointA)
                                verticesGlobal.Add(t.pointB)
                                verticesGlobal.Add(t.pointC)

                for i in range(0, verticesGlobal.Count, 3):

                    # it can be that the IFC contains "triangles" where all three points are on one line, we don't want those
                    # that would lead to a division by zero in the algorithm that checks if the perpendicular solution is within the triangle
                    # checking it here is doubling up some computations, but it could also mess up the plane and normal vector creation
                    vertex1 = verticesGlobal[i+0]
                    vertex2 = verticesGlobal[i+1]
                    vertex3 = verticesGlobal[i+2]

                    a = Vector3D(vertex1, vertex2)
                    b = Vector3D(vertex1, vertex3)
                    
                    aa = Vector3D.DotProduct(a,a)[0]
                    ab = Vector3D.DotProduct(a,b)[0]
                    bb = Vector3D.DotProduct(b,b)[0]
                    
                    d = round(ab * ab - aa * bb, 14) # otherwise it could still happen that "triangles" slip through

                    if d != 0:
                        vertexlist.Add(vertex1)
                        vertexlist.Add(vertex2)
                        vertexlist.Add(vertex3)

        return vertexlist
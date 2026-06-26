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
    "layerpicker":   0,
    "surfacepicker": 0,
    "dropgrid":      False,
    "gridinterval":  1.0,
    "multipoint":    False,
    "singlepoint":   True,
    "labelonly":     False,
    "coordpick":     0,
}

def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_Rainfall"
    cmdData.CommandName = "SCR_Rainfall"
    cmdData.Caption = "_SCR_Rainfall"
    cmdData.UIForm = "SCR_Rainfall"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "DTM"
        cmdData.ShortCaption = "Rainfall"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.15
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Rainfall"
        cmdData.ToolTipTextFormatted = "follow a Raindrop on the DTM"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png") 
        cmdData.ImageSmall = b
    except:
        pass


class SCR_Rainfall(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_Rainfall.xaml") as s:
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
        self.surfacepicker.FilterByEntityTypes = types    # we fill the dropdownlist by applying that types array as filter
        self.surfacepicker.AllowNone = False              # our list shall not show an empty field

        self.objs.IsEntityValidCallback = self.IsValidPoints
        optionMenu = SelectionContextMenuHandler()
        # remove options that don't apply here
        optionMenu.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        self.objs.ButtonContextMenu = optionMenu
        self.point3dType = clr.GetClrType(Point3D)
        self.coordpointType = clr.GetClrType(CoordPoint)
        self.cadpointType = clr.GetClrType(CadPoint)

        self.coordpick.ShowElevationIf3D = True
        self.coordpick.ValueChanged += self.CoordPickChanged
        self.coordpick.AutoTab = False

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        self.lfp = self.lunits.Properties.Copy()
        self.linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation

        self.labelinterval.Content = 'Rain-Grid Interval [' + self.linearsuffix + ']'

        self.gridinterval.DistanceMin = 0.001
        self.gridinterval.Distance = 1.0

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def CoordPickChanged(self, ctrl, e):
        if e.Cause == InputMethod.Mouse:     
            self.OkClicked(None, None)

    def IsValidPoints(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.coordpointType):
            return True
        if isinstance(o, self.cadpointType):
            return True
        return False

    def SetDefaultOptions(self):
        SCROptions.LoadMacroOptions(self, "SCR_Rainfall", _OPTIONS, self.currentProject)

    def SaveOptions(self):
        SCROptions.SaveMacroOptions(self, "SCR_Rainfall", _OPTIONS)

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''
        self.success.Content = ""

        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)

        wv = self.currentProject [Project.FixedSerial.WorldView]
        wv.PauseGraphicsCache(True)

        try:
            # the "with" statement will unroll any changes if something go wrong
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                surface = wv.Lookup(self.surfacepicker.SelectedSerial)
                nt = surface.NumberOfTriangles
                nv = surface.GetNumberOfTriangulatedVertices()
                # precompute once — avoids repeated TBC API calls inside the hot trace loop
                self.trivertices  = [(surface.GetTriangleIVertex(i,0), surface.GetTriangleIVertex(i,1), surface.GetTriangleIVertex(i,2)) for i in range(nt)]
                self.vertexpoints = [surface.GetVertexPoint(i) for i in range(nv)]
                self.triplane     = [None] * nt
                self.nullmat      = surface.NullMaterialIndex()
                self.trimaterial  = [surface.GetTriangleMaterial(i) for i in range(nt)]

                vertextrirelations = [[] for i in range(nv)]
                for i in range(nt):
                    for v in self.trivertices[i]:
                        vertextrirelations[v].append(i)

                vertexrelations = [None] * nv
                deadends = []

                if self.labelonly.IsChecked:
                    # for debugging - label surface vertices and triangles
                    
                    # label the triangles
                    for i in range(surface.NumberOfTriangles):
                        if surface.IsTriangleMaterialPresent(i):
                            # get the triangle vertices
                            verticelist = []
                            if surface.GetTriangleMaterial(i) != surface.NullMaterialIndex():
                                verticelist.Add(surface.GetVertexPoint(surface.GetTriangleIVertex(i,0)))
                                verticelist.Add(surface.GetVertexPoint(surface.GetTriangleIVertex(i,1)))
                                verticelist.Add(surface.GetVertexPoint(surface.GetTriangleIVertex(i,2)))
                    
                            if verticelist.Count == 3:
                                # label the triangle centers
                                p = Triangle3D.CenterOfMass(verticelist[0], verticelist[1], verticelist[2])
                                t = wv.Add(clr.GetClrType(MText))
                                t.AlignmentPoint = p[0]
                                t.AttachPoint = AttachmentPoint.MiddleMid
                                t.Color = Color.Red
                                t.Height = 0.2
                                t.Layer = self.layerpicker.SelectedSerialNumber
                                t.TextString = str(i)
                
                                # label the triangle corners
                                for j in range(0, 3):
                                    v = Vector3D(verticelist[j], p[0])
                                    v.Length = 1
                                    t = wv.Add(clr.GetClrType(MText))
                                    t.AlignmentPoint = verticelist[j] + v
                                    t.AttachPoint = AttachmentPoint.MiddleMid
                                    t.Color = Color.Magenta
                                    t.Height = 0.2
                                    t.Layer = self.layerpicker.SelectedSerialNumber
                                    t.TextString = str(j)
                                
                                # label all vertices with the overall number
                                for j in range(0, 3):
                                    t = wv.Add(clr.GetClrType(MText))
                                    t.AlignmentPoint = verticelist[j]
                                    t.AttachPoint = AttachmentPoint.BottomLeft
                                    t.Color = Color.Blue
                                    t.Height = 0.2
                                    t.Layer = self.layerpicker.SelectedSerialNumber
                                    t.TextString = str(surface.GetTriangleIVertex(i,j))

                else:

                    pointlist = []

                    if self.multipoint.IsChecked and self.objs.SelectedMembers(self.currentProject).Count > 0:
                        for o in self.objs.SelectedMembers(self.currentProject): # go through all selected points
                            if isinstance(o, CoordPoint) or isinstance(o, self.cadpointType):
                                pointlist.Add(o.Position)

                    if self.singlepoint.IsChecked:
                        pointlist.Add(self.coordpick.Coordinate)

                    if self.dropgrid.IsChecked:
                        s_bndy_seg = surface.Boundary
                        s_box = s_bndy_seg.BoundingBox
                        min_x = s_box.ptMin.X
                        max_x = s_box.ptMax.X
                        min_y = s_box.ptMin.Y
                        max_y = s_box.ptMax.Y
                        curr_x = min_x
                        curr_y = min_y
                        while curr_y < max_y:
                            while curr_x < max_x:
                                testp = Point3D(curr_x, curr_y,0)
                                if s_bndy_seg.PointInPolyseg(testp):
                                    pointlist.Add(testp)
                                curr_x += self.gridinterval.Distance
                            curr_x = min_x
                            curr_y += self.gridinterval.Distance

                    
                    if pointlist.Count > 0 and not isinstance(surface, ProjectedSurface):
                        ProgressBar.TBC_ProgressBar.Title = "raining " + str(pointlist.Count) + " drops"
                        self.time1 = datetime.now()

                        for i in range(pointlist.Count):
                            p = pointlist[i]

                            if (datetime.now() - self.time1).seconds > 0.5:
                                if ProgressBar.TBC_ProgressBar.SetProgress(i * 100 // pointlist.Count):
                                    break   # function returns true if user pressed cancel
                                self.time1 = datetime.now()

                            # land on the surface for the first time
                            outPoint = clr.StrongBox[Point3D]()
                            outPrimitive = clr.StrongBox[Primitive]()
                            outInt = clr.StrongBox[Int32]()
                            outByte = clr.StrongBox[Byte]()
                            visitedtri = []
                            tracevert = []

                            l = None
                            self.debugiterator = 0
                            self.recursionlimiterror = False

                            if surface.PickSurface(p, outPrimitive, outInt, outByte, outPoint):
                                if outPrimitive.Value == Primitive.Point:
                                    giv = surface.GetTriangleIVertex(outInt.Value, outByte.Value)
                                    if not l:
                                        l = self.newtraceline(outPoint.Value, None)
                                    tracevert = []
                                    deadends, tracevert, vertexrelations = self.followterrain(surface, outPoint.Value, giv, None, None, l, deadends, tracevert, vertexrelations, vertextrirelations)

                                if outPrimitive.Value == Primitive.Triangle:
                                    if surface.GetTriangleMaterial(outInt.Value) != surface.NullMaterialIndex():
                                        if not l:
                                            l = self.newtraceline(outPoint.Value, None)
                                        deadends, tracevert, vertexrelations = self.followterrain(surface, outPoint.Value, None, outInt.Value, None, l, deadends, tracevert, vertexrelations, vertextrirelations)

                tt = self.debugiterator
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
        wv.PauseGraphicsCache(False)

        if self.singlepoint.IsChecked:
            Keyboard.Focus(self.coordpick)


    def followterrain(self, surface, raindrop, giv, tri, in_side, droptrace, deadends, tracevert, vertexrelations, vertextrirelations):
        # iterative replacement for the former recursive implementation
        # every branch either returns or tail-calls itself, so we just loop and update locals
        while True:
            self.debugiterator += 1
            if self.debugiterator > 100000:
                self.recursionlimiterror = True
                return deadends, tracevert, vertexrelations
            # --- corner/vertex case ---
            if giv != None and tri == None:
                if giv in tracevert:
                    deadends.Add(giv)
                    return deadends, tracevert, vertexrelations

                tracevert.Add(giv)

                if vertexrelations[giv] == None:
                    vertexrelations = self.connectedvertices(surface, giv, vertexrelations, vertextrirelations)
                if vertexrelations[giv] == None:
                    deadends.Add(giv)
                    return deadends, tracevert, vertexrelations

                p, nextindex, in_side, nextisvertex = self.decideatcorner(surface, giv, vertexrelations, vertextrirelations)
                droptrace = self.addtotraceline(droptrace, p)

                if nextisvertex:
                    raindrop = p; giv = nextindex; tri = None; in_side = None
                elif nextindex != None:
                    raindrop = p; giv = None; tri = nextindex  # in_side already set by decideatcorner
                else:
                    return deadends, tracevert, vertexrelations
                continue

            # --- triangle interior case ---
            if tri != None and giv == None:
                tv = self.trivertices[tri]
                vp = self.vertexpoints
                verticelist = (vp[tv[0]], vp[tv[1]], vp[tv[2]])

                plane = self.triplane[tri]
                if plane is None:
                    plane = Plane3D(verticelist[0], verticelist[1], verticelist[2])[0]
                    self.triplane[tri] = plane

                if plane.normal.Length2D != 0: # if plane.Slope != 0: is too inaccurate for super thin triangles
                    out_side, in_side, ip = self.followtriangleslope(raindrop, tri, in_side, verticelist, plane)

                    if out_side != None:
                        droptrace = self.addtotraceline(droptrace, ip)

                        isOuter = surface.GetTriangleOuterSide(tri, out_side)
                        if not isOuter:
                            (b, nexttri, in_side) = surface.GetTriangleAdjacent(tri, out_side)
                            if self.trimaterial[nexttri] == self.nullmat:
                                isOuter = True

                        if isOuter:
                            # outer edge - follow down to the lower of the two boundary vertices
                            v0 = out_side  # v0/out_side is the same as the vertex index inside the triangle
                            v1 = v0 + 1
                            if v1 == 3: v1 = 0
                            if self.vectorslope(ip, verticelist[v0]) < self.vectorslope(ip, verticelist[v1]):
                                droptrace = self.addtotraceline(droptrace, verticelist[v0])
                                raindrop = ip; giv = tv[v0]; tri = None; in_side = None
                            else:
                                droptrace = self.addtotraceline(droptrace, verticelist[v1])
                                raindrop = ip; giv = tv[v1]; tri = None; in_side = None
                        else:
                            # inner edge - check if we landed exactly on a vertex
                            # squared-distance avoids creating Vector3D objects — threshold 1e-12 ~ (1e-6)^2
                            nextindex = None
                            for j in range(0, 3):
                                vj = verticelist[j]
                                if (ip.X-vj.X)**2 + (ip.Y-vj.Y)**2 + (ip.Z-vj.Z)**2 < 1e-12:
                                    nextindex = tv[j]
                            if nextindex != None:
                                raindrop = ip; giv = nextindex; tri = None; in_side = None
                            else:
                                raindrop = ip; giv = None; tri = nexttri  # in_side already set by GetTriangleAdjacent
                        continue

                    elif in_side != None:
                        # triangle slopes opposite to our direction - follow the entry edge downhill
                        v0 = in_side  # v0/in_side is the same as the vertex index inside the triangle
                        v1 = v0 + 1
                        if v1 == 3: v1 = 0
                        if self.vectorslope(raindrop, verticelist[v0]) < self.vectorslope(raindrop, verticelist[v1]):
                            droptrace = self.addtotraceline(droptrace, verticelist[v0])
                            raindrop = ip; giv = tv[v0]; tri = None; in_side = None
                        else:
                            droptrace = self.addtotraceline(droptrace, verticelist[v1])
                            raindrop = ip; giv = tv[v1]; tri = None; in_side = None
                        continue

            return deadends, tracevert, vertexrelations


        
    def followtriangleslope(self, raindrop, tri, in_side, verticelist, plane): 

        #intersections = Intersections()
        #ip = clr.StrongBox[Point3D]()
        out_side = None

        dropdir = plane.normal
        dropdir.To2D()
        dropdir.Length = 1000.0
        # we can't use segment.Intersect - it's not accurate enough for super thin triangles - it returns too quickly a T=0

        # dropsegment.Extend(10000.0) # can't use that one, it extends the line both ways

        intersections = []
        tolerance = 0.0000001 / dropdir.Length
        if in_side == None:
            #check which triangle side we hit
            # we need to avoid zero intersections due to rounding errors, but we should only get one other solution
            debugwhile = 0
            while intersections.Count < 1:
                debugwhile += 1
                for v0 in range(0, 3):
                    v1 = v0 + 1
                    if v1 == 3: v1 = 0
                    t1, t2, pp = self.find_intersection(verticelist[v0], verticelist[v1], raindrop, raindrop + dropdir, tolerance)
                    if pp != None:
                        intersections.Add([t2, pp, v0]) 
                
                tolerance *= 10
                if debugwhile > 10000:
                    tt = 1
                    break

            intersections.sort()
            
            tt = Vector3D(intersections[0][1], raindrop).Length
            out_side = intersections[0][2]
            # in_side stays None
            return out_side, in_side, intersections[0][1]

        else: # we know where we are coming from
            v0 = in_side
            v1 = v0 + 1
            if v1 == 3: v1 = 0
            v2 = v0 -1
            if v2 == -1: v2 = 2
            t1, t2, pp = self.find_intersection(verticelist[v1], verticelist[v2], raindrop, raindrop + dropdir, tolerance)
            if pp != None:
                intersections.Add([t2, pp, v1]) 
            t1, t2, pp = self.find_intersection(verticelist[v0], verticelist[v2], raindrop, raindrop + dropdir, tolerance)
            if pp != None:
                intersections.Add([t2, pp, v2]) 
            intersections.sort()
            
            # if we have no intersection the triangle falls the opposite way
            # in_side stays what it was, won't be None
            if intersections.Count == 0:
                out_side = None
                return out_side, in_side, None
            elif Vector3D(intersections[0][1], raindrop).Length != 0:
                out_side = intersections[0][2]
                return out_side, in_side, intersections[0][1]
            else:
                return None, None, None




    def find_intersection(self, p0, p1, p2, p3, tolerance ) :

        s10_x = p1.X - p0.X
        s10_y = p1.Y - p0.Y
        s32_x = p3.X - p2.X
        s32_y = p3.Y - p2.Y

        roundto = 14

        denom = s10_x * s32_y - s32_x * s10_y

        if round(denom, 13) == 0 : return None, None, None # collinear


        s02_x = p0.X - p2.X
        s02_y = p0.Y - p2.Y

        s_numer = s10_x * s02_y - s10_y * s02_x
        t_numer = s32_x * s02_y - s32_y * s02_x
        t1 = t_numer / denom
        t2 = s_numer / denom

        if t1 >= 0.-tolerance and t1 <= 1.+tolerance and t2 >= 0.-tolerance and t2 <= 1.+tolerance:

            pp = Point3D(p0.X + (t1 * s10_x), p0.Y + (t1 * s10_y))
            #pp.Z = p0.Z + (Vector3D(p0, p1).Z * t1)  # Vector3D object only needed for .Z which is just p1.Z - p0.Z
            pp.Z = p0.Z + (p1.Z - p0.Z) * t1

            return t1, t2, pp

        else:

            return None, None, None

    def vectorslope(self, p1, p2):
        # inline arithmetic avoids creating a Vector3D object on every call
        #v = Vector3D(p1, p2)
        #if v.Length2D > 0:
        #    slope = v.Z / v.Length2D * 100
        #else:
        #    slope = 0
        #return slope
        dx = p2.X - p1.X
        dy = p2.Y - p1.Y
        len2d = math.sqrt(dx*dx + dy*dy)
        return ((p2.Z - p1.Z) / len2d * 100) if len2d > 0 else 0


    def addtotraceline(self, l, p):
        # continue the traceline
        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
        e.Position = p
        l.AppendElement(e)

        return l

    def newtraceline(self, p1, p2):
        wv = self.currentProject [Project.FixedSerial.WorldView]

        # start new trace line
        l = wv.Add(clr.GetClrType(Linestring))
        l.Layer = self.layerpicker.SelectedSerialNumber
        l.Color = Color.Blue
        l.Weight = 80
        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
        e.Position = p1
        l.AppendElement(e)
        if p2 != None:
            e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
            e.Position = p2
            l.AppendElement(e)
        return l

    #def findtrianglesviacornervertex(self, surface, giv):
    #
    #    nexttovisit = []
    #
    #    # go through all the triangles of the surface
    #    for i in range(surface.NumberOfTriangles):
    #        if surface.GetTriangleMaterial(i) == surface.NullMaterialIndex():
    #            continue
    #        # check if any of the triangle vertices is the one we are looking for
    #        if surface.GetTriangleIVertex(i,0) == giv or \
    #           surface.GetTriangleIVertex(i,1) == giv or \
    #           surface.GetTriangleIVertex(i,2) == giv:
    #            ### if we found a triangle with our vertex we need to check if we've been there already
    #            ##if not i in visitedtri:
    #            ##    nexttovisit.Add(i)
    #            nexttovisit.Add(i)
    #
    #    nexttovisit = list(set(nexttovisit)) # remove duplicates
    #
    #    return nexttovisit

    def connectedvertices(self, surface, giv, vertexrelations, vertextrirelations):

        vertexconnections = []
        # keeping an array with the already computed relations will probably save time when a lot of
        # points rain
        if vertexrelations[giv] == None:
            p_giv = self.vertexpoints[giv]

            # now find the indexes of all visible triangles that touch this corner
            nexttriangles = vertextrirelations[giv]

            # now go through those touching triangles
            for tri in nexttriangles:
                if self.trimaterial[tri] != self.nullmat:
                    tv = self.trivertices[tri]
                    # we iterate through the corners
                    for ci0 in range(0, 3):

                        if giv == tv[ci0]: # if the global index matches
                            ci1 = ci0 + 1
                            if ci1 == 3: ci1 = 0
                            ci2 = ci0 - 1
                            if ci2 == -1: ci2 = 2

                            # get the global indexes and the coordinates for the other 2 corners
                            giv_ci1 = tv[ci1];  p_ci1 = self.vertexpoints[giv_ci1]
                            giv_ci2 = tv[ci2];  p_ci2 = self.vertexpoints[giv_ci2]

                            # we only want to return the connected vertices if they are lower or flat
                            if self.vectorslope(p_giv, p_ci1) <= 0:
                                vertexconnections.Add(giv_ci1)
                            if self.vectorslope(p_giv, p_ci2) <= 0:
                                vertexconnections.Add(giv_ci2)

            # now that we've checked all the touching triangles we finish this vertex
            vertexconnections = list(set(vertexconnections)) # remove duplicates
            vertexconnections.sort()
            if vertexconnections.Count == 0:
                vertexrelations[giv] = None
            else:
                vertexrelations[giv] = vertexconnections
        else:
            pass
        return vertexrelations

    def decideatcorner(self, surface, giv, vertexrelations, vertextrirelations):

        # check the connections to vertices
        # get the corner as Point3D
        p_giv = self.vertexpoints[giv]

        min_side_grade = 10000
        min_side_grade_i = None

        # get all connections to that vertex - global vertex indexes
        vrel = vertexrelations[giv]
        # check which one the steepest is
        for i in range(vrel.Count):
            sl = self.vectorslope(p_giv, self.vertexpoints[vrel[i]])
            if sl < min_side_grade:
                min_side_grade = sl
                min_side_grade_i = vrel[i]

        # check the touching triangles
        touchingtriangles = vertextrirelations[giv]
        min_t_grade = 10000
        ip = clr.StrongBox[Point3D]()
        tracetri = None
        intersections = Intersections()

        for t in touchingtriangles:
            if self.trimaterial[t] != self.nullmat:
                tv = self.trivertices[t]
                vp = self.vertexpoints
                verticelist = (vp[tv[0]], vp[tv[1]], vp[tv[2]])

                # we have to find out if the triangle falls away from us, has a negative slope
                # that's the case if at least on other corner is lower
                for v0 in range(0, 3):
                    v1 = v0 + 1
                    if v1 == 3: v1 = 0
                    v2 = v0 - 1
                    if v2 == -1: v2 = 2
                    # first we need to find out which triangle corner we've just hit
                    if tv[v0] == giv:
                        # since the vertexrelations list only contains lower points we can just look it up
                        if tv[v1] in vertexrelations[giv] or tv[v2] in vertexrelations[giv]:
                            # now we have to find out if a trace line would actually run across the triangle

                            # create a plane from the triangle vertices - use cache to avoid recomputing
                            plane = self.triplane[t]
                            if plane is None:
                                plane = Plane3D(verticelist[0], verticelist[1], verticelist[2])[0]
                                self.triplane[t] = plane
                            dropdir = plane.normal
                            dropdir.To2D()
                            dropdir.Length = 10000.0
                            # create an internal line from the incoming point with the direction of the plane slope
                            dropsegment = SegmentLine(p_giv, p_giv + dropdir)
                            # create an internal line for the opposite triangle side
                            sidesegment = SegmentLine(verticelist[v1], verticelist[v2])
                            if sidesegment.Intersect(dropsegment, False, intersections):
                                if intersections[0].T1 >= 0 and intersections[0].T1 <= 1:
                                    if -(plane.Slope) < min_t_grade:
                                        sidesegment.ComputePoint(intersections[0].T1, ip)
                                        min_t_grade = -(plane.Slope) # plane slope is always positive
                                        min_t_ip = ip.Value
                                        tracetri = t
                                        min_ip_side = v1
                            intersections.Clear()


        # handover back - p, nextindex, in_side, nextisvertex
        if round(min_side_grade, 8) <= round(min_t_grade, 8):
            return self.vertexpoints[min_side_grade_i], min_side_grade_i, None, True
        else:
            # we have to find the next triangle number in order to hand it over to follow terrain
            # check if we hit a surface edge
            isOuter = surface.GetTriangleOuterSide(tracetri, min_ip_side)
            if not isOuter:
                (b, nexttri, ic) = surface.GetTriangleAdjacent(tracetri, min_ip_side)
                # double check if the next triangle might have no material set
                if self.trimaterial[nexttri] != self.nullmat:
                    return min_t_ip, nexttri, ic, False
                else:
                    return min_t_ip, None, None, False
            else:
                return min_t_ip, None, None, False
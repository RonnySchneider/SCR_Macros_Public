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
    cmdData.Key = "SCR_IntersectSurfaceAnySurface"
    cmdData.CommandName = "SCR_IntersectSurfaceAnySurface"
    cmdData.Caption = "_SCR_IntersectSurfaceAnySurface"
    cmdData.UIForm = "SCR_IntersectSurfaceAnySurface"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "DTM"
        cmdData.ShortCaption = "Intersect Surfaces"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.04
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Intersect any Surfaces"
        cmdData.ToolTipTextFormatted = "Intersect any kind Surface"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_IntersectSurfaceAnySurface(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_IntersectSurfaceAnySurface.xaml") as s:
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
        types = Array [Type] ([clr.GetClrType (ProjectedSurface)]) + Array [Type] (SurfaceTypeLists.AllWithCutFillMap)    # we fill an array with TBC object types, we could combine different types

                                                                                                                                                                                                                                                                                                                           # +Array[Type](SurfaceTypeLists.AllWithCutFillMap)
        self.surface1droplist.FilterByEntityTypes = types    # we fill the dropdownlist by applying that types array as filter
        self.surface2droplist.FilterByEntityTypes = types

        self.surface1droplist.AllowNone = False              # our list shall not show an empty field
        self.surface2droplist.AllowNone = False              # our list shall not show an empty field
                                                             # I haven't found
 # a way yet to auto-select the top most value in the list, it says it's read only

        self.surface1droplist.ValueChanged += self.surfacedroplistChanged    # elevation field and the dropdown list the ability to react to changes
        self.surface2droplist.ValueChanged += self.surfacedroplistChanged    # elevation field and the dropdown list the ability to react to changes
        

    def surfacedroplistChanged(self, sender, e):        # in case we select a new surface from the list we update the min/max
                                                          # textfields
        wv = self.currentProject [Project.FixedSerial.WorldView]
        if self.surface1droplist.SelectedSerial != 0 and self.surface2droplist.SelectedSerial != 0:
            surface1 = wv.Lookup (self.surface1droplist.SelectedSerial) # we get our selected surface as object
            surface2 = wv.Lookup (self.surface2droplist.SelectedSerial) # we get our selected surface as object
            nTri1 = surface1.NumberOfTrianglesWithMaterial
            nTri2 = surface2.NumberOfTrianglesWithMaterial
            
            self.Label_nTri1.Content = 'Surface 1: ' + str (nTri1) # we get some surface values into some labels
            self.Label_nTri2.Content = 'Surface 2: ' + str (nTri2)
            self.Label_iterations.Content = 'Iterations: ' + str (nTri1 * nTri2)

        
    def CancelClicked(self, thisCmd, args):
        thisCmd.CloseUICommand ()


    def OkClicked(self, thisCmd, e):
        self.label_benchmark.Content = ''

        start_t = timer ()
        wv = self.currentProject [Project.FixedSerial.WorldView]

        

        surface1 = wv.Lookup (self.surface1droplist.SelectedSerial)    # we get our selected surface as object
        nTri1 = surface1.NumberOfTriangles    # we need the number of triangles in the surface to count through them
        s1_limits = surface1.GetLimits ()
        s1_limits = Limits3D (s1_limits [0],s1_limits [1])

        surface2 = wv.Lookup (self.surface2droplist.SelectedSerial)    # we get our selected surface as object
        nTri2 = surface2.NumberOfTriangles    # we need the number of triangles in the surface to count through them
        s2_limits = surface2.GetLimits ()
        s2_limits = Limits3D (s2_limits [0],s2_limits [1])
        
        if s1_limits.LimitsInLimits (s2_limits,True) [0] == Side.Out: 
            return

        intersectpoints = Array [Point3D] ([Point3D ()] * 6)
        linepoints = Array[Point3D]([Point3D()]*2)

        linetuples = DynArray()

        # progress=0.0
        # max_iterations=nTri1*nTri2
        # i2=0

        # we try to reduce the work we have to do with the second surface, not
        # converting all the time
        s2_array = DynArray ()
        for i in range (nTri2):
            # s2_used[i]=False
            # depending on the type of surface we have to get the triangle
            # vertices in a different way
            if isinstance (surface2,ProjectedSurface):
                tri_arr = Point3DArray (3)
                tri_arr [0] = surface2.TransformPointToWorldDelegate (surface2.GetVertexPoint (surface2.GetTriangleIVertex (i,0)))
                tri_arr [1] = surface2.TransformPointToWorldDelegate (surface2.GetVertexPoint (surface2.GetTriangleIVertex (i,1)))
                tri_arr [2] = surface2.TransformPointToWorldDelegate (surface2.GetVertexPoint (surface2.GetTriangleIVertex (i,2)))
                s2_array.Add (tri_arr)
            else:    
                tri_arr = Point3DArray (3)
                tri_arr [0] = surface2.GetVertexPoint (surface2.GetTriangleIVertex (i,0))
                tri_arr [1] = surface2.GetVertexPoint (surface2.GetTriangleIVertex (i,1))
                tri_arr [2] = surface2.GetVertexPoint (surface2.GetTriangleIVertex (i,2))
                s2_array.Add (tri_arr)

        # now we check the triangles against each other
        for i in range (nTri1): # go through all triangles in Surface 1
            # tt=((((i+1)*(i2+1))/float(max_iterations))*100.)-progress
            # if ((((i+1)*(i2+1))/float(max_iterations))*100.)-progress>0.01:
            #     progress=progress+0.01
            #     self.Progress.Content='Progress: ' + str("%3.2f" % progress)
            #     + ' %'
            
            # if surface1.IsTriangleMaterialPresent (i) == False: continue      # if the triangle is invisible we don't want to waste time
            if surface1.GetTriangleMaterial(i) == surface1.NullMaterialIndex(): continue

            # depending on the type of surface we have to get the triangle
            # vertices in a different way
            if isinstance (surface1,ProjectedSurface):
                s1_v1 = surface1.TransformPointToWorldDelegate (surface1.GetVertexPoint (surface1.GetTriangleIVertex (i,0)))
                s1_v2 = surface1.TransformPointToWorldDelegate (surface1.GetVertexPoint (surface1.GetTriangleIVertex (i,1)))
                s1_v3 = surface1.TransformPointToWorldDelegate (surface1.GetVertexPoint (surface1.GetTriangleIVertex (i,2)))
            else:
                s1_v1 = surface1.GetVertexPoint (surface1.GetTriangleIVertex (i,0))
                s1_v2 = surface1.GetVertexPoint (surface1.GetTriangleIVertex (i,1))
                s1_v3 = surface1.GetVertexPoint (surface1.GetTriangleIVertex (i,2))

            t1_limit = Limits3D (s1_v1, s1_v2)     # we get us the simple limits of that triangle
            t1_limit.ExpandToPoint (s1_v3)       # don't ask me why, but calling Limits3D with 3 points doesn't work,
                                                   # returns a tuple

            plane1 = Plane3D (s1_v1,s1_v2,s1_v3) [0] # the plane itself is returned as first element

            

            for i2 in range (nTri2): # go through all triangles in Surface 2
                
                # if ((((i+1)*(i2+1))/max_iterations)*100.)-progress>0.01:
                #     progress=progress+0.01
                #     self.Progress.Content='Progress: ' + str("%3.2f" %
                #     progress) + ' %'
                
                # if s2_used[i2]==True: continue # if we've already drawn a
                # line in that triangle we skip it
                # if surface2.IsTriangleMaterialPresent (i2) == False: continue  # if the triangle is invisible we don't want to waste time
                if surface2.GetTriangleMaterial(i2) == surface2.NullMaterialIndex(): continue

                # Trimble.Vce.Gem.Model3D.NullMaterialIndex()
                # Trimble.Vce.Gem.Gem.GetTriangleMaterialIndex(int)
                
                # try to reduce the workload, if the whole triangle is outside
                # the other triangle limits skip it
                t2_limit = Limits3D (s2_array [i2] [0], s2_array [i2] [1])     # we get us the simple limits of that triangle
                t2_limit.ExpandToPoint (s2_array [i2] [2])       # don't ask me why, but calling Limits3D with 3 points doesn't work,
                                                                       # returns atuple

                # if the limits of both triangles don't overlap we don't want
                # to waste any more time
                if t1_limit.LimitsInLimits (t2_limit,True) [0] == Side.Out: continue


                plane2 = Plane3D (s2_array [i2] [0], s2_array [i2] [1], s2_array [i2] [2]) [0] # the plane itself is returned as first element

                intersectpoints [0] = plane2.IntersectLine (s1_v1,s1_v2) [5].point # we intersect all sides of Triangle 1 with Plane 2
                intersectpoints [1] = plane2.IntersectLine (s1_v1,s1_v3) [5].point # we intersect all sides of Triangle 1 with Plane 2
                intersectpoints [2] = plane2.IntersectLine (s1_v2,s1_v3) [5].point # we intersect all sides of Triangle 1 with Plane 2

                intersectpoints [3] = plane1.IntersectLine (s2_array [i2] [0], s2_array [i2] [1]) [5].point # we intersect all sides of Triangle 2 with Plane 1
                intersectpoints [4] = plane1.IntersectLine (s2_array [i2] [0], s2_array [i2] [2]) [5].point # we intersect all sides of Triangle 2 with Plane 1
                intersectpoints [5] = plane1.IntersectLine (s2_array [i2] [1], s2_array [i2] [2]) [5].point # we intersect all sides of Triangle 2 with Plane 1

                # the two points that lie inside both triangles form the line
                # segment we are after
                # we go through all intersect points and check if they are
                # inside of them, on the sides actually
                ptcount=0
                for i3 in range (0,6):

                    # TBC only has a built-in function to check for 2D
                    # triangles if a point is inside/on the side
                    # that might not work correctly for perfectly vertical
                    # triangles, so we have to use our own function
                    # if
                    # Triangle2D.IsPointInside(s1_v1,s1_v2,s1_v3,intersectpoints[i3])==True
                    # and \
                    #    Triangle2D.IsPointInside(s2_v1,s2_v2,s2_v3,intersectpoints[i3])==True:
                    if PointInsideTriangle3D (s1_v1,s1_v2,s1_v3,intersectpoints [i3]) == True and \
                       PointInsideTriangle3D (s2_array [i2] [0], s2_array [i2] [1], s2_array [i2] [2],intersectpoints [i3]) == True:
                        if ptcount == 0:
                            linepoints[ptcount] = intersectpoints [i3]
                        else:
                            if ptcount == 1 and linepoints[0] != intersectpoints [i3]:
                                linepoints[ptcount] = intersectpoints [i3]
                        ptcount+=1
                
                # we only start the line drawing in case we have two points        
                if ptcount==2:
                    # we initialize the line we want to draw
                    # linevertex = ElementFactory.Create (clr.GetClrType (IStraightSegment), clr.GetClrType (IXYZLocation))
                    # drawline = wv.Add (clr.GetClrType (Linestring))
                    # linevertex.Position = linepoints [0]
                    # drawline.AppendElement (linevertex)
                    # linevertex.Position = linepoints [1]
                    # drawline.AppendElement (linevertex)
                
                    linetuples.Add(linepoints.Clone())
        
        self.connect_linetuples(linetuples) # we connect the lines as much as possible
        
        end_t = timer ()
        self.label_benchmark.Content = 'elapsed time: ' + str (timedelta (seconds=end_t - start_t))
    
    def connect_linetuples(self, linetuples):                  
    
        wv = self.currentProject[Project.FixedSerial.WorldView]
        layer_sn = self.layerpicker.SelectedSerialNumber

        linestart=Point3D()
        lineend=Point3D()

        while linetuples.Count>0:  # we've got at least one tuple/line
            
            if linetuples.Count==1: # if it's just one we make it simple
                l = wv.Add(clr.GetClrType(Linestring))
                l.Layer = layer_sn
                e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                e.Position = linetuples[0][0]
                l.AppendElement(e)
                e.Position = linetuples[0][1]
                l.AppendElement(e)
                linetuples.RemoveAt(0)
            
            else:
                # if we have more we start with adding the first line and store start/end
                l = wv.Add(clr.GetClrType(Linestring))
                l.Layer = layer_sn
                e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                e.Position = linetuples[0][0]
                linestart=linetuples[0][0]
                l.AppendElement(e)
                e.Position = linetuples[0][1]
                lineend=linetuples[0][1]
                l.AppendElement(e)
                linetuples.RemoveAt(0) # and delete it from the list

                # and now we try to find tuples that are connected to this one
                foundprevious=True # we use this to jump over voids, if our list still contains tuples, 
                                   # but we couldn't find one that is attached we better start a new line
                                   # we have to set it True here since we just started a new line and want to find attached tuples

                while linetuples.Count>0 and foundprevious==True: # if our tuple list still contains values, and we did find an attached one in the previous loop,
                                                                  # we try to find another one
                                                                  # if we can't find another one, but the list contains values,
                                                                  # we revert to the loop above and start a new line
                    
                    foundprevious=False
                    
                    for i in range(0,linetuples.Count): 

                        if linestart.IsDuplicate(linetuples[i][0]): # a tuple that connects to the start of the line
                           e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                           e.Position = linetuples[i][1]                        
                           l.InsertElementAt(e,0)
                           linestart=linetuples[i][1]
                           linetuples.RemoveAt(i)
                           foundprevious=True # we set this to true in order to keep the loop running and try to find another attached tuple
                           break # we stop searching since there can't be another one in this application

                        if linestart.IsDuplicate(linetuples[i][1]): # a tuple that connects to the start of the line
                           e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                           e.Position = linetuples[i][0]                        
                           l.InsertElementAt(e,0)
                           linestart=linetuples[i][0]
                           linetuples.RemoveAt(i)
                           foundprevious=True # we set this to true in order to keep the loop running and try to find another attached tuple
                           break # we stop searching since there can't be another one in this application

                        if lineend.IsDuplicate(linetuples[i][0]): # a tuple that connects to the end of the line
                           e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                           e.Position = linetuples[i][1]                        
                           l.AppendElement(e)
                           lineend=linetuples[i][1]
                           linetuples.RemoveAt(i)
                           foundprevious=True # we set this to true in order to keep the loop running and try to find another attached tuple
                           break # we stop searching since there can't be another one in this application

                        if lineend.IsDuplicate(linetuples[i][1]): # a tuple that connects to the end of the line
                           e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                           e.Position = linetuples[i][0]                        
                           l.AppendElement(e)
                           lineend=linetuples[i][0]
                           linetuples.RemoveAt(i)
                           foundprevious=True # we set this to true in order to keep the loop running and try to find another attached tuple
                           break # we stop searching since there can't be another one in this application
             


def PointInsideTriangle3D(p1,p2,p3,p4):
    a = Vector3D (p2.X - p1.X, p2.Y - p1.Y, p2.Z - p1.Z)
    b = Vector3D (p3.X - p1.X, p3.Y - p1.Y, p3.Z - p1.Z)
    w = Vector3D (p4.X - p1.X, p4.Y - p1.Y, p4.Z - p1.Z)

    aa = Vector3D.DotProduct (a,a) [0]
    ab = Vector3D.DotProduct (a,b) [0]
    bb = Vector3D.DotProduct (b,b) [0]
    wa = Vector3D.DotProduct (w,a) [0]
    wb = Vector3D.DotProduct (w,b) [0]
    
    d = ab * ab - aa * bb

    inside = True
    s = round ((ab * wb - bb * wa) / d, 5) # rounding that value a bit down gives us better results when very close to
                                               # the triangle side, what we are
    if s < 0 or s > 1:                    # otherwise we might miss a value
        inside = False
    t = round ((ab * wa - aa * wb) / d, 5)
    if t < 0 or (s + t) > 1:
        inside = False



    return inside

                     
                
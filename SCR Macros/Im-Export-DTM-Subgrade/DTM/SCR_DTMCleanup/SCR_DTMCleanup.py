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
    cmdData.Key = "SCR_DTMCleanup"
    cmdData.CommandName = "SCR_DTMCleanup"
    cmdData.Caption = "_SCR_DTMCleanup"
    cmdData.UIForm = "SCR_DTMCleanup"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = ""
    cmdData.HelpTopic = ""

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "DTM"
        cmdData.ShortCaption = "DTM Cleanup"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.05
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "DTM Cleanup"
        cmdData.ToolTipTextFormatted = "average close Vertices"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png") 
        cmdData.ImageSmall = b
    except:
        pass


class SCR_DTMCleanup(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_DTMCleanup.xaml") as s:
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

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        self.lfp = self.lunits.Properties.Copy()
        self.linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation

        self.hztol.DistanceMin = 0.00001
        self.hztol.DistanceMax = 1.00000
        self.labelhztol.Content = 'horizontal Radius [' + self.linearsuffix + ']'

        self.vztol.DistanceMin = 0.00001
        self.vztol.DistanceMax = 1.00000
        self.labelvztol.Content = '+- vert Difference [' + self.linearsuffix + ']'

        #self.hztol.MinValue = self.lunits.Convert(self.lunits.InternalType, 0.00001, self.lunits.DisplayType)
        #self.hztol.MaxValue = self.lunits.Convert(self.lunits.InternalType, 1.0, self.lunits.DisplayType)
        #self.hztol.NumberOfDecimals = self.lunits.Units[self.lunits.DisplayType].NumberOfDecimals
        #self.hztol.Value = 0.025
        #
        #self.vztol.MinValue = self.lunits.Convert(self.lunits.InternalType, 0.00001, self.lunits.DisplayType)
        #self.vztol.MaxValue = self.lunits.Convert(self.lunits.InternalType, 1.0, self.lunits.DisplayType)
        #self.vztol.NumberOfDecimals = self.lunits.Units[self.lunits.DisplayType].NumberOfDecimals
        #self.vztol.Value = 0.001

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass


    def SetDefaultOptions(self):
        try:    self.surfacepicker.SelectIndex(OptionsManager.GetInt("SCR_DTMCleanup.surfacepicker", 0))
        except: self.surfacepicker.SelectIndex(0)
        self.hztol.Distance = OptionsManager.GetDouble("SCR_DTMCleanup.hztol", 0.001)
        self.vztol.Distance = OptionsManager.GetDouble("SCR_DTMCleanup.vztol", 0.001)
        self.highlight.IsChecked = OptionsManager.GetBool("SCR_DTMCleanup.highlight", False)

    def SaveOptions(self):
        try:    # if nothing is selected it would throw an error
            OptionsManager.SetValue("SCR_DTMCleanup.surfacepicker", self.surfacepicker.SelectedIndex)
        except:
            pass
        OptionsManager.SetValue("SCR_DTMCleanup.hztol", self.hztol.Distance)
        OptionsManager.SetValue("SCR_DTMCleanup.vztol", self.vztol.Distance)
        OptionsManager.SetValue("SCR_DTMCleanup.highlight", self.highlight.IsChecked)

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''

        self.success.Content = ""

        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())

        wv = self.currentProject [Project.FixedSerial.WorldView]

        try:
            # the "with" statement will unroll any changes if something go wrong
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                surface = wv.Lookup(self.surfacepicker.SelectedSerial)

                if surface.NumberOfTriangles > 2:
                    # for each global vertex number it has the attached triangles
                    vertextrirelations = [[] for i in range(surface.GetNumberOfTriangulatedVertices())]
                    # for each triangle it has the global vertex number
                    trianglevertrelations = [[] for i in range(surface.NumberOfTriangles)]
                    # vertices that are visible
                    visiblevertices = []
                    
                    ProgressBar.TBC_ProgressBar.Title = "preparing Tables"

                    for i in range(surface.NumberOfTriangles):
                        if surface.GetTriangleMaterial(i) != surface.NullMaterialIndex(): # only add triangles that are visible
                            for j in range(0, 3):
                                tt = surface.GetTriangleIVertex(i,j)
                                tt2 = vertextrirelations[tt]
                                tt2.append(i)
                                #vertextrirelations[tt] = tt2
                                tt2 = trianglevertrelations[i]
                                tt2.append(surface.GetTriangleIVertex(i, j))
                                visiblevertices.Add(surface.GetTriangleIVertex(i, j))
                    
                    visiblevertices = list(set(visiblevertices))
                    visiblevertices.sort()
                    # create a copy where we apply our changes
                    newtrianglevertrelations = copy.deepcopy(trianglevertrelations)

                    # in large surface we need a table to quickly find closeby vertices instead of going through all of them
                    lookupx = {}
                    for i in range(0, visiblevertices.Count):
                        p = surface.GetVertexPoint(visiblevertices[i])
                        key = self.truncate(p.X, 0)
                        if key in lookupx:
                            lst = lookupx[key]
                            lst.Add(visiblevertices[i])
                        else:
                            lst = [visiblevertices[i]]
                            lookupx.Add(key, lst)

                    lookupy = {}
                    for i in range(0, visiblevertices.Count):
                        p = surface.GetVertexPoint(visiblevertices[i])
                        key = self.truncate(p.Y, 0)
                        if key in lookupy:
                            lst = lookupy[key]
                            lst.Add(visiblevertices[i])
                        else:
                            lst = [visiblevertices[i]]
                            lookupy.Add(key, lst)


                    # skip vertices we have already averaged
                    skipvertices = [False] * surface.GetNumberOfTriangulatedVertices()

                    # go through all visible vertices and check if they need to be averaged
                    averagepoints = []   # list with the 3D Points
                    ProgressBar.TBC_ProgressBar.Title = "Analyzing " + str(visiblevertices.Count) + " Vertices"
                    buildsurface = True
                    for cnt in range(visiblevertices.Count - 1): # 1 less than the wholelist

                        v1 = visiblevertices[cnt]

                        if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(cnt * 100 / visiblevertices.Count)):
                            buildsurface = False
                            break   # function returns true if user pressed cancel

                        if skipvertices[v1] == True: continue

                        combinevertices = []
                        combinevertices.Add(v1)
                        averagep = surface.GetVertexPoint(v1)
                        
                        closebyvertices = self.quickfindclosevertices(lookupx, lookupy, averagep)

                        for i in range(closebyvertices.Count):
                            if skipvertices[closebyvertices[i]] == True: continue
                            testp = surface.GetVertexPoint(closebyvertices[i])
                            if closebyvertices[i] != v1 and \
                               Vector3D(averagep, testp).Length2D <= self.hztol.Distance and \
                               abs(averagep.Z - testp.Z) <= self.vztol.Distance:

                                # compute new average point
                                combinevertices.Add(closebyvertices[i])
                                averagep = self.averagevertices(surface, combinevertices)
                        
                        # now that we have a new average point we have to update the trianglevertrelations
                        if combinevertices.Count > 1:
                            if self.highlight.IsChecked:
                                cadPoint = wv.Add(clr.GetClrType(CadPoint))
                                cadPoint.Point0 = averagep
                                newcircle = wv.Add(clr.GetClrType(CadCircle))
                                newcircle.CenterPoint = averagep
                                newcircle.Radius = 0.2
                                newcircle.Color = Color.Magenta
                                newcircle.Weight = 80
                        
                            # add the new point to list
                            averagepoints.Add(averagep)
                            
                            # compile a list with the triangles we need to update
                            tritoupdate = []
                            for c in range(combinevertices.Count):
                                v = combinevertices[c]
                                attachedtri = vertextrirelations[v]
                                for tri in range(attachedtri.Count):
                                    tritoupdate.Add(attachedtri[tri])
                            tritoupdate = list(set(tritoupdate))
                        
                            # now we go through that list and update newtrianglevertrelations
                            for tri in range(tritoupdate.Count):
                                ttorg = trianglevertrelations[tritoupdate[tri]] # the list with the original vertices
                                ttnew = newtrianglevertrelations[tritoupdate[tri]] # the list with the new vertices
                                for i in range(0, 3):
                                    for j in range(combinevertices.Count):
                                        if ttorg[i] == combinevertices[j]:
                                            ttnew[i] = "A." + str(averagepoints.Count - 1)
                            
                            # update the skip vertices list
                            for i in range(combinevertices.Count):
                                skipvertices[combinevertices[i]] = True


                    # now that we have the updated lists we can start to compile a new surface
                    # most of the code now is from the CombineSurface macro

                    if averagepoints.Count > 0 and buildsurface == True:
                        ProgressBar.TBC_ProgressBar.Title = "Building new DTM"
                    #if 1 == 2:
                        surName = Model3D.GetUniqueName(surface.Name, None, wv) #make sure name is unique
                        newSurface = wv.Add(clr.GetClrType(Model3D))
                        newSurface.Name = surName
                        builder = newSurface.GetGemBatchBuilder()

                        nTri = surface.NumberOfTriangles
                        mapVertices = {}
                        removedtri = []


                        for t in range(nTri):
                            if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(t * 100 / nTri)):
                                break   # function returns true if user pressed cancel
                            if not surface.IsTriangleMaterialPresent(t):
                                continue
                            # now check if we completely replace this triangle, if any of the vertices are duplicate
                            triverts = newtrianglevertrelations[t]
                            if len(triverts) == len(set(triverts)): # if we don't have duplicates

                                # get the global vertex indices one by one and retrieve the Point3D
                                for i in range(3):
                                    # get the vertex value in order to analyze it
                                    gvi = triverts[i]
                                    if "A." in str(gvi):
                                        tt = gvi.replace('A.', '')
                                        tt = int(tt)
                                        vp = averagepoints[tt]
                                    else:
                                        vp = surface.GetVertexPoint(triverts[i])
                                    
                                    v = builder.AddVertex(vp) # that results the index in the new surface
                                    mapVertices.Add(surface.GetTriangleIVertex(t, i), v[0]) # map the old vertex numbers to the new one
                        
                                for side in range(3):
                                    isOuter = surface.GetTriangleOuterSide(t, side)
                                    if not isOuter:
                                        (ok, tAdj, sideAdj) = surface.Gem.GetTriangleAdjacent(t, side)
                                        if not surface.IsTriangleMaterialPresent(tAdj):
                                            isOuter = True # treat edges next to null as valid
                                    iVertexA = surface.GetTriangleIVertex(t, side)
                                    nextSide = side + 1
                                    if nextSide == 3:
                                        nextSide = 0
                                    iVertexB = surface.GetTriangleIVertex(t, nextSide)
                                    if isOuter or iVertexA < iVertexB or tAdj in removedtri:
                                        b = DTMSharpness.eSoft
                                        if isOuter:
                                            b = DTMSharpness.eSharpAndTextureBndy
                                        builder.AddBreakline(Byte(b), mapVertices[iVertexA], mapVertices[iVertexB])
                            else:
                                removedtri.Add(t)
                        builder.Construction()
                        builder.Commit() 
                    else:
                        self.success.Content = 'no vertices within tolerance found'

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
        #self.success.Content += '\nold Surface - '
        #self.success.Content += str(trianglevertrelations.Count) + " Triangles - " + \
        #                        str(visiblevertices.Count) + " Vertices"
        #self.success.Content += '\nnew Surface - '
        #self.success.Content += str(newSurface.NumberOfTriangles) + " Triangles - " + \
        #                        str(newSurface.GetNumberOfTriangulatedVertices()) + " Vertices"
        ProgressBar.TBC_ProgressBar.Title = ""
        #self.currentProject.Calculate(False)
        self.SaveOptions()
   
    def quickfindclosevertices(self, lookupx, lookupy, p):

        px = self.truncate(p.X, 0)
        py = self.truncate(p.Y, 0)
        closev = []
        xlist = []
        ylist = []
        # check which lookup list has more values - chance is that the for-loop will be shorter
        
        for x in range(px-2, px+3):
            if x in lookupx:
                xlist = xlist + lookupx[x]

        for y in range(py-2, py+3):
            if y in lookupy:
                ylist = ylist + lookupy[y]

        closev = list(set(xlist) & set(ylist))
        return closev

    def averagevertices(self, surface, combinevertices):
        avx = 0
        avy = 0
        avz = 0
        for i in range(combinevertices.Count):
            avx += surface.GetVertexPoint(combinevertices[i]).X
            avy += surface.GetVertexPoint(combinevertices[i]).Y
            avz += surface.GetVertexPoint(combinevertices[i]).Z
        avx = avx / combinevertices.Count
        avy = avy / combinevertices.Count
        avz = avz / combinevertices.Count
        averagep = Point3D(avx, avy, avz)
        return averagep

    def truncate(self, number, decimals=0):
        """
        Returns a value truncated to a specific number of decimal places.
        """
        if not isinstance(decimals, int):
            raise TypeError("decimal places must be an integer.")
        elif decimals < 0:
            raise ValueError("decimal places has to be 0 or more.")
        elif decimals == 0:
            return math.trunc(number)
    
        factor = 10.0 ** decimals
        return math.trunc(number * factor) / factor

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
import multiprocessing
exec(open(r"C:\ProgramData\Trimble\MacroCommands3\SCR Macros\SCR_Imports.py").read())

_OPTIONS = {
    "samplingdistance": 0.005,
    "samplingcount":    15000.0,
    "layerpicker":      8,
    "fitmode":          True,
    "hzoffset":         0.0,
    "tolerance":        0.001,
    "bestfittype":      1.0,
    "averagemode":      False,
    "average":          True,
    "layerpicker2":     8,
    "median":           True,
    "layerpicker3":     8,
    "interval":         0.1,
    "intervaloffset":   0.05,
    "hzoffset2":        0.015,
    "filterdtm":        False,
    "surfacepicker":    0,
    "dtmoffset":        0.02,
}

def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_AnalyzeCloudAlongLine"
    cmdData.CommandName = "SCR_AnalyzeCloudAlongLine"
    cmdData.Caption = "_SCR_AnalyzeCloudAlongLine"
    cmdData.UIForm = "SCR_AnalyzeCloudAlongLine"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "Pointcloud"
        cmdData.ShortCaption = "Analyze along Line"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.14
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Analyze Cloud along Line"
        cmdData.ToolTipTextFormatted = "creates Bestfit-Lines into a Pointcloud Corridor defined by Source-Lines"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png") 
        cmdData.ImageSmall = b
    except:
        pass

class SCR_AnalyzeCloudAlongLine(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_AnalyzeCloudAlongLine.xaml") as s:
            wpf.LoadComponent (self, s)
        self.currentProject = currentProject
        self.macroFileFolder = macroFileFolder

    def HelpClicked(self, cmd, e):
        webbrowser.open("C:\ProgramData\Trimble\MacroCommands3\SCR Macros\MacroHelp\MacroHelp.htm#" + type(self).__name__)


    def OnLoad(self, cmd, buttons, event):
        
        # fix broken project
        #failGuard.Commit()
        #UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
        #self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)

        
        self.okBtn = buttons[0]
        buttons[2].Content = "Help"
        buttons[2].Visibility = Visibility.Visible
        buttons[2].Click += self.HelpClicked
        self.Caption = cmd.Command.Caption

        self.lType = clr.GetClrType(IPolyseg)
        self.pointcloudType = clr.GetClrType(PointCloudRegion)

        optionMenu = SelectionContextMenuHandler()
        # remove options that don't apply here
        optionMenu.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        self.objs.ButtonContextMenu = optionMenu
        self.objs.IsEntityValidCallback = self.IsValid

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        #self.lfp = self.lunits.Properties.Copy()
        linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation

        self.hzoffsetlabel.Content = "filter Cloud Left/Right of Line [" + linearsuffix + "]"
        self.tolerancelabel.Content = "Bestfit chording Tolerance [" + linearsuffix + "]"

        self.bestfittype.NumberOfDecimals = 0
        self.bestfittype.MinValue = 1

        self.samplingcount.NumberOfDecimals = 0
        self.samplingcount.MinValue = 1
        self.samplingcount.MaxValue = 15000

        types = Array [Type] (SurfaceTypeLists.AllWithCutFillMap)    # we fill an array with TBC object types, we could combine different types
        self.surfacepicker.FilterByEntityTypes = types    # we fill the dropdownlist by applying that types array as filter
        self.surfacepicker.AllowNone = False              # our list shall not show an empty field

        SCRExpanders.wire_pairs([
            (self.expander_fitmode, self.fitmode),
            (self.expander_averagemode, self.averagemode),
        ])
		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        self.lfp = self.lunits.Properties.Copy() # create a copy in order to set the decimals and enable/disable the suffix
        self.lfp.AddSuffix = False # disable suffix, we need to set it manually, it would always add the current projects units


    def tooutputunit(self, v):
        
        #self.lfp.AddSuffix = self.addunitsuffix.IsChecked
        self.lfp.AddSuffix = False
        #return self.lunits.Format(LinearType.Meter, v, self.lfp, LinearType(self.outputunitenum))
        return self.lunits.Format(LinearType.Meter, v, self.lfp, LinearType.Meter)

    def SetDefaultOptions(self):
        SCROptions.LoadMacroOptions(self, "SCR_AnalyzeCloudAlongLine", _OPTIONS, self.currentProject)

    def SaveOptions(self):
        SCROptions.SaveMacroOptions(self, "SCR_AnalyzeCloudAlongLine", _OPTIONS)

    def IsValid(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        ##if isinstance(o, self.pointcloudType):
        ##    return True
        if isinstance(o, CoordPoint) or isinstance(o, CadPoint):
            return True
        return False

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content = ''
        self.success.Content = ''
        
        wv = self.currentProject [Project.FixedSerial.WorldView]

        activeForm = TrimbleOffice.TheOffice.MainWindow.AppViewManager.ActiveView
        if not isinstance(activeForm, clr.GetClrType(Hoops2dView)):
            self.error.Content += "\nthe active view must be a 2D-View"
            return
        if not activeForm.View.LeftMouseOperation == LeftMouseModeType.PolygonSelection:
            self.error.Content = "\nyou must be in 'Polygon Select Mode'"
            return

        if self.averagemode.IsChecked and not (self.average.IsChecked or self.median.IsChecked):
            self.error.Content = '\nyou need to select at least one average mode'
            return

        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)

        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
        
                self.linelist = []
                self.cloudlist = []
                normalpoints = []

                for o in self.objs:   
                    if isinstance(o, self.lType):
                        self.linelist.Add(o)
                    if isinstance(o, CoordPoint) or isinstance(o, CadPoint):
                        normalpoints.Add(o.Position)
                    if isinstance(o, self.pointcloudType):
                        self.cloudlist.Add(o)

                if self.linelist.Count > 0:

                    self.line = 0
                    # go through all the selected lines
                    for l in self.linelist:
                        self.line += 1
                        ProgressBar.TBC_ProgressBar.Title = "filter Cloud-Points along Line " + str(self.line) + '/' + str(self.linelist.Count)

                        # retrieve the points from the cloud - get list with Point3D's
                        cloudpoints = self.selectpointsfromcloud(l)
                        # add normal cad, survey, stake points
                        for p in normalpoints:
                            cloudpoints.Add(p)

                        if cloudpoints.Count > 0:
                            
                            self.fullbreak = False
                                    
                            if self.fitmode.IsChecked:
                                self.bestfitline(cloudpoints, l)
                                if self.fullbreak:
                                    break

                            elif self.averagemode.IsChecked:
                                self.computeaverage(cloudpoints, l)
                                if self.fullbreak:
                                    break

                        else:
                            self.error.Content = 'no Cloudpoints found'
                else:
                    self.error.Content = 'no valid Lines found'

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


        ProgressBar.TBC_ProgressBar.Title = ''
        self.SaveOptions()


    def selectpointsfromcloud(self, l):

        wv = self.currentProject [Project.FixedSerial.WorldView]
        fakemouse = True
        cloudpoints = []

        if fakemouse:

            activeForm = TrimbleOffice.TheOffice.MainWindow.AppViewManager.ActiveView

            polyseg = l.ComputePolySeg().Clone()
            polyseg = polyseg.ToWorld()
            
            # to be on the save side we select more points from the cloud
            # the proper filtering is done later
            if self.fitmode.IsChecked:
                hzoffset = abs(self.hzoffset.Distance) * 1.5
            elif self.averagemode.IsChecked:
                hzoffset = abs(self.hzoffset2.Distance) * 1.5

            if not isinstance(activeForm, clr.GetClrType(LimitSliceWindow)): # normal PlanView
                # moving lines sideways
                polyseg_r = polyseg.Offset(Side.Right, hzoffset)[1]
                polyseg_l = polyseg.Offset(Side.Left, hzoffset)[1]
                polyseg_r.Extend(hzoffset)
                polyseg_l.Extend(hzoffset)
            else:   # Cutting Plane View
                # moving lines upwards instead of sideways
                polyseg_r = polyseg.Clone()
                polyseg_r.Translate(polyseg_r.FirstNode, polyseg_r.LastNode, Vector3D(0, 0, -hzoffset))
                polyseg_l = polyseg.Clone()
                polyseg_l.Translate(polyseg_l.FirstNode, polyseg_l.LastNode, Vector3D(0, 0, hzoffset))

            polyseg_l.Reverse()

            polyseg_r.Add(polyseg_l.FirstNode.Point)

            finalpolyseg = polyseg_r.Clone()
            tt2 = finalpolyseg.Join(polyseg_l.Clone())

            if not finalpolyseg.IsClockWise(): # have a clockwise motion, avoid selecting things by just crossing them
                finalpolyseg.Reverse()
            
            # make sure we don't have arcs
            finalpolyseg = finalpolyseg.Linearize(0.001, 0.001, 50, None, False)

            ptsList = List[Point3D]() # we need a generic list, not an array
            pts = finalpolyseg.ToPoint3DArray()
            for p in pts:
                if p.Is2D: p.Z = 0 # the points must be 3D, otherwise the selection fails
                ptsList.Add(p)
            p = ptsList[0]
            ptsList.Add(p)

            #newpolyseg = PolySeg.PolySeg()
            #newpolyseg.Add(ptsList)
            #ltest = wv.Add(clr.GetClrType(Linestring))
            #ltest.Append(newpolyseg, None, False, False)
            #ltest.Layer = self.layerpicker.SelectedSerialNumber

            #GlobalSelection.Clear()

            op =  activeForm.View.CurrentOperator # should be PolygonSelectionOperator
            op.PolygonStarted = True
            op.PolygonPoints = ptsList
            activeForm.View.PolygonSelectionStarted = True # need to fake a mouse gesture
            op.PerformSelection()
            #activeForm.View.PolygonSelectionStarted = False # need to fake a mouse gesture

            selectionSet = GlobalSelection.Items(self.currentProject)
            cloudselectionids = []
            for o in selectionSet:   
                # compile a list with the ID's of the selected cloud portions
                if isinstance(o, self.pointcloudType):
                    cloudselectionids.Add(o.Integration.GetSelectedCloudId())
                    cloudintegration = o.Integration.PointCloudDatabase.Integration

            if cloudselectionids.Count > 0:
                #rwcloudpoints = []
                ProgressBar.TBC_ProgressBar.Title = "retrieve selected Cloud-Points"
                cps = cloudintegration.GetSelectedPoints(cloudselectionids)
                self.error.Content += '\nselected Cloud Point: ' + str(cps.Count)
                
                #for o in selectionSet:
                #    if isinstance(o, self.pointcloudType):
                        #integration = o.Integration  # = SdePointCloudRegionIntegration
                        #selectedid = integration.GetSelectedCloudId() # it seems the selected points form a sub-cloud
                        #regiondb = integration.PointCloudDatabase # PointCloudDatabase
                        #sdedb = regiondb.Integration # SdePointCloudDatabaseIntegration
                        #cps = sdedb.GetPoints(selectedid)
                        
                        #cps = cloudintegration.GetPoints(o.Integration.GetSelectedCloudId())
                        #t1 = o.Integration.PointCloudDatabase.Integration.GetPointInfos(o.Integration.GetSelectedCloudId())
                        #t1 = o.Integration.PointCloudDatabase.Integration.GetScanIds(o.Integration.GetSelectedCloudId())

                        #lookupx = {}
                        #lookupy = {}
                        #lookupz = {}
                        
                if self.fitmode.IsChecked and cps.Count > 15000:

                    ProgressBar.TBC_ProgressBar.Title = "downsampling Cloud-Points"
                    # throw that list of all the selected portions onto the sampling algorithm
                    # it returns another ID
                    cpssampleid = cloudintegration.CreateSpatiallySampledCloud(cloudselectionids, 0.002, 15000)
                    # now get the points from that specific subcloud
                    cps2 = cloudintegration.GetPoints(cpssampleid)

                    for p in cps2:
                        cloudpoints.Add(p)

                    self.error.Content += '\nsampled to : ' + str(cloudpoints.Count)

                elif self.averagemode.IsChecked:
                    
                    for p in cps:
                        cloudpoints.Add(p)

                            #rwcloudpoints.Add(RwPoint3D(p.X, p.Y, p.Z))
                            ### test regards quick filtering of cloud points - dictionary approach from DTMCleanup
                            ##key = self.truncate(p.X, 2)
                            ##if key in lookupx:
                            ##    lst = lookupx[key]
                            ##    lst.Add(cloudpoints.Count - 1)
                            ##else:
                            ##    lst = [cloudpoints.Count - 1]
                            ##    lookupx.Add(key, lst)
                            ##
                            ##key = self.truncate(p.Y, 2)
                            ##if key in lookupy:
                            ##    lst = lookupy[key]
                            ##    lst.Add(cloudpoints.Count - 1)
                            ##else:
                            ##    lst = [cloudpoints.Count - 1]
                            ##    lookupy.Add(key, lst)
                            ##
                            ##key = self.truncate(p.Z, 2)
                            ##if key in lookupz:
                            ##    lst = lookupz[key]
                            ##    lst.Add(cloudpoints.Count - 1)
                            ##else:
                            ##    lst = [cloudpoints.Count - 1]
                            ##    lookupz.Add(key, lst)
                        #tt = RwPlane3D.FitPlaneTo3DPoints(rwcloudpoints)
                        #tt = 1

        ## currently not used
        ## no fake mouse gesture - use built in FilterCloudByPolyline instead
        ##else:   
        ##    ProgressBar.TBC_ProgressBar.Title = "retrieve Points from Cloud"
        ##    cloudpoints = []
        ##    
        ##    polyseg = l.ComputePolySeg()
        ##    polyseg = polyseg.ToWorld()
        ##    polyseg_v = l.ComputeVerticalPolySeg()
        ##
        ##    # muse have a valid vertical polyseg for FilterCloudByPolyline
        ##    if not polyseg_v:
        ##        polyseg_v = PolySeg.PolySeg()
        ##        polyseg_v.Add(Point3D(0, 0, 0))
        ##
        ##    for o in self.cloudlist:
        ##        #progress = 0
        ##        #ProgressBar.TBC_ProgressBar.SetProgress(progress)
        ## 
        ##        integration = o.Integration  # = SdePointCloudRegionIntegration
        ##        selectedid = integration.GetSelectedCloudId() # it seems the selected points form a sub-cloud
        ##        regiondb = integration.PointCloudDatabase # PointCloudDatabase
        ##        sdedb = regiondb.Integration # SdePointCloudDatabaseIntegration
        ##
        ##        useonlyradius = False   # if both are True it overrides left/right values but not above/below
        ##                                # does include points directly on the line
        ##        leftvalue = 0.015       # left/right won't include points directly on the line
        ##        rightvalue = 0.015  
        ##        radiusvalue = 0.03
        ##        useelevation = False    # if both are False then a vertical section through the cloud is selected 
        ##        abovevalue = 0.005
        ##        belowvalue = 0.0020
        ##
        ##        testlist = []
        ##        testlist.Add(selectedid)
        ##        filterresult = sdedb.FilterCloudByPolyline(testlist, polyseg, polyseg_v, useonlyradius, leftvalue, rightvalue, radiusvalue, useelevation, abovevalue, belowvalue)
        ##
        ##        cps = sdedb.GetPoints(filterresult.IncludedByFilterCloudId)
        ##
        ##        for p in cps:
        ##            cadPoint = wv.Add(clr.GetClrType(CadPoint))
        ##            cadPoint.Layer = self.layerpicker.SelectedSerialNumber
        ##            cadPoint.Point0 = p
        ##
        ##            cloudpoints.Add(p)
   
        return cloudpoints

    def computeaverage(self, cloudpoints, o):

        wv = self.currentProject [Project.FixedSerial.WorldView]
        
        polyseg = o.ComputePolySeg()
        polyseg = polyseg.ToWorld()
        polyseg_v = o.ComputeVerticalPolySeg()

        if self.filterdtm.IsChecked:
            surface = wv.Lookup(self.surfacepicker.SelectedSerial)
            surfacegem = surface.Gem.Clone()
        else:
            surfacegem = None

        pointstoaverage = []

        chainagelist = []
        pointcollection = []
        ch = abs(self.intervaloffset.Distance)
        i = 0
        # compile chainagelist and create indices into pointcollection
        while not ch > polyseg.ComputeStationing():
            # 2 alternating entries
            # chainagelist(start-ch)
            # chainagelist(end-ch)
            chainagelist.Add([ch - abs(self.intervaloffset.Distance)])
            chainagelist.Add([ch + abs(self.intervaloffset.Distance)])
            pointcollection.Add([])
            ch += abs(self.interval.Distance)
            i += 1

        #threadcount = cpu_count()
        threadcount = 1
        self.exc_info = None

        hzoffset2 = self.hzoffset2.Distance
        dtmoffset = self.dtmoffset.Distance
        filterdtm = self.filterdtm.IsChecked

        ProgressBar.TBC_ProgressBar.Title = "filtering points along the line"
        threads = [Thread(target = self.filteralongline, args = (polyseg, surfacegem, filterdtm, cloudpoints, pointcollection, chainagelist, hzoffset2, dtmoffset, threadcount, threadindex,)) for threadindex in range(threadcount)]
        
        # start threads
        for thread in threads:
            thread.start()
        # wait for all threads to terminate
        for thread in threads:
            thread.join()

        if self.exc_info:
            exc_type, exc_obj, exc_tb = self.exc_info
            self.error.Content += '\nan Error occurred - Result probably incomplete\n' + str(exc_type) + '\n' + str(exc_obj) + '\nLine ' + str(exc_tb.tb_lineno)


        #self.filteralongline(polyseg, cloudpoints, pointcollection, chainagelist, 1, 0)


        if self.fullbreak: 
            return

        ProgressBar.TBC_ProgressBar.Title = "computing average"
        for sample in pointcollection:
            if sample.Count > 0:
                finalp = Point3D()
                
                if self.average.IsChecked:
                    for p in sample:
                        finalp += p

                    finalp.X /= sample.Count
                    finalp.Y /= sample.Count
                    finalp.Z /= sample.Count

                    # draw the point
                    cadPoint = wv.Add(clr.GetClrType(CadPoint))
                    cadPoint.Layer = self.layerpicker2.SelectedSerialNumber
                    cadPoint.Point0 = finalp
                    #cadPoint.Name = 'Average from ' + str(pointstoaverage.Count) + ' Points'

                if self.median.IsChecked:
                    allx = []
                    ally = []
                    allz = []

                    for p in sample:
                        allx.Add(p.X)
                        ally.Add(p.Y)
                        allz.Add(p.Z)

                    allx.sort()
                    ally.sort()
                    allz.sort()

                    # get the median value
                    finalp.X = allx[int(math.floor(sample.Count / 2))]
                    finalp.Y = ally[int(math.floor(sample.Count / 2))]
                    finalp.Z = allz[int(math.floor(sample.Count / 2))]

                    # draw the point
                    cadPoint = wv.Add(clr.GetClrType(CadPoint))
                    cadPoint.Layer = self.layerpicker3.SelectedSerialNumber
                    cadPoint.Point0 = finalp
                    #cadPoint.Name = 'Average from ' + str(pointstoaverage.Count) + ' Points'

        return

    def filteralongline(self, polyseg, surface, filterdtm, cloudpoints, pointcollection, chainagelist, hzoffset2, dtmoffset ,threadcount, threadindex):

        try:
        
            #self.error.Content += str(threadcount) + " - " + str(threadindex)
            

            activeForm = TrimbleOffice.TheOffice.MainWindow.AppViewManager.ActiveView
            wv = self.currentProject [Project.FixedSerial.WorldView]

            outSegment=clr.StrongBox[Segment]()
            out_t=clr.StrongBox[float]()
            outPointOnCL = clr.StrongBox[Point3D]()
            station = clr.StrongBox[float]()
            perpVector3D=clr.StrongBox[Vector3D]()
            testDist=clr.StrongBox[float]()
            testside=clr.StrongBox[Side]()

            surfaceoutPoint = clr.StrongBox[Point3D]()

            pointsperthread = math.ceil(cloudpoints.Count / threadcount)

            jstart = int(0 + threadindex * pointsperthread)
            jend = int(0 + threadindex * pointsperthread + pointsperthread)
            if jend > cloudpoints.Count:
                jend = cloudpoints.Count

            for j in range(jstart, jend):

                p = cloudpoints[j]
                
                if polyseg.FindPointFromPoint(p, outSegment, out_t, outPointOnCL, station, perpVector3D, testDist, testside): 
                    offsettestedok = False
                    if not isinstance(activeForm, clr.GetClrType(LimitSliceWindow)):
                        if testDist.Value <= abs(hzoffset2):
                            offsettestedok = True
                    else: 
                        pcompute = outPointOnCL.Value
                        #if polyseg_v != None:
                        #    pcompute.Z = polyseg_v.ComputeVerticalSlopeAndGrade(station.Value)[1]
                        if abs(p.Z - pcompute.Z) <= abs(hzoffset2):
                            offsettestedok = True
            
                    if offsettestedok:
                        for i in range(0, chainagelist.Count, 2):
                            if station.Value > chainagelist[i][0] and station.Value < chainagelist[i + 1][0]:
                                if filterdtm:
                                    if surface.PickSurface(p, surfaceoutPoint):
                                        tt = Vector3D(p, surfaceoutPoint.Value).Length
                                        if Vector3D(p, surfaceoutPoint.Value).Length <= abs(dtmoffset):
                                            
                                            pointcollection[int(i / 2)].Add(p) # chainagelist is twice as long

                                            break
                                    else:
                                        break
                                else:

                                    pointcollection[int(i / 2)].Add(p) # chainagelist is twice as long

                                    break
        except Exception as e:
            self.exc_info = sys.exc_info()


    def bestfitline(self, cloudpoints, o):
        
        activeForm = TrimbleOffice.TheOffice.MainWindow.AppViewManager.ActiveView

        wv = self.currentProject [Project.FixedSerial.WorldView]

        polyseg = o.ComputePolySeg()
        polyseg = polyseg.ToWorld()
        #polyseg_v = o.ComputeVerticalPolySeg()
        #polyseg_v.Transform(o.OSCToWorld)
        
        outSegment=clr.StrongBox[Segment]()
        out_t=clr.StrongBox[float]()
        outPointOnCL = clr.StrongBox[Point3D]()
        station = clr.StrongBox[float]()
        perpVector3D=clr.StrongBox[Vector3D]()
        testDist=clr.StrongBox[float]()
        testside=clr.StrongBox[Side]()
        
        pointstofitto = []
        
        # filter along the line
        j = 0
        time1 = datetime.now()
        for p in cloudpoints:
            j += 1
            if (datetime.now() - time1).seconds > 0.5:
                if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(j * 100 / cloudpoints.Count)):
                    self.fullbreak = True
                    break   # function returns true if user pressed cancel
                time1 = datetime.now()
            
            if polyseg.FindPointFromPoint(p, outSegment, out_t, outPointOnCL, station, perpVector3D, testDist, testside): 
                if not isinstance(activeForm, clr.GetClrType(LimitSliceWindow)):
                    if testDist.Value <= abs(self.hzoffset.Distance):
                        pointstofitto.Add(p)
                else: 
                    pcompute=outPointOnCL.Value
                    #if polyseg_v != None:
                    #    pcompute.Z = polyseg_v.ComputeVerticalSlopeAndGrade(station.Value)[1]
                    if abs(p.Z - pcompute.Z) <= abs(self.hzoffset.Distance):
                        pointstofitto.Add(p)

        if self.fullbreak: 
            return

        if pointstofitto.Count > 1:
            ProgressBar.TBC_ProgressBar.Title = "computing Best-Fit Line " + str(self.line) + '/' + str(self.linelist.Count)
            # create the best fit function object
            bestfitline = CreateBestFitLine()
            t = int(self.bestfittype.Value)
            e1 = Action[String, String]
            e2 = abs(self.tolerance.Distance)
            
            # compute the best fit solution
            bestfitenum = bestfitline.CalculatePolynomialCurve3D(pointstofitto, t, e1, e2)
            bestfitlinepoints = List[Point3D]()

            # copy the points into a list object that we can use as input for a new polyseg
            for p in bestfitenum:
                bestfitlinepoints.Add(p)
            
            newpolyseg = PolySeg.PolySeg()
            newpolyseg.Add(bestfitlinepoints)

            debughog = False

            if debughog:
                maxhog = self.getmaxhog(newpolyseg.Clone())
                if maxhog[0]:
                    cadPoint = wv.Add(clr.GetClrType(CadPoint))
                    cadPoint.Layer = self.layerpicker.SelectedSerialNumber
                    cadPoint.Point0 = maxhog[2]
                    #t = wv.Add(clr.GetClrType(MText))
                    #t.AlignmentPoint = maxhog[2]
                    #t.TextString = self.tooutputunit(maxhog[1])
                    #t.Layer = self.layerpicker.SelectedSerialNumber
                    #t.Height = 0.5


            l = wv.Add(clr.GetClrType(Linestring))
            l.Append(newpolyseg, None, False, False)
            l.Layer = self.layerpicker.SelectedSerialNumber
            l.Name = 'Bestfit-Type: ' + str(t) + ' - ' + str(pointstofitto.Count) + ' Points'
            #tt = Parabola.CreateParabolaFrom3Points(bestfitlist[0], bestfitlist[int(bestfitlist.Count / 2)], bestfitlist[bestfitlist.Count - 1])
            #tt = e2

        return

    def getmaxhog(self, polyseg):

        
        p1 = polyseg.FirstSegment.BeginPoint
        p2 = polyseg.LastSegment.EndPoint

        rottozero = Spinor3D.ComputeRotation(Vector3D(p1, p2), Vector3D(0,0,1), Vector3D(1,0,0), Vector3D(0,0,1))
        matrixtozero = Matrix4D.BuildTransformMatrix(Vector3D(p1), Vector3D(p1, Point3D(0, 0, 0)), rottozero, Vector3D(1,1,1))
        matrixback = Matrix4D.Inverse(matrixtozero)

        polyseg.Transform(matrixtozero)
        
        maxhog = 0.0
        maxnode = None

        for n in polyseg.Nodes():
            try:
                if n.Visible:
                    if n.Point.Is3D and n.Point.Z > maxhog:
                        maxhog = n.Point.Z
                        maxnode = n
            except:
                pass

        if maxnode:
            return [True, maxhog, matrixback.TransformPoint(maxnode.Point)]
        else:
            return [False, None, None]


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

# custom class wrapping a list in order to make it thread safe
class ThreadSafeList():
    # constructor
    def __init__(self):
        # initialize the list
        self._list = list()
        # initialize the lock
        self._lock = Lock()
 
    # add a value to the list
    def Add(self, value):
        # acquire the lock
        with self._lock:
            # append the value
            self._list.Add(value)
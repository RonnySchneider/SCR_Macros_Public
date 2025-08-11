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
    cmdData.Key = "SCR_PerpDistToDTM"
    cmdData.CommandName = "SCR_PerpDistToDTM"
    cmdData.Caption = "_SCR_PerpDistToDTM"
    cmdData.UIForm = "SCR_PerpDistToDTM" # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "DTM"
        cmdData.ShortCaption = "Dist to DTM/IFC"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3
     
        cmdData.Version = 1.19
        cmdData.MacroAuthor = "SCR"
        cmdData.ToolTipTitle = "3D distance to DTM/IFC"
        cmdData.ToolTipTextFormatted = "get the perpendicular (or plumb) 3D distance from point(s) to a DTM or IFC-Object"
    except:
        pass
    try:
        b = Bitmap(macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass

# the name of this class must match name from cmdData.UIForm (above)
class SCR_PerpDistToDTM(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader(macroFileFolder + r"\SCR_PerpDistToDTM.xaml") as s:
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
        types = Array[Type](SurfaceTypeLists.AllWithCutFillMap)+Array[Type]([clr.GetClrType(ProjectedSurface)])
        #types.extend (Array[Type]([clr.GetClrType(ProjectedSurface)]))
        self.surfacepicker.FilterByEntityTypes = types
        self.surfacepicker.AllowNone = False

        self.coordCtl1.ValueChanged += self.Coord1Changed
        
        self.cadpointType = clr.GetClrType(CadPoint)
        self.coordpointType = clr.GetClrType(CoordPoint)
        self.pointcloudType = clr.GetClrType(PointCloudRegion)

        self.textheight.DistanceMin = 0

        # self.ifcs.UseLocalSelection = True
        # self.objs.UseLocalSelection = True
        # self.ifcs.ProcessGlobalSelectionChanges = False
        # self.objs.ProcessGlobalSelectionChanges = False

        self.ifcs.IsEntityValidCallback = self.IsValidIFC
        # optionMenuifcs = SelectionContextMenuHandler()
        # remove options that don't apply here
        # optionMenuifcs.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        # self.ifcs.ButtonContextMenu = optionMenuifcs
        
        self.objs.IsEntityValidCallback = self.IsValid
        # optionMenuobjs = SelectionContextMenuHandler()
        # remove options that don't apply here
        # optionMenuobjs.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        # self.objs.ButtonContextMenu = optionMenuobjs

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass
        
        self.unitssetup(None, None)        
        self.ignoreGotFocus = False
        self.selectionControls = [self.objs, self.ifcs]

    def Selection_PreviewGotFocus(self, sender, e):
        self.ignoreGotFocus = True
        for ctrl in self.selectionControls:
            value = (sender == ctrl)
            ctrl.ProcessGlobalSelectionChanges = value
            ctrl.UpdateTextOnSelectionChange = value
        pass

    def Selection_ValueChanged(self, sender, e):
        #There is a bug in the Selection control, it resets the selection on the "GotFocus" event and then sets it back to the proper selection
        if self.ignoreGotFocus:
            self.ignoreGotFocus = False
            return

    def IsValidIFC(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, BIMEntity):
            return True
        if isinstance(o, Shell3D):
            return True
        return False
        
    def IsValid(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.coordpointType):
            return True
        if isinstance(o, self.cadpointType):
            return True
        if isinstance(o, self.pointcloudType):
            return True
        return False
    
    def usedtmChanged(self, sender, e):
        if self.usedtm.IsChecked:
            self.surfacepicker.IsEnabled = True
            # self.computeplumb.IsEnabled = True
            self.ifcs.IsEnabled = False
        else:
            self.surfacepicker.IsEnabled = False
            # self.computeplumb.IsEnabled = False
            self.ifcs.IsEnabled = True

    def usesinglepointChanged(self, sender, e):
        if self.usesinglepoint.IsChecked:
            self.coordCtl1.IsEnabled = True
            self.objs.IsEnabled = False
        else:
            self.coordCtl1.IsEnabled = False
            self.objs.IsEnabled = True

    def unitssetup(self, sender, e):
        # setup everything for the unit conversions
        self.outputunitenum = 0
        self.textdecimals.NumberOfDecimals = 0

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        self.lfp = self.lunits.Properties.Copy() # create a copy in order to set the decimals and enable/disable the suffix
        self.lfp.AddSuffix = False # disable suffix, we need to set it manually, it would always add the current projects units

        # fill the unitpicker
        for u in self.lunits.Units:
            item = ComboBoxItem()
            item.Content = u.Key
            item.FontSize = 1
            self.unitpicker.Items.Add(item)

        tt = self.unitpicker.Text
        self.unitpicker.SelectedIndex = 0
        if tt != "":
            self.unitpicker.Text = tt
        self.unitpicker.SelectionChanged += self.unitschanged
        self.textdecimals.MinValue = 0.0
        self.textdecimals.ValueChanged += self.unitschanged

        self.lfp.NumberOfDecimals = int(self.textdecimals.Value)
        self.unitschanged(None, None)
    
    def unitschanged(self, sender, e):

        # find the enum for the selected LinearType
        for e in range(0, 19):
            if LinearType(e) == self.unitpicker.SelectedItem.Content:
                self.outputunitenum = e
        
        # loop through all objects of self and set the properties for all DistanceEdits
        # the code is slower than doing it manually for each single one
        # but more convenient since we don't have to worry about how many DistanceEdit Controls we have in the UI
        tt = self.__dict__.items()
        for i in self.__dict__.items():
            if i[1].GetType() == TBCWpf.DistanceEdit().GetType():
                i[1].DisplayUnit = LinearType(self.outputunitenum)
                i[1].ShowControlIcon(False)
                i[1].FormatProperty.AddSuffix = ControlBoolean(1)
                i[1].FormatProperty.NumberOfDecimals = int(self.textdecimals.Value)

    def decdecimals_Click(self, sender, e):
        if not self.textdecimals.Value == 0:
            self.textdecimals.Value -= 1
             # setup the linear format properties
            self.lfp.NumberOfDecimals = int(self.textdecimals.Value)
            self.unitschanged(None, None)

    def incdecimals_Click(self, sender, e):
        self.textdecimals.Value += 1
        # setup the linear format properties
        self.lfp.NumberOfDecimals = int(self.textdecimals.Value)
        self.unitschanged(None, None)

    def tooutputunit(self, v):
        
        self.lfp.AddSuffix = self.addunitsuffix.IsChecked
        return self.lunits.Format(LinearType.Meter, v, self.lfp, LinearType(self.outputunitenum))

    def SetDefaultOptions(self):

        # layer
        lserial = OptionsManager.GetUint("SCR_PerpDistToDTM.layerpicker", 8)
        o = self.currentProject.Concordance.Lookup(lserial)
        if o != None:
            if isinstance(o.GetSite(), LayerCollection):    
                self.layerpicker.SetSelectedSerialNumber(lserial, InputMethod(3))
            else:                       
                self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:                       
            self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        
        # Select surface
        try:    self.surfacepicker.SelectIndex(OptionsManager.GetInt("SCR_PerpDistToDTM.surfacepicker", 0))
        except: self.surfacepicker.SelectIndex(0)


        self.usedtm.IsChecked = OptionsManager.GetBool("SCR_PerpDistToDTM.usedtm", True)
        self.computeplumb.IsChecked = OptionsManager.GetBool("SCR_PerpDistToDTM.computeplumb", False)
        try:    self.outcolorpicker.SelectedColor = Color.FromArgb(OptionsManager.GetInt("SCR_PerpDistToDTM.outcolorpicker"))
        except: self.outcolorpicker.SelectedColor = Color.Red
        
        self.usesinglepoint.IsChecked = OptionsManager.GetBool("SCR_PerpDistToDTM.usesinglepoint", False)
        self.showallsolutions.IsChecked = OptionsManager.GetBool("SCR_PerpDistToDTM.showallsolutions", False)
        self.checkBox_point.IsChecked = OptionsManager.GetBool("SCR_PerpDistToDTM.checkBox_point", False)
        self.checkBox_line.IsChecked = OptionsManager.GetBool("SCR_PerpDistToDTM.checkBox_line", True)

        self.checkBox_text.IsChecked = OptionsManager.GetBool("SCR_PerpDistToDTM.checkBox_text", True)
        self.unitpicker.Text = OptionsManager.GetString("SCR_PerpDistToDTM.unitpicker", "Meter")
        self.addunitsuffix.IsChecked = OptionsManager.GetBool("SCR_PerpDistToDTM.addunitsuffix", False)
        self.textdecimals.Value = OptionsManager.GetDouble("SCR_PerpDistToDTM.textdecimalsfloat", 4)
        self.textheight.Distance = OptionsManager.GetDouble("SCR_PerpDistToDTM.textheightfloat", 0.1)
        self.drawleader.IsChecked = OptionsManager.GetBool("SCR_PerpDistToDTM.drawleader", False)
        self.checkBox_showpointrl.IsChecked = OptionsManager.GetBool("SCR_PerpDistToDTM.checkBox_showpointrl", False)
        self.pointrlprefix.Text = OptionsManager.GetString("SCR_PerpDistToDTM.pointrlprefix", "GirderUS=")

        self.checkBox_showcomparerl.IsChecked = OptionsManager.GetBool("SCR_PerpDistToDTM.checkBox_showcomparerl", False)
        self.comparerlprefix.Text = OptionsManager.GetString("SCR_PerpDistToDTM.comparerlprefix", "DeckTOC=")

        self.checkBox_showdh.IsChecked = OptionsManager.GetBool("SCR_PerpDistToDTM.checkBox_showdh", True)

        self.checkBox_resultoffset.IsChecked = OptionsManager.GetBool("SCR_PerpDistToDTM.checkBox_resultoffset", False)
        self.resultoffset.Distance = OptionsManager.GetDouble("SCR_PerpDistToDTM.resultoffset", 0.0)

        self.checkBox_showoffset.IsChecked = OptionsManager.GetBool("SCR_PerpDistToDTM.checkBox_showoffset", False)
        self.offsetprefix.Text = OptionsManager.GetString("SCR_PerpDistToDTM.offsetprefix", "dhCorr=")

        self.dhprefix.Text = OptionsManager.GetString("SCR_PerpDistToDTM.dhprefix", "dh=")

    def SaveOptions(self):
        OptionsManager.SetValue("SCR_PerpDistToDTM.layerpicker", self.layerpicker.SelectedSerialNumber)

        OptionsManager.SetValue("SCR_PerpDistToDTM.usedtm", self.usedtm.IsChecked)
        OptionsManager.SetValue("SCR_PerpDistToDTM.computeplumb", self.computeplumb.IsChecked)
        OptionsManager.SetValue("SCR_PerpDistToDTM.surfacepicker", self.surfacepicker.SelectedIndex)
        OptionsManager.SetValue("SCR_PerpDistToDTM.outcolorpicker", self.outcolorpicker.SelectedColor.ToArgb())

        OptionsManager.SetValue("SCR_PerpDistToDTM.usesinglepoint", self.usesinglepoint.IsChecked)
        OptionsManager.SetValue("SCR_PerpDistToDTM.showallsolutions", self.showallsolutions.IsChecked)
        OptionsManager.SetValue("SCR_PerpDistToDTM.checkBox_point", self.checkBox_point.IsChecked)
        OptionsManager.SetValue("SCR_PerpDistToDTM.checkBox_line", self.checkBox_line.IsChecked)
 
        OptionsManager.SetValue("SCR_PerpDistToDTM.checkBox_text", self.checkBox_text.IsChecked)
        OptionsManager.SetValue("SCR_PerpDistToDTM.unitpicker", self.unitpicker.Text)
        OptionsManager.SetValue("SCR_PerpDistToDTM.addunitsuffix", self.addunitsuffix.IsChecked)
        OptionsManager.SetValue("SCR_PerpDistToDTM.textdecimalsfloat", abs(self.textdecimals.Value))
        OptionsManager.SetValue("SCR_PerpDistToDTM.textheightfloat", self.textheight.Distance)
        OptionsManager.SetValue("SCR_PerpDistToDTM.drawleader", self.drawleader.IsChecked)
        OptionsManager.SetValue("SCR_PerpDistToDTM.checkBox_showpointrl", self.checkBox_showpointrl.IsChecked)
        OptionsManager.SetValue("SCR_PerpDistToDTM.pointrlprefix", self.pointrlprefix.Text)

        OptionsManager.SetValue("SCR_PerpDistToDTM.checkBox_showcomparerl", self.checkBox_showcomparerl.IsChecked)
        OptionsManager.SetValue("SCR_PerpDistToDTM.comparerlprefix", self.comparerlprefix.Text)

        OptionsManager.SetValue("SCR_PerpDistToDTM.checkBox_showdh", self.checkBox_showdh.IsChecked)

        OptionsManager.SetValue("SCR_PerpDistToDTM.checkBox_resultoffset", self.checkBox_resultoffset.IsChecked)
        OptionsManager.SetValue("SCR_PerpDistToDTM.resultoffset", self.resultoffset.Distance)

        OptionsManager.SetValue("SCR_PerpDistToDTM.checkBox_showoffset", self.checkBox_showoffset.IsChecked)
        OptionsManager.SetValue("SCR_PerpDistToDTM.offsetprefix", self.offsetprefix.Text)

        OptionsManager.SetValue("SCR_PerpDistToDTM.dhprefix", self.dhprefix.Text)

    def Coord1Changed(self, ctrl, e):
        if e.Cause == InputMethod.Mouse:     
            self.OkClicked(None, None)
        Keyboard.Focus(self.coordCtl1)
        
    def CancelClicked(self, thisCmd, args):
        thisCmd.CloseUICommand()

    def OkClicked(self, thisCmd, e):
        Keyboard.Focus(self.okBtn)

        self.error.Content = ''
        self.success.Content = ''
        self.label_benchmark.Content = ''

        wv = self.currentProject[Project.FixedSerial.WorldView]
        wv.PauseGraphicsCache(True)

        pnew = Point3D()

        inputok=True
        try:
            textdecimals = str(abs(int(self.textdecimals.Value)))
            #tt=abs(float(self.textheight.Distance))
        except:
            inputok = False

        if inputok:
            self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
            UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
            ProgressBar.TBC_ProgressBar.Title = self.Caption
            try:
                with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                    pointlist = [] # Point3DArray()

                    # create a list of points we want to check
                    if self.usesinglepoint.IsChecked: 
                        vertice1_sel = self.coordCtl1.Coordinate
                        if math.isnan(vertice1_sel.Z):
                            self.coordCtl1.StatusMessage = "No coordinate defined"
                            # return
                        else:
                            pointlist.Add(vertice1_sel)
                    else:
                        for o in self.objs.SelectedMembers(self.currentProject):
                            if isinstance(o, self.cadpointType):
                                pointlist.Add(o.Point0)
                            elif isinstance(o, self.coordpointType):
                                pointlist.Add(o.Position)
                            elif isinstance(o, self.pointcloudType):
                                integration = o.Integration  # = SdePointCloudRegionIntegration
                                selectedid = integration.GetSelectedCloudId() # it seems the selected points form a sub-cloud

                                regiondb = integration.PointCloudDatabase # PointCloudDatabase
                                sdedb = regiondb.Integration # SdePointCloudDatabaseIntegration
                                scanpointlist = sdedb.GetPoints(selectedid)
                                for p in scanpointlist:
                                    pointlist.Add(p)

                    # create a list of triangle vertices
                    vertexlist = self.createvertexlist()

                    # compute the distances
                    start_t = timer()

                    surface = self.currentProject.Concordance.Lookup(self.surfacepicker.SelectedSerial)
                    if isinstance(surface, ProjectedSurface):
                        projected=True
                    else:
                        projected=False

                    j = 0
                    for vertice1_sel in pointlist:
                        if not math.isnan(vertice1_sel.Z):
                            pnew = Point3D(0,0,0)
                            
                            dtmresults = []
                            closest_pnew = None

                            # compute plumb if it's a dtm
                            if self.usedtm.IsChecked and self.computeplumb.IsChecked and projected == False:
                                outPoint = clr.StrongBox[Point3D]()
                                if surface.PickSurface(vertice1_sel, outPoint):
                                    dtmresults.Add(outPoint.Value)
                                    closest_pnew = outPoint.Value

                            else: # compute perpendicular with all other types of objects

                                multithread = False
                                computeplumb = self.computeplumb.IsChecked
                                listlock = Lock()
                                debuglist = []

                                if multithread:

                                    # use multithreading to check all triangles for a solution
                                    threadcount = cpu_count()
                                    
                                    # limit threadcount if there are more than work to do, otherwise we'd end up with double-ups in the result list
                                    if threadcount > vertexlist.Count / 3:
                                        threadcount = int(vertexlist.Count / 3)
                                    
                                    #threadcount = 20 # debug manual thread limiter
                                    self.error.Content += '\n' + str(vertexlist.Count / 3) + ' Triangles - ' + str(threadcount) + ' Threads used'

                                    self.exc_info = None

                                    start_t_threads = timer()

                                    threads = [Thread(target = self.perpdisttotriangle, args = (debuglist, listlock, computeplumb, vertice1_sel, vertexlist, dtmresults, threadcount, threadindex,)) for threadindex in range(threadcount)]

                                    # start threads
                                    for thread in threads:
                                        thread.start()

                                    end_t_threads = timer()
                                    self.label_benchmark.Content += '\nThreads Start - ' + str(timedelta(seconds=end_t_threads - start_t_threads))

                                    start_t_join = timer()

                                    # wait for all threads to terminate
                                    for thread in threads:
                                        thread.join()

                                    end_t_join = timer()
                                    self.label_benchmark.Content += '\nThreads Join - ' + str(timedelta(seconds=end_t_join - start_t_join))


                                    if self.exc_info:
                                        exc_type, exc_obj, exc_tb = self.exc_info
                                        self.error.Content += '\nan Error occurred - Result probably incomplete\n' + str(exc_type) + '\n' + str(exc_obj) + '\nLine ' + str(exc_tb.tb_lineno)
                                    
                                else: # no multithreading

                                    self.perpdisttotriangle(debuglist, listlock, computeplumb, vertice1_sel, vertexlist, dtmresults, 1, 0)

                                for t in debuglist:
                                    self.error.Content += t

                                for pnew in dtmresults:

                                    if self.showallsolutions.IsChecked: # if show all solutions is checked then draw the text etc. for all found solutions
                                        self.drawtext(vertice1_sel, pnew, Vector3D(vertice1_sel, pnew).Length)
                                    else:   # if we only want the text for the shortest solution
                                        if not closest_pnew: # if it's the first result we find
                                            closest_pnew = Point3D(pnew)
                                        else:
                                            if Vector3D(vertice1_sel, pnew).Length < Vector3D(vertice1_sel, closest_pnew).Length:
                                                closest_pnew = Point3D(pnew)


                            # now that we checked all triangles we create the text if we only want the shortest solution
                            if self.showallsolutions.IsChecked == False and self.computeplumb.IsChecked == False and closest_pnew:
                                self.drawtext(vertice1_sel, closest_pnew, Vector3D(vertice1_sel, closest_pnew).Length)
                            if self.usedtm.IsChecked and self.computeplumb.IsChecked and projected == False and closest_pnew:
                                self.drawtext(vertice1_sel, closest_pnew, Vector3D(vertice1_sel, closest_pnew).Length)

                            # we update the progress bar
                            j += 1
                            if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(j * 100 / pointlist.Count)):
                                break   # function returns true if user pressed cancel

                        end_t=timer()
                        self.label_benchmark.Content += '\nOverall - ' + str(timedelta(seconds=end_t-start_t))
                    
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
        
        wv.PauseGraphicsCache(False)
        self.SaveOptions()

    def perpdisttotriangle(self, debuglist, listlock, computeplumb, vertice1_sel, vertexlist, dtmresults, threadcount, threadindex):

        try:
            trianglesperthread = math.ceil(vertexlist.Count / 3.0 / threadcount) # must be 3.0 otherwise it won't be a float and is rounded down before I can round it up
            with listlock:
                debuglist.Add('\nvertexlistcount ' + str(vertexlist.Count) + ' - Triangles per Thread ' + str(trianglesperthread))

            istart = int(0 + threadindex * trianglesperthread * 3)
            iend = int(0 + threadindex * trianglesperthread * 3 + trianglesperthread * 3)


            if iend > vertexlist.Count:
                iend = int(vertexlist.Count)

            with listlock:
                debuglist.Add('\nThreadindex ' + str(threadindex) + ' - ' + str(istart) + ' - ' + str(iend))

            for i in range(istart, iend, 3):

                ### v1 = vertexlist[i]
                ### v2 = vertexlist[i+1]
                ### v3 = vertexlist[i+2]

                ### p = Plane3D(v1,v2,v3)
                p = Plane3D(vertexlist[i], vertexlist[i+1], vertexlist[i+2])[0] # the plane is returned as first element

                if not p.IsValid:
                    return
                ### nv = p.normal
                ### nv.Normalize()
                ### #   we use the normal vector to fill the "Hesse" plane formula
                ### #   since we use the normalized vector the denominator is 1 and we can ignore it 
                ### 
                ### #   v.x * x  +  v.y * y + v.z * z
                ### #   -----------------------------  =  d  =  nx * x  +  ny * y + nz * z
                ### #             |v| = 1
                ### 
                ### #  now we use our point instead one of the vertexes in the plane formula and the
                ### #  resulting value is the 3D distance to that plane
                ### 
                ### #  d3d =  d - nx * x  +  ny * y + nz * z
                ### 
                ### 
                ### d = (v1.X * nv.X + v1.Y * nv.Y + v1.Z * nv.Z)
                ### d3d = (d - vertice1_sel.X * nv.X - vertice1_sel.Y * nv.Y - vertice1_sel.Z * nv.Z)
                ### if d3d == 0:
                ###     tt = d3d
                ### # if round(d3d,6) != 0:
                ### pnew.X = vertice1_sel.X + d3d * nv.X
                ### pnew.Y = vertice1_sel.Y + d3d * nv.Y
                ### pnew.Z = vertice1_sel.Z + d3d * nv.Z
                
                if computeplumb:
                    pnew = Plane3D.IntersectWithRay(p, vertice1_sel, Vector3D(0, 0, 1))
                else:
                    pnew = Plane3D.IntersectWithRayPerpendicular(p, vertice1_sel)

                # we only want results were the intersecting point is within the tested triangle
                # otherwise we get hundreds of false results which don't lie on the DTM

                #if Triangle2D.IsPointInside(v1,v2,v3,pnew)[0] == True:
                ### if PointInsideTriangle3D(v1,v2,v3,pnew):
                if pnew != Point3D.Undefined and PointInsideTriangle3D(vertexlist[i], vertexlist[i+1], vertexlist[i+2], pnew):    
                    
                    with listlock:   
                        dtmresults.Add(pnew)

        except Exception as e:
            self.exc_info = sys.exc_info()


    def drawtext(self, vertice1_sel, pnew, d3d):
        
        wv = self.currentProject[Project.FixedSerial.WorldView]
        if self.checkBox_point.IsChecked:
            cadPoint = wv.Add(clr.GetClrType(CadPoint))
            cadPoint.Point0 = pnew
            cadPoint.Layer = self.layerpicker.SelectedSerialNumber
            cadPoint.Color = self.outcolorpicker.SelectedColor

        if self.checkBox_line.IsChecked:
            l = wv.Add(clr.GetClrType(Linestring))
            e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
            e.Position = vertice1_sel 
            l.AppendElement(e)
            e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
            e.Position = pnew
            l.AppendElement(e)
            l.Layer = self.layerpicker.SelectedSerialNumber
            l.Color = self.outcolorpicker.SelectedColor

        if self.checkBox_text.IsChecked:
            t = wv.Add(clr.GetClrType(MText))

            # add point RL if ticked
            if self.checkBox_showpointrl.IsChecked:
                t.TextString = self.addtotextstring(t.TextString, self.pointrlprefix.Text)
                t.TextString += self.tooutputunit(vertice1_sel.Z)

            # add compare RL if ticked
            if self.checkBox_showcomparerl.IsChecked:
                t.TextString = self.addtotextstring(t.TextString, self.comparerlprefix.Text)
                t.TextString += self.tooutputunit(pnew.Z)

            # add dh if ticked
            if self.checkBox_showdh.IsChecked:

                # add value offset if ticked
                if self.checkBox_resultoffset.IsChecked:
                    d3d += self.resultoffset.Distance
                #convert from metric to whatever the display unit is
                #d3d = self.lunits.Convert(self.lunits.InternalType, abs(d3d), self.lunits.DisplayType)

                # add correction value if ticked
                if self.checkBox_resultoffset.IsChecked and self.checkBox_showoffset.IsChecked:
                    t.TextString = self.addtotextstring(t.TextString, self.offsetprefix.Text)
                    t.TextString += self.tooutputunit(self.resultoffset.Distance)

                # add actual dh
                t.TextString = self.addtotextstring(t.TextString, self.dhprefix.Text)
                t.TextString += self.tooutputunit(d3d)


            th = abs(self.textheight.Distance) # project units
            # we need to convert it to internal metres
            #th = self.lunits.Convert(self.lunits.DisplayType, th, self.lunits.InternalType)
            t.Height = th
            t.Point0 = pnew
            t.Layer = self.layerpicker.SelectedSerialNumber
            t.Color = self.outcolorpicker.SelectedColor

            if self.drawleader.IsChecked:
                leaderpoints = List[Point3D]()
                leaderpoints.Add(t.AlignmentPoint)
                t.AlignmentPoint = Point3D(t.AlignmentPoint.X + 0.1, t.AlignmentPoint.Y - 0.1, t.AlignmentPoint.Z)
                leaderpoints.Add(t.AlignmentPoint)
                
                l = wv.Add(clr.GetClrType(Leader))
                l.AnnotationSerial = t.SerialNumber
                l.ScaleFactor = 1
                l.Points = leaderpoints
                l.Layer = self.layerpicker.SelectedSerialNumber
                l.LeaderType = LeaderType.LineNoArrow
                l.ArrowheadType = DimArrowheadType.ArrowDefault
                l.ArrowheadSize = 0.0
                l.TextGap = 0.1
                l.Color = Color.DarkGray
                #l.AnnotationTargetDelta = t.AlignmentPoint

    def addtotextstring(self, t, addstring):
        if len(t) > 0:
            t += '\\P' + addstring
        else:
            t += addstring

        return t

    def createvertexlist(self):
        # create a list of triangle vertices
        vertexlist = []
        if self.usedtm.IsChecked:

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
    
                 
def PointInsideTriangle3D(vertice1,p2,p3,p4):
    # see also https://blackpawn.com/texts/pointinpoly/

    a = Vector3D(p2.X-vertice1.X, p2.Y-vertice1.Y, p2.Z-vertice1.Z)     # v1
    b = Vector3D(p3.X-vertice1.X, p3.Y-vertice1.Y, p3.Z-vertice1.Z)     # v0
    w = Vector3D(p4.X-vertice1.X, p4.Y-vertice1.Y, p4.Z-vertice1.Z)     # v2

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


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
    cmdData.Key = "SCR_ViewFilter"
    cmdData.CommandName = "SCR_ViewFilter"
    cmdData.Caption = "_SCR_ViewFilter"
    #cmdData.UIForm = "SCR_ViewFilter"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
                                                      # if you enable or disable this line, you MUST restart TBC
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.Version = 1.05
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.DefaultTabKey = "SCR Expld-SNR-Relay-Prop"
        cmdData.DefaultTabGroupKey = "Properties"
        cmdData.ShortCaption = "ViewFilter"
        cmdData.ToolTipTitle = "advanced ViewFilter setting"
        cmdData.ToolTipTextFormatted = "advanced ViewFilter setting with list filtering"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3


    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass

def Execute(cmd, currentProject, macroFileFolder, parameters):
    form = SCR_ViewFilterDialog(currentProject, macroFileFolder).Show()
    return
    # .Show() - is non modal - you can interact with the drawing window
    # .ShowDialog() - is modal - you CAN NOT interact with the drawing window

class SCR_ViewFilterDialog(Window): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):

        with StreamReader(macroFileFolder + r"\SCR_ViewFilter.xaml") as s:
            wpf.LoadComponent(self, s) 

        ElementHost.EnableModelessKeyboardInterop(self)
        
        self.currentProject = currentProject
        self.macroFileFolder = macroFileFolder

        # Description is the actual property name inside the item in ticklist.Content.Items
        sd = System.ComponentModel.SortDescription("Description", System.ComponentModel.ListSortDirection.Ascending)

        lc = self.currentProject.Concordance.Lookup(Project.FixedSerial.LayerContainer)
        wv = self.currentProject[Project.FixedSerial.WorldView]
        self.oldlayercount = lc.List[Layer]().Count
        self.oldsurfacecount = wv.List[Model3D]().Count

        self.layerticklist1.SearchContainer = Project.FixedSerial.LayerContainer
        self.layerticklist1.UseSelectionEngine = False
        self.layerticklist1.SetEntityType(clr.GetClrType(Layer), self.currentProject)
        self.layerticklist1.Content.Items.SortDescriptions.Add(sd)
        self.layerticklist2.SearchContainer = Project.FixedSerial.LayerContainer
        self.layerticklist2.UseSelectionEngine = False
        self.layerticklist2.SetEntityType(clr.GetClrType(Layer), self.currentProject)
        self.layerticklist2.Content.Items.SortDescriptions.Add(sd)

        self.surfacetypes = Array[Type](SurfaceTypeLists.AllWithCutFillMap)+Array[Type]([clr.GetClrType(ProjectedSurface)])
        self.surfaceticklist1.SearchContainer = Project.FixedSerial.WorldView
        self.surfaceticklist1.UseSelectionEngine = False
        self.surfaceticklist1.SetEntityType(self.surfacetypes, self.currentProject)
        self.surfaceticklist1.Content.Items.SortDescriptions.Add(sd)
        self.surfaceticklist2.SearchContainer = Project.FixedSerial.WorldView
        self.surfaceticklist2.UseSelectionEngine = False
        self.surfaceticklist2.SetEntityType(self.surfacetypes, self.currentProject)
        self.surfaceticklist2.Content.Items.SortDescriptions.Add(sd)

        self.layerfilter1.TextChanged += self.userfilterchangeevent
        self.layerfilter2.TextChanged += self.userfilterchangeevent
        self.surfacefilter1.TextChanged += self.userfilterchangeevent
        self.surfacefilter2.TextChanged += self.userfilterchangeevent

        self.buttongetstarted.Click += self.debuggetstartedevent

        self.layerticklist1.ValueChanged += self.tickboxchangeevent
        self.layerticklist2.ValueChanged += self.tickboxchangeevent
        self.surfaceticklist1.ValueChanged += self.tickboxchangeevent
        self.surfaceticklist2.ValueChanged += self.tickboxchangeevent

        self.previous_event_time = timer()
        self.previous_ticks_event_time = timer()
        self.previous_tickbox_content_event_time = timer()

        self.activeForm = TrimbleOffice.TheOffice.MainWindow.AppViewManager.ActiveView
        try:
            self.activeViewFilter = self.currentProject.Concordance.Lookup(self.activeForm.ViewFilter)  
            self.activeViewFilter.FilterChanged += self.viewfiltersettingschangeevent
        except:
            pass

        self.entityisdeleting = False # must establish here before all the event triggers that rely on it
        self.undoredoinprogress = False

        TrimbleOffice.TheOffice.MainWindow.AppViewManager.ViewActivated += self.activeviewchanged
        UIEvents.UIViewFilterChanged += self.filternameforactiveviewchanged # i.e. if View Filter for a window changed
        self.activeviewchanged(None, None) # get the active window on startup

        ModelEvents.ProjectUnloading += self.projectischanging # close macro on project change
        ModelEvents.EntityDeleting += self.modeldeletestart
        CalculateProjectEvents.CalculateProjectEvent += self.projectcalc

        UIEvents.BeforeDataProcessing += self.backgroundupdatestart
        UIEvents.AfterDataProcessing += self.backgroundupdateend


        #self.firstrun = True

        self.Loaded += self.SetDefaultOptions
        self.Closing += self.SaveOptions


        # calculate project is triggered before EntityAdded
        # create layer - calculate, calculate
        # create surface - calculate, calculate
        # delete layer triggers - entityadded, EntityDeleting, calculate,
        # delete surface triggers - calculate, entityadded, EntityDeleting, entityadded, EntityDeleting, EntityAdded, EntityDeleting, calculate

    def backgroundupdatestart(self, sender, e):
        tt = 1
        try:
            if sender.CommandName == "Undo" or sender.CommandName == "Redo":
                self.undoredoinprogress = True
        except:
            pass

    def backgroundupdateend(self, sender, e):
        tt = 1
        try:
            if sender.CommandName == "Undo" or sender.CommandName == "Redo":
                self.undoredoinprogress = False
                #self.viewfiltersettingschangeevent(None, None) # will trigger self.updateticksinlists("1,2,3,4")

                # trigger resetting ticks via manual triggering a "userfilterchangeevent"
                # need to do it this way to properly trigger ticks after the itemlist was rebuild
                self.tickboxfilterchange(self.layerfilter1)
                self.tickboxfilterchange(self.layerfilter2)
                self.tickboxfilterchange(self.surfacefilter1)
                self.tickboxfilterchange(self.surfacefilter2)
        except:
            pass

    def modeldeletestart(self, sender, e):
        # only worry about changes to Layers or Surfaces
        # and only deletion - need to avoid addressing serial numbers that may not exist anymore
        # adding entities isn't really a problem
        try:
            sn = e.EntitySerialNumber
            o = self.currentProject.Concordance.Lookup(sn)
            if isinstance(o, Layer) or isinstance(o, Model3D):
                self.disableeventtriggers(True)
                self.entityisdeleting = True
        except:
            pass

    def projectcalc(self, sender, e):

        if self.entityisdeleting: # if the recalc follows a delete event
            self.disableeventtriggers(False)
            self.entityisdeleting = False

        else: # otherwise it could be creating a new layer or surface
            lc = self.currentProject.Concordance.Lookup(Project.FixedSerial.LayerContainer)
            wv = self.currentProject[Project.FixedSerial.WorldView]
            currentlayercount = lc.List[Layer]().Count
            currentsurfacecount = wv.List[Model3D]().Count
            if self.oldlayercount < currentlayercount or self.oldsurfacecount < currentsurfacecount:
                self.oldlayercount = currentlayercount
                self.oldsurfacecount = currentsurfacecount

                self.tickboxfilterchange(self.layerfilter1)
                self.tickboxfilterchange(self.layerfilter2)
                self.tickboxfilterchange(self.surfacefilter1)
                self.tickboxfilterchange(self.surfacefilter2)
                

    def Dispose(self, thisCmd, disposing):
        pass

    def debuggetstartedevent(self, sender, e):

        #self.debug.Content += 'get started, '
        
        self.viewfiltersettingschangeevent(None, None)

    def projectischanging(self, sender, e):
        self.Close()

    def cleanupeventtriggers(self):
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.ViewActivated -= self.activeviewchanged
        UIEvents.UIViewFilterChanged -= self.filternameforactiveviewchanged
        self.activeViewFilter.FilterChanged -= self.viewfiltersettingschangeevent

        ModelEvents.EntityDeleting -= self.modeldeletestart
        CalculateProjectEvents.CalculateProjectEvent -= self.projectcalc

        UIEvents.BeforeDataProcessing -= self.backgroundupdatestart
        UIEvents.AfterDataProcessing -= self.backgroundupdateend

        ModelEvents.ProjectUnloading -= self.projectischanging

        self.layerfilter1.TextChanged -= self.userfilterchangeevent
        self.layerfilter2.TextChanged -= self.userfilterchangeevent
        self.surfacefilter1.TextChanged -= self.userfilterchangeevent
        self.surfacefilter2.TextChanged -= self.userfilterchangeevent

        self.buttongetstarted.Click -= self.debuggetstartedevent

        self.layerticklist1.ValueChanged -= self.tickboxchangeevent
        self.layerticklist2.ValueChanged -= self.tickboxchangeevent
        self.surfaceticklist1.ValueChanged -= self.tickboxchangeevent
        self.surfaceticklist2.ValueChanged -= self.tickboxchangeevent

        tt = 1

    def tickboxchangeevent(self, sender, e):

        #self.debug.Content += 'layer set by macro, '

        # a layer delete is not always trigering the model delete event
        # need to catch it tis way
        # if the items list and the serial number count are different it's very likely that a delete is in progress

        if sender.EntitySerialNumbers.Count == sender.Content.Items.Count:
            
            #self.disableeventtriggers(True)

            if not sender.LastInputMethod == InputMethod.Programmatically:
                self.setlayersfrommacro(sender)

                if sender.Name == "layerticklist1":
                    self.updateticksinlists("2")
                elif sender.Name == "layerticklist2":
                    self.updateticksinlists("1")
                elif sender.Name == "surfaceticklist1":
                    self.updateticksinlists("4")
                elif sender.Name == "surfaceticklist2":
                    self.updateticksinlists("3")

            sender.LastInputMethod = InputMethod(0)
            tt = 1

            #self.disableeventtriggers(False)

    def userfilterchangeevent(self, sender, e):
        
        if self.activeViewFilter and not self.entityisdeleting and not self.undoredoinprogress:
            
            self.tickboxfilterchange(sender) # compiles and applies exclude list to listbox


    def activeviewchanged(self, sender, e):

        #self.debug.Content += 'active view was changed, '
        hasviewfilter = True
        try:
            newactiveForm = None
            newactiveForm = TrimbleOffice.TheOffice.MainWindow.AppViewManager.ActiveView
            tt = self.currentProject.Concordance.Lookup(newactiveForm.ViewFilter)
            if newactiveForm.ViewFilter == 0 or newactiveForm.ViewFilter == 22: # 22 is View Filter "All" which we can't change
                hasviewfilter = False
        except:
            hasviewfilter = False
        
        if hasviewfilter:
            self.debug2.Content = ""
            self.filterandlists.IsEnabled = True
            self.activeForm = TrimbleOffice.TheOffice.MainWindow.AppViewManager.ActiveView

            try:
                self.activeViewFilter.FilterChanged -= self.viewfiltersettingschangeevent
            except:
                tt = 1
                pass
            self.activeViewFilter = self.currentProject.Concordance.Lookup(self.activeForm.ViewFilter)   
            self.activeViewFilter.FilterChanged += self.viewfiltersettingschangeevent

            self.labelproject.Content = "Project: " + self.currentProject.FileName
            self.labelviewfilter.Content = "View Filter: " + self.activeViewFilter.Name            

            self.viewfiltersettingschangeevent(None, None) # a full view filter change is the same as if a layer was ticked on/off in the view filter manager

        else:
            self.filterandlists.IsEnabled = False # grey out the user interface
            self.debug2.Content = "current View doesn't have a changeable View-Filter"

        
    def filternameforactiveviewchanged(self, sender, e):
        
        # is in the end like changing the view
        self.activeviewchanged(None, None)

    def viewfiltersettingschangeevent(self, sender, e):
        
        #self.debug.Content += 'viewfilter setting changed within TBC, '
        if self.activeViewFilter and not self.entityisdeleting and not self.undoredoinprogress:
            
            #self.disableeventtriggers(True)
            
            self.updateticksinlists("1,2,3,4")
            
            #self.disableeventtriggers(False)

    def disableeventtriggers(self, dis):

        if dis:
            self.layerticklist1.ValueChanged -= self.tickboxchangeevent
            self.layerticklist2.ValueChanged -= self.tickboxchangeevent
            self.surfaceticklist1.ValueChanged -= self.tickboxchangeevent
            self.surfaceticklist2.ValueChanged -= self.tickboxchangeevent

            self.activeViewFilter.FilterChanged -= self.viewfiltersettingschangeevent
            UIEvents.UIViewFilterChanged -= self.filternameforactiveviewchanged
        else:
            self.layerticklist1.ValueChanged += self.tickboxchangeevent
            self.layerticklist2.ValueChanged += self.tickboxchangeevent
            self.surfaceticklist1.ValueChanged += self.tickboxchangeevent
            self.surfaceticklist2.ValueChanged += self.tickboxchangeevent

            self.activeViewFilter.FilterChanged += self.viewfiltersettingschangeevent
            UIEvents.UIViewFilterChanged += self.filternameforactiveviewchanged

    def setlayersfrommacro(self, sender):

        self.disableeventtriggers(True)


        current_event_time = timer()
        tt = self.previous_event_time
        if current_event_time - self.previous_event_time > 0.1 and not self.entityisdeleting and not self.undoredoinprogress:

            self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, "SCR_ViewFilter")#self.Caption)
            UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
            try:
                # the "with" statement will unroll any changes if something go wrong
                with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                    if "layer" in sender.Name:                    
                        lc = self.currentProject.Concordance.Lookup(Project.FixedSerial.LayerContainer)
                        # go through all existing layers
                        for l in lc:
                            # if the current viewfilter isn't aware of the layer yet then add it first - can be if it's brand new
                            if not self.activeViewFilter.LayerOverrides.Contains(l.SerialNumber):
                                self.activeViewFilter.LayerOverrides.Add(l.SerialNumber)
                            
                        # now set the visibility based on the items list from the event sender, which is filtered
                        for item in sender.Content.ItemsSource:
                                            
                            layeroverride = self.activeViewFilter.LayerOverrides[item.EntitySerialNumber] # get the layer settings for the layer in the current view filter
                            # set the visibility based on the appearance in the list from the macro
                            if layeroverride.Visible != item.Checked:
                                layeroverride.Visible = item.Checked

                    elif "surface" in sender.Name: 
                        wv = self.currentProject[Project.FixedSerial.WorldView]
                        surfaces = wv.List[Model3D]()
        
                        imicoll = self.currentProject.Concordance.Lookup(20) # is a fixed serial
                        try:
                            for s in surfaces:
                                # https://community.trimble.com/discussion/how-to-disableenable-layerssurfaces-with-a-macro
                                imitypeserial = imicoll[s.ImitationTypeGuid].SerialNumber

                                # if the current viewfilter isn't aware of the layer yet then add it first - can be if it's brand new
                                if not self.activeViewFilter.ImitationTypeOverrides.Contains(imitypeserial):
                                    self.activeViewFilter.ImitationTypeOverrides.Add(imitypeserial)

                            # now set the visibility based on the items list from the event sender, which is filtered
                            for item in sender.Content.ItemsSource:
                                
                                s = self.currentProject.Concordance.Lookup(item.EntitySerialNumber)
                                imitypeserial = imicoll[s.ImitationTypeGuid].SerialNumber

                                layeroverride = self.activeViewFilter.ImitationTypeOverrides[imitypeserial] # get the layer settings for the layer in the current view filter
                                # set the visibility based on the appearance in the list from the macro
                                if layeroverride.Visible != item.Checked:
                                    layeroverride.Visible = item.Checked
                        except:
                            pass

                    failGuard.Commit()
                    UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
                    self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
            
            except Exception as err:
                tt = sys.exc_info()
                exc_type, exc_obj, exc_tb = sys.exc_info()
                # EndMark MUST be set no matter what
                # otherwise TBC won't work anymore and needs to be restarted
                self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
                UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
                self.debug2.Content += '\nan Error occurred - Result probably incomplete\n' + str(exc_type) + '\n' + str(exc_obj) + '\nLine ' + str(exc_tb.tb_lineno)


            self.previous_event_time = timer()

        self.disableeventtriggers(False)


    def updateticksinlists(self, boxtoupdate):

        self.disableeventtriggers(True)
        
        #self.debug.Content += 'updateticksinlist, '

        # in case the list gets rebuild by a filter change

        tt = self.previous_ticks_event_time
        current_ticks_event_time = timer()
        if current_ticks_event_time - self.previous_ticks_event_time > 0.1  and not self.entityisdeleting and not self.undoredoinprogress:

            if "1" in boxtoupdate or "2" in boxtoupdate:
            
                lc = self.currentProject.Concordance.Lookup(Project.FixedSerial.LayerContainer)
                selectedserials = List[UInt32]()
                # go through all layers and make sure the viewfilter knows about it
                for l in lc:
                    if not self.activeViewFilter.LayerOverrides.Contains(l.SerialNumber):
                        self.activeViewFilter.LayerOverrides.Add(l.SerialNumber)
                    # now that the viewfilter definitely knows the layer we can get its visibility state
                    layeroverride = self.activeViewFilter.LayerOverrides[l.SerialNumber] 
                    if layeroverride.Visible:
                        selectedserials.Add(l.SerialNumber) # if it's visible than add it to the list we use to set the ticks in the list
                
                if "1" in boxtoupdate:
                    self.layerticklist1.SelectedSerialsX = selectedserials
                    #self.layerticklist1.ReloadEntities()
                if "2" in boxtoupdate:
                    self.layerticklist2.SelectedSerialsX = selectedserials
                    #self.layerticklist2.ReloadEntities()

            if "3" in boxtoupdate or "4" in boxtoupdate:

                # go through the surfaces
                selectedserials = List[UInt32]()
                wv = self.currentProject[Project.FixedSerial.WorldView]
                surfaces = wv.List[Model3D]()
                tt = self.activeViewFilter.ImitationTypeOverrides
                imicoll = self.currentProject.Concordance.Lookup(20) # is a fixed serial
                fails = 0
                for s in surfaces:
                    # https://community.trimble.com/discussion/how-to-disableenable-layerssurfaces-with-a-macro
                    try:
                        imitypeserial = imicoll[s.ImitationTypeGuid].SerialNumber

                        # if the current viewfilter isn't aware of the layer yet then add it first - can be if it's brand new
                        if not self.activeViewFilter.ImitationTypeOverrides.Contains(imitypeserial):
                            self.activeViewFilter.ImitationTypeOverrides.Add(imitypeserial)
                        # now that the viewfilter definitely knows the layer we can get its visibility state
                        layeroverride = self.activeViewFilter.ImitationTypeOverrides[imitypeserial] 
                        if layeroverride.Visible:
                            selectedserials.Add(s.SerialNumber) # if it's visible than add it to the list we use to set the ticks in the list
                    except:
                        fails += 1
                        pass

                if "3" in boxtoupdate:
                    #for sn in selectedserials:
                    #    self.surfaceticklist1.SelectBySerialNumber(sn)
                    self.surfaceticklist1.SelectedSerialsX = selectedserials
                    #self.surfaceticklist1.ReloadEntities()
                if "4" in boxtoupdate:
                    #for sn in selectedserials:
                    #    self.surfaceticklist2.SelectBySerialNumber(sn)
                    self.surfaceticklist2.SelectedSerialsX = selectedserials
                    #self.surfaceticklist2.ReloadEntities()

            self.previous_ticks_event_time = timer()

        self.disableeventtriggers(False)

    def tickboxfilterchange(self, ctrl):

        # compiles and applies exclude list to listbox

        if ctrl.Name == "layerfilter1":
            exclude = []
            self.layerticklist1.LastInputMethod = InputMethod.Programmatically
            self.disableeventtriggers(True)
            self.layerticklist1.SetExcludedEntities(exclude)
            self.disableeventtriggers(False)
            
            tt = self.layerfilter1.Text.lower()
            ticklistfilter = tt.split()
            
            for i in self.layerticklist1.EntitySerialNumbers:
                for f in ticklistfilter:
                   if not f in i.Key.Description.lower():
                        exclude.Add(i.Value)
            
            tt = self.layerticklist1
            self.layerticklist1.Content.ItemContainerGenerator.StatusChanged += self.layer1populated            
            self.disableeventtriggers(True)
            self.layerticklist1.SetExcludedEntities(exclude)
            self.disableeventtriggers(False)

        elif ctrl.Name == "layerfilter2":
            exclude = []
            self.layerticklist2.LastInputMethod = InputMethod.Programmatically
            self.disableeventtriggers(True)
            self.layerticklist2.SetExcludedEntities(exclude)
            self.disableeventtriggers(False)
            
            tt = self.layerfilter2.Text.lower()
            ticklistfilter = tt.split()
            
            for i in self.layerticklist2.EntitySerialNumbers:
                for f in ticklistfilter:
                   if not f in i.Key.Description.lower():
                        exclude.Add(i.Value)

            self.layerticklist2.Content.ItemContainerGenerator.StatusChanged += self.layer2populated  
            self.disableeventtriggers(True)
            self.layerticklist2.SetExcludedEntities(exclude)
            self.disableeventtriggers(False)

        elif ctrl.Name == "surfacefilter1":
            exclude = []
            self.surfaceticklist1.LastInputMethod = InputMethod.Programmatically
            self.disableeventtriggers(True)
            self.surfaceticklist1.SetExcludedEntities(exclude)
            self.disableeventtriggers(False)
            
            tt = self.surfacefilter1.Text.lower()
            ticklistfilter = tt.split()
            
            for i in self.surfaceticklist1.EntitySerialNumbers:
                for f in ticklistfilter:
                   if not f in i.Key.Description.lower():
                        exclude.Add(i.Value)

            self.surfaceticklist1.Content.ItemContainerGenerator.StatusChanged += self.surface1populated  
            self.disableeventtriggers(True)
            self.surfaceticklist1.SetExcludedEntities(exclude)
            self.disableeventtriggers(False)

        elif ctrl.Name == "surfacefilter2":
            exclude = []
            self.surfaceticklist2.LastInputMethod = InputMethod.Programmatically
            self.disableeventtriggers(True)
            self.surfaceticklist2.SetExcludedEntities(exclude)
            self.disableeventtriggers(False)
            
            tt = self.surfacefilter2.Text.lower()
            ticklistfilter = tt.split()
            
            for i in self.surfaceticklist2.EntitySerialNumbers:
                for f in ticklistfilter:
                   if not f in i.Key.Description.lower():
                        exclude.Add(i.Value)
            
            self.surfaceticklist2.Content.ItemContainerGenerator.StatusChanged += self.surface2populated  
            self.disableeventtriggers(True)
            self.surfaceticklist2.SetExcludedEntities(exclude)
            self.disableeventtriggers(False)


    def layer1populated(self, sender, e):
        # generator status - 0-not started; 1-generating; 2-generated/done; 3-error
        if int(sender.Status) == 2:
            self.layerticklist1.Content.ItemContainerGenerator.StatusChanged -= self.layer1populated   
            self.updateticksinlists("1")
            time.sleep(0.1)
            self.layerticklist1.LastInputMethod = InputMethod(0)

    def layer2populated(self, sender, e):
        # generator status - 0-not started; 1-generating; 2-generated/done; 3-error
        if int(sender.Status) == 2:
            self.layerticklist2.Content.ItemContainerGenerator.StatusChanged -= self.layer2populated   
            self.updateticksinlists("2")
            time.sleep(0.1)
            self.layerticklist2.LastInputMethod = InputMethod(0)

    def surface1populated(self, sender, e):
        # generator status - 0-not started; 1-generating; 2-generated/done; 3-error
        if int(sender.Status) == 2:
            self.surfaceticklist1.Content.ItemContainerGenerator.StatusChanged -= self.surface1populated   
            self.updateticksinlists("3")
            time.sleep(0.1)
            self.surfaceticklist1.LastInputMethod = InputMethod(0)

    def surface2populated(self, sender, e):
        # generator status - 0-not started; 1-generating; 2-generated/done; 3-error
        if int(sender.Status) == 2:
            self.surfaceticklist2.Content.ItemContainerGenerator.StatusChanged -= self.surface2populated   
            self.updateticksinlists("4")
            time.sleep(0.1)
            self.surfaceticklist2.LastInputMethod = InputMethod(0)
            tt = 1


    def SetDefaultOptions(self, sender, e):

        self.debug.Content = ''

        try:
            self.Left = OptionsManager.GetDouble("SCR_ViewFilterDialog.windowleft", 100)
            self.Top = OptionsManager.GetDouble("SCR_ViewFilterDialog.windowtop", 100)
            self.Height = OptionsManager.GetDouble("SCR_ViewFilterDialog.windowheight", 500)
            self.Width = OptionsManager.GetDouble("SCR_ViewFilterDialog.windowwidth", 1000)
            self.WindowState = System.Windows.WindowState(OptionsManager.GetInt("SCR_ViewFilterDialog.windowstateint", 0))
            #if not OptionsManager.GetInt("SCR_ViewFilterDialog.windowstateint", 0) == 2:
            #self.Height = 500
            #self.Width = 1000
            #self.Left = -2000
            #self.Top = -2000
            if not self.IsWindowOnAnyScreen(): #!!! this a self written method and must be copied to a new macro
                self.Left = 100
                self.Top = 100
                self.WindowState = System.Windows.WindowState(0)
        except Exception as e:
            tt = sys.exc_info()
            exc_type, exc_obj, exc_tb = sys.exc_info()
            #self.error.Content += '\nan Error occurred - Result probably incomplete\n' + str(exc_type) + '\n' + str(exc_obj) + '\nLine ' + str(exc_tb.tb_lineno)

        
        settings = ConstructionCommandsSettings.ProvideObject(self.currentProject)
        self.layerfilter1.Text = settings.GetString("SCR_ViewFilterDialog.layerfilter1", "")
        self.layerfilter2.Text = settings.GetString("SCR_ViewFilterDialog.layerfilter2", "")
        self.surfacefilter1.Text = settings.GetString("SCR_ViewFilterDialog.surfacefilter1", "")
        self.surfacefilter2.Text = settings.GetString("SCR_ViewFilterDialog.surfacefilter2", "")


        #tt = settings.GetDouble("SCR_ViewFilterDialog.col1width", 9999999)
        #if tt == 9999999:
        #    self.gridcolumn1.ColumnDefinitions[0].Width = GridLength(1, GridUnitType.Star)
        #else:
        #    self.gridcolumn1.ColumnDefinitions[0].ActualWidth = GridLength(tt, GridUnitType.Pixel)
        #tt = settings.GetDouble("SCR_ViewFilterDialog.col2width", 9999999)
        #if tt == 9999999:
        #    self.gridcolumn2.ColumnDefinitions[0].Width = GridLength(1, GridUnitType.Star)
        #else:
        #    self.gridcolumn2.ColumnDefinitions[0].ActualWidth = GridLength(tt, GridUnitType.Pixel)
        #tt = settings.GetDouble("SCR_ViewFilterDialog.col3width", 9999999)
        #if tt == 9999999:
        #    self.gridcolumn3.ColumnDefinitions[0].Width = GridLength(1, GridUnitType.Star)
        #else:
        #    self.gridcolumn3.ColumnDefinitions[0].ActualWidth = GridLength(tt, GridUnitType.Pixel)
        #tt = settings.GetDouble("SCR_ViewFilterDialog.col4width", 9999999)
        #if tt == 9999999:
        #    self.gridcolumn4.ColumnDefinitions[0].Width = GridLength(1, GridUnitType.Star)
        #else:
        #    self.gridcolumn4.ColumnDefinitions[0].ActualWidth = GridLength(tt, GridUnitType.Pixel)

        #self.layerticklist1.Content.Items.ProcessPendingChanges()

        #self.debuggetstartedevent(None, None)


    def SaveOptions(self, sender, e):
        OptionsManager.SetValue("SCR_ViewFilterDialog.windowheight", self.Height)
        OptionsManager.SetValue("SCR_ViewFilterDialog.windowwidth", self.Width)
        OptionsManager.SetValue("SCR_ViewFilterDialog.windowleft", self.Left)
        OptionsManager.SetValue("SCR_ViewFilterDialog.windowtop", self.Top)
        OptionsManager.SetValue("SCR_ViewFilterDialog.windowstateint", int(self.WindowState))

        settings = ConstructionCommandsSettings.ProvideObject(self.currentProject)
        settings.SetString("SCR_ViewFilterDialog.layerfilter1", self.layerfilter1.Text)
        settings.SetString("SCR_ViewFilterDialog.layerfilter2", self.layerfilter2.Text)
        settings.SetString("SCR_ViewFilterDialog.surfacefilter1", self.surfacefilter1.Text)
        settings.SetString("SCR_ViewFilterDialog.surfacefilter2", self.surfacefilter2.Text)

        #settings.SetDouble("SCR_ViewFilterDialog.col1width", self.gridcolumn1.ColumnDefinitions[0].ActualWidth)
        #settings.SetDouble("SCR_ViewFilterDialog.col2width", self.gridcolumn2.ColumnDefinitions[0].ActualWidth)
        #settings.SetDouble("SCR_ViewFilterDialog.col3width", self.gridcolumn3.ColumnDefinitions[0].ActualWidth)
        #settings.SetDouble("SCR_ViewFilterDialog.col4width", self.gridcolumn4.ColumnDefinitions[0].ActualWidth)

        self.cleanupeventtriggers()

    def HelpClicked(self, cmd, e):
        webbrowser.open("C:\ProgramData\Trimble\MacroCommands3\SCR Macros\MacroHelp\MacroHelp.htm#" + type(self).__name__)



    def IsWindowOnAnyScreen(self):
    
        self.doeswindowexist()
        
    #     0        negative Y is upwards and negative X is left
    #   0-|----
    #     |
            
        #tt = System.Windows.Forms.Screen.AllScreens
        for screen in System.Windows.Forms.Screen.AllScreens:
            if self.Left >= screen.WorkingArea.Left and self.Left < screen.WorkingArea.Left + screen.WorkingArea.Width and \
               self.Top >= screen.WorkingArea.Top and self.Top < screen.WorkingArea.Top + screen.WorkingArea.Height:

                return True

        return False

    def doeswindowexist(self):
        


        pass

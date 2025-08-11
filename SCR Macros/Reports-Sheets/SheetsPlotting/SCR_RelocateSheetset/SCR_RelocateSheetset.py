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
#   along with this program.  If not, see <htttargetplanset://www.gnu.org/licenses/>

from System.Collections.Generic import List, IEnumerable # import here, otherwise there is a weird issue with Count and Add for lists
import os
exec(open("C:\ProgramData\Trimble\MacroCommands3\SCR Macros\SCR_Imports.py").read())

def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_RelocateSheetset"
    cmdData.CommandName = "SCR_RelocateSheetset"
    cmdData.Caption = "_SCR_RelocateSheetset"
    cmdData.UIForm = "SCR_RelocateSheetset"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Reports"
        cmdData.DefaultTabGroupKey = "Sheets and Dynaviews"
        cmdData.ShortCaption = "Relocate Sheetset"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.04
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "relocate a Sheetset to a different Planset"
        cmdData.ToolTipTextFormatted = "relocate a Sheetset to a different Planset"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_RelocateSheetset(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_RelocateSheetset.xaml") as s:
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


        #self.sheetsetpicker.SearchContainer = 1 # ProjectFixedSerial 1 is the project itself
        #self.sheetsetpicker.SearchSubContainer = True
        #self.sheetsetpicker.UseSelectionEngine = False
        #self.sheetsetpicker.FilterByEntityTypes = Array[Type]([clr.GetClrType(SheetSet), clr.GetClrType(XSSheetSet), clr.GetClrType(ProfileSheetSet)])

        #sd = System.ComponentModel.SortDescription("Description", System.ComponentModel.ListSortDirection.Ascending)
        #
        #self.sheetsetpicker.SearchContainer = 1
        #self.sheetsetpicker.UseSelectionEngine = False
        #self.sheetsetpicker.SearchSubContainer = True
        #self.sheetsetpicker.SetEntityType(Array[Type]([clr.GetClrType(SheetSet), clr.GetClrType(XSSheetSet), clr.GetClrType(ProfileSheetSet)]), self.currentProject)
        ## Description is the actual property name inside the item in ticklist.Content.Items
        #self.sheetsetpicker.Content.Items.SortDescriptions.Add(sd)

        self.plansetpicker.SearchContainer = 1 # ProjectFixedSerial 1 is the project itself
        self.plansetpicker.SearchSubContainer = True
        self.plansetpicker.UseSelectionEngine = False
        self.plansetpicker.FilterByEntityTypes = Array[Type]([clr.GetClrType(PlanSetSheetView)])

    #def SetDefaultOptions(self):
    #    settings = ConstructionCommandsSettings.ProvideObject(self.currentProject)
    #
    #    try:
    #        self.sheetsetpicker.SelectBySerialNumber(settings.GetInt32("SCR_RelocateSheetset.sheetsetpicker", 0))
    #    except:
    #        pass
    #    
    #    settings.SetInt32("SCR_RelocateSheetset.sheetsetpicker", self.sheetsetpicker.SelectedSerial)
    #    
    #    try:
    #        self.plansetpicker.SelectBySerialNumber(settings.GetInt32("SCR_RelocateSheetset.plansetpicker", 0))
    #    except:
    #        pass
    #
    #    settings.SetInt32("SCR_RelocateSheetset.plansetpicker", self.plansetpicker.SelectedSerial)
    #
    #def SaveOptions(self):
    #    settings = ConstructionCommandsSettings.ProvideObject(self.currentProject)
    #
    #    settings.SetInt32("SCR_RelocateSheetset.sheetsetpicker", self.sheetsetpicker.SelectedSerial)
    #    settings.SetInt32("SCR_RelocateSheetset.plansetpicker", self.plansetpicker.SelectedSerial)
    

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()


    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''

        self.success.Content = ''

        wv = self.currentProject [Project.FixedSerial.WorldView]
        #wv.PauseGraphicsCache(True)

        #self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.SubStep, self.Caption)
        #UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        #
        try:
            # the "with" statement will unroll any changes if something go wrong
            #with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
            selectionSet = GlobalSelection.Items(self.currentProject)
            for o in GlobalSelection.SelectedMembers(self.currentProject):

                #if isinstance(o, SheetSet) or isinstance(o, XSSheetSet) or isinstance(o, ProfileSheetSet) or  isinstance(o, PlanGridSheetSet):
                if isinstance(o, SheetSetBase):
                    # move the sheetset in the project data tree
                    
                    ss = o
                    sourceplanset = ss.GetSite()
                    targetplanset = self.currentProject.Concordance.Lookup(self.plansetpicker.SelectedSerial)

                    if not sourceplanset.SerialNumber == targetplanset.SerialNumber:
                        ss.SetSite(targetplanset)
                        self.currentProject.Concordance.AddObserver(ss.SerialNumber, targetplanset.SerialNumber)
                        self.currentProject.Concordance.RemoveObserver(ss.SerialNumber, sourceplanset.SerialNumber)#RemoveObserver(uint target, uint source) source is the observer

            # now update the project explorer manually - haven't found a better way yet
            te = ExplorerData.TheExplorer
            
            ## find the plansetcollection - node
            for i in te.Items.AllItems:
                if isinstance(i, PlanSetCollection):
                    pscn = i

            pscn.Expanded = True # must expand it, otherwise the allitems list won't be populated
            
            tt = [] # need temp list since we may run into an error while removing data from a live enumerator
            for i in pscn.Items.AllItems:
                if isinstance(i, PlanSetNode):
                    tt.Add(i)
            
            # now call removeitem on all plansetnodes - that basically empties the planset node in the project explorer
            for i in tt:
                pscn.RemoveItem(i)
            
            # trigger repopulate of all child items from plansetnode level
            # this will recreate the items list based on the project tree database we altered with addobserver/removeobserver
            pscn.Populate()

            # re-expand the project explorer tree
            for i in pscn.Items.AllItems:
                if isinstance(i, PlanSetNode):
                    i.Expanded = True

            # redraw the project explorer ui
            te.UpdateData()

            # update the dropdown list in the sheetview window
            for v in TrimbleOffice.TheOffice.MainWindow.AppViewManager.Views:
                if isinstance(v, clr.GetClrType(HoopsSheetView)):
                    for c1 in v.Controls:
                        if "TableLayoutPanel" in c1.GetType().Name:
                            for c2 in c1.Controls:
                                if "ComboBoxEntityPicker" in c2.GetType().Name:
                                    c2.ReloadEntities()



            #failGuard.Commit()
            
            #UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
            #self.currentProject.TransactionManager.AddEndMark(CommandGranularity.SubStep)
        
        except Exception as e:
            tt = sys.exc_info()
            exc_type, exc_obj, exc_tb = sys.exc_info()
            # EndMark MUST be set no matter what
            # otherwise TBC won't work anymore and needs to be restarted
            #self.currentProject.TransactionManager.AddEndMark(CommandGranularity.SubStep)
            #UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
            self.error.Content += '\nan Error occurred - Result probably incomplete\n' + str(exc_type) + '\n' + str(exc_obj) + '\nLine ' + str(exc_tb.tb_lineno)
        
        self.success.Content += '\nDone'

        #wv.PauseGraphicsCache(False)

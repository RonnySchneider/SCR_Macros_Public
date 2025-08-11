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
    cmdData.Key = "SCR_Relayer"
    cmdData.CommandName = "SCR_Relayer"
    cmdData.Caption = "_SCR_Relayer"
    cmdData.UIForm = "SCR_Relayer"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Expld-SNR-Relay-Prop"
        cmdData.DefaultTabGroupKey = "Relayer"
        cmdData.ShortCaption = "Relayer Objects"
        cmdData.DefaultRibbonToolSize = 0 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.15
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Quickly Relayer Objects"
        cmdData.ToolTipTextFormatted = "Quickly Relayer Objects"

    except:
        pass
    
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_Relayer(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_Relayer.xaml") as s:
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


        self.layerticklist.SearchContainer = Project.FixedSerial.LayerContainer
        self.layerticklist.UseSelectionEngine = False
        self.layerticklist.SetEntityType(clr.GetClrType(Layer), self.currentProject)
        # Description is the actual property name inside the item in ticklist.Content.Items
        sd = System.ComponentModel.SortDescription("Description", System.ComponentModel.ListSortDirection.Ascending)
        self.layerticklist.Content.Items.SortDescriptions.Add(sd)

        self.relayerticklist.SearchContainer = Project.FixedSerial.LayerContainer
        self.relayerticklist.UseSelectionEngine = False
        self.relayerticklist.SetEntityType(clr.GetClrType(Layer), self.currentProject)
        # Description is the actual property name inside the item in ticklist.Content.Items
        self.relayerticklist.Content.Items.SortDescriptions.Add(sd)
        
        relayerexclude = []
        for i in self.relayerticklist.Content.Items:
            relayerexclude.Add(i.EntitySerialNumber)
        self.relayerticklist.SetExcludedEntities(relayerexclude)

        self.ticklistfilter.TextChanged += self.FilterChanged

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def FilterChanged(self, ctrl, e):
        
        exclude = []
        self.layerticklist.SetExcludedEntities(exclude)

        tt = self.ticklistfilter.Text.lower()
        ticklistfilter = tt.split()

        for i in self.layerticklist.EntitySerialNumbers:
            for f in ticklistfilter:
               if not f in i.Key.Description.lower():
                    exclude.Add(i.Value)

        self.layerticklist.SetExcludedEntities(exclude)

        self.SaveOptions()

    def combinedfiltercheckboxChanged(self, sender, e):
        if self.combinedfiltercheckbox.IsChecked:
            self.separatefilters.IsEnabled = False
        else:
            self.separatefilters.IsEnabled = True

    def addbutton_Click(self, sender, e):
        relayerexclude = self.relayerticklist.ExcludedSerials
        relayerexclude = list(relayerexclude)

        for i in self.layerticklist.Content.SelectedItems:
            if i.EntitySerialNumber in relayerexclude:
                relayerexclude.remove(i.EntitySerialNumber)

        self.relayerticklist.SetExcludedEntities([])
        self.relayerticklist.SetExcludedEntities(relayerexclude)

        self.SaveOptions()

    def layerticklist_doubleclick(self, sender, e):
        relayerexclude = list(self.relayerticklist.ExcludedSerials) # we need it as list, not an array

        if self.layerticklist.Content.SelectedValue.EntitySerialNumber in relayerexclude:
            relayerexclude.remove(self.layerticklist.Content.SelectedValue.EntitySerialNumber)

        self.relayerticklist.SetExcludedEntities([])
        self.relayerticklist.SetExcludedEntities(relayerexclude)

        self.SaveOptions()

    def removebutton_Click(self, sender, e):
        relayerexclude = list(self.relayerticklist.ExcludedSerials) # we need it as list, not an array

        for i in self.relayerticklist.Content.SelectedItems:
            relayerexclude.Add(i.EntitySerialNumber)

        self.relayerticklist.SetExcludedEntities([])
        self.relayerticklist.SetExcludedEntities(relayerexclude)

        self.SaveOptions()

    def removeallbutton_Click(self, sender, e):
        relayerexclude = list(self.relayerticklist.ExcludedSerials) # we need it as list, not an array
        
        for i in self.relayerticklist.Content.Items:
            relayerexclude.Add(i.EntitySerialNumber)
        
        self.relayerticklist.SetExcludedEntities([])
        self.relayerticklist.SetExcludedEntities(relayerexclude)

        self.SaveOptions()

    def relayerticklist_doubleclick(self, sender, e):
        self.OkClicked(None, None)
    
    def SetDefaultOptions(self):
        settings = ConstructionCommandsSettings.ProvideObject(self.currentProject)
        
        # we have to restore the layer list first, currently the listbox shows everything
        relayerexclude = settings.GetString("SCR_Relayer.relayerexclude", "")
        if relayerexclude != "":
            relayerexclude = relayerexclude.split(",")
            relayerexclude = [ System.UInt32.Parse(element) for element in relayerexclude]
            self.relayerticklist.SetExcludedEntities([])
            self.relayerticklist.SetExcludedEntities(relayerexclude)
        
        tt = settings.GetBoolean("SCR_Relayer.filter_objs")
        self.filter_objs.IsChecked = settings.GetBoolean("SCR_Relayer.filter_objs", False)
        self.combinedfiltercheckbox.IsChecked = settings.GetBoolean("SCR_Relayer.combinedfiltercheckbox", False)
        self.combinedfilter.Text = settings.GetString("SCR_Relayer.combinedfilter", "")
        self.pointnamefiltercheckbox.IsChecked = settings.GetBoolean("SCR_Relayer.pointnamefiltercheckbox", False)
        self.pointnamefilter.Text = settings.GetString("SCR_Relayer.pointnamefilter", "")
        self.pointcodefiltercheckbox.IsChecked = settings.GetBoolean("SCR_Relayer.pointcodefiltercheckbox", False)
        self.pointcodefilter.Text = settings.GetString("SCR_Relayer.pointcodefilter", "")

        # there is a save settings instruction after each time we change the filter
        # if we start restoring the filter text before the relayerlist is restored, we'd overwrite it with the whole list
        self.ticklistfilter.Text = settings.GetString("SCR_Relayer.ticklistfilter", "")

    def SaveOptions(self):
        settings = ConstructionCommandsSettings.ProvideObject(self.currentProject)
        settings.SetString("SCR_Relayer.ticklistfilter", self.ticklistfilter.Text)

        relayerexclude = self.relayerticklist.ExcludedSerials
        relayerexclude = [str(element) for element in relayerexclude]
        relayerexclude = ",".join(relayerexclude)
        settings.SetString("SCR_Relayer.relayerexclude", relayerexclude)

        tt = self.filter_objs.IsChecked
        settings.SetBoolean("SCR_Relayer.filter_objs", self.filter_objs.IsChecked)
        settings.SetBoolean("SCR_Relayer.combinedfiltercheckbox", self.combinedfiltercheckbox.IsChecked)
        settings.SetString("SCR_Relayer.combinedfilter", self.combinedfilter.Text)
        settings.SetBoolean("SCR_Relayer.pointnamefiltercheckbox", self.pointnamefiltercheckbox.IsChecked)
        settings.SetString("SCR_Relayer.pointnamefilter", self.pointnamefilter.Text)
        settings.SetBoolean("SCR_Relayer.pointcodefiltercheckbox", self.pointcodefiltercheckbox.IsChecked)
        settings.SetString("SCR_Relayer.pointcodefilter", self.pointcodefilter.Text)


    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

  
    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''
        self.success.Content = ''

        # filter the objects if necessary
        relayerobjs = []
        for o in self.objs:

            if self.filter_objs.IsChecked:
                addok = True
                
                if isinstance(o.GetSite(), PointCollection):

                    # if both are ticked than pointname and code need to match
                    if self.pointnamefiltercheckbox.IsChecked:
                        if self.pointnamefilter.Text.lower() not in o.AnchorName.lower():
                            addok = False

                    if self.pointcodefiltercheckbox.IsChecked:
                        if self.pointcodefilter.Text.lower() not in o.FeatureCode.lower():
                            addok = False

                    
                    if self.combinedfiltercheckbox.IsChecked:
                        addok = False
                        if self.combinedfilter.Text.lower() in o.AnchorName.lower(): addok = True
                        if self.combinedfilter.Text.lower() in o.FeatureCode.lower(): addok = True
                
                else: # isn't a Coordinatepoint
                    addok = False

                if addok:
                    relayerobjs.Add(o)
            
            else: # no point-filter
                relayerobjs.Add(o)

        if self.relayerticklist.Content.SelectedItems.Count > 0 and relayerobjs.Count > 0:
            UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
            self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
            try:

                with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                    for o in relayerobjs:
                        o.Layer = self.relayerticklist.Content.SelectedValue.EntitySerialNumber
                    self.success.Content = str(relayerobjs.Count) + " Objects re-layered to " + self.relayerticklist.Content.SelectedValue.Description
                    failGuard.Commit()

                UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
                self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
            
            except:
                # EndMark MUST be set no matter what
                # otherwise TBC won't work anymore and needs to be restarted
                self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
                UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
                self.error.Content += '\nan Error occurred - Result probably incomplete'

        else:
            self.error.Content += '\nno target layer or objects selected'

        Keyboard.Focus(self.objs)
        self.SaveOptions()

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
    "filter_objs":             False,
    "combinedfiltercheckbox":  False,
    "combinedfilter":          "",
    "pointnamefiltercheckbox": False,
    "pointnamefilter":         "",
    "pointcodefiltercheckbox": False,
    "pointcodefilter":         "",
    "ticklistfilter":          "",    # must be last — TextChanged fires SaveOptions via FilterChanged
}

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

        cmdData.Version = 1.21
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


        self.layerticklist, self.innerlist_layers = SCREntityPicker.create_layers(self.layerlisthost)
        if self.innerlist_layers is not None:
            self.innerlist_layers.DoubleClick += self.layerticklist_doubleclick

        self.relayerticklist, self.innerlist_relayer = SCREntityPicker.create_layers(
            self.relayerlisthost, multi_select=False)
        if self.innerlist_relayer is not None:
            self.innerlist_relayer.DoubleClick += self.relayerticklist_doubleclick

        relayerexclude = [kvp.Value for kvp in self.relayerticklist.EntitySerialNumbers]
        self.relayerticklist.SetExcludedEntities([])
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
        SCREntityPicker.reapply_highlights(self.innerlist_layers)

        self.SaveOptions()

    def combinedfiltercheckboxChanged(self, sender, e):
        if self.combinedfiltercheckbox.IsChecked:
            self.separatefilters.IsEnabled = False
        else:
            self.separatefilters.IsEnabled = True

    def addbutton_Click(self, sender, e):
        relayerexclude = list(self.relayerticklist.ExcludedSerials)
        for sn in SCREntityPicker.get_selected_serials(self.layerticklist, self.innerlist_layers):
            if sn in relayerexclude:
                relayerexclude.remove(sn)
        self.relayerticklist.SetExcludedEntities([])
        self.relayerticklist.SetExcludedEntities(relayerexclude)
        self.SaveOptions()

    def layerticklist_doubleclick(self, sender, e):
        relayerexclude = list(self.relayerticklist.ExcludedSerials)
        serials = SCREntityPicker.get_selected_serials(self.layerticklist, self.innerlist_layers)
        if not serials:
            return
        sn = serials[0]
        if sn in relayerexclude:
            relayerexclude.remove(sn)
        self.relayerticklist.SetExcludedEntities([])
        self.relayerticklist.SetExcludedEntities(relayerexclude)
        self.SaveOptions()

    def removebutton_Click(self, sender, e):
        relayerexclude = list(self.relayerticklist.ExcludedSerials)
        for sn in SCREntityPicker.get_selected_serials(self.relayerticklist, self.innerlist_relayer):
            relayerexclude.Add(sn)
        self.relayerticklist.SetExcludedEntities([])
        self.relayerticklist.SetExcludedEntities(relayerexclude)
        self.SaveOptions()

    def removeallbutton_Click(self, sender, e):
        if self.innerlist_relayer is None:
            return
        relayerexclude = list(self.relayerticklist.ExcludedSerials)
        nameToSN = {kvp.Key.Description: kvp.Value for kvp in self.relayerticklist.EntitySerialNumbers}
        for item in self.innerlist_relayer.Items:
            sn = nameToSN.get(item.Text)
            if sn is not None:
                relayerexclude.Add(sn)
        self.relayerticklist.SetExcludedEntities([])
        self.relayerticklist.SetExcludedEntities(relayerexclude)
        self.SaveOptions()

    def relayerticklist_doubleclick(self, sender, e):
        self.OkClicked(None, None)
    
    def SetDefaultOptions(self):
        settings = ConstructionCommandsSettings.ProvideObject(self.currentProject)
        # relayerexclude must be restored first — restoring ticklistfilter fires SaveOptions,
        # which would overwrite relayerexclude with the full unfiltered list if not set beforehand
        relayerexclude = settings.GetString("SCR_Relayer.relayerexclude", "")
        if relayerexclude != "":
            relayerexclude = [System.UInt32.Parse(e) for e in relayerexclude.split(",")]
            self.relayerticklist.SetExcludedEntities([])
            self.relayerticklist.SetExcludedEntities(relayerexclude)
        SCROptions.LoadProjectOptions(self, "SCR_Relayer", _OPTIONS, self.currentProject)

    def SaveOptions(self):
        settings = ConstructionCommandsSettings.ProvideObject(self.currentProject)
        settings.SetString("SCR_Relayer.relayerexclude", ",".join(str(e) for e in self.relayerticklist.ExcludedSerials))
        SCROptions.SaveProjectOptions(self, "SCR_Relayer", _OPTIONS, self.currentProject)


    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

  
    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''
        self.success.Content = ''

        wv = self.currentProject [Project.FixedSerial.WorldView]
        wv.PauseGraphicsCache(True)

        # filter the objects if necessary
        relayerobjs = []
        selectionserials = []

        # save the serials before doing something to them
        # should be faster than updating properties and shoe line direction every time
        for o in self.objs:
            selectionserials.Add(o.SerialNumber)

        GlobalSelection.Clear()

        for sn in selectionserials:

            o = self.currentProject.Concordance[sn]

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

        serials = SCREntityPicker.get_selected_serials(self.relayerticklist, self.innerlist_relayer)
        targetsn = serials[0] if serials else None

        if targetsn is not None and relayerobjs.Count > 0:

            UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
            self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
            try:

                with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                    for o in relayerobjs:
                        o.Layer = targetsn
                    self.success.Content = str(relayerobjs.Count) + " Objects re-layered to " + self.currentProject.Concordance[targetsn].Name
                    failGuard.Commit()

                self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
                UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
            
            except:
                # EndMark MUST be set no matter what
                # otherwise TBC won't work anymore and needs to be restarted
                self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
                UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
                self.error.Content += '\nan Error occurred - Result probably incomplete'

        else:
            self.error.Content += '\nno target layer or objects selected'


        # reinstate old selection
        ProgressBar.TBC_ProgressBar.Title = "reinstating selection"
        GlobalSelection.Items(self.currentProject).Set(selectionserials)

        ProgressBar.TBC_ProgressBar.Title = ""

        Keyboard.Focus(self.objs)
        self.SaveOptions()
        wv.PauseGraphicsCache(False)

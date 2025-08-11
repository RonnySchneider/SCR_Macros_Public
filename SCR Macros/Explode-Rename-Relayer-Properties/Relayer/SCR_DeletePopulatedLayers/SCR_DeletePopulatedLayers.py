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
    cmdData.Key = "SCR_DeletePopulatedLayers"
    cmdData.CommandName = "SCR_DeletePopulatedLayers"
    cmdData.Caption = "_SCR_DeletePopulatedLayers"
    cmdData.UIForm = "SCR_DeletePopulatedLayers"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Expld-SNR-Relay-Prop"
        cmdData.DefaultTabGroupKey = "Relayer"
        cmdData.ShortCaption = "Delete Layers"
        cmdData.DefaultRibbonToolSize = 0 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.02
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Quickly Delete Populated Layers"
        cmdData.ToolTipTextFormatted = "Quickly Delete Populated Layers or just the Content, without actual selecting"

    except:
        pass
    
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_DeletePopulatedLayers(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_DeletePopulatedLayers.xaml") as s:
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
		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass
        self.layerpicker.ValueChanged += self.layerpickerChanged
        self.layerpickerChanged(self, cmd) # call it once to get the number for the initial selection

    def layerpickerChanged(self, ctrl, e):
        try:
            self.membercount.Content = "Membercount: " + str(self.currentProject.Concordance.Lookup(self.layerpicker.SelectedSerialNumber).Members.Count)
        except:
            tt = 1
            pass

    def SetDefaultOptions(self):
        self.deletelayer.IsChecked = OptionsManager.GetBool("SCR_DeletePopulatedLayers.deletelayer", True)
        self.picklayer.IsChecked = OptionsManager.GetBool("SCR_DeletePopulatedLayers.picklayer", True)
        self.searchlayer.IsChecked = OptionsManager.GetBool("SCR_DeletePopulatedLayers.searchlayer", False)
        self.searchtext.Text = OptionsManager.GetString("SCR_DeletePopulatedLayers.searchtext", "")
        self.multilayer.IsChecked = OptionsManager.GetBool("SCR_DeletePopulatedLayers.multilayer", False)

    def SaveOptions(self):
        OptionsManager.SetValue("SCR_DeletePopulatedLayers.deletelayer", self.deletelayer.IsChecked)
        OptionsManager.SetValue("SCR_DeletePopulatedLayers.picklayer", self.picklayer.IsChecked)
        OptionsManager.SetValue("SCR_DeletePopulatedLayers.searchlayer", self.searchlayer.IsChecked)
        OptionsManager.SetValue("SCR_DeletePopulatedLayers.searchtext", self.searchtext.Text)
        OptionsManager.SetValue("SCR_DeletePopulatedLayers.multilayer", self.multilayer.IsChecked)

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

  
    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''



        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        
        ProgressBar.TBC_ProgressBar.Title = self.Caption
        wv = self.currentProject [Project.FixedSerial.WorldView]
        wlc = self.currentProject[Project.FixedSerial.LayerContainer]
        wv.PauseGraphicsCache(True)

        self.success.Content=''

        #wlc = self.currentProject[Project.FixedSerial.LayerContainer] # we get all the layers into an object, LayerCollection
        #wl = wlc[self.layerpicker.SelectedSerialNumber]    # we get just the source layer as an object
        
        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                layerserials = []

                # get either just one layer serial from the layerpicker or multiple if they are ticked in the list
                if self.multilayer.IsChecked:
                    for item in self.layerticklist.Content.ItemsSource:
                        if item.Checked:
                            layerserials.Add(item.EntitySerialNumber)
                
                # one single layer from layerpicker
                if self.picklayer.IsChecked:
                    layerserials.Add(self.layerpicker.SelectedSerialNumber)
                
                # search for specific name part
                if self.searchlayer.IsChecked:
                    if self.searchtext.Text != "": # with an empty string it deletes all layers
                        for l in wlc:
                            if self.searchtext.Text in l.Name:
                                layerserials.Add(l.SerialNumber)

                # go through all the layerserials
                for layerserial in layerserials:

                    layerobject = self.currentProject.Concordance.Lookup(layerserial) # get the layer as object
                    layermembers = layerobject.Members # we get serial number list of all the elements on that layer
                    
                    ProgressBar.TBC_ProgressBar.Title = "working on Layer: " + layerobject.Name # set the progressbar description
                    
                    j=0
                    layercleaned = True

                    for serialn in layermembers:
                        try:    # use try to recognize if there was an error, in that case the layer won't be empty at the end and shouldn't be deleted
                            j += 1
                            if ((j * 100 / layermembers.Count) % 5 == 0):
                                if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(j * 100 / layermembers.Count)):
                                    break   # function returns true if user pressed cancel

                            o = self.currentProject.Concordance.Lookup(serialn)
                            osite = o.GetSite()    # we find out in which container the serial number reside
                            osite.Remove(serialn)   # we delete the object from that container
                        except:
                            layercleaned = False    # if it throws any error the layer might not be empty at the end
                            #if not o == None:
                            #    #self.success.Content += "\nCouldn't delete Memberserial: " + str(serialn)

                    # we double check the layer content
                    # even if there might have been a serial that threw an error the layer might be empty after all
                    if layerobject.Members.Count == 0: layercleaned = True

                    if self.deletelayer.IsChecked and layercleaned: # if the delete layer tickbox is ticked and it is empty try to delete it
                        layersite = layerobject.GetSite()
                        try: # use try because it could be layer "0" or "Points", which can't be deleted
                            layersite.Remove(layerserial)
                        except:
                            self.success.Content += "\nCouldn't delete Layer - " + layerobject.Name + " - SN: " + str(layerobject.SerialNumber)
                            if self.multilayer.IsChecked == False:
                                self.layerpickerChanged(self, cmd) # call it once to update the number in case we couldn't delete the layer and the Layerselection didn't change
                
                failGuard.Commit()
                UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
                self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
        
        except:
            # EndMark MUST be set no matter what
            # otherwise TBC won't work anymore and needs to be restarted
            self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
            UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
            self.error.Content += '\nan Error occurred - Result probably incomplete'
        
        ProgressBar.TBC_ProgressBar.Title = ""
        wv.PauseGraphicsCache(False)

        self.SaveOptions()
        Keyboard.Focus(self.layerpicker)

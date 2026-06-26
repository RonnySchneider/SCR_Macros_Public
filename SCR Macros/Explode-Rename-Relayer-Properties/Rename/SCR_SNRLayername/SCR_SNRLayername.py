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
    "selectlayers":   False,
    "searchtext":     "",
    "replacetext":    "",
    "createcopies":   False,
    "copyobjects":    False,
    "renamegroup":    False,
    "ticklistfilter": "",    # must be last — TextChanged fires SaveOptions via FilterChanged
}

def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_SNRLayername"
    cmdData.CommandName = "SCR_SNRLayername"
    cmdData.Caption = "_SCR_SNRLayername"
    cmdData.UIForm = "SCR_SNRLayername"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Expld-SNR-Relay-Prop"
        cmdData.DefaultTabGroupKey = "Renaming"
        cmdData.ShortCaption = "SNR in Layernames"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.20
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "SNR in Layernames"
        cmdData.ToolTipTextFormatted = "search and replace in layer names"

    except:
        pass

    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_SNRLayername(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_SNRLayername.xaml") as s:
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
        
        self.layerticklist, self.innerlist = SCREntityPicker.create_layers(self.layerlisthost)
        self.ticklistfilter.TextChanged += self.FilterChanged

        SCRExpanders.wire_pairs([
            (self.expander_selectlayers, self.selectlayers),
        ])
		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def SetDefaultOptions(self):
        SCROptions.LoadProjectOptions(self, "SCR_SNRLayername", _OPTIONS, self.currentProject)

    def SaveOptions(self):
        SCROptions.SaveProjectOptions(self, "SCR_SNRLayername", _OPTIONS, self.currentProject)


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
        SCREntityPicker.reapply_highlights(self.innerlist)

        self.SaveOptions()


    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()


    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''


        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        
        wv = self.currentProject [Project.FixedSerial.WorldView]
        wl = self.currentProject[Project.FixedSerial.LayerContainer]
        dp = self.currentProject.CreateDuplicator()
            
        try:
            # the "with" statement will unroll any changes if something go wrong
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                
                layerserials = []
                if self.selectlayers.IsChecked:
                    layerserials = SCREntityPicker.get_selected_serials(self.layerticklist, self.innerlist)
                else:
                    for o in wl: # go through everything in the LayerContainer
                        layerserials.Add(o.SerialNumber)

                    
                # string.replace(oldvalue, newvalue, count)
                for sn in layerserials:
                    o = self.currentProject.Concordance.Lookup(sn)
                    if self.searchtext.Text in o.Name:
                        if not self.createcopies.IsChecked:
                            if o.SerialNumber != 8 and o.SerialNumber != 23: # we don't want to rename Layer Zero and Points
                                o.Name = o.Name.replace(self.searchtext.Text, self.replacetext.Text)
                        else: # create new layers
                            outputlayer = Layer.FindOrCreateLayer(self.currentProject, o.Name.replace(self.searchtext.Text, self.replacetext.Text))
                            outputlayer.DefaultColor = o.DefaultColor
                            outputlayer.FilterCategory = o.FilterCategory
                            outputlayer.LineStyle = o.LineStyle
                            outputlayer.LineWeight = o.LineWeight
                            outputlayer.Print = o.Print
                            outputlayer.Priority = o.Priority
                            #outputlayer.Visible = o.Visible # for some reason that would set the layer enabled/ticked in all View Filters
                            outputlayer.Protected = o.Protected
                            
                            if self.renamegroup.IsChecked:
                                currentlg = self.currentProject.Concordance.Lookup(o.LayerGroupSerial)
                                if currentlg:
                                    newlgname = currentlg.Name.replace(self.searchtext.Text, self.replacetext.Text)
                                    lgCollection = LayerGroupCollection.GetLayerGroupCollection(self.currentProject, False)
                                    newlg = lgCollection.FindOrCreateLayerGroup(newlgname)
                                    outputlayer.LayerGroupSerial = newlg.SerialNumber
                            else:
                                outputlayer.LayerGroupSerial = o.LayerGroupSerial

                            if self.copyobjects.IsChecked:
                                
                                for sn in o.Members:    # go through the layer objects of the original layer

                                    lo = self.currentProject.Concordance.Lookup(sn)
                                    # create a copy of each object
                                    try:
                                        o_new = lo.GetSite().Add(clr.GetClrType(lo.GetType()))
                                        o_new.CopyBody(self.currentProject.Concordance, self.currentProject.TransactionManager, lo, dp)                        
                                        o_new.Layer = outputlayer.SerialNumber
                                    except:
                                        pass

                failGuard.Commit()
                self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
                UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
        
        except:
            tt = sys.exc_info()
            exc_type, exc_obj, exc_tb = sys.exc_info()
            # EndMark MUST be set no matter what
            # otherwise TBC won't work anymore and needs to be restarted
            self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
            UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
            self.error.Content += '\nan Error occurred - Result probably incomplete\n' + str(exc_type) + '\n' + str(exc_obj) + '\nLine ' + str(exc_tb.tb_lineno)
        
        self.SaveOptions()


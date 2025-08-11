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
    cmdData.Key = "SCR_ChangeDynaviewFilters"
    cmdData.CommandName = "SCR_ChangeDynaviewFilters"
    cmdData.Caption = "_SCR_ChangeDynaviewFilters"
    cmdData.UIForm = "SCR_ChangeDynaviewFilters"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Reports"
        cmdData.DefaultTabGroupKey = "Sheets and Dynaviews"
        cmdData.ShortCaption = "change Dynaview-Filters"
        cmdData.DefaultRibbonToolSize = 0 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.08
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "change Dynaview Filters on Multiple Sheets"
        cmdData.ToolTipTextFormatted = "change Dynaview Filters on Multiple Sheets"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_ChangeDynaviewFilters(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_ChangeDynaviewFilters.xaml") as s:
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
        #types = Array [Type] ([CadPoint]) + Array [Type] ([Point3D])    # we fill an array with TBC object types, we could combine different types
        
        vfc = self.currentProject [Project.FixedSerial.ViewFilterCollection] # getting all View Filters in one list
        
        # to have the drop down list sorted we fill a templist first, sort it and then fill the dropdownbox
        templist = []
        for f in vfc:
            templist.Add(f.Name)
        templist.sort()
        
        # fill the viewfilter drop down lists
        for f in templist:
            item = ComboBoxItem()
            item.Content = f
            item.FontSize = 12
            self.viewfilterpicker.Items.Add(item)
        for f in templist:
            item = ComboBoxItem()
            item.Content = f
            item.FontSize = 12
            self.ignorefilterpicker.Items.Add(item)

        self.sortstart.NumberOfDecimals = 0
        #self.sortstart.MinValue = 0

        self.sortinc.NumberOfDecimals = 0
        
		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def SetDefaultOptions(self):
        
        self.changefilter.IsChecked = OptionsManager.GetBool("SCR_ChangeDynaviewFilters.changefilter", True)
        self.changesort.IsChecked = OptionsManager.GetBool("SCR_ChangeDynaviewFilters.changesort", False)

        self.sortstart.Value = OptionsManager.GetDouble("SCR_ChangeDynaviewFilters.sortstart", 0)
        self.sortinc.Value = OptionsManager.GetDouble("SCR_ChangeDynaviewFilters.sortinc", 1)

    def SaveOptions(self):

        OptionsManager.SetValue("SCR_ChangeDynaviewFilters.changefilter", self.changefilter.IsChecked)    
        OptionsManager.SetValue("SCR_ChangeDynaviewFilters.changesort", self.changesort.IsChecked)    

        OptionsManager.SetValue("SCR_ChangeDynaviewFilters.sortstart", self.sortstart.Value)
        OptionsManager.SetValue("SCR_ChangeDynaviewFilters.sortinc", self.sortinc.Value)


    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()


    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''

        self.success.Content = ''

        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)

        try:

            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                # self.label_benchmark.Content = ''

                # start_t = timer ()
                # wv = self.currentProject [Project.FixedSerial.WorldView]
                # bc = self.currentProject [Project.FixedSerial.BlockCollection]    # getting all blocks as collection
                # lsc = self.currentProject [Project.FixedSerial.LabelStyleContainer]    # getting all blocks as collection
                vfc = self.currentProject [Project.FixedSerial.ViewFilterCollection] # getting all View Filters in one list
                
                if self.changefilter.IsChecked:
                    # find the selected view filter serial number
                    # for some reason vfc[self.viewfilterpicker.SelectedIndex] doesn't work
                    if self.viewfilterpicker.SelectedIndex >= 0:
                        for vf in vfc:
                            if vf.Name == self.viewfilterpicker.SelectionBoxItem:
                                vfserial = vf.SerialNumber
                    else:
                        vfserial = -1
                                
                    if self.ignorefilterpicker.SelectedIndex >= 0:
                        for vf in vfc:
                            if vf.Name == self.ignorefilterpicker.SelectionBoxItem:
                                ignorevfserial = vf.SerialNumber
                    else:
                        ignorevfserial = -1

                    for o in self.objs: # go through all selected objects and test if it is a sheet
                        if isinstance(o, BasicSheet):
                            for i in o: # go through all objects on the sheet and test if it is a DynaView
                                if isinstance(i, DynaView):
                                    # test if the ignore filter is set and if it doesn't match
                                    if ignorevfserial == -1 or i.ViewFilter != ignorevfserial:
                                        # test if the change filter is set and apply it
                                        if vfserial != -1:
                                            i.ViewFilter = vfserial
                        else:
                            self.error.Content += "\nSelection wasn't a simple Sheet"
                else:
                    sheetlist = {}
                    for o in self.objs: # go through all selected objects and test if it is a sheet
                        if isinstance(o, BasicSheet):
                            sheetlist.Add(o.SortRank, o.SerialNumber)
                    
                    sheetlist_sorted = sorted(sheetlist.items(), key=lambda x: x[0], reverse=False)
                    
                    newrank = math.floor(self.sortstart.Value)
                    for s in sheetlist_sorted:
                        o = self.currentProject.Concordance.Lookup(s[1])
                        o.SortRank = newrank
                        newrank += math.floor(self.sortinc.Value)

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
        self.SaveOptions()

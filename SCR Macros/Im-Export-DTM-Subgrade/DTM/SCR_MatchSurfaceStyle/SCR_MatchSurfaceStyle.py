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
    cmdData.Key = "SCR_MatchSurfaceStyle"
    cmdData.CommandName = "SCR_MatchSurfaceStyle"
    cmdData.Caption = "_SCR_MatchSurfaceStyle"
    cmdData.UIForm = "SCR_MatchSurfaceStyle"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "DTM"
        cmdData.ShortCaption = "Match Surface Style"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.03
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "match the surface representation style"
        cmdData.ToolTipTextFormatted = "match the surface representation style"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_MatchSurfaceStyle(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_MatchSurfaceStyle.xaml") as s:
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
        types = Array [Type] ([clr.GetClrType (ProjectedSurface)]) + Array [Type] (SurfaceTypeLists.AllWithCutFillMap)    # we fill an array with TBC object types, we could combine different types

                                                                                                                                                                                                                                                                                                                           # +Array[Type](SurfaceTypeLists.AllWithCutFillMap)
        self.surfacepicker1.FilterByEntityTypes = types    # we fill the dropdownlist by applying that types array as filter
        self.surfacepicker2.FilterByEntityTypes = types

        self.surfacepicker1.AllowNone = False              # our list shall not show an empty field
        self.surfacepicker2.AllowNone = False              # our list shall not show an empty field

        self.surface2ticklist.SearchContainer = Project.FixedSerial.WorldView
        self.surface2ticklist.UseSelectionEngine = False
        self.surface2ticklist.SetEntityType(types, self.currentProject)

        # Description is the actual property name inside the item in ticklist.Content.Items
        sd = System.ComponentModel.SortDescription("Description", System.ComponentModel.ListSortDirection.Ascending)
        self.surface2ticklist.Content.Items.SortDescriptions.Add(sd)

        self.ticklistfilter.TextChanged += self.FilterChanged

        self.SetDefaultOptions()

    def SetDefaultOptions(self):

        # Select surface
        try:    self.surfacepicker1.SelectIndex(OptionsManager.GetInt("SCR_MatchSurfaceStyle.surfacepicker1", 0))
        except: self.surfacepicker1.SelectIndex(0)
        try:    self.surfacepicker2.SelectIndex(OptionsManager.GetInt("SCR_MatchSurfaceStyle.surfacepicker2", 0))
        except: self.surfacepicker2.SelectIndex(0)

        self.single.IsChecked = OptionsManager.GetBool("SCR_MatchSurfaceStyle.single", True)
        self.multi.IsChecked = OptionsManager.GetBool("SCR_MatchSurfaceStyle.multi", False)
        self.color.IsChecked = OptionsManager.GetBool("SCR_MatchSurfaceStyle.color", False)
        self.maxedgelength.IsChecked = OptionsManager.GetBool("SCR_MatchSurfaceStyle.maxedgelength", False)
        self.maxedgeangle.IsChecked = OptionsManager.GetBool("SCR_MatchSurfaceStyle.maxedgeangle", False)
        self.rebuildmethod.IsChecked = OptionsManager.GetBool("SCR_MatchSurfaceStyle.rebuildmethod", False)
        self.transparency.IsChecked = OptionsManager.GetBool("SCR_MatchSurfaceStyle.transparency", False)

        # need to restore that one last, since it also saves, it would clear all the other fields
        self.ticklistfilter.Text = OptionsManager.GetString("SCR_MatchSurfaceStyle.ticklistfilter", "")

    def SaveOptions(self):
        OptionsManager.SetValue("SCR_MatchSurfaceStyle.surfacepicker1", self.surfacepicker1.SelectedIndex)
        OptionsManager.SetValue("SCR_MatchSurfaceStyle.surfacepicker2", self.surfacepicker2.SelectedIndex)

        OptionsManager.SetValue("SCR_MatchSurfaceStyle.single", self.single.IsChecked)
        OptionsManager.SetValue("SCR_MatchSurfaceStyle.multi", self.multi.IsChecked)
        OptionsManager.SetValue("SCR_MatchSurfaceStyle.ticklistfilter", self.ticklistfilter.Text)

        OptionsManager.SetValue("SCR_MatchSurfaceStyle.color", self.color.IsChecked)
        OptionsManager.SetValue("SCR_MatchSurfaceStyle.maxedgelength", self.maxedgelength.IsChecked)
        OptionsManager.SetValue("SCR_MatchSurfaceStyle.maxedgeangle", self.maxedgeangle.IsChecked)
        OptionsManager.SetValue("SCR_MatchSurfaceStyle.rebuildmethod", self.rebuildmethod.IsChecked)
        OptionsManager.SetValue("SCR_MatchSurfaceStyle.transparency", self.transparency.IsChecked)

    def CancelClicked(self, thisCmd, args):
        thisCmd.CloseUICommand ()

    def FilterChanged(self, ctrl, e):
        
        exclude = []
        self.surface2ticklist.SetExcludedEntities(exclude)

        tt = self.ticklistfilter.Text.lower()
        ticklistfilter = tt.split()

        for i in self.surface2ticklist.EntitySerialNumbers:
            for f in ticklistfilter:
               if not f in i.Key.Description.lower():
                    exclude.Add(i.Value)

        self.surface2ticklist.SetExcludedEntities(exclude)

        self.SaveOptions()

    def OkClicked(self, thisCmd, e):
        self.error.Content = ''
        self.success.Content = ''

        surface1 = self.currentProject.Concordance.Lookup(self.surfacepicker1.SelectedSerial)    # we get our selected surface as object
        #surface2 = self.currentProject.Concordance.Lookup(self.surfacepicker2.SelectedSerial)    # we get our selected surface as object

        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())

        targetserials = []

        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                # prepare a list with the target serials - depending on multi or single surface change
                if self.single.IsChecked:
                    targetserials.Add(self.surfacepicker2.SelectedSerial)
                else:
                    for i in self.surface2ticklist.Content.SelectedItems:
                        targetserials.Add(i.EntitySerialNumber)
                
                # apply trhe changes to all serials
                for sn in targetserials:

                    surface2 = self.currentProject.Concordance.Lookup(sn)
                    
                    if surface2 and surface1 != surface2:
                        
                        surface2.Mode = surface1.Mode

                        if self.color.IsChecked:
                            surface2.Color = surface1.Color
                        if self.rebuildmethod.IsChecked:
                            surface2.RebuildMethod = surface1.RebuildMethod
                        if self.transparency.IsChecked:
                            surface2.TransparencyPercentage = surface1.TransparencyPercentage

                        if not (isinstance(surface1, ProjectedSurface) or isinstance(surface2, ProjectedSurface)):

                            if self.maxedgelength.IsChecked:
                                surface2.MaxEdgeLength = surface1.MaxEdgeLength
                            if self.maxedgeangle.IsChecked:
                                surface2.MaxEdgeAngle = surface1.MaxEdgeAngle

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

        self.SaveOptions()

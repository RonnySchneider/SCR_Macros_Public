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
    cmdData.Key = "SCR_QuickDynaToSheet"
    cmdData.CommandName = "SCR_QuickDynaToSheet"
    cmdData.Caption = "_SCR_QuickDynaToSheet"
    cmdData.UIForm = "SCR_QuickDynaToSheet"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Reports"
        cmdData.DefaultTabGroupKey = "Sheets and Dynaviews"
        cmdData.ShortCaption = "Dynaview to Sheet"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.26
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Dynaview to Sheet"
        cmdData.ToolTipTextFormatted = "Dynaview to Sheet"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_QuickDynaToSheet(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_QuickDynaToSheet.xaml") as s:
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

        optionMenu = SelectionContextMenuHandler()
        # remove options that don't apply here
        optionMenu.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        self.objs.ButtonContextMenu = optionMenu

        self.linepicker1.IsEntityValidCallback=self.IsValid
        
        self.lType = clr.GetClrType(IPolyseg)

        self.coordpick2.ValueChanged += self.coordpick2Changed

        self.dynascale.MinValue = 0
        self.dynascale.NumberOfDecimals = 2

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear

        #self.lfp = self.lunits.Properties.Copy()
        self.linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation
        if self.lunits.DisplayType == LinearType.Meter:
            self.linearsuffixsmall = self.lunits.Units[LinearType.Millimeter].Abbreviation
        else:
            self.linearsuffixsmall = self.lunits.Units[LinearType.Inch].Abbreviation

        self.labelx.Content = 'X in  [' + self.linearsuffixsmall + ']'
        self.labely.Content = 'Y in  [' + self.linearsuffixsmall + ']'

        self.dynax.DisplayUnit = LinearType.DisplaySmall
        self.dynay.DisplayUnit = LinearType.DisplaySmall

        #self.dynax.NumberOfDecimals = 1
        #self.dynay.NumberOfDecimals = 1

        self.sheetnameincstart.MinValue = 0
        self.sheetnameincstart.NumberOfDecimals = 0
        self.sheetnameinc.MinValue = 0
        self.sheetnameinc.NumberOfDecimals = 0

        self.linepicker1.ValueChanged += self.lineChanged

        vfc = self.currentProject [Project.FixedSerial.ViewFilterCollection] # getting all View Filters in one list
        
        self.viewfilterpicker.SearchContainer = Project.FixedSerial.ViewFilterCollection
        self.viewfilterpicker.UseSelectionEngine = False
        self.viewfilterpicker.FilterByEntityTypes = Array[Type]([clr.GetClrType(ViewFilter)])

        ## to have the drop down list sorted we fill a templist first, sort it and then fill the dropdownbox
        #templist = []
        #for f in vfc:
        #    templist.Add(f.Name)
        #templist.sort()
        #
        ## fill the viewfilter drop down list from the sorted temp list
        #for f in templist:
        #    item = ComboBoxItem()
        #    item.Content = f
        #    item.FontSize = 12
        #    self.viewfilterpicker.Items.Add(item)
        

        self.sheetsetpicker.SearchContainer = 1 # ProjectFixedSerial 1 is the project itself
        self.sheetsetpicker.SearchSubContainer = True
        self.sheetsetpicker.UseSelectionEngine = False
        self.sheetsetpicker.FilterByEntityTypes = Array[Type]([clr.GetClrType(SheetSet)])
        self.sheetsetpicker.ValueChanged += self.resetsheetpicker

        self.sheetpicker.SearchContainer = self.sheetsetpicker.SelectedSerial
        self.sheetpicker.SearchSubContainer = True
        self.sheetpicker.UseSelectionEngine = False
        self.sheetpicker.FilterByEntityTypes = Array[Type]([clr.GetClrType(BasicSheet)])

        self.resetsheetpicker(None, None)
        
        # Filter combine test
        #self.sheetsetpicker.FilterByEntityTypes = Array[Type]([clr.GetClrType(PlanSetSheetView), clr.GetClrType(SheetSet), clr.GetClrType(BasicSheet)])

        ## to have the drop down list sorted we fill a templist first, sort it and then fill the dropdownbox
        #templist = []
        ## fill the sheetset drop down list
        #for o in self.currentProject:
        #    if isinstance(o, PlanSetSheetViews): # Top Entry in Project Explorer "Plan Sets"
        #        allplansets = o
        #        for planset in allplansets:
        #            if isinstance(planset, PlanSetSheetView): # single Plan Set
        #                for sheetset in planset:
        #                    if isinstance(sheetset, SheetSet): # single Sheet Set
        #                        templist.Add(planset.Name + ' - ' + sheetset.Name)
        #templist.sort()
        ## fill the sheetset list from the sorted temp list
        #for f in templist:
        #    item = ComboBoxItem()
        #    item.Content = f
        #    item.FontSize = 12
        #    self.sheetsetpicker.Items.Add(item)

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def resetsheetpicker(self, ctrl, e):

        self.sheetpicker.SearchContainer = self.sheetsetpicker.SelectedSerial # ProjectFixedSerial 1 is the project itself


    def SetDefaultOptions(self):
        settings = ConstructionCommandsSettings.ProvideObject(self.currentProject)

        try:
            self.sheetsetpicker.SelectBySerialNumber(settings.GetInt32("SCR_QuickDynaToSheet.sheetsetpicker", 0))
        except:
            pass
        
        settings.SetInt32("SCR_QuickDynaToSheet.sheetsetpicker", self.sheetsetpicker.SelectedSerial)
        
        try:
            self.sheetpicker.SelectBySerialNumber(settings.GetInt32("SCR_QuickDynaToSheet.sheetpicker", 0))
        except:
            pass
        
        try:
            self.viewfilterpicker.SelectBySerialNumber(settings.GetInt32("SCR_QuickDynaToSheet.viewfilterpicker", 0))
        except:
            pass

        settingserial = settings.GetUInt32("SCR_QuickDynaToSheet.dynalayerpicker", 8) # 8 is FixedSerial for Layer Zero
        o = self.currentProject.Concordance.Lookup(settingserial)
        if o != None:
            if isinstance(o.GetSite(), LayerCollection):    
                self.dynalayerpicker.SetSelectedSerialNumber(settingserial, InputMethod(3))
            else:                       
                self.dynalayerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:                       
            self.dynalayerpicker.SetSelectedSerialNumber(8, InputMethod(3))

        self.dynax.Distance = settings.GetDouble("SCR_QuickDynaToSheet.dynax", 10.000)
        self.dynay.Distance = settings.GetDouble("SCR_QuickDynaToSheet.dynay", 35.000)
        self.dynascale.Value = settings.GetDouble("SCR_QuickDynaToSheet.dynascale", 20.000)
        self.sheetnameprefix.Text = settings.GetString("SCR_QuickDynaToSheet.sheetnameprefix", "Frame ")
        self.sheetnameincstart.Value = settings.GetDouble("SCR_QuickDynaToSheet.sheetnameincstart", 1)
        self.sheetnameinc.Value = settings.GetDouble("SCR_QuickDynaToSheet.sheetnameinc", 1)
        self.easymassframe.IsChecked = settings.GetBoolean("SCR_QuickDynaToSheet.easymassframe", True)
        self.useexistingsheet.IsChecked = settings.GetBoolean("SCR_QuickDynaToSheet.useexistingsheet", False)
        self.manualpick.IsChecked = settings.GetBoolean("SCR_QuickDynaToSheet.manualpick", False)


    def SaveOptions(self):
        settings = ConstructionCommandsSettings.ProvideObject(self.currentProject)

        settings.SetBoolean("SCR_QuickDynaToSheet.useexistingsheet", self.useexistingsheet.IsChecked)
        settings.SetInt32("SCR_QuickDynaToSheet.sheetpicker", self.sheetpicker.SelectedSerial)
        settings.SetInt32("SCR_QuickDynaToSheet.viewfilterpicker", self.viewfilterpicker.SelectedSerial)
        settings.SetUInt32("SCR_QuickDynaToSheet.dynalayerpicker", self.dynalayerpicker.SelectedSerialNumber)
        settings.SetDouble("SCR_QuickDynaToSheet.dynax", self.dynax.Distance)
        settings.SetDouble("SCR_QuickDynaToSheet.dynay", self.dynay.Distance)
        settings.SetDouble("SCR_QuickDynaToSheet.dynascale", self.dynascale.Value)
        settings.SetString("SCR_QuickDynaToSheet.sheetnameprefix", self.sheetnameprefix.Text)
        settings.SetDouble("SCR_QuickDynaToSheet.sheetnameincstart", self.sheetnameincstart.Value)
        settings.SetDouble("SCR_QuickDynaToSheet.sheetnameinc", self.sheetnameinc.Value)
        settings.SetBoolean("SCR_QuickDynaToSheet.easymassframe", self.easymassframe.IsChecked)
        settings.SetBoolean("SCR_QuickDynaToSheet.manualpick", self.manualpick.IsChecked)
    
    def IsValid(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        return False

    def easymassframeChanged(self, sender, e):
        if self.easymassframe.IsChecked:
            self.objs.IsEnabled = True
            self.singleframe.IsEnabled = False
        else:
            self.objs.IsEnabled = False
            self.singleframe.IsEnabled = True

    def coordpick2Changed(self, ctrl, e):
        # set keyboard focus if change was due to mouse pick
        if e.Cause == InputMethod.Mouse:     
            self.OkClicked(None, None)

    def lineChanged(self, ctrl, e):
        # set keyboard focus if change was due to mouse pick
        if e.Cause == InputMethod.Mouse and self.manualpick.IsChecked == False: # when we don't want to pick the corners
            self.OkClicked(None, None)
            self.linepicker1.AutoTab = False
        else:
            self.linepicker1.AutoTab = True
            Keyboard.Focus(self.coordpick1)

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()


    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''


        self.success.Content = ''

        wv = self.currentProject [Project.FixedSerial.WorldView]
        #wv.PauseGraphicsCache(True)

        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())

        try:
            # the "with" statement will unroll any changes if something go wrong
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                frameserials = []
                if self.easymassframe.IsChecked:
                    for o in self.objs:
                        if isinstance(o, self.lType) and not isinstance(o, BlockReference):
                            frameserials.Add(o.SerialNumber)
                else:
                    if isinstance(self.linepicker1.Entity, self.lType):
                        frameserials.Add(self.linepicker1.Entity.SerialNumber)

                frameserials.sort()

                inputok=True

                if (self.viewfilterpicker.SelectedIndex < 0) or (self.sheetsetpicker.SelectedIndex < 0):
                    self.error.Content += '\nSelect a Sheetset/Viewfilter'
                    inputok = False
                else:
                    vfserial = self.viewfilterpicker.SelectedSerial
                    ssserial = self.sheetsetpicker.SelectedSerial

                if self.useexistingsheet.IsChecked and (self.sheetpicker.SelectedIndex < 0):
                    self.error.Content += '\nSelect a Sheet'
                    inputok = False
                else:
                    sserial = self.sheetpicker.SelectedSerial

                if frameserials.Count > 0 and inputok:

                    for serial in frameserials: # go through the serial number list

                        # get the line as object
                        l1 = self.currentProject.Concordance.Lookup(serial)
                        
                        # retrieve two points for the frame orientation
                        #automatically
                        if self.easymassframe.IsChecked or (self.easymassframe.IsChecked == False and self.manualpick.IsChecked == False):
                            # even if the polyline starts in the bottom left it could be counterclockwise
                            # either way we get the 2 coordinates in the right order
                            polyseg1 = l1.ComputePolySeg()
                            n = polyseg1.ToPoint3DArray()
                            if polyseg1.IsClockWise():
                                p1 = n[n.Count - 1]
                                p2 = n[n.Count - 2]
                            else:
                                p1 = n[0]
                                p2 = n[1]
                        # manually
                        if self.manualpick.IsChecked and self.easymassframe.IsChecked == False :
                            p1 = self.coordpick1.Coordinate
                            p2 = self.coordpick2.Coordinate

                        # fill the sheets
                        # get the sheetset as object
                        sheetset = self.currentProject.Concordance.Lookup(ssserial)
                        
                        if self.useexistingsheet.IsChecked:
                            newsheet = self.currentProject.Concordance.Lookup(sserial)
                        else:
                            #create a new sheet in the sheetset
                            newsheet = sheetset.Add(clr.GetClrType(BasicSheet))
                            if not math.isnan(self.sheetnameincstart.Value):
                                newsheet.Name = self.sheetnameprefix.Text + str(int(self.sheetnameincstart.Value))
                            else:
                                newsheet.Name = self.sheetnameprefix.Text

                            if not math.isnan(self.sheetnameincstart.Value):
                                newsheet.SortRank = int(self.sheetnameincstart.Value * 10)
                            else:
                                newsheet.SortRank = 0

                        # create a dynaview on the new sheet
                        newdynaview = newsheet.Add(clr.GetClrType(DynaView))
                        newdynaview.ViewFilter = vfserial
                        newdynaview.Boundary = l1.SerialNumber
                        newdynaview.Layer = self.dynalayerpicker.SelectedSerialNumber
                        newdynaview.Location = Point3D(self.dynax.Distance, self.dynay.Distance, 0)
                        if not math.isnan(self.sheetnameincstart.Value):
                            newdynaview.Name = self.sheetnameprefix.Text + str(int(self.sheetnameincstart.Value))
                            if not math.isnan(self.sheetnameinc.Value):
                                self.sheetnameincstart.Value += self.sheetnameinc.Value
                        else:
                            newdynaview.Name = self.sheetnameprefix.Text
    
                        newdynaview.Scale = self.dynascale.Value
                        newdynaview.Rotation = Vector3D(p1, p2).Azimuth - math.pi/2



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

        if self.easymassframe.IsChecked:
            Keyboard.Focus(self.objs)
        else:
            Keyboard.Focus(self.linepicker1)

        #wv.PauseGraphicsCache(False)

        #self.currentProject.Calculate(False)

        self.SaveOptions()

    


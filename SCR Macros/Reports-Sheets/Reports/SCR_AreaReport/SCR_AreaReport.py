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
    cmdData.Key = "SCR_AreaReport"
    cmdData.CommandName = "SCR_AreaReport"
    cmdData.Caption = "_SCR_AreaReport"
    cmdData.UIForm = "SCR_AreaReport"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Reports"
        cmdData.DefaultTabGroupKey = "Reports"
        cmdData.ShortCaption = "Area Report"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.07
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "reports the Area and Chainages of a closed Polygon along an Alignment"
        cmdData.ToolTipTextFormatted = "reports the Area and Chainages of a closed Polygon along an Alignment"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png") 
        cmdData.ImageSmall = b
    except:
        pass

class SCR_AreaReport(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_AreaReport.xaml") as s:
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

        self.lType = clr.GetClrType(IPolyseg)

        self.linepicker1.IsEntityValidCallback = self.IsValidLine
        self.linepicker1.ValueChanged += self.lineChanged
        self.linepicker2.IsEntityValidCallback = self.IsValidLine

        optionMenu = SelectionContextMenuHandler()
        # remove options that don't apply here
        optionMenu.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        self.objs.ButtonContextMenu = optionMenu
        self.objs.IsEntityValidCallback = self.IsValidLine

        # get the units for chainage distance
        self.chunits = self.currentProject.Units.Station
        self.chfp = self.chunits.Properties.Copy()

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def SetDefaultOptions(self):
        self.automode.IsChecked = OptionsManager.GetBool("SCR_SCR_AreaReport.automode", True)
        self.manualmode.IsChecked = OptionsManager.GetBool("SCR_SCR_AreaReport.manualmode", False)

    def SaveOptions(self):
        OptionsManager.SetValue("SCR_SCR_AreaReport.automode", self.automode.IsChecked)  
        OptionsManager.SetValue("SCR_SCR_AreaReport.manualmode", self.manualmode.IsChecked)  


    def lineChanged(self, ctrl, e):
        l1 = self.linepicker1.Entity
        if l1 != None:
            self.stationframe.IsEnabled = True
            self.startstation.StationProvider = l1
            self.endstation.StationProvider = l1
        else:
            self.stationframe.IsEnabled = False


    def IsValidLine(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        return False

    def CancelClicked(self, sender, args):
        sender.CloseUICommand()

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''
        self.success.Content=''

        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        
        wv = self.currentProject [Project.FixedSerial.WorldView]
        
        try:
            # the "with" statement will unroll any changes if something go wrong
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                
                l2list = []
                if self.automode.IsChecked:
                    for l2 in self.objs:
                        if isinstance(l2, self.lType):
                            l2list.Add(l2)
                else:
                    l2list.Add(self.linepicker2.Entity)
                    stations = []
                    stations.Add(self.startstation.Distance)
                    stations.Add(self.endstation.Distance)
                    stations.sort()

                # open/create the temp CSV-File and write the header information
                filename = os.path.expanduser('~/Downloads/Polygon-Areas.csv')
                if File.Exists(filename):
                        File.Delete(filename)
                with open(filename, 'w') as f:

                    l1 = self.linepicker1.Entity

                    polyseg1 = l1.ComputePolySeg()
                    polyseg1 = polyseg1.ToWorld()

                    f.write(",Area Report\n")
                    f.write("\n")
                    f.write("Layer:\n")
                    f.write("Paving Date:\n")
                    f.write('Reference line:,' + IName.Name.__get__(l1) + "\n")
                    f.write("\n")
                    f.write("Report" + ',' + "Area" + ',' +  "Chainage Start" + ',' + "Chainage End\n")


                    # compute the intersections/areas between all outlines and the reference Alignment
                    intersections = Intersections()

                    for l2 in l2list:
                        if isinstance(l2, self.lType):

                            # get the line data as polyseg, in world coordinates
                            polyseg2 = l2.ComputePolySeg()
                            polyseg2 = polyseg2.ToWorld()
                            # just in case, close it if it's open
                            if not polyseg2.IsClosed:
                                polyseg2.Close(True)

                            (p2area, centroid, side) = polyseg2.Area()

                            if self.automode.IsChecked:
                                polyseg1.Intersect(polyseg2, True, intersections)

                                stations = []
                                if intersections.Count == 2:
                                    for i in intersections:

                                        (found, pnt, sta) = polyseg1.FindPointFromPoint(i.Point)
                                        stations.Add(sta)
                                    stations.sort()
                                    f.write(IName.Name.__get__(l2) + ',' + str(p2area) + ',' + str(stations[0]) + ',' + str(stations[1]) + "\n")
                                    #pass
                                else:
                                    self.error.Content += '\n' + IName.Name.__get__(l2) + ' - found wrong number of intersections '
                            
                            else:
                                # manually selected stations
                                f.write(IName.Name.__get__(l2) + ',' + str(p2area) + ',' + str(stations[0]) + ',' + str(stations[1]) + "\n")

                            #self.success.Content += '\nArea - ' + str(p2area)
                            #self.success.Content += '\nCh-Start - ' + str(stations[0])
                            #self.success.Content += '\nArea - ' + str(stations[1])

                            # write the results into the CSV
                         
                        intersections.Clear()


                # start excel if we can and show the CSV
                os.system('start excel.exe "%s"' % (filename,))

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
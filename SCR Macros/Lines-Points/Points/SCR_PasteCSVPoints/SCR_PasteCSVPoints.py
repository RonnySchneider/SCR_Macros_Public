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
    cmdData.Key = "SCR_PasteCSVPoints"
    cmdData.CommandName = "SCR_PasteCSVPoints"
    cmdData.Caption = "_SCR_PasteCSVPoints"
    cmdData.UIForm = "SCR_PasteCSVPoints"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "Points"
        cmdData.ShortCaption = "paste CSV Points"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.02
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "paste CSV data from Clipboard"
        cmdData.ToolTipTextFormatted = "paste CSV data from Clipboard"

    except:
        pass
    
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_PasteCSVPoints(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_PasteCSVPoints.xaml") as s:
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

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass


    def SetDefaultOptions(self):
        self.c1isID.IsChecked = OptionsManager.GetBool("SCR_PasteCSVPoints.c1isID", True)
        self.c1isEasting.IsChecked = OptionsManager.GetBool("SCR_PasteCSVPoints.c1isEasting", False)
        self.createnamed.IsChecked = OptionsManager.GetBool("SCR_PasteCSVPoints.createnamed", True)
        self.createsimple.IsChecked = OptionsManager.GetBool("SCR_PasteCSVPoints.createsimple", False)


    def SaveOptions(self):
        OptionsManager.SetValue("SCR_PasteCSVPoints.c1isID", self.c1isID.IsChecked)
        OptionsManager.SetValue("SCR_PasteCSVPoints.c1isEasting", self.c1isEasting.IsChecked)
        OptionsManager.SetValue("SCR_PasteCSVPoints.createnamed", self.createnamed.IsChecked)
        OptionsManager.SetValue("SCR_PasteCSVPoints.createsimple", self.createsimple.IsChecked)


    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''

        #if File.Exists(self.openfilename.Text) == False:
        #    self.error.Content += 'no file selected\n'
        
        
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())

        wv = self.currentProject [Project.FixedSerial.WorldView]
        wv.PauseGraphicsCache(True)

        #find PointManager as object
        for o in self.currentProject:
            if isinstance(o, PointManager):
                pm = o

        uda = self.currentProject.Lookup(27) # User Defined Attributes are always Serial# 27
            
        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                pnew_name = "PasteCSV_0001"
                fc = ""
                
                # if we don't have a id column we could live with just x and y
                if self.c1isID.IsChecked:
                    mindata = 3
                else:
                    mindata = 2
                #tt0 = Clipboard.GetText()

                data = [re.split('\t|,', r) for r in [r for r in Clipboard.GetText().split("\r\n")]]
                if data and data.Count > 0:
                    for pointdata in data:
                        enoughdata = False
                        if pointdata.Count >= mindata:
                            if mindata == 2 and pointdata[0].replace(".", "").isnumeric() and pointdata[1].replace(".", "").isnumeric():
                                enoughdata = True
                            elif mindata == 3 and pointdata[1].replace(".", "").isnumeric() and pointdata[2].replace(".", "").isnumeric():
                                enoughdata = True

                        if enoughdata:
                            pnew = Point3D()
                            
                            if self.c1isID.IsChecked:
                                pnew_name = pointdata[0]
                                pnew.X = float(pointdata[1])
                                pnew.Y = float(pointdata[2])
                                if pointdata.Count >= 4 and pointdata[3].replace(".", "").isnumeric():
                                    pnew.Z = float(pointdata[3])
                                if pointdata.Count >= 5:
                                    fc = pointdata[4]
                            else:
                                pnew.X = float(pointdata[0])
                                pnew.Y = float(pointdata[1])
                                if pointdata.Count >= 3 and pointdata[2].replace(".", "").isnumeric():
                                    pnew.Z = float(pointdata[2])
                                if pointdata.Count >= 4:
                                    fc = pointdata[3]

                            if self.c1isEasting.IsChecked and self.createnamed.IsChecked:
                                # find free point number
                                while PointHelper.Helpers.Find(self.currentProject, pnew_name).Count != 0:
                                    pnew_name = PointHelper.Helpers.Increment(pnew_name, None, True)
                                    
                            if self.createnamed.IsChecked:
                                pnew_named = CoordPoint.CreatePoint(self.currentProject, pnew_name)
                                #pnew_named.AddPosition(pnew)
                                keyed_coord = KeyedIn(CoordSystem.eGrid, pnew.Y, pnew.X, CoordQuality.eSurvey, pnew.Z, CoordQuality.eSurvey, CoordComponentType.eElevation, System.DateTime.UtcNow)
                                OfficeEnteredCoord.AddOfficeEnteredCoord(self.currentProject, pnew_named, keyed_coord)
                                
                                try:
                                    pm.SetFeatureCodeAtPoint(pnew_named.SerialNumber, fc)
                                except:
                                    pass

                            else:

                                cadPoint = wv.Add(clr.GetClrType(CadPoint))
                                cadPoint.Point0 = pnew
                                attdic = {"Id":pnew_name,"FeatureCode":fc,"Description1":"","Description2":"","FeatureAttributes":None,"TwelveDAttributes":None,"Key":"ConvertedCADPointDisplay"}
                                paa = uda.GetAssociatedAttributes(cadPoint.SerialNumber)
                                paa['ConvertedCADPointDisplay'] = json.dumps(attdic)


                
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
        
        ProgressBar.TBC_ProgressBar.Title = ""

        wv.PauseGraphicsCache(False)
        self.SaveOptions()


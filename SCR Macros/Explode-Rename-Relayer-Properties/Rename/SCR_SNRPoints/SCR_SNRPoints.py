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
    cmdData.Key = "SCR_SNRPoints"
    cmdData.CommandName = "SCR_SNRPoints"
    cmdData.Caption = "_SCR_SNRPoints"
    cmdData.UIForm = "SCR_SNRPoints"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Expld-SNR-Relay-Prop"
        cmdData.DefaultTabGroupKey = "Renaming"
        cmdData.ShortCaption = "SNR Points+FC"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.02
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "SNR Points or Featurecodes"
        cmdData.ToolTipTextFormatted = "search and replace in Point-Names or Feature-Codes"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_SNRPoints(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_SNRPoints.xaml") as s:
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
        self.objs.IsEntityValidCallback=self.IsValid
        optionMenu = SelectionContextMenuHandler()
        # remove options that don't apply here
        optionMenu.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        self.objs.ButtonContextMenu = optionMenu

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def SetDefaultOptions(self):
        settings = ConstructionCommandsSettings.ProvideObject(self.currentProject)
        self.snrpoint.IsChecked = settings.GetBoolean("SCR_SNRPoints.snrpoint", True)
        self.snrfc.IsChecked = settings.GetBoolean("SCR_SNRPoints.snrfc", False)
        self.radio1.IsChecked = settings.GetBoolean("SCR_SNRPoints.radio1", True)
        self.radio2.IsChecked = settings.GetBoolean("SCR_SNRPoints.radio2", False)

        self.textBox1.Text = settings.GetString("SCR_SNRPoints.textBox1", "")
        self.textBox2.Text = settings.GetString("SCR_SNRPoints.textBox2", "")
        self.textBox3.Text = settings.GetString("SCR_SNRPoints.textBox3", "")
        self.textBox4.Text = settings.GetString("SCR_SNRPoints.textBox4", "")

    def SaveOptions(self):
        settings = ConstructionCommandsSettings.ProvideObject(self.currentProject)
        settings.SetBoolean("SCR_SNRPoints.snrpoint", self.snrpoint.IsChecked)
        settings.SetBoolean("SCR_SNRPoints.snrfc", self.snrfc.IsChecked)
        settings.SetBoolean("SCR_SNRPoints.radio1", self.radio1.IsChecked)
        settings.SetBoolean("SCR_SNRPoints.radio2", self.radio2.IsChecked)

        settings.SetString("SCR_SNRPoints.textBox1", self.textBox1.Text)
        settings.SetString("SCR_SNRPoints.textBox2", self.textBox2.Text)
        settings.SetString("SCR_SNRPoints.textBox3", self.textBox3.Text)
        settings.SetString("SCR_SNRPoints.textBox4", self.textBox4.Text)

    def IsValid(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, CoordPoint):
            return True
        return False
        

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()


    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Text = ''
        self.success.Content = ''

        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        
        #wv = self.currentProject [Project.FixedSerial.WorldView]
        wl = self.currentProject[Project.FixedSerial.LayerContainer]
        
        try:
            # the "with" statement will unroll any changes if something go wrong
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                for o in self.currentProject:
                #find PointManager as object
                    if isinstance(o, PointManager):
                        pm = o

                for o in self.objs:

                    if isinstance(o, CoordPoint):

                        if self.snrpoint.IsChecked:
                            
                            if self.fctopointname.IsChecked:

                                o.PointID = ""

                                features = pm.AssociatedRDFeatures(o.SerialNumber)

                                for fc in features:

                                    o.PointID += fc.Code
                                

                            else:
                                
                                oldstring = o.PointID
                        
                                # string.replace(oldvalue, newvalue, count)
                                if self.radio1.IsChecked and self.textBox1.Text != self.textBox2.Text:
                                    oldstring = oldstring.replace(self.textBox1.Text, self.textBox2.Text)

                                if self.radio2.IsChecked:
                                    if self.radio1.IsChecked and self.textBox1.Text == self.textBox2.Text:
                                        if self.textBox1.Text in oldstring:
                                            oldstring = self.textBox3.Text + oldstring + self.textBox4.Text
                                    else:
                                        oldstring = self.textBox3.Text + oldstring + self.textBox4.Text
                                o.PointID = oldstring
                            
                        else:

                            features = pm.AssociatedRDFeatures(o.SerialNumber)
                            for fc in features:
                                oldstring = fc.Code
                            
                                # string.replace(oldvalue, newvalue, count)
                                if self.radio1.IsChecked and self.textBox1.Text != self.textBox2.Text:
                                    oldstring = oldstring.replace(self.textBox1.Text, self.textBox2.Text)

                                if self.radio2.IsChecked:
                                    if self.radio1.IsChecked and self.textBox1.Text == self.textBox2.Text:
                                        if self.textBox1.Text in oldstring:
                                            oldstring = self.textBox3.Text + oldstring + self.textBox4.Text
                                    else:
                                        oldstring = self.textBox3.Text + oldstring + self.textBox4.Text

                                fc.Code = oldstring
                        
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
            self.error.Text += '\nan Error occurred - Result probably incomplete\n' + str(exc_type) + '\n' + str(exc_obj) + '\nLine ' + str(exc_tb.tb_lineno)

        self.SaveOptions()
            

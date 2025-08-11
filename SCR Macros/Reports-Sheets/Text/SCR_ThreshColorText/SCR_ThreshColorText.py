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
    cmdData.Key = "SCR_ThreshColorText"
    cmdData.CommandName = "SCR_ThreshColorText"
    cmdData.Caption = "_SCR_ThreshColorText"
    cmdData.UIForm = "SCR_ThreshColorText"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Reports"
        cmdData.DefaultTabGroupKey = "Text"
        cmdData.ShortCaption = "Recolor Text"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.14
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Recolor Text based on threshold"
        cmdData.ToolTipTextFormatted = "Recolor Text based on threshold"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_ThreshColorText(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_ThreshColorText.xaml") as s:
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

        self.objs.IsEntityValidCallback = self.IsValidText
        
        optionMenu = SelectionContextMenuHandler()
        # remove options that don't apply here
        optionMenu.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        self.objs.ButtonContextMenu = optionMenu

        self.mtextType = clr.GetClrType(MText)
        self.cadtextType = clr.GetClrType(CadText)


		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def SetDefaultOptions(self):
        self.searchtext.Text = OptionsManager.GetString("SCR_ThreshColorText.searchtext", "dZ=")

        self.thresh1.Text = OptionsManager.GetString("SCR_ThreshColorText.thresh1", "0.005")
        self.thresh2.Text = OptionsManager.GetString("SCR_ThreshColorText.thresh2", "0.010")
        self.thresh3.Text = OptionsManager.GetString("SCR_ThreshColorText.thresh3", "0.015")
        try:    self.thresh1colorpicker.SelectedColor = Color.FromArgb(OptionsManager.GetInt("SCR_ThreshColorText.thresh1colorpicker"))
        except: self.thresh1colorpicker.SelectedColor=Color.Green
        try:    self.thresh2colorpicker.SelectedColor = Color.FromArgb(OptionsManager.GetInt("SCR_ThreshColorText.thresh2colorpicker"))
        except: self.thresh2colorpicker.SelectedColor=Color.Yellow
        try:    self.thresh3colorpicker.SelectedColor = Color.FromArgb(OptionsManager.GetInt("SCR_ThreshColorText.thresh3colorpicker"))
        except: self.thresh3colorpicker.SelectedColor=Color.Magenta
        try:    self.thresh4colorpicker.SelectedColor = Color.FromArgb(OptionsManager.GetInt("SCR_ThreshColorText.thresh4colorpicker"))
        except: self.thresh4colorpicker.SelectedColor=Color.Red


    def SaveOptions(self):
        OptionsManager.SetValue("SCR_ThreshColorText.searchtext", self.searchtext.Text)

        OptionsManager.SetValue("SCR_ThreshColorText.thresh1", self.thresh1.Text)
        OptionsManager.SetValue("SCR_ThreshColorText.thresh2", self.thresh2.Text)
        OptionsManager.SetValue("SCR_ThreshColorText.thresh3", self.thresh3.Text)

        OptionsManager.SetValue("SCR_ThreshColorText.thresh1colorpicker", self.thresh1colorpicker.SelectedColor.ToArgb())
        OptionsManager.SetValue("SCR_ThreshColorText.thresh2colorpicker", self.thresh2colorpicker.SelectedColor.ToArgb())
        OptionsManager.SetValue("SCR_ThreshColorText.thresh3colorpicker", self.thresh3colorpicker.SelectedColor.ToArgb())
        OptionsManager.SetValue("SCR_ThreshColorText.thresh4colorpicker", self.thresh4colorpicker.SelectedColor.ToArgb())

    def IsValidText(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.mtextType) or isinstance(o, self.cadtextType) or isinstance(o, CadLabel):
            return True
        return False

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''

        self.success.Content += ''
        
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())

        wv = self.currentProject [Project.FixedSerial.WorldView]
        bc = self.currentProject [Project.FixedSerial.BlockCollection]    # getting all blocks as collection

        #check thresholds
        thresholds=True
        try:
            thresh1 = float(self.thresh1.Text)
            thresh2 = float(self.thresh2.Text)
            thresh3 = float(self.thresh3.Text)
        except:
            self.success.Content += '\ncheck Threshold Values'

            thresholds=False
        
        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                
                recolorcount = 0
                chars = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ!"#$%&\'()*,/:;<=>?@[\\]^_`{|}~ \t\n\r\x0b\x0c'

                for o in self.objs:
                    if isinstance(o, self.mtextType) or isinstance(o, self.cadtextType) or isinstance(o, CadLabel):
                        text = o.TextString
                        textlist = o.TextString.split('\\P')
                        for t in textlist:
                            if self.searchtext.Text in t:
                                t = t[t.find(self.searchtext.Text) + len(self.searchtext.Text) : len(t)] # get the string part behind the search expression
                                #t = t.translate(None, chars) # remove any non-numeric characters that might be left
                                t = re.sub(r'[^\d.+-]', '', t)
                                try:
                                    textnumber = float(t)
                
                                    # find the right color for the text
                                    if textnumber <= thresh1: abcolor = self.thresh1colorpicker.SelectedColor
                                    if textnumber > thresh1: abcolor = self.thresh2colorpicker.SelectedColor
                                    if textnumber > thresh2: abcolor = self.thresh3colorpicker.SelectedColor
                                    if textnumber > thresh3: abcolor = self.thresh4colorpicker.SelectedColor
    
                                    o.Color = abcolor

                                    recolorcount += 1
                                except:
                                    pass
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
        
        if recolorcount == 0:
            self.success.Content += "\ndidn't find any Text with the search Expression"
        else:
            self.success.Content += '\nre-colored ' + str(recolorcount) + ' Texts'


        self.SaveOptions()           

        
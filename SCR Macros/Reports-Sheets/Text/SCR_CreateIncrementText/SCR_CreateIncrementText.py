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
    cmdData.Key = "SCR_CreateIncrementText"
    cmdData.CommandName = "SCR_CreateIncrementText"
    cmdData.Caption = "_SCR_CreateIncrementText"
    cmdData.UIForm = "SCR_CreateIncrementText"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Reports"
        cmdData.DefaultTabGroupKey = "Text"
        cmdData.ShortCaption = "Create Incrementing Text"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.08
        
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Create Incrementing Text"
        cmdData.ToolTipTextFormatted = "Create Incrementing Text"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_CreateIncrementText(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_CreateIncrementText.xaml") as s:
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
        
        #self.textheight.NumberOfDecimals = 2
        #self.textheight.MinValue = 0.001
        #self.textoffsetx.NumberOfDecimals = 3
        #self.textoffsety.NumberOfDecimals = 3
        
        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        #self.lfp = self.lunits.Properties.Copy()
        self.linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation

        self.labeltextheight.Content = 'Text Height [' + self.linearsuffix + ']'
        self.labelx.Content = 'offset X-Direction [' + self.linearsuffix + ']'
        self.labely.Content = 'offset Y-Direction [' + self.linearsuffix + ']'

        self.textheight.DistanceMin = 0
        
        self.startnumber.NumberOfDecimals = 0
        self.startnumber.MinValue = 0
        self.increment.NumberOfDecimals = 0
        self.increment.MinValue = 0
        self.numberlength.NumberOfDecimals = 0
        self.numberlength.MinValue = 1


        self.coordpick1.ShowElevationIf3D = True
        self.coordpick2.ShowElevationIf3D = True
        self.coordpick1.ValueChanged += self.coordpick1changed
        self.coordpick2.ValueChanged += self.coordpick2changed
        if self.usefixedbearing.IsChecked:        
            self.coordpick1.AutoTab = False
        else:
            self.coordpick1.AutoTab = True
        
		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def usefixedbearingChanged(self, sender, e):
        if self.usefixedbearing.IsChecked:
            self.bearingpicker.IsEnabled = True
            self.coordpick1.AutoTab = False
            self.coordpick2.IsEnabled = False
        else:
            self.bearingpicker.IsEnabled = False
            self.coordpick1.AutoTab = True
            self.coordpick2.IsEnabled = True

    def SetDefaultOptions(self):

        settingserial = OptionsManager.GetUint("SCR_CreateIncrementText.layerpicker", 8) # 8 is FixedSerial for Layer Zero
        o = self.currentProject.Concordance.Lookup(settingserial)
        if o != None:
            if isinstance(o.GetSite(), LayerCollection):    
                self.layerpicker.SetSelectedSerialNumber(settingserial, InputMethod(3))
            else:                       
                self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:                       
            self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        
        # changed some inputs from string to double, in order to avoid issues I added a "v" to the variable names  in the settings
        self.textheight.Distance = OptionsManager.GetDouble("SCR_CreateIncrementText.textheightv", 0.1)
        self.textoffsetx.Distance = OptionsManager.GetDouble("SCR_CreateIncrementText.textoffsetxv", 0.0)
        self.textoffsety.Distance = OptionsManager.GetDouble("SCR_CreateIncrementText.textoffsetyv", 0.0)
        self.prefix.Text = OptionsManager.GetString("SCR_CreateIncrementText.prefix", "")
        self.startnumber.Value = OptionsManager.GetDouble("SCR_CreateIncrementText.startnumberv", 0)
        self.increment.Value = OptionsManager.GetDouble("SCR_CreateIncrementText.incrementv", 1)
        self.numberlength.Value = OptionsManager.GetDouble("SCR_CreateIncrementText.numberlengthv", 1)

        self.usefixedbearing.IsChecked = OptionsManager.GetBool("SCR_CreateIncrementText.usefixedbearing", True)
        self.bearingpicker.Direction = OptionsManager.GetDouble("SCR_CreateIncrementText.bearingpicker", 0.000)
    
    def SaveOptions(self):

        OptionsManager.SetValue("SCR_CreateIncrementText.layerpicker", self.layerpicker.SelectedSerialNumber)

        # changed some inputs from string to double, in order to avoid issues I added a "v" to the variable names  in the settings
        OptionsManager.SetValue("SCR_CreateIncrementText.textheightv", self.textheight.Distance)
        OptionsManager.SetValue("SCR_CreateIncrementText.textoffsetxv", self.textoffsetx.Distance)
        OptionsManager.SetValue("SCR_CreateIncrementText.textoffsetyv", self.textoffsety.Distance)
        OptionsManager.SetValue("SCR_CreateIncrementText.prefix", self.prefix.Text)
        OptionsManager.SetValue("SCR_CreateIncrementText.startnumberv", self.startnumber.Value)
        OptionsManager.SetValue("SCR_CreateIncrementText.incrementv", self.increment.Value)
        OptionsManager.SetValue("SCR_CreateIncrementText.numberlengthv", self.numberlength.Value)

        OptionsManager.SetValue("SCR_CreateIncrementText.usefixedbearing", self.usefixedbearing.IsChecked)
        OptionsManager.SetValue("SCR_CreateIncrementText.bearingpicker", self.bearingpicker.Direction)

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def coordpick1changed(self, ctrl, e):
        if e.Cause == InputMethod.Mouse and self.usefixedbearing.IsChecked:
            self.OkClicked(None, None)
    
    def coordpick2changed(self, ctrl, e):
        # set keyboard focus if change was due to mouse pick
        if e.Cause == InputMethod.Mouse:     
            self.OkClicked(None, None)

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''

        self.success.Content = ""

        wv = self.currentProject [Project.FixedSerial.WorldView]

        inputok = True
        prefix = self.prefix.Text

        try:
            textheight = self.textheight.Distance
            textoffsetx = self.textoffsetx.Distance
            textoffsety = self.textoffsety.Distance
            startnumber = int(self.startnumber.Value)
            increment = int(self.increment.Value)
            numberlength = int(self.numberlength.Value)
        except:
            self.success.Content = "check Number Settings"
            inputok = False

        p1 = self.coordpick1.Coordinate

        if self.usefixedbearing.IsChecked:
            bearing = self.bearingpicker.Direction
            offsetvector = Vector3D(bearing)
        else:
            p2 = self.coordpick2.Coordinate
            bearing = math.pi/2 - Vector3D(p1, p2).Azimuth
            offsetvector = Vector3D(p1, p2)
        
        insertp = p1
        offsetvector.Length = textoffsetx
        insertp += offsetvector
        if textoffsetx > 0:
            offsetvector.Rotate90(Side.Left)
        else:
            offsetvector.Rotate90(Side.Right)
        offsetvector.Length = textoffsety
        insertp += offsetvector
       
        if inputok:
            self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
            UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
            try:
                with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                    t = wv.Add(clr.GetClrType(MText))
                    t.AlignmentPoint = insertp
                    t.AttachPoint = AttachmentPoint(7) # bottom left
                    t.TextString = prefix + str(int(self.startnumber.Value)).zfill(numberlength)

                    t.Height = textheight
                    t.Layer = self.layerpicker.SelectedSerialNumber
                    t.Rotation = bearing

                    self.startnumber.Value = startnumber + increment
                    
                    Keyboard.Focus(self.coordpick1)

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

        
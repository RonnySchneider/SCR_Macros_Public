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
    "changelinelayer": False,
    "newlinelayerpicker": 8,
    "changelinecolor": True,
    "changelinestyle": False,
    "changelinestylescale": False,
    "linestylescale": 1.0,
    "changelineweight": False,
    "changetextlayer": False,
    "newtextlayerpicker": 8,
    "changetextcolor": False,
    "changetextstyle": False,
    "changetextheight": False,
    "textheight": 1.0,
    "changetextweight": False,
}

def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_ChangePropNoSelect"
    cmdData.CommandName = "SCR_ChangePropNoSelect"
    cmdData.Caption = "_SCR_ChangePropNoSelect"
    cmdData.UIForm = "SCR_ChangePropNoSelect"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Expld-SNR-Relay-Prop"
        cmdData.DefaultTabGroupKey = "Properties"
        cmdData.ShortCaption = "Change Properties without Select"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.06
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Change Properties without Select"
        cmdData.ToolTipTextFormatted = "Change Properties without Select"

    except:
        pass
    
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_ChangePropNoSelect(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_ChangePropNoSelect.xaml") as s:
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
        self.unitssetup(None, None)

    def unitssetup(self, sender, e):
        # setup everything for the unit conversions

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        self.lfp = self.lunits.Properties.Copy() # create a copy in order to set the decimals and enable/disable the suffix
        self.lfp.AddSuffix = False # disable suffix, we need to set it manually, it would always add the current projects units

        self.unitschanged(None, None)
    
    def unitschanged(self, sender, e):

        # loop through all objects of self and set the properties for all DistanceEdits
        # the code is slower than doing it manually for each single one
        # but more convenient since we don't have to worry about how many DistanceEdit Controls we have in the UI
        tt = self.__dict__.items()
        for i in self.__dict__.items():
            if i[1].GetType() == TBCWpf.DistanceEdit().GetType():
                i[1].DisplayUnit = LinearType(self.lunits.DisplayType)
                i[1].ShowControlIcon(False)
                i[1].FormatProperty.AddSuffix = ControlBoolean(1)
                #i[1].FormatProperty.NumberOfDecimals = int(self.lfp.NumberOfDecimals)

        
    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()
    
    # code for enabling disabling the pickers can found below the main loop


    # main loop
    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''


        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        
        wv = self.currentProject [Project.FixedSerial.WorldView]
        #wv.PauseGraphicsCache(True)

        ProgressBar.TBC_ProgressBar.Title = self.Caption
        layer_sn = self.layerpicker.SelectedSerialNumber    # we get the source layer serial number



        wlc = self.currentProject[Project.FixedSerial.LayerContainer] # we get all the layers into an object
        wl=wlc[layer_sn]    # we get just the source layer as an object
        members=wl.Members  # we get serial number list of all the elements on that layer
        
        
        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                if members.Count>0:   # if we have at least one elemen
                    j = 0
                    for i in members:   # we go through the serial numbers
                        j += 1
                        if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(j * 100 / members.Count)): # is false if user pressed Cancel
                            break
                        o=self.currentProject.Concordance.Lookup(i)     # we get as the single elements as object using their serial number

                        # we use try to change the attribute since we don't know if the element actually has that kind of attribute
                        # if it does have the attribute it will be changed
                        # if it doesn't, it would raise an error message, but with try/except we just ignore that
                        # we also make sure if we work on a text or not

                        # changing non-Text Objects
                        if self.changelinelayer.IsChecked and not (isinstance(o, CadText) or isinstance(o, MText)):
                            try: o.Layer = self.newlinelayerpicker.SelectedSerialNumber
                            except: pass
                        
                        if self.changelinecolor.IsChecked and not (isinstance(o, CadText) or isinstance(o, MText)):
                            try: o.Color = self.linecolorpicker.SelectedColor
                            except: pass

                        if self.changelineweight.IsChecked and not (isinstance(o, CadText) or isinstance(o, MText)):
                            try: o.Weight = self.lineweightpicker.Lineweight
                            except: pass

                        if self.changelinestyle.IsChecked and not (isinstance(o, CadText) or isinstance(o, MText)):
                            try: o.LineStyle = self.linestylepicker.SelectedSerialNumber
                            except: pass

                        if self.changelinestylescale.IsChecked and not (isinstance(o, CadText) or isinstance(o, MText)):
                            try: o.LineTypeScale = abs(self.linestylescale.Value)
                            except: pass

                        # changing Text Objects
                        if self.changetextlayer.IsChecked and (isinstance(o, CadText) or isinstance(o, MText)):
                            try: o.Layer = self.newtextlayerpicker.SelectedSerialNumber
                            except: pass

                        if self.changetextcolor.IsChecked and (isinstance(o, CadText) or isinstance(o, MText)):
                            try: o.Color = self.textcolorpicker.SelectedColor
                            except: pass

                        if self.changetextstyle.IsChecked and (isinstance(o, CadText) or isinstance(o, MText)):
                            try: o.TextStyleSerial = self.textstylepicker.SelectedSerialNumber
                            except: pass

                        if self.changetextheight.IsChecked and (isinstance(o, CadText) or isinstance(o, MText)):
                            try: o.Height = abs(self.textheight.Distance)
                            except: pass
 
                        if self.changetextweight.IsChecked and (isinstance(o, CadText) or isinstance(o, MText)):
                            try: o.Weight = self.textweightpicker.Lineweight
                            except: pass
                        
                failGuard.Commit()
                self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
                UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
        
        except:
            # EndMark MUST be set no matter what
            # otherwise TBC won't work anymore and needs to be restarted
            self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
            UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
            self.error.Content += '\nan Error occurred - Result probably incomplete'

        ProgressBar.TBC_ProgressBar.Title = ""
        #wv.PauseGraphicsCache(False)
        #self.currentProject.Calculate(False) # otherwise Project Explorer wouldn't show the changed attributes
        self.SaveOptions()           

    
        
    def SetDefaultOptions(self):
        SCROptions.LoadMacroOptions(self, "SCR_ChangePropNoSelect", _OPTIONS, self.currentProject)
        self.linecolorpicker.SelectedColor = Color.FromArgb(OptionsManager.GetInt("SCR_ChangePropNoSelect.linecolorpicker", Color.Red.ToArgb()))
        settingserial = OptionsManager.GetUint("SCR_ChangePropNoSelect.textstylepicker", 24)
        o = self.currentProject.Concordance[settingserial]
        if o is None or not isinstance(o.GetSite(), TextStyleCollection):
            settingserial = 24
        self.textstylepicker.SetSelectedSerialNumber(settingserial, InputMethod(3))
        self.lineweightpicker.Lineweight = OptionsManager.GetInt("SCR_ChangePropNoSelect.lineweightpicker", 0)
        self.textcolorpicker.SelectedColor = Color.FromArgb(OptionsManager.GetInt("SCR_ChangePropNoSelect.textcolorpicker", Color.Red.ToArgb()))
        self.textweightpicker.Lineweight = OptionsManager.GetInt("SCR_ChangePropNoSelect.textweightpicker", 0)

    def SaveOptions(self):
        SCROptions.SaveMacroOptions(self, "SCR_ChangePropNoSelect", _OPTIONS)
        OptionsManager.SetValue("SCR_ChangePropNoSelect.linecolorpicker", self.linecolorpicker.SelectedColor.ToArgb())
        OptionsManager.SetValue("SCR_ChangePropNoSelect.linestylepicker", self.linestylepicker.SelectedSerialNumber)
        OptionsManager.SetValue("SCR_ChangePropNoSelect.lineweightpicker", self.lineweightpicker.Lineweight)
        OptionsManager.SetValue("SCR_ChangePropNoSelect.textstylepicker", self.textstylepicker.SelectedSerialNumber)
        OptionsManager.SetValue("SCR_ChangePropNoSelect.textcolorpicker", self.textcolorpicker.SelectedColor.ToArgb())
        OptionsManager.SetValue("SCR_ChangePropNoSelect.textweightpicker", self.textweightpicker.Lineweight)

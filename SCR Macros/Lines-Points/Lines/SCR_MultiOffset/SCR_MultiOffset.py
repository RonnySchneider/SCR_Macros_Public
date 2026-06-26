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
    "layerpicker": 8, "hzoffset": 0.0, "voffset": 0.0,
    "left": True, "right": False, "both": False,
}

def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_MultiOffset"
    cmdData.CommandName = "SCR_MultiOffset"
    cmdData.Caption = "_SCR_MultiOffset"
    cmdData.UIForm = "SCR_MultiOffset"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "Lines"
        cmdData.ShortCaption = "Multi Offset"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.07
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "offset multiple lines at once"
        cmdData.ToolTipTextFormatted = "offset multiple lines at once"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_MultiOffset(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_MultiOffset.xaml") as s:
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
        self.objs.IsEntityValidCallback = self.IsValid

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        #self.lfp = self.lunits.Properties.Copy()
        linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation
        #self.lfp.AddSuffix = False
        self.hzoffsetlabel.Content = "horizontal Offset [" + linearsuffix + "]"
        self.voffsetlabel.Content = "vertical Offset [" + linearsuffix + "]"
        
		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def SetDefaultOptions(self):
        SCROptions.LoadMacroOptions(self, "SCR_MultiOffset", _OPTIONS, self.currentProject)

    def SaveOptions(self):
        SCROptions.SaveMacroOptions(self, "SCR_MultiOffset", _OPTIONS)

    def IsValid(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        return False

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()


    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content = ''
        self.success.Content = ''

        wv = self.currentProject [Project.FixedSerial.WorldView]
        
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())

        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                voffset = self.voffset.Distance
                if math.isnan(voffset): voffset = 0.0
                    
                for o in self.objs:
                    if isinstance(o, self.lType):

                        
                        polyseg = o.ComputePolySeg().Clone()
                        polyseg = polyseg.ToWorld()
                        polyseg_v = o.ComputeVerticalPolySeg()
                        
                        if polyseg.NumberOfNodes >= 3 and polyseg.IsClosed and polyseg.IsClockWise():
                            polyseg.Reverse()
                            if polyseg_v:
                                polyseg_v.Reverse()

                        polyseg_v_lin = polyseg.Linearize(0.0001, 0.0001, 10000.0, polyseg_v, True)
                        polyseg_v_lin.ConvertPolysegToStationElevation(1.0)

                        if self.left.IsChecked or self.both.IsChecked:
                            l = wv.Add(clr.GetClrType(Linestring))
                            polyseg_ol = polyseg.Offset(Side.Left, abs(self.hzoffset.Distance))[1]
                            polyseg_ol_v = self.verticalatoffset(polyseg_ol, polyseg, polyseg_v_lin)
                            l.Layer = self.layerpicker.SelectedSerialNumber
                            l.Append(polyseg_ol, polyseg_ol_v, False, False)

                        if self.right.IsChecked or self.both.IsChecked:
                            l = wv.Add(clr.GetClrType(Linestring))
                            polyseg_or = polyseg.Offset(Side.Right, abs(self.hzoffset.Distance))[1]
                            polyseg_or_v = self.verticalatoffset(polyseg_or, polyseg, polyseg_v_lin)
                            l.Layer = self.layerpicker.SelectedSerialNumber
                            l.Append(polyseg_or, polyseg_or_v, False, False)

                failGuard.Commit()
                self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
                UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
        
        except Exception as e:
            tt = sys.exc_info()
            exc_type, exc_obj, exc_tb = sys.exc_info()
            # EndMark MUST be set no matter what
            # otherwise TBC won't work anymore and needs to be restarted
            self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
            UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
            self.error.Content += '\nan Error occurred - Result probably incomplete\n' + str(exc_type) + '\n' + str(exc_obj) + '\nLine ' + str(exc_tb.tb_lineno)
        
            
        self.SaveOptions()

    def verticalatoffset(self, polyseg_os, polyseg_hz, polyseg_v):
        return polyseg_os.ComputeElevationOverrideOnOffsetPolyseg(polyseg_hz, polyseg_v)

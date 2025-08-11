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
    cmdData.Key = "SCR_BatchExportRxl"
    cmdData.CommandName = "SCR_BatchExportRxl"
    cmdData.Caption = "_SCR_BatchExportRxl"
    cmdData.UIForm = "SCR_BatchExportRxl"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "Import/Export"
        cmdData.ShortCaption = "Export Rxl"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.02
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Batch export Alignments to RXL"
        cmdData.ToolTipTextFormatted = "Batch export Alignments to RXL"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass

class SCR_BatchExportRxl(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_BatchExportRxl.xaml") as s:
            wpf.LoadComponent (self, s)
        self.currentProject = currentProject
        self.macroFileFolder = macroFileFolder
        if self.openfilename.Text=='': self.openfilename.Text = macroFileFolder

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

        self.profileType = clr.GetClrType(ProfileView)

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def IsValid(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.profileType):
            return True
        return False

    def SetDefaultOptions(self):
        self.openfilename.Text = OptionsManager.GetString("SCR_BatchExportRxl.openfilename", os.path.expanduser('~\\Downloads'))

    def SaveOptions(self):
        OptionsManager.SetValue("SCR_BatchExportRxl.openfilename", self.openfilename.Text)

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def openbutton_Click(self, sender, e):
        dialog=FolderBrowserDialog()
        dialog.Reset()
        dialog.RootFolder = Environment.SpecialFolder.Desktop
        dialog.SelectedPath = self.openfilename.Text
        tt=dialog.ShowDialog()
        if tt==DialogResult.OK:
            self.openfilename.Text = dialog.SelectedPath

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''



        wv = self.currentProject [Project.FixedSerial.WorldView]
        self.success.Content=''
        
        if self.objs.SelectedMembers(self.currentProject).Count>0:

            for o in self.objs.SelectedMembers(self.currentProject): # go through the selection

                if isinstance(o, ProfileView): # double check if it is an alignment
                    
                    alignmentname = str.upper(o.HorizontalAlignment.Name) # get the alignment name
                    filename = self.openfilename.Text + "\\" + alignmentname + ".rxl" # put the filename together
                    hal=o.HorizontalAlignment   # get the horizontal alignment as object
                    
                    i=0
                    for o2 in o:    # an alignment can contain multiple vertical alignments 
                        i+=1        # this approach will currently only work properly if it's just one
                        val=o2      # in the end we'll always just get the last VAL, hence it might be the wrong one
                    
                    # if we've found multiple VAL we at least show an error message, and skip it
                    if i>1:
                        self.success.Content+='multiple VALs in ' + alignmentname + ' -> was skipped\n'
                        continue

                    # Write(string fileName, string roadName, double stationInterval, Trimble.Vce.Alignment.HorizontalAlignment hal, Trimble.Vce.Alignment.VerticalAlignment val)
                    interval=1.000 # the interval doesn't really change the RXL, only how often labels are shown in Access
                    owriter=RxlFileWriter(RxlVersions.SixDotFour) # we create the output function
                    
                    if i==0:
                        owriter.Write(filename, alignmentname, interval, hal) # in case we don't have a VAL
                    else:
                        owriter.Write(filename, alignmentname, interval, hal, val) # Alignment with HAL and VAL

            self.success.Content+='Export done'
        else:
            self.success.Content+='nothing selected'

        self.SaveOptions()

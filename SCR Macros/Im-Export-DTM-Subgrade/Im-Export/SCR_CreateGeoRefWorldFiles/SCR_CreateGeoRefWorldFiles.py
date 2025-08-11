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
    cmdData.Key = "SCR_CreateGeoRefWorldFiles"
    cmdData.CommandName = "SCR_CreateGeoRefWorldFiles"
    cmdData.Caption = "_SCR_CreateGeoRefWorldFiles"
    cmdData.UIForm = "SCR_CreateGeoRefWorldFiles"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "Import/Export"
        cmdData.ShortCaption = "Write GeoRef Worldfiles"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.07
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Write GeoRef Worldfiles for multiple images at once"
        cmdData.ToolTipTextFormatted = "Write GeoRef Worldfiles for multiple images at once"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png") 
        cmdData.ImageSmall = b
    except:
        pass

class SCR_CreateGeoRefWorldFiles(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_CreateGeoRefWorldFiles.xaml") as s:
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

        self.objs.IsEntityValidCallback = self.IsValidImage

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def SetDefaultOptions(self):
        self.copytofolder.IsChecked = OptionsManager.GetBool("SCR_CreateGeoRefWorldFiles.copytofolder", False)
        self.renamefile.IsChecked = OptionsManager.GetBool("SCR_CreateGeoRefWorldFiles.renamefile", False)
        self.targetfolder.Text = OptionsManager.GetString("SCR_CreateGeoRefWorldFiles.targetfolder", "C:\\")

    def SaveOptions(self):
        OptionsManager.SetValue("SCR_CreateGeoRefWorldFiles.copytofolder", self.copytofolder.IsChecked)
        OptionsManager.SetValue("SCR_CreateGeoRefWorldFiles.renamefile", self.renamefile.IsChecked)
        OptionsManager.SetValue("SCR_CreateGeoRefWorldFiles.targetfolder", self.targetfolder.Text)

    def IsValidImage(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, GeoReferencedImage):
            return True
        return False

    def openbutton_Click(self, sender, e):
        dialog = FolderBrowserDialog()
        dialog.Reset()
        dialog.RootFolder = Environment.SpecialFolder.Desktop
        dialog.SelectedPath = self.targetfolder.Text
        tt = dialog.ShowDialog()
        if tt==DialogResult.OK:
            self.targetfolder.Text = dialog.SelectedPath

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content = ''
        self.success.Text = ''

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear

        wff = WorldFileFormat()

        for o in self.objs:
            if isinstance(o, GeoReferencedImage) or isinstance(o, GeoReferencedImage3D):  # Trimble.Vce.Data.RawData.GeoReferencedImage
                fullimagefn = o.Filename
                imagefiletype = os.path.splitext(os.path.basename(fullimagefn))[1]
                
                # there is a bug in CreateWorldFile if the file suffix is tiff with double f
                if str.upper(imagefiletype) == ".TIFF":
                    fullimagefnfixed = fullimagefn[:-1]
                else:
                    fullimagefnfixed = fullimagefn

                ul = self.toprojectunit(o.UpperLeftGrid)
                ur = self.toprojectunit(o.UpperRightGrid)
                ll = self.toprojectunit(o.LowerLeftGrid)
                lr = self.toprojectunit(o.LowerRightGrid)

                tiepoints = List[TiePoint]()

                tiepoints.Add(TiePoint(0, 0, ul.X, ul.Y))
                tiepoints.Add(TiePoint(o.WidthPixels, 0, ur.X, ur.Y))
                tiepoints.Add(TiePoint(0, o.HeightPixels, ll.X, ll.Y))
                tiepoints.Add(TiePoint(o.WidthPixels, o.HeightPixels, lr.X, lr.Y))

                fullwrldfn = wff.CreateWorldFile(fullimagefnfixed, tiepoints)   # needs 3 or more points, that's why I use the 4 corner points
                if fullwrldfn == "":
                    self.error.Content += "can't handle image format"
                else:
                    wrldfiletype = os.path.splitext(os.path.basename(fullwrldfn))[1]

                    self.success.Text += 'created\n' + fullwrldfn + '\n\n'
                    
                    if self.copytofolder.IsChecked and not self.renamefile.IsChecked:
                        shutil.copy2(fullimagefn, self.targetfolder.Text)
                        shutil.copy2(fullwrldfn, self.targetfolder.Text)

                        self.success.Text += 'and copied to\n' + self.targetfolder.Text + '\n\n'
                    
                    if self.copytofolder.IsChecked and self.renamefile.IsChecked:
                        shutil.copy2(fullimagefn, os.path.join(self.targetfolder.Text, o.Name + imagefiletype))
                        shutil.copy2(fullwrldfn, os.path.join(self.targetfolder.Text, o.Name + wrldfiletype))

                        self.success.Text += 'and copied as\n' + o.Name + ' to ' + self.targetfolder.Text + '\n\n'



        self.SaveOptions()

    def toprojectunit(self, p):

        p.X = self.lunits.Convert(self.lunits.InternalType, p.X, self.lunits.DisplayType)
        p.Y = self.lunits.Convert(self.lunits.InternalType, p.Y, self.lunits.DisplayType)
        p.Z = self.lunits.Convert(self.lunits.InternalType, p.Z, self.lunits.DisplayType)

        return p
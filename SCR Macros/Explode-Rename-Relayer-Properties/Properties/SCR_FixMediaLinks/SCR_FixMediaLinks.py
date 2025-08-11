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
    cmdData.Key = "SCR_FixMediaLinks"
    cmdData.CommandName = "SCR_FixMediaLinks"
    cmdData.Caption = "_SCR_FixMediaLinks"
    cmdData.UIForm = "SCR_FixMediaLinks"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Expld-SNR-Relay-Prop"
        cmdData.DefaultTabGroupKey = "Properties"
        cmdData.ShortCaption = "Fix Media Links"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.04
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Fix Media Links"
        cmdData.ToolTipTextFormatted = "try to find Media and set absolute path's"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_FixMediaLinks(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_FixMediaLinks.xaml") as s:
            wpf.LoadComponent (self, s)
        self.currentProject = currentProject
        self.macroFileFolder = macroFileFolder
        if self.openfilename.Text=='': self.openfilename.Text = macroFileFolder
        if self.targetfolder.Text=='': self.targetfolder.Text = macroFileFolder

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

    def IsValid(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        try:
            if o.FileProperties.Count > 0:
                return True
        except:
            pass
        return False

    def SetDefaultOptions(self):
        self.useprojectfolder.IsChecked = OptionsManager.GetBool("SCR_FixMediaLinks.useprojectfolder", True)
        self.openfilename.Text = OptionsManager.GetString("SCR_FixMediaLinks.openfilename", "Z:\\2709 - Day Jobs")
        self.fixspecificmedia.IsChecked = OptionsManager.GetBool("SCR_FixMediaLinks.fixspecificmedia", False)
        self.copy.IsChecked = OptionsManager.GetBool("SCR_FixMediaLinks.copy", False)
        self.makerelativatnewlocation.IsChecked = OptionsManager.GetBool("SCR_FixMediaLinks.makerelativatnewlocation", False)
        self.targetfolder.Text = OptionsManager.GetString("SCR_FixMediaLinks.targetfolder", "Z:\\")

    def SaveOptions(self):
        OptionsManager.SetValue("SCR_FixMediaLinks.useprojectfolder", self.useprojectfolder.IsChecked)
        OptionsManager.SetValue("SCR_FixMediaLinks.openfilename", self.openfilename.Text)
        OptionsManager.SetValue("SCR_FixMediaLinks.fixspecificmedia", self.fixspecificmedia.IsChecked)
        OptionsManager.SetValue("SCR_FixMediaLinks.copy", self.copy.IsChecked)
        OptionsManager.SetValue("SCR_FixMediaLinks.makerelativatnewlocation", self.makerelativatnewlocation.IsChecked)
        OptionsManager.SetValue("SCR_FixMediaLinks.targetfolder", self.targetfolder.Text)
    
    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()
    
    def useprojectfolderChanged(self, sender, e):
        if self.useprojectfolder.IsChecked:
            self.openbutton.IsEnabled = False
            self.openfilename.IsEnabled = False
        else:
            self.openbutton.IsEnabled = True
            self.openfilename.IsEnabled = True


    def openbutton_Click(self, sender, e):
        dialog = FolderBrowserDialog()
        dialog.Reset()
        dialog.RootFolder = Environment.SpecialFolder.Desktop
        dialog.SelectedPath = self.openfilename.Text
        tt = dialog.ShowDialog()
        if tt==DialogResult.OK:
            self.openfilename.Text = dialog.SelectedPath

    def openbutton2_Click(self, sender, e):
        dialog = FolderBrowserDialog()
        dialog.Reset()
        dialog.RootFolder = Environment.SpecialFolder.Desktop
        dialog.SelectedPath = self.targetfolder.Text
        tt = dialog.ShowDialog()
        if tt==DialogResult.OK:
            self.targetfolder.Text = dialog.SelectedPath

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''



        wv = self.currentProject [Project.FixedSerial.WorldView]
        self.success.Content=''
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)

        if self.useprojectfolder.IsChecked:
            searchpath = os.path.dirname(self.currentProject.ProjectName)
        else:
            searchpath = self.openfilename.Text

        ProgressBar.TBC_ProgressBar.Title = "searching directory structure" # set the progressbar description
        
        # get a recursive list of all files within the search folder, list contains path and file
        allfiles = [os.path.join(root, name)
                    for root, dirs, files in os.walk(searchpath)
                    for name in files]
                    #if name.endswith((".*"))]
        
        replacedcount = 0
        tobereplacedcount = 0

        try:

            # find the ProjectExplorer-MediaFilesFolder as object
            fpc = FilePropertiesContainer.ProvideFilePropertiesContainer(self.currentProject)
            
            if self.fixspecificmedia.IsChecked:
                for o in self.objs.SelectedMembers(self.currentProject):
                    if o.FileProperties.Count > 0:
                        tobereplacedcount += 1
            else:
                for o in fpc:
                    if isinstance(o, FileProperties):
                        tobereplacedcount += 1

            if self.copy.IsChecked and self.makerelativatnewlocation.IsChecked:
                try:
                    os.mkdir(os.path.join(self.targetfolder.Text, "Photos"))
                except:
                    pass

            if self.fixspecificmedia.IsChecked:
                for o in self.objs.SelectedMembers(self.currentProject): # go through the selection
                    try:
                        if o.FileProperties.Count > 0:  # if the object has a file property
                            for fp in o:    # loop through them
                                filename = os.path.basename(fp.FilePath) # get just the filename

                                for singlefile in allfiles: # try to find it in the file list
                                    # loop through the file list and see if any of them contains the filename string
                                    if filename in singlefile: 
                                        
                                        if self.copy.IsChecked: # if we want to copy first
                                            if self.makerelativatnewlocation.IsChecked:
                                                shutil.copy2(singlefile, os.path.join(self.targetfolder.Text, "Photos")) # copy
                                                fp.FilePath = "..\\Photos\\" + filename # and set link to this location
                                            else:
                                                shutil.copy2(singlefile, self.targetfolder.Text) # copy
                                                fp.FullName = os.path.join(self.targetfolder.Text, filename) # and set link to this location
                                        else:
                                            fp.FullName = singlefile # if yes we replace the media path with that string
                                        replacedcount += 1
                                        



                                        ProgressBar.TBC_ProgressBar.Title = "fixed Media Files " + str(replacedcount) + "/" + str(tobereplacedcount) # set the progressbar description
                                        if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(replacedcount * 100 / tobereplacedcount)):
                                            break   # function returns true if user pressed cancel
                                        
                                        break

                    except:
                        pass
            else:
                for o in fpc: # go through the whole file properties container
                    if isinstance(o, FileProperties):
                                      
                        filename = os.path.basename(o.FilePath)

                        for singlefile in allfiles:
                            if filename in singlefile:
                                if self.copy.IsChecked: # if we want to copy first
                                    if self.makerelativatnewlocation.IsChecked:
                                        shutil.copy2(singlefile, os.path.join(self.targetfolder.Text, "Photos")) # copy
                                        o.FullName = "..\\Photos\\" + filename # and set link to this location
                                    else:
                                        shutil.copy2(singlefile, self.targetfolder.Text) # copy
                                        o.FullName = os.path.join(self.targetfolder.Text, filename) # and set link to this location

                                else:
                                    o.FullName = singlefile
                                replacedcount += 1

                                ProgressBar.TBC_ProgressBar.Title = "fixed Media Files " + str(replacedcount) + "/" + str(tobereplacedcount) # set the progressbar description
                                if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(replacedcount * 100 / tobereplacedcount)):
                                    break   # function returns true if user pressed cancel
                                
                                break
            
        
        finally:

            ProgressBar.TBC_ProgressBar.Title = ""
            self.success.Content +='Replaced the Link in ' + str(replacedcount) + ' of ' + str(tobereplacedcount) + ' Files'
            self.success.Content +='\nDone'

            self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
            UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())

        self.SaveOptions()

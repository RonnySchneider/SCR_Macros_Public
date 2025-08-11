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
    cmdData.Key = "SCR_BatchRenameAlignments"
    cmdData.CommandName = "SCR_BatchRenameAlignments"
    cmdData.Caption = "_SCR_BatchRenameAlignments"
    cmdData.UIForm = "SCR_BatchRenameAlignments"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Expld-SNR-Relay-Prop"
        cmdData.DefaultTabGroupKey = "Renaming"
        cmdData.ShortCaption = "Rename Alignments"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.02
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Batch rename Alignments"
        cmdData.ToolTipTextFormatted = "Batch rename Alignments based on their name and a List"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + "_i1.png") 
        cmdData.ImageSmall = b
    except:
        pass


class SCR_BatchRenameAlignments(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_BatchRenameAlignments.xaml") as s:
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
		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    
    def SetDefaultOptions(self):
        self.openfilename.Text = OptionsManager.GetString("SCR_BatchRenameAlignments.openfilename", os.path.expanduser('~\\Downloads'))

    def SaveOptions(self):
        OptionsManager.SetValue("SCR_BatchRenameAlignments.openfilename", self.openfilename.Text)
        

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def savebutton_Click(self, sender, e):
        wv = self.currentProject [Project.FixedSerial.WorldView]

        filename = os.path.expanduser('~/Downloads/TBC-output-Alignment-names.csv')
        if File.Exists(filename):
                File.Delete(filename)
        with open(filename, 'w') as f:            
            for o in wv: # go through everything in the project
                if isinstance(o, ProfileView): # in case it is an alignment write the name into the file
                    f.write(o.HorizontalAlignment.Name + "\n")
            f.close()
  
    def openbutton_Click(self, sender, e):

        os.path.dirname(self.openfilename.Text)


        dialog=OpenFileDialog()
        dialog.InitialDirectory = os.path.dirname(self.openfilename.Text)
        dialog.Filter=("CSV|*.csv")
        
        tt=dialog.ShowDialog()
        if tt==DialogResult.OK:
            self.openfilename.Text = dialog.FileName



    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''


        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())

        wv = self.currentProject [Project.FixedSerial.WorldView]
        
        #tt = self.openfilename.Text


        try:
            # the "with" statement will unroll any changes if something go wrong
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                if File.Exists(self.openfilename.Text): # if file still exists
                    
                    namelist=[]

                    with open(self.openfilename.Text,'r') as csvfile: 
                        reader = csv.reader(csvfile, delimiter=',', quotechar='|') 
                        for row in reader:
                            namelist.Add(row)
                    #tt = namelist.Count
                    

                    for o in wv: # go through everything in the project

                        if isinstance(o, ProfileView): # in case it is an alignment rename it based on the list
                            
                            for i in range(0, namelist.Count): # go through the whole name list and try to find a match
                                if namelist[i].Count>0:
                                    searchstring = str.upper(namelist[i][0])          # use Upper case in case of typos
                                    alignmentstring = str.upper(o.HorizontalAlignment.Name)
                                    
                                    if alignmentstring.find(searchstring) > -1:   # check if the Alignment name contains any of the values from the list
                                        if namelist[i][1] != '':                 # if yes, and the replacement string is not '' then replace it with the second name from the list
                                            o.HorizontalAlignment.Name = namelist[i][1]   
                                            for o2 in o:            # go through the Labelling and vertical alignments
                                                try:                # use try, on some objects you cant set the layer or name - would throw an error
                                                                    # which would get caught by the with-failguard loop, and no changes at all would commit
                                                    o2.Name = namelist[i][1]       # try changing the name
                                                    o2.Layer = o.Layer    # try changing the layer
                                                except:
                                                    pass
                    
                failGuard.Commit()
                UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
                self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
        
        except:
            # EndMark MUST be set no matter what
            # otherwise TBC won't work anymore and needs to be restarted
            self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
            UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
            self.error.Content += '\nan Error occurred - Result probably incomplete'

        self.SaveOptions()


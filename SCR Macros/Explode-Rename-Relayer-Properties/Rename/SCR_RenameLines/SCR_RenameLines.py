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
    cmdData.Key = "SCR_RenameLines"
    cmdData.CommandName = "SCR_RenameLines"
    cmdData.Caption = "_SCR_RenameLines"
    cmdData.UIForm = "SCR_RenameLines"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Expld-SNR-Relay-Prop"
        cmdData.DefaultTabGroupKey = "Renaming"
        cmdData.ShortCaption = "Rename Lines"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.03  # included metre to project conversion
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Rename Lines"
        cmdData.ToolTipTextFormatted = "Rename Lines"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_RenameLines(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_RenameLines.xaml") as s:
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

        self.polyType = clr.GetClrType(IPolyseg)

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
        self.globalselection.IsChecked = OptionsManager.GetBool("SCR_RenameLines.globalselection", True)
        self.selectmanually.IsChecked = OptionsManager.GetBool("SCR_RenameLines.selectmanually", False)

        self.linenameelev.IsChecked = OptionsManager.GetBool("SCR_RenameLines.linenameelev", True)
        self.linenamelayername.IsChecked = OptionsManager.GetBool("SCR_RenameLines.linenamelayername", False)
        self.linenamemanual.IsChecked = OptionsManager.GetBool("SCR_RenameLines.linenamemanual", False)

        self.newlinename.Text = OptionsManager.GetString("SCR_RenameLines.newlinename", "")

    def SaveOptions(self):
        OptionsManager.SetValue("SCR_RenameLines.globalselection", self.globalselection.IsChecked)
        OptionsManager.SetValue("SCR_RenameLines.selectmanually", self.selectmanually.IsChecked)

        OptionsManager.SetValue("SCR_RenameLines.linenameelev", self.linenameelev.IsChecked)
        OptionsManager.SetValue("SCR_RenameLines.linenamelayername", self.linenamelayername.IsChecked)
        OptionsManager.SetValue("SCR_RenameLines.linenamemanual", self.linenamemanual.IsChecked)

        OptionsManager.SetValue("SCR_RenameLines.newlinename", self.newlinename.Text)

    def IsValid(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.polyType):
            return True
        return False

        
    def CancelClicked(self, cmd, args):

        cmd.CloseUICommand ()

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''


        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())

        wv = self.currentProject [Project.FixedSerial.WorldView]
        #bc = self.currentProject [Project.FixedSerial.BlockCollection]    # getting all blocks as collection
        
        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        # we don't want the units to be included (so we make copy and turn that off). Otherwise get something like "12.50 ft"
        self.lfp = self.lunits.Properties.Copy()
        self.lfp.AddSuffix = False

        objectserialnumbers = DynArray()

        # different ways to compile a list with the drawing object serial numbers
        if self.globalselection.IsChecked:
            wlc = self.currentProject[Project.FixedSerial.LayerContainer] # we get all the layers into an object
            for i in wlc:   # we go through all the layers
                wl=wlc[i.SerialNumber]    # we get just the source layer as an object
                layermembers=wl.Members  # we get serial number list of all the elements on that layer
                for i2 in layermembers:
                    objectserialnumbers.Add(i2) # we compile a list of all drawing objects

        if self.selectmanually.IsChecked:
            for o in self.objs.SelectedMembers(self.currentProject):
                objectserialnumbers.Add(o.SerialNumber)

        try:
            # the "with" statement will unroll any changes if something go wrong
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                # go through the serial numbers and change the text if possible                
                if objectserialnumbers.Count>0:
                    for i in objectserialnumbers:
                        o=self.currentProject.Concordance.Lookup(i)
                        if isinstance(o, self.polyType):
                            if self.linenameelev.IsChecked:
                                try:
                                    tt = CommandHelper.GetMinMaxElevations(o)
                                except: pass
                                try:
                                    minel = self.lunits.Format(tt[1], self.lfp)
                                    maxel = self.lunits.Format(tt[2], self.lfp)
                                    if minel==maxel:
                                        o.Name = 'El: ' + minel
                                    else:
                                        o.Name = 'Min: ' + minel + ' Max: ' + maxel
                                except:
                                    try:
                                        o.Name = 'El: ' + o.Elevation
                                    except: pass
                                continue
                            if self.linenamelayername.IsChecked:
                                try:
                                    o.Name = self.currentProject.Concordance.Lookup(o.Layer).Name
                                except: pass
                                continue
                            if self.linenamemanual.IsChecked:
                                try:
                                    o.Name = self.newlinename.Text
                                except: pass
                                continue                                  
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

        
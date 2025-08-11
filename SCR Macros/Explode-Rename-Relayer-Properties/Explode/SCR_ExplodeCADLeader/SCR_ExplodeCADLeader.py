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
    cmdData.Key = "SCR_ExplodeCADLeader"
    cmdData.CommandName = "SCR_ExplodeCADLeader"
    cmdData.Caption = "_SCR_ExplodeCADLeader"
    cmdData.UIForm = "SCR_ExplodeCADLeader"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Expld-SNR-Relay-Prop"
        cmdData.DefaultTabGroupKey = "Explode"
        cmdData.ShortCaption = "Explode CAD-Leader"
        cmdData.DefaultRibbonToolSize = 0 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.12
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Explode CAD Leader to simple Polyline"
        cmdData.ToolTipTextFormatted = "Explode CAD Leader to simple Polyline"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_ExplodeCADLeader(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_ExplodeCADLeader.xaml") as s:
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

        self.lType = clr.GetClrType(IPolyseg)
        
        self.objs.IsEntityValidCallback = self.IsValid
        optionMenu = SelectionContextMenuHandler()
        # remove options that don't apply here
        optionMenu.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        self.objs.ButtonContextMenu = optionMenu

    def IsValid(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, Leader):
            return True
        return False


        
    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def Coord1Changed(self, ctrl, e):
        # set keyboard focus if change was due to mouse pick
        if e.Cause == InputMethod.Mouse:     
            self.OkClicked(None, None)


    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''


        # self.label_benchmark.Content = ''

        # start_t = timer ()
        wv = self.currentProject [Project.FixedSerial.WorldView]
        bc = self.currentProject [Project.FixedSerial.BlockCollection]    # getting all blocks as collection
        lsc = self.currentProject [Project.FixedSerial.LabelStyleContainer]    # getting all blocks as collection

        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)

        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                ProgressBar.TBC_ProgressBar.Title = "exploding CAD-Leaders"

                # get the current list so we can clear global selection, otherwise it will update the selection each time we explode and that is very slow
                entities = Array[ISnapIn](self.objs.SelectedMembers(self.currentProject))
                GlobalSelection.Clear()
                j = 0
                for o in entities:
                    if isinstance(o, Leader):

                        j += 1
                        if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(j * 100 / entities.Count)):
                            break   # function returns true if user pressed cancel

                        leaderpolyline = o.ComputePolySeg()
                        l = wv.Add(clr.GetClrType(Linestring))
                        l.Layer = o.Layer
                        l.Color = o.Color
                        for i in range(0, leaderpolyline.NumberOfNodes):
                            drawpoint = leaderpolyline[i].Point
                            e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                            e.Position = drawpoint
                            l.AppendElement(e)

                        osite = o.GetSite()    # we find out in which container the serial number reside
                        osite.Remove(o.SerialNumber)   # we delete the object from that container

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


        ProgressBar.TBC_ProgressBar.Title = ""
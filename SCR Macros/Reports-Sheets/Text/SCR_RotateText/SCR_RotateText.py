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
    cmdData.Key = "SCR_RotateText"
    cmdData.CommandName = "SCR_RotateText"
    cmdData.Caption = "_SCR_RotateText"
    cmdData.UIForm = "SCR_RotateText"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Reports"
        cmdData.DefaultTabGroupKey = "Text"
        cmdData.ShortCaption = "Rotate Text"
        cmdData.DefaultRibbonToolSize = 0 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.13
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Rotate Text"
        cmdData.ToolTipTextFormatted = "Rotate Text"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_RotateText(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_RotateText.xaml") as s:
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
        self.objs.IsEntityValidCallback=self.IsValid
        optionMenu = SelectionContextMenuHandler()
        # remove options that don't apply here
        optionMenu.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        self.objs.ButtonContextMenu = optionMenu
        self.mtextType = clr.GetClrType(MText)
        self.cadtextType = clr.GetClrType(CadText)

        self.coordpick1.PointSelected += self.coordpick1mouseclick

        self.coordpick2.ShowGdiCursor += self.coordpick2changed
        self.coordpick2.PointSelected += self.coordpick2mouseclick
        
        self.coordpick2.AutoTab = False

        if self.objs.Count > 0:
            def SetFocusToControl():
                Keyboard.Focus(self.coordpick1)
            Dispatcher.BeginInvoke(Dispatcher.CurrentDispatcher, Action(SetFocusToControl))

            #Keyboard.Focus(self.coordpick1) # seems to unreliable on OnLoad

    def coordpick1mouseclick(self, ctrl, e):
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        #UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())

    def coordpick2mouseclick(self, ctrl, e):
        Keyboard.Focus(self.coordpick1)

        self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
        #UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())


    def coordpick2changed(self, sender, e):

        self.success.Content = ""
        self.error.Content = ""

        wv = self.currentProject [Project.FixedSerial.WorldView]

        #self.coordpick2.CursorStyle = CursorStyle.CrossHair | CursorStyle.RubberLine
        #if self.coordpick1.ResultCoordinateSystem:
        #    self.coordpick2.AnchorPoint = MousePosition(self.coordpick1.ClickWindow, self.coordpick1.Coordinate, self.coordpick1.ResultCoordinateSystem)
        #else:
        #    self.coordpick2.AnchorPoint = None

        p1 = self.coordpick1.Coordinate
        p1.Z = 0
        
        if not isinstance(e.MessagingView, clr.GetClrType(I2DProjection)):
            return

        worldmousexy = e.MessagingView.ConvertPageToProject(Point3D(e.MousePosition.X, e.MousePosition.Y, 0)) # delivers only a 2D Point
        worldmousexy.Z = 0 # need to set it at least to Zero, otherwise the td matrix will be wrong
        
        td = TransformData(Matrix4D(Vector3D(p1, worldmousexy)), Matrix4D(Vector3D.Zero))
        
        # committing here every time when we move the mouse would be too slow
        # we set the start and end mark in coordpick1mouseclick further up
        for o in self.objs:
            try:
                o.Transform(td)
            except:
                pass
                
        self.coordpick1.Coordinate = worldmousexy


    def IsValid(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.mtextType) or isinstance(o, self.cadtextType) or isinstance(o, CadLabel):
            return True
        return False
        
    def ccw45_Click(self, sender, e):
        self.rotate_text(math.pi/4)
        Keyboard.Focus(self.coordpick1)

    def cw45_Click(self, sender, e):
        self.rotate_text(-math.pi/4)
        Keyboard.Focus(self.coordpick1)

    def flip180_Click(self, sender, e):
        self.rotate_text(math.pi)
        Keyboard.Focus(self.coordpick1)

    def changetotl(self, sender, e):
        self.changejustification(1, 6)

    def changetotc(self, sender, e):
        self.changejustification(2, 7)

    def changetotr(self, sender, e):
        self.changejustification(3, 8)

    def changetoml(self, sender, e):
        self.changejustification(4, 3)

    def changetomc(self, sender, e):
        self.changejustification(5, 4)

    def changetomr(self, sender, e):
        self.changejustification(6, 5)

    def changetobl(self, sender, e):
        self.changejustification(7, 0)

    def changetobc(self, sender, e):
        self.changejustification(8, 1)

    def changetobr(self, sender, e):
        self.changejustification(9, 2)

    
    # who is inventing this kind of crap
    # for old style text you have to use TextStyle.TextJustification
    # 6 7 8
    # 3 4 5
    # 0 1 2
    # for MText you have to use Trimble.Vce.ForeignCad.AttachmentPoint
    # 1 2 3
    # 4 5 6
    # 7 8 9
        
    def changejustification(self, justMText, justCadText):
        self.success.Content = ""
        self.error.Content = ""

        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        
        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                
                for o in self.objs:
                    if isinstance(o, self.mtextType):
                        o.AttachPoint = AttachmentPoint(justMText)
                    if isinstance(o, self.cadtextType):
                        o.Alignment = TextStyle.TextJustification(justCadText)
                    if isinstance(o, CadLabel):
                        self.error.Content += "\nLabels don't have a Justification Setting"
                        self.error.Content += "\nyou'd need to change the Textstyle itself"

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
        
        Keyboard.Focus(self.coordpick1)

    def rotate_text(self, rotation):
        self.success.Content = ""
        self.error.Content = ""

        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        
        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                
                for o in self.objs:
                    if isinstance(o, self.mtextType) or isinstance(o, CadLabel):
                        o.Rotation += rotation
                    if isinstance(o, self.cadtextType):
                        o.RotateAngle += rotation

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

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.success.Content = ""
        self.error.Content = ""



        pass


            

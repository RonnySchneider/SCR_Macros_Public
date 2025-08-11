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
    cmdData.Key = "SCR_CreateRectangleCLs"
    cmdData.CommandName = "SCR_CreateRectangleCLs"
    cmdData.Caption = "_SCR_CreateRectangleCLs"
    cmdData.UIForm = "SCR_CreateRectangleCLs"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "Lines"
        cmdData.ShortCaption = "Rectangle CLs"
        cmdData.DefaultRibbonToolSize = 0 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.03
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Create Center Lines in 4 side rectangles"
        cmdData.ToolTipTextFormatted = "Create Center Lines in 4 side rectangles"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_CreateRectangleCLs(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_CreateRectangleCLs.xaml") as s:
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
        self.objs.IsEntityValidCallback=self.IsValid
        optionMenu = SelectionContextMenuHandler()
        # remove options that don't apply here
        optionMenu.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        self.objs.ButtonContextMenu = optionMenu
        
        self.lType = clr.GetClrType(IPolyseg)

    def IsValid(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        return False

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()


    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''


        # self.label_benchmark.Content = ''

        # start_t = timer ()
        wv = self.currentProject [Project.FixedSerial.WorldView]
        pts = Point3DArray()
        outPointOnCL = clr.StrongBox[Point3D]()
        outstation = clr.StrongBox[float]()

        for o in self.objs:
            if isinstance(o, self.lType):
                segline = o.ComputePolySeg()
                segline_v = o.ComputeVerticalPolySeg()

                if segline.NumberOfNodes == 5 and segline.IsClosed:
                    pts = segline.ToPoint3DArray()
                    
                    # create the CLs
                    l = wv.Add(clr.GetClrType(Linestring))
                    l.Layer = o.Layer
                    e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                    mid = Point3D.MidPoint(pts[0], pts[1])
                    if segline_v != None:
                        if segline.FindPointFromPoint(mid, outPointOnCL, outstation):
                            mid.Z = segline_v.ComputeVerticalSlopeAndGrade(outstation.Value)[1]
                    e.Position = mid
                    l.AppendElement(e)       

                    e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                    mid = Point3D.MidPoint(pts[2], pts[3])
                    if segline_v != None:
                        if segline.FindPointFromPoint(mid, outPointOnCL, outstation):
                            mid.Z = segline_v.ComputeVerticalSlopeAndGrade(outstation.Value)[1]
                    e.Position = mid
                    l.AppendElement(e)       

                    l = wv.Add(clr.GetClrType(Linestring))
                    l.Layer = o.Layer
                    e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                    mid = Point3D.MidPoint(pts[1], pts[2])
                    if segline_v != None:
                        if segline.FindPointFromPoint(mid, outPointOnCL, outstation):
                            mid.Z = segline_v.ComputeVerticalSlopeAndGrade(outstation.Value)[1]
                    e.Position = mid
                    l.AppendElement(e)       
                    e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                    mid = Point3D.MidPoint(pts[3], pts[4])
                    if segline_v != None:
                        if segline.FindPointFromPoint(mid, outPointOnCL, outstation):
                            mid.Z = segline_v.ComputeVerticalSlopeAndGrade(outstation.Value)[1]
                    e.Position = mid
                    l.AppendElement(e)       

        
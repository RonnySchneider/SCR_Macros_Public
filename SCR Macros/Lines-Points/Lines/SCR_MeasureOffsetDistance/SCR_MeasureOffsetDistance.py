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
    cmdData.Key = "SCR_MeasureOffsetDistance"
    cmdData.CommandName = "SCR_MeasureOffsetDistance"
    cmdData.Caption = "_SCR_MeasureOffsetDistance"
    cmdData.UIForm = "SCR_MeasureOffsetDistance"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "Lines"
        cmdData.ShortCaption = "Measure Distance/Offset"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.12
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Measure Distance/Offset"
        cmdData.ToolTipTextFormatted = "Measure Distance/Offset"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png") 
        cmdData.ImageSmall = b
    except:
        pass


class SCR_MeasureOffsetDistance(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_MeasureOffsetDistance.xaml") as s:
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
        #self.coordCtl1.ValueChanged += self.Coord1Changed
        #self.coordCtl2.ValueChanged += self.Coord2Changed
        self.coordCtl3.ValueChanged += self.Coord3Changed
        self.coordCtl3.AutoTab = False
        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        # we don't want the units to be included (so we make copy and turn that off). Otherwise get something like "12.50 ft"
        self.lfp = self.lunits.Properties.Copy()
        self.lfp.AddSuffix = False
        linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation
        self.textdistancelabel.Content = "Distance along [" + linearsuffix + "]"
        self.textoffsetlabel.Content = "Offset [" + linearsuffix + "]"
        self.textdhlabel.Content = "vertical dh to Line [" + linearsuffix + "]"

        self.coordCtl1.ValueChanged += self.Coord1Changed

        self.textdecimals.MinValue = 0 # otherwise the volume computation fails
        self.textdecimals.NumberOfDecimals = 0


		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    # def Coord1Changed(self, ctrl, e):
    #     self.coordCtl2.CursorStyle = CursorStyle.CrossHair | CursorStyle.RubberLine
    #     if self.coordCtl1.ResultCoordinateSystem:
    #         self.coordCtl2.AnchorPoint = MousePosition(self.coordCtl1.ClickWindow, self.coordCtl1.Coordinate, self.coordCtl1.ResultCoordinateSystem)
    #     else:
    #         self.coordCtl2.AnchorPoint = None
    # 
    # def Coord2Changed(self, ctrl, e):
    #     self.coordCtl3.CursorStyle = CursorStyle.CrossHair | CursorStyle.RubberLine
    #     if self.coordCtl2.ResultCoordinateSystem:
    #         self.coordCtl3.AnchorPoint = MousePosition(self.coordCtl2.ClickWindow, self.coordCtl2.Coordinate, self.coordCtl2.ResultCoordinateSystem)
    #     else:
    #         self.coordCtl3.AnchorPoint = None

    def SetDefaultOptions(self):

        self.drawtext.IsChecked = OptionsManager.GetBool("SCR_MeasureOffsetDistance.drawtext", False)
        self.textdecimals.Value = OptionsManager.GetDouble("SCR_MeasureOffsetDistance.textdecimals", 4.0)
        self.includedist.IsChecked = OptionsManager.GetBool("SCR_MeasureOffsetDistance.includedist", False)
        self.includeos.IsChecked = OptionsManager.GetBool("SCR_MeasureOffsetDistance.includeos", False)
        self.includedh.IsChecked = OptionsManager.GetBool("SCR_MeasureOffsetDistance.includedh", False)

        lserial = OptionsManager.GetUint("SCR_MeasureOffsetDistance.textlayerpicker", 8)
        o = self.currentProject.Concordance.Lookup(lserial)
        if o != None:
            if isinstance(o.GetSite(), LayerCollection):    
                self.textlayerpicker.SetSelectedSerialNumber(lserial, InputMethod(3))
            else:                       
                self.textlayerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:                       
            self.textlayerpicker.SetSelectedSerialNumber(8, InputMethod(3))


    def SaveOptions(self):
        OptionsManager.SetValue("SCR_MeasureOffsetDistance.drawtext", self.drawtext.IsChecked)
        OptionsManager.SetValue("SCR_MeasureOffsetDistance.textdecimals", self.textdecimals.Value)
        OptionsManager.SetValue("SCR_MeasureOffsetDistance.includedist", self.includedist.IsChecked)
        OptionsManager.SetValue("SCR_MeasureOffsetDistance.includeos", self.includeos.IsChecked)
        OptionsManager.SetValue("SCR_MeasureOffsetDistance.includedh", self.includedh.IsChecked)
        OptionsManager.SetValue("SCR_MeasureOffsetDistance.textlayerpicker", self.textlayerpicker.SelectedSerialNumber)


    def Coord1Changed(self, ctrl, e):
        self.coordCtl2.CursorStyle = CursorStyle.CrossHair | CursorStyle.RubberLine
        if self.coordCtl1.ResultCoordinateSystem:
            self.coordCtl2.AnchorPoint = MousePosition(self.coordCtl1.ClickWindow, self.coordCtl1.Coordinate, self.coordCtl1.ResultCoordinateSystem)
        else:
            self.coordCtl2.AnchorPoint = None

    def Coord3Changed(self, ctrl, e):
        if e.Cause == InputMethod.Mouse:     
            self.OkClicked(None, None)
        Keyboard.Focus(self.coordCtl3)

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''

        wv = self.currentProject [Project.FixedSerial.WorldView]

        p1 = self.coordCtl1.Coordinate      
        p2 = self.coordCtl2.Coordinate      
        p3 = self.coordCtl3.Coordinate      
        
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())

        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                # setup the line
                seg1 = SegmentLine(p1, p2)

                # prepare variables
                out_t = clr.StrongBox[float]()
                outPointOnCL = clr.StrongBox[Point3D]()
                testside = clr.StrongBox[Side]()

                # project the point
                seg1.ProjectPoint(p3, out_t, outPointOnCL, testside)
                offset_v = Vector3D(outPointOnCL.Value, p3)

                # write the result into the textboxes
                self.resultdistance.Text = self.lunits.Format(out_t.Value * seg1.Length2D , self.lfp)

                if Vector3D.HzAngleBetween(Vector3D(p1, p2), offset_v) < 0: # right side
                    self.resultoffset.Text = self.lunits.Format(offset_v.Length2D , self.lfp)
                else: # left side
                    self.resultoffset.Text = self.lunits.Format(-offset_v.Length2D , self.lfp)

                seg_dh_per_m = (p2.Z - p1.Z) / seg1.Length2D
                pointhog = p3.Z - (p1.Z + (out_t.Value * seg1.Length2D * seg_dh_per_m))
                self.resultdh.Text = self.lunits.Format(pointhog , self.lfp)
                
       
                if self.drawtext.IsChecked and (self.includedist.IsChecked or self.includeos.IsChecked or  self.includedh.IsChecked):
                    t = wv.Add(clr.GetClrType(MText))
                    t.AlignmentPoint = p3
                    t.AttachPoint = AttachmentPoint.BottomLeft
                    t.Rotation = 2 * math.pi/2 - Vector3D(p1, p2).Azimuth - math.pi/2
                    t.Height = 0.1
                    t.Layer = self.textlayerpicker.SelectedSerialNumber
                    
                    textdecimals = int(self.textdecimals.Value)
                    
                    if self.includedist.IsChecked: t.TextString = self.addtotextstring(t.TextString, "d=" + str("{:.{}f}".format(float(self.resultdistance.Text), textdecimals)))
                    if self.includeos.IsChecked: t.TextString = self.addtotextstring(t.TextString, "dO=" + str("{:.{}f}".format(float(self.resultoffset.Text), textdecimals)))
                    if self.includedh.IsChecked: t.TextString = self.addtotextstring(t.TextString, "dH=" + str("{:.{}f}".format(float(self.resultdh.Text), textdecimals)))

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


        self.SaveOptions()           

    def addtotextstring(self, t, addstring):
        if len(t) > 0:
            t += '\\P' + addstring
        else:
            t += addstring

        return t

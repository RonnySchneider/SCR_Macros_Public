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
    cmdData.Key = "SCR_BearingFlatnessCheck"
    cmdData.CommandName = "SCR_BearingFlatnessCheck"
    cmdData.Caption = "_SCR_BearingFlatnessCheck"
    cmdData.UIForm = "SCR_BearingFlatnessCheck" # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!

    try:
        cmdData.DefaultTabKey = "SCR Reports"
        cmdData.DefaultTabGroupKey = "Reports"
        cmdData.ShortCaption = "Bearing Flatness"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.05
        cmdData.MacroAuthor = "SCR"
        cmdData.ToolTipTitle = "Bearing Flatness Check"
        cmdData.ToolTipTextFormatted = "Bearing Flatness Check"
    except:
        pass
    try:
        b = Bitmap(macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


# the name of this class must match name from cmdData.UIForm (above)
class SCR_BearingFlatnessCheck(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader(macroFileFolder + r"\SCR_BearingFlatnessCheck.xaml") as s:
            wpf.LoadComponent(self, s)
        self.currentProject = currentProject
        self.macroFileFolder = macroFileFolder

    def HelpClicked(self, cmd, e):
        webbrowser.open("C:\ProgramData\Trimble\MacroCommands3\SCR Macros\MacroHelp\MacroHelp.htm#" + type(self).__name__)

    def OnLoad(self, cmd, buttons, event):
        self.okBtn = buttons[0]
        buttons[2].Content = "Help"
        buttons[2].Visibility = Visibility.Visible
        buttons[2].Click += self.HelpClicked
        buttons[0].Content = "Apply"
        buttons[1].Content = "Close"
        self.Caption = cmd.Command.Caption

        self.coordCtl1.ValueChanged += self.Coord1Changed
        self.coordCtl2.ValueChanged += self.Coord2Changed
        self.coordCtl3.ValueChanged += self.Coord3Changed

        self.cadpointType = clr.GetClrType(CadPoint)
        self.coordpointType = clr.GetClrType(CoordPoint)

        self.objs.IsEntityValidCallback = self.IsValid

        self.textdecimals.NumberOfDecimals = 0
        self.flatnessdecimals.NumberOfDecimals = 0

        #self.compareflatness.MinValue = 0
        #self.compareflatness.NumberOfDecimals = 0
        #self.compareflatness.Value = 300.0

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass
        self.unitssetup(None, None)

    def IsValid(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.coordpointType):
            return True
        if isinstance(o, self.cadpointType):
            return True
        return False

    def SetDefaultOptions(self):
        lserial = OptionsManager.GetUint("SCR_BearingFlatnessCheck.layerpicker", 8)
        o = self.currentProject.Concordance.Lookup(lserial)
        if o != None:
            if isinstance(o.GetSite(), LayerCollection):    
                self.layerpicker.SetSelectedSerialNumber(lserial, InputMethod(3))
            else:                       
                self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:                       
            self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        try:    self.incolorpicker.SelectedColor = Color.FromArgb(OptionsManager.GetInt("SCR_BearingFlatnessCheck.incolorpicker"))
        except: self.incolorpicker.SelectedColor = Color.Red
        try:    self.outcolorpicker.SelectedColor = Color.FromArgb(OptionsManager.GetInt("SCR_BearingFlatnessCheck.outcolorpicker"))
        except: self.outcolorpicker.SelectedColor = Color.Red

        self.textheight.Distance = OptionsManager.GetDouble("SCR_BearingFlatnessCheck.textheightfloat", 0.01)
        self.unitpicker.Text = OptionsManager.GetString("SCR_BearingFlatnessCheck.unitpicker", "Meter")
        self.addunitsuffix.IsChecked = OptionsManager.GetBool("SCR_BearingFlatnessCheck.addunitsuffix", False)
        self.textdecimals.Value = OptionsManager.GetDouble("SCR_BearingFlatnessCheck.textdecimalsfloat", 4)

        self.compareunitpicker.Text = OptionsManager.GetString("SCR_BearingFlatnessCheck.compareunitpicker", "Meter")
        self.compareaddunitsuffix.IsChecked = OptionsManager.GetBool("SCR_BearingFlatnessCheck.compareaddunitsuffix", False)
        self.flatnessdecimals.Value = OptionsManager.GetDouble("SCR_BearingFlatnessCheck.flatnessdecimals", 0)
        self.compareflatness1.Distance = OptionsManager.GetDouble("SCR_BearingFlatnessCheck.compareflatness1", 0.001)
        self.compareflatness2.Distance = OptionsManager.GetDouble("SCR_BearingFlatnessCheck.compareflatness2", 0.300)

    def SaveOptions(self):
        OptionsManager.SetValue("SCR_BearingFlatnessCheck.layerpicker", self.layerpicker.SelectedSerialNumber)
        OptionsManager.SetValue("SCR_BearingFlatnessCheck.incolorpicker", self.incolorpicker.SelectedColor.ToArgb())
        OptionsManager.SetValue("SCR_BearingFlatnessCheck.outcolorpicker", self.outcolorpicker.SelectedColor.ToArgb())

        OptionsManager.SetValue("SCR_BearingFlatnessCheck.textheightfloat", self.textheight.Distance)
        OptionsManager.SetValue("SCR_BearingFlatnessCheck.unitpicker", self.unitpicker.Text)
        OptionsManager.SetValue("SCR_BearingFlatnessCheck.addunitsuffix", self.addunitsuffix.IsChecked)
        OptionsManager.SetValue("SCR_BearingFlatnessCheck.textdecimalsfloat", self.textdecimals.Value)
        
        OptionsManager.SetValue("SCR_BearingFlatnessCheck.compareunitpicker", self.compareunitpicker.Text)
        OptionsManager.SetValue("SCR_BearingFlatnessCheck.compareaddunitsuffix", self.compareaddunitsuffix.IsChecked)
        OptionsManager.SetValue("SCR_BearingFlatnessCheck.flatnessdecimals", self.flatnessdecimals.Value)
        OptionsManager.SetValue("SCR_BearingFlatnessCheck.compareflatness1", self.compareflatness1.Distance)
        OptionsManager.SetValue("SCR_BearingFlatnessCheck.compareflatness2", self.compareflatness2.Distance)

    def unitssetup(self, sender, e):
        # setup everything for the unit conversions
        self.outputunitenum = 0
        self.textdecimals.NumberOfDecimals = 0
        self.compareoutputunitenum = 0
        self.flatnessdecimals.NumberOfDecimals = 0

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        self.lfp = self.lunits.Properties.Copy() # create a copy in order to set the decimals and enable/disable the suffix
        self.lfpcompare = self.lunits.Properties.Copy() # create a copy in order to set the decimals and enable/disable the suffix
        self.lfp.AddSuffix = False # disable suffix, we need to set it manually, it would always add the current projects units
        self.lfpcompare.AddSuffix = False # disable suffix, we need to set it manually, it would always add the current projects units

        # fill the unitpicker
        for u in self.lunits.Units:
            item = ComboBoxItem()
            item.Content = u.Key
            item.FontSize = 1
            self.unitpicker.Items.Add(item)

        # fill the unitpicker
        for u in self.lunits.Units:
            item = ComboBoxItem()
            item.Content = u.Key
            item.FontSize = 1
            self.compareunitpicker.Items.Add(item)

        tt = self.unitpicker.Text
        self.unitpicker.SelectedIndex = 0
        if tt != "":
            self.unitpicker.Text = tt

        tt = self.compareunitpicker.Text
        self.compareunitpicker.SelectedIndex = 0
        if tt != "":
            self.compareunitpicker.Text = tt

        self.unitpicker.SelectionChanged += self.unitschanged
        self.compareunitpicker.SelectionChanged += self.unitschanged
        
        self.textdecimals.MinValue = 0.0
        self.textdecimals.ValueChanged += self.unitschanged
        self.flatnessdecimals.MinValue = 0.0
        self.flatnessdecimals.ValueChanged += self.unitschanged

        self.lfp.NumberOfDecimals = int(self.textdecimals.Value)
        self.lfpcompare.NumberOfDecimals = int(self.flatnessdecimals.Value)
        self.unitschanged(None, None)
    
    def unitschanged(self, sender, e):

        # find the enum for the selected LinearType
        for e in range(0, 19):
            if LinearType(e) == self.unitpicker.SelectedItem.Content:
                self.outputunitenum = e
        # find the enum for the selected LinearType
        for e in range(0, 19):
            if LinearType(e) == self.compareunitpicker.SelectedItem.Content:
                self.compareoutputunitenum = e
        

        self.textheight.DisplayUnit = LinearType(self.outputunitenum)
        self.textheight.ShowControlIcon(False)
        self.textheight.FormatProperty.AddSuffix = ControlBoolean(1)
        self.textheight.FormatProperty.NumberOfDecimals = int(self.textdecimals.Value)

        self.compareflatness1.DisplayUnit = LinearType(self.compareoutputunitenum)
        self.compareflatness1.ShowControlIcon(False)
        self.compareflatness1.FormatProperty.AddSuffix = ControlBoolean(1)
        self.compareflatness1.FormatProperty.NumberOfDecimals = int(self.flatnessdecimals.Value)
        self.compareflatness2.DisplayUnit = LinearType(self.compareoutputunitenum)
        self.compareflatness2.ShowControlIcon(False)
        self.compareflatness2.FormatProperty.AddSuffix = ControlBoolean(1)
        self.compareflatness2.FormatProperty.NumberOfDecimals = int(self.flatnessdecimals.Value)

    def decdecimals_Click(self, sender, e):
        if not self.textdecimals.Value == 0:
            self.textdecimals.Value -= 1
             # setup the linear format properties
            self.lfp.NumberOfDecimals = int(self.textdecimals.Value)
            self.unitschanged(None, None)

    def incdecimals_Click(self, sender, e):
        self.textdecimals.Value += 1
        # setup the linear format properties
        self.lfp.NumberOfDecimals = int(self.textdecimals.Value)
        self.unitschanged(None, None)

    def decflatdecimals_Click(self, sender, e):
        if not self.flatnessdecimals.Value == 0:
            self.flatnessdecimals.Value -= 1
             # setup the linear format properties
            self.lfpcompare.NumberOfDecimals = int(self.flatnessdecimals.Value)
            self.unitschanged(None, None)

    def incflatdecimals_Click(self, sender, e):
        self.flatnessdecimals.Value += 1
        # setup the linear format properties
        self.lfpcompare.NumberOfDecimals = int(self.flatnessdecimals.Value)
        self.unitschanged(None, None)

    def tooutputunit(self, v):
        
        self.lfp.AddSuffix = self.addunitsuffix.IsChecked
        return self.lunits.Format(LinearType.Meter, v, self.lfp, LinearType(self.outputunitenum))

    def tooutputunitcompare(self, v):
        
        self.lfpcompare.AddSuffix = self.compareaddunitsuffix.IsChecked
        return self.lunits.Format(LinearType.Meter, v, self.lfpcompare, LinearType(self.compareoutputunitenum))

    def usemultipointChanged(self, sender, e):
        if self.usemultipoint.IsChecked:
            self.coordCtl4.IsEnabled = False
            self.objs.IsEnabled = True
        else:
            self.coordCtl4.IsEnabled = True
            self.objs.IsEnabled = False


    def Coord1Changed(self, ctrl, e):
        self.coordCtl2.CursorStyle = CursorStyle.CrossHair | CursorStyle.RubberLine
        if self.coordCtl1.ResultCoordinateSystem:
            self.coordCtl2.AnchorPoint = MousePosition(self.coordCtl1.ClickWindow, self.coordCtl1.Coordinate, self.coordCtl1.ResultCoordinateSystem)
        else:
            self.coordCtl2.AnchorPoint = None

    def Coord2Changed(self, ctrl, e):
        self.coordCtl3.CursorStyle = CursorStyle.CrossHair | CursorStyle.RubberLine
        if self.coordCtl2.ResultCoordinateSystem:
            self.coordCtl3.AnchorPoint = MousePosition(self.coordCtl2.ClickWindow, self.coordCtl2.Coordinate, self.coordCtl2.ResultCoordinateSystem)
        else:
            self.coordCtl3.AnchorPoint = None

    def Coord3Changed(self, ctrl, e):
        self.coordCtl1.CursorStyle = CursorStyle.CrossHair | CursorStyle.RubberLine
        if self.coordCtl3.ResultCoordinateSystem:
            self.coordCtl1.AnchorPoint = MousePosition(self.coordCtl3.ClickWindow, self.coordCtl3.Coordinate, self.coordCtl3.ResultCoordinateSystem)
        else:
            self.coordCtl1.AnchorPoint = None

        
    def CancelClicked(self, thisCmd, args):
        thisCmd.CloseUICommand()

    def OkClicked(self, thisCmd, e):
        Keyboard.Focus(self.okBtn)  # a trick to evaluate all input fields before execution, otherwise you'd have to click in another field first

        self.success.Content=''

        wv = self.currentProject[Project.FixedSerial.WorldView]

        # get the units for linear distance
        # self.lunits = self.currentProject.Units.Linear
        # we don't want the units to be included (so we make copy and turn that off). Otherwise get something like "12.50 ft"
        # self.lfp = self.lunits.Properties.Copy()
        # self.lfp.AddSuffix = False
        pointlist = []
        for o in self.objs.SelectedMembers(self.currentProject):
            if isinstance(o, self.cadpointType):
                pointlist.Add(o.Point0)
            elif isinstance(o, self.coordpointType):
                pointlist.Add(o.Position)



        inputok=True

        textdecimals = abs(int(self.textdecimals.Value))
        flatnessdecimals = abs(int(self.flatnessdecimals.Value))
        th = abs(self.textheight.Distance)
        if th == 0.0:
            self.error.Content='Text Height set to Zero'

        p1_sel = self.coordCtl1.Coordinate
        if not p1_sel.Is3D :
            self.coordCtl1.StatusMessage = "No coordinate defined"
            return False
        self.coordCtl1.StatusMessage = ""

        p2_sel = self.coordCtl2.Coordinate
        if not p2_sel.Is3D :
            self.coordCtl2.StatusMessage = "No coordinate defined"
            return False
        self.coordCtl2.StatusMessage = ""

        p3_sel = self.coordCtl3.Coordinate
        if not p3_sel.Is3D :
            self.coordCtl3.StatusMessage = "No coordinate defined"
            return False
        self.coordCtl3.StatusMessage = ""
        
        ## make sure that p1_sel to p2_sel is the long side
        #if Vector3D(p1_sel, p2_sel).Length < Vector3D(p2_sel, p3_sel).Length:
        #    p1_sel, p3_sel = p3_sel, p1_sel # swap if necessary
        
        length_long_side = Vector3D(p1_sel, p2_sel).Length
        length_short_side = Vector3D(p2_sel, p3_sel).Length

        if inputok:
            self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
            UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())

            try:

                # draw the ref text markers
                shiftvec = Vector3D(p1_sel, p2_sel)
                shiftvec.Length = length_long_side / 40
                shiftvec.RotateAboutZ(math.pi/2 + math.pi/4)
                self.drawreftext(p1_sel + shiftvec, "1", Vector3D(p2_sel, p3_sel).Azimuth, AttachmentPoint.TopRight)
                shiftvec.Rotate90(Side.Right)
                self.drawreftext(p2_sel + shiftvec, "2", Vector3D(p2_sel, p3_sel).Azimuth, AttachmentPoint.BottomRight)
                shiftvec.Rotate90(Side.Right)
                self.drawreftext(p3_sel + shiftvec, "3", Vector3D(p2_sel, p3_sel).Azimuth, AttachmentPoint.BottomLeft)

                long_side_segment = SegmentLine(p1_sel, p2_sel)

                with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                    l1_exists, m1_exists, m2_exists, m3_exists, r1_exists, r2_exists = (False,) * 6
                    if pointlist.Count > 0:
                        
                        #newSurface = wv.Add(clr.GetClrType(Model3D))
                        #newSurface.Name = "test"
                        #builder = newSurface.GetGemBatchBuilder()
                        #builder.AddVertex(p1_sel)
                        #builder.AddVertex(p2_sel)
                        #builder.AddVertex(p3_sel)
                        
                        for p4_sel in pointlist:
                            #builder.AddVertex(p4_sel)

                            #   p2_sel  m3      p3_sel
                            #
                            #   l1      m2      r2
                            #
                            #   p1_sel  m1      r1

                            # go through all the points in the list and try to assign a name from the abvove scheme
                            if not p4_sel.Is3D:
                                return False

                            if not Vector3D(p4_sel, p1_sel).Length == 0 or Vector3D(p4_sel, p2_sel).Length == 0 or Vector3D(p4_sel, p3_sel).Length == 0:
                                
                                testplane = Plane3D(p1_sel, p2_sel, p3_sel)[0] # the plane is returned as first element
                                if not testplane.IsValid:
                                    return

                                out_t = clr.StrongBox[float]()
                                outPointOnCL = clr.StrongBox[Point3D]()
                                out_side = clr.StrongBox[Side]()
                                
                                if long_side_segment.ProjectPoint(p4_sel, out_t, outPointOnCL, out_side):
                                    # check if it could be l1
                                    if out_t.Value > 0.3 and out_t.Value < 0.7 and Vector3D(p4_sel, outPointOnCL.Value).Length < 0.05:
                                        l1 = p4_sel
                                        l1_exists = True
                                    # check if it could be m1
                                    if out_t.Value > -0.2 and out_t.Value < 0.2 and abs(length_short_side/2 - Vector3D(p4_sel, outPointOnCL.Value).Length) < 0.05:
                                        m1 = p4_sel
                                        m1_exists = True
                                    # check if it could be m2
                                    if out_t.Value > 0.3 and out_t.Value < 0.7 and abs(length_short_side/2 - Vector3D(p4_sel, outPointOnCL.Value).Length) < 0.05:
                                        m2 = p4_sel
                                        m2_exists = True
                                    # check if it could be m3
                                    if out_t.Value > 0.8 and out_t.Value < 1.2 and abs(length_short_side/2 - Vector3D(p4_sel, outPointOnCL.Value).Length) < 0.05:
                                        m3 = p4_sel
                                        m3_exists = True
                                    # check if it could be r1
                                    if out_t.Value > -0.2 and out_t.Value < 0.2 and abs(length_short_side - Vector3D(p4_sel, outPointOnCL.Value).Length) < 0.05:
                                        r1 = p4_sel
                                        r1_exists = True
                                    # check if it could be r2
                                    if out_t.Value > 0.3 and out_t.Value < 0.7 and abs(length_short_side - Vector3D(p4_sel, outPointOnCL.Value).Length) < 0.05:
                                        r2 = p4_sel
                                        r2_exists = True
                        
                        #   p2_sel  m3      p3_sel
                        #
                        #   l1      m2      r2
                        #
                        #   p1_sel  m1      r1

                        # now that we know where the points are we can compute the flatness and draw the result

                        if r1_exists:
                            if r2_exists: 
                                self.drawflatness(r1, r2, testplane, AttachmentPoint.BottomMid)
                            else:
                                self.drawflatness(r1, p3_sel, testplane, AttachmentPoint.BottomMid)
                            
                            if m1_exists:
                                self.drawflatness(m1, r1, testplane, AttachmentPoint.BottomMid)
                            else:
                                self.drawflatness(p1_sel, r1, testplane, AttachmentPoint.BottomMid)

                            if m2_exists:
                                self.drawflatness(r1, m2, testplane, AttachmentPoint.MiddleMid)
                            else:
                                self.drawflatness(p2_sel, r1, testplane, AttachmentPoint.MiddleMid)


                        if m1_exists:
                            self.drawflatness(p1_sel, m1, testplane, AttachmentPoint.BottomMid)
                            if m2_exists:
                                self.drawflatness(m1, m2, testplane, AttachmentPoint.MiddleMid)
                            else:
                                if m3_exists:
                                    self.drawflatness(m1, m3, testplane, AttachmentPoint.MiddleMid)

                        if r2_exists:
                            self.drawflatness(r2, p3_sel, testplane, AttachmentPoint.BottomMid)
                            if m2_exists:
                                self.drawflatness(m2, r2, testplane, AttachmentPoint.MiddleMid)
                            else:
                                if l1_exists:
                                    self.drawflatness(l1, r2, testplane, AttachmentPoint.MiddleMid)

                        #   p2_sel  m3      p3_sel
                        #
                        #   l1      m2      r2
                        #
                        #   p1_sel  m1      r1

                        if m2_exists:
                            self.drawflatness(p1_sel, m2, testplane, AttachmentPoint.MiddleMid)
                            self.drawflatness(m2, p2_sel, testplane, AttachmentPoint.MiddleMid)
                            self.drawflatness(m2, p3_sel, testplane, AttachmentPoint.MiddleMid)
                            if m3_exists:
                                self.drawflatness(m2, m3, testplane, AttachmentPoint.MiddleMid)
                            if l1_exists:
                                self.drawflatness(l1, m2, testplane, AttachmentPoint.MiddleMid)

                        
                        if l1_exists:
                            self.drawflatness(p1_sel, l1, testplane, AttachmentPoint.TopCenter)
                            self.drawflatness(l1, p2_sel, testplane, AttachmentPoint.TopCenter)
                        else:
                            self.drawflatness(p1_sel, p2_sel, testplane, AttachmentPoint.TopCenter)

                        if m3_exists:
                            self.drawflatness(p2_sel, m3, testplane, AttachmentPoint.TopCenter)
                            self.drawflatness(m3, p3_sel, testplane, AttachmentPoint.TopCenter)
                        else:
                            self.drawflatness(p2_sel, p3_sel, testplane, AttachmentPoint.TopCenter)
                        
                        #if builder:
                        #    builder.Construction()
                        #    builder.Commit()

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
        
        Keyboard.Focus(self.coordCtl1)
        self.SaveOptions()
 
    def drawflatness(self, p1, p2, testplane, attach):

        wv = self.currentProject[Project.FixedSerial.WorldView]

        ptemp = Plane3D.IntersectWithRayPerpendicular(testplane, p1) # intersection point with plane
        p1_z = Vector3D(p1, ptemp).Length # is a signless perpendicular distance to the plane
        if p1.Z < ptemp.Z: p1_z *= -1 # check if the original point is below the plane and set sign negative if necessary

        ptemp = Plane3D.IntersectWithRayPerpendicular(testplane, p2) # intersection point with plane
        p2_z = Vector3D(p2, ptemp).Length # is a signless perpendicular distance to the plane
        if p2.Z < ptemp.Z: p2_z *= -1 # check if the original point is below the plane and set sign negative if necessary

        flatnessperm = (p2_z - p1_z) / Vector3D(p1, p2).Length # slope per metre
        flatnesspercompare = flatnessperm * self.compareflatness2.Distance # slope per compare distance, still in metres

        textObj = wv.Add(clr.GetClrType(MText))
        self.lfp.NumberOfDecimals = abs(int(self.textdecimals.Value))
        textObj.TextString = self.tooutputunit(p2_z - p1_z) + " -->"

        self.lfp.NumberOfDecimals = abs(int(self.flatnessdecimals.Value))
        textObj.TextString += "\n" + self.tooutputunitcompare(flatnesspercompare) + "/" + self.tooutputunitcompare(self.compareflatness2.Distance)
        textObj.Point0 = Point3D.MidPoint(p1, p2)
        textObj.AttachPoint = attach # or BottomMid
        # text rotation works counter clockwise, with 0 in positive x direction
        textObj.Rotation = (2*math.pi) - Vector3D(p1, p2).Azimuth + math.pi/2
        textObj.Height = abs(float(self.textheight.Distance))
        textObj.Layer = self.layerpicker.SelectedSerialNumber
        textObj.AddWhiteOut = True

        t1 = self.lunits.Convert(LinearType.Meter, abs(flatnesspercompare), LinearType(self.compareoutputunitenum))
        t2 = self.lunits.Convert(LinearType.Meter, abs(self.compareflatness1.Distance), LinearType(self.compareoutputunitenum))
        if t1 > t2:
            textObj.Color = self.outcolorpicker.SelectedColor
        else:
            textObj.Color = self.incolorpicker.SelectedColor

    def drawreftext(self, p1, count, rot, attach):

        wv = self.currentProject[Project.FixedSerial.WorldView]

        textObj = wv.Add(clr.GetClrType(MText))
        textObj.TextString = "Ref " + count
        textObj.Point0 = p1
        textObj.AttachPoint = attach
        textObj.Rotation = (2*math.pi) - rot + math.pi/2
        textObj.Height = abs(float(self.textheight.Distance)) * 0.75
        textObj.Layer = self.layerpicker.SelectedSerialNumber
        textObj.AddWhiteOut = True
    
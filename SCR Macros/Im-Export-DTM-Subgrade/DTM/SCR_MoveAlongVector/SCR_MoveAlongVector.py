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
    cmdData.Key = "SCR_MoveAlongVector"
    cmdData.CommandName = "SCR_MoveAlongVector"
    cmdData.Caption = "_SCR_MoveAlongVector"
    cmdData.UIForm = "SCR_MoveAlongVector" # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "DTM"
        cmdData.ShortCaption = "move along Vector"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.09
        cmdData.MacroAuthor = "SCR"
        cmdData.ToolTipTitle = "Move Objects along Vector"
        cmdData.ToolTipTextFormatted = "Move Objects along Vector"
    except:
        pass
    try:
        b = Bitmap(macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


# the name of this class must match name from cmdData.UIForm (above)
class SCR_MoveAlongVector(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader(macroFileFolder + r"\SCR_MoveAlongVector.xaml") as s:
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

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        self.lfp = self.lunits.Properties.Copy()
        self.linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation
        #self.lfp.AddSuffix = False

        self.coordCtl1.ValueChanged += self.Coord1Changed
        self.coordCtl2.ValueChanged += self.Coord2Changed
        self.coordCtl3.ValueChanged += self.Coord3Changed

        #self.offset.NumberOfDecimals = 4
        self.offset.Allow3DDistance = True
        self.offset.DistanceType = DistanceType.XYZ

        self.cadpointType = clr.GetClrType(CadPoint)
        self.coordpointType = clr.GetClrType(CoordPoint)

        self.lType = clr.GetClrType(IPolyseg)

        types = Array [Type] (SurfaceTypeLists.AllWithCutFillMap)    # we fill an array with TBC object types, we could combine different types
        self.touchsurfacepicker.FilterByEntityTypes = types    # we fill the dropdownlist by applying that types array as filter
        self.touchsurfacepicker.AllowNone = False              # our list shall not show an empty field

        # self.objs.IsEntityValidCallback = self.IsValid
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
        if isinstance(o, self.coordpointType):
            return True
        if isinstance(o, self.cadpointType):
            return True
        return False

    def SetDefaultOptions(self):

        self.alongvector.IsChecked = OptionsManager.GetBool("SCR_MoveAlongVector.alongvector", True)
        self.fromplane.IsChecked = OptionsManager.GetBool("SCR_MoveAlongVector.fromplane", False)
        self.fixedoffset.IsChecked = OptionsManager.GetBool("SCR_MoveAlongVector.fixedoffset", True)
        self.offset.Distance = OptionsManager.GetDouble("SCR_MoveAlongVector.offset", 0.0)
        self.fromplane.IsChecked = OptionsManager.GetBool("SCR_MoveAlongVector.fromplane", False)
        try:    self.touchsurfacepicker.SelectIndex(OptionsManager.GetInt("SCR_MoveAlongVector.touchsurfacepicker", 0))
        except: self.touchsurfacepicker.SelectIndex(0)
        self.returntocoordCtl1.IsChecked = OptionsManager.GetBool("SCR_MoveAlongVector.returntocoordCtl1", False)

    def SaveOptions(self):

        OptionsManager.SetValue("SCR_MoveAlongVector.alongvector", self.alongvector.IsChecked)    
        OptionsManager.SetValue("SCR_MoveAlongVector.fromplane", self.fromplane.IsChecked)    
        OptionsManager.SetValue("SCR_MoveAlongVector.fixedoffset", self.fixedoffset.IsChecked)    
        OptionsManager.SetValue("SCR_MoveAlongVector.offset", self.offset.Distance)
        OptionsManager.SetValue("SCR_MoveAlongVector.touchsurface", self.touchsurface.IsChecked)    
        try:    # if nothing is selected it would throw an error
            OptionsManager.SetValue("SCR_MoveAlongVector.touchsurfacepicker", self.touchsurfacepicker.SelectedIndex)
        except:
            pass
        OptionsManager.SetValue("SCR_MoveAlongVector.returntocoordCtl1", self.returntocoordCtl1.IsChecked)    

    def Coord1Changed(self, ctrl, e):
        self.coordCtl2.CursorStyle = CursorStyle.CrossHair | CursorStyle.RubberLine
        if self.coordCtl1.ResultCoordinateSystem:
            self.coordCtl2.AnchorPoint = MousePosition(self.coordCtl1.ClickWindow, self.coordCtl1.Coordinate, self.coordCtl1.ResultCoordinateSystem)
        else:
            self.coordCtl2.AnchorPoint = None
        Keyboard.Focus(self.coordCtl2)

    def Coord2Changed(self, ctrl, e):
        self.coordCtl3.CursorStyle = CursorStyle.CrossHair | CursorStyle.RubberLine
        if self.coordCtl2.ResultCoordinateSystem:
            self.coordCtl3.AnchorPoint = MousePosition(self.coordCtl2.ClickWindow, self.coordCtl2.Coordinate, self.coordCtl2.ResultCoordinateSystem)
        else:
            self.coordCtl3.AnchorPoint = None

        if self.alongvector.IsChecked:
            Keyboard.Focus(self.offset)
        else:
            Keyboard.Focus(self.coordCtl3)

    def Coord3Changed(self, ctrl, e):
        Keyboard.Focus(self.objs)

    def planeorvectorChanged(self, sender, e):
        if self.fromplane.IsChecked:
            self.fixedoffset.Content = "distance from plane [" + self.linearsuffix + "]"
        else:
            self.fixedoffset.Content = "distance along vector [" + self.linearsuffix + "]"

    def CancelClicked(self, thisCmd, args):
        thisCmd.CloseUICommand()

    def OkClicked(self, thisCmd, e):
        Keyboard.Focus(self.okBtn)  # a trick to evaluate all input fields before execution, otherwise you'd have to click in another field first

        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())

        self.success.Content=''
        self.error.Content=''

        wv = self.currentProject[Project.FixedSerial.WorldView]

        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                inputok = True
                p1_sel = self.coordCtl1.Coordinate
                if not p1_sel.Is3D :
                    self.coordCtl1.StatusMessage = "No coordinate defined"
                    inputok = False
                self.coordCtl1.StatusMessage = ""

                p2_sel = self.coordCtl2.Coordinate
                if not p2_sel.Is3D :
                    self.coordCtl2.StatusMessage = "No coordinate defined"
                    inputok = False
                self.coordCtl2.StatusMessage = ""

                if self.fromplane.IsChecked:
                    p3_sel = self.coordCtl3.Coordinate
                    if not p3_sel.Is3D:
                        self.coordCtl3.StatusMessage = "No coordinate defined"
                        inputok = False
                    self.coordCtl3.StatusMessage = ""

                    if inputok:
                        p = Plane3D(p1_sel, p2_sel, p3_sel)
                        p = p[0] # the plane is returned as first element
                        if not p.IsValid:
                            inputok = False
                        else:
                            v_perp_face = p.normal
                else:
                    if inputok:
                        v_perp_face = Vector3D(p1_sel, p2_sel)
                # v.Length = self.offset.Value
                
                surface = wv.Lookup(self.touchsurfacepicker.SelectedSerial)
                tiesolfromabove = 0
                tiesolfrombelow = 0
                

                outPointOnCL = clr.StrongBox[Point3D]()
                station = clr.StrongBox[float]()
                
                tuple = Array[Point3D]([Point3D()]*2)
                tuples = DynArray()

                if inputok:
                    if self.touchsurface.IsChecked:
                        v_perp_face.Length = 1

                        for o in self.objs:
                            if isinstance(o, self.lType):
                                polyseg1 = o.ComputePolySeg()
                                polyseg1 = polyseg1.ToWorld()
                                polyseg1_v = o.ComputeVerticalPolySeg()
                                
                                # go through all the nodes in the linestring
                                for testpoint in polyseg1.Nodes():

                                    polyseg1.FindPointFromPoint(testpoint.Point, outPointOnCL, station)
                                    testpoint = outPointOnCL.Value
                                    if polyseg1_v != None:
                                        testpoint.Z = polyseg1_v.ComputeVerticalSlopeAndGrade(station.Value)[1]

                                    tiepoint = clr.StrongBox[Point3D]()
                                    # slope in Computetie is zenith angle with upwards=0
                                    # Vector3D.Horizon is positive above the horizon and negative below
                                    if surface.ComputeTie(testpoint, v_perp_face, math.pi/2 - (v_perp_face.Horizon), 100, tiepoint):
                                        if Vector3D(testpoint, tiepoint.Value).Horizon < 0: # point is above target surface
                                            if tiesolfromabove == 0:
                                                tiesolfromabove +=1
                                                tiepointfromabove = tiepoint.Value
                                                abovetiestart = testpoint
                                            else:
                                                if Vector3D(testpoint, tiepoint.Value).Length > Vector3D(abovetiestart, tiepointfromabove).Length: # from above we search the longest offset
                                                    tiesolfromabove +=1
                                                    tiepointfromabove = tiepoint.Value
                                                    abovetiestart = testpoint

                                        else: # point is below target surface
                                            if tiesolfrombelow == 0:
                                                tiesolfrombelow +=1
                                                tiepointfrombelow = tiepoint.Value
                                                belowtiestart = testpoint
                                            else:
                                                if Vector3D(testpoint, tiepoint.Value).Length < Vector3D(belowtiestart, tiepointfrombelow).Length: # from below the minimum distance
                                                    tiesolfrombelow +=1
                                                    tiepointfrombelow = tiepoint.Value
                                                    belowtiestart = testpoint


                        if tiesolfromabove > 0:
                            tt = Vector3D(abovetiestart, tiepointfromabove).Length
                            td = TransformData(Matrix4D(Vector3D(abovetiestart, tiepointfromabove)), Matrix4D(Vector3D.Zero))                
                        else:
                            if tiesolfrombelow > 0:
                                tt = Vector3D(belowtiestart, tiepointfrombelow).Length
                                td = TransformData(Matrix4D(Vector3D(belowtiestart, tiepointfrombelow)), Matrix4D(Vector3D.Zero))
                        
                        for o in self.objs:
                            try:
                                o.Transform(td)
                            except:
                                pass

                    if self.fixedoffset.IsChecked:

                        v_perp_face.Length = self.offset.Distance
                        td = TransformData(Matrix4D(v_perp_face), Matrix4D(Vector3D.Zero))

                        for o in self.objs:
                            try:
                                o.Transform(td)
                            except:
                                pass
                else:
                    self.error.Content += '\nProblem with the Input Coordinates - are they 3D?'
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
        
        if self.returntocoordCtl1.IsChecked:
            Keyboard.Focus(self.coordCtl1)
        self.SaveOptions()
 

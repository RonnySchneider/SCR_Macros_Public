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
    cmdData.Key = "SCR_RotateIFC"
    cmdData.CommandName = "SCR_RotateIFC"
    cmdData.Caption = "_SCR_RotateIFC"
    cmdData.UIForm = "SCR_RotateIFC"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "IFC"
        cmdData.ShortCaption = "Rotate IFC"
        cmdData.DefaultRibbonToolSize = 0 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.07
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Rotate IFC"
        cmdData.ToolTipTextFormatted = "Rotate IFC"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_RotateIFC(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_RotateIFC.xaml") as s:
            wpf.LoadComponent (self, s)
        self.currentProject = currentProject
        self.macroFileFolder = macroFileFolder
        self.savedserials = []
        self.computeorigin = Point3D()
        self.computex = Point3D()
        self.computey = Point3D()
        self.computez = Point3D()

        self.ViewOverlay = Guid.NewGuid()
        self.overlayBag = OverlayBag(self.ViewOverlay)

    def HelpClicked(self, cmd, e):
        webbrowser.open("C:\ProgramData\Trimble\MacroCommands3\SCR Macros\MacroHelp\MacroHelp.htm#" + type(self).__name__)

    def OnLoad(self, cmd, buttons, event):
        self.okBtn = buttons[0]
        buttons[2].Content = "Help"
        buttons[2].Visibility = Visibility.Visible
        buttons[2].Click += self.HelpClicked
        self.Caption = cmd.Command.Caption

        self.objs.IsEntityValidCallback = self.IsValid

        optionMenu = SelectionContextMenuHandler()
        # remove options that don't apply here
        optionMenu.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        self.objs.ButtonContextMenu = optionMenu

        self.coordpick1.ShowElevationIf3D = True
        self.coordpick1.ValueChanged += self.CoordPick1Changed
        self.coordpick2.ShowElevationIf3D = True
        self.coordpick2.ValueChanged += self.CoordPick2Changed
        self.coordpick3.ShowElevationIf3D = True
        self.coordpick3.ValueChanged += self.CoordPick3Changed
        self.coordpick4.ShowElevationIf3D = True
        self.coordpick4.ValueChanged += self.CoordPick4Changed

        self.log = []

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass
        #self.coordpick1.PointSelected += self.coordpick1mouseclick
        #
        #self.coordpick2.ShowGdiCursor += self.coordpick2changed
        #self.coordpick2.PointSelected += self.coordpick2mouseclick
        #
        #self.coordpick2.AutoTab = False
        #
        #if self.objs.Count > 0:
        #    Keyboard.Focus(self.coordpick1)

    #def coordpick1mouseclick(self, ctrl, e):
    #    self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
    #    #UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
    #
    #def coordpick2mouseclick(self, ctrl, e):
    #    Keyboard.Focus(self.coordpick1)
    #
    #    self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
    #    #UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
    #


    def SetDefaultOptions(self):
        self.squarez.IsChecked = OptionsManager.GetBool("SCR_RotateIFC.squarez", True)
        self.independentz.IsChecked = OptionsManager.GetBool("SCR_RotateIFC.independentz", False)
        self.rotateaxes.IsChecked = OptionsManager.GetBool("SCR_RotateIFC.rotateaxes", True)
        self.xrot.Angle = OptionsManager.GetDouble("SCR_RotateIFC.xrot", 0.1)
        self.yrot.Angle = OptionsManager.GetDouble("SCR_RotateIFC.yrot", 0.1)
        self.zrot.Angle = OptionsManager.GetDouble("SCR_RotateIFC.zrot", 0.1)
        # need to set the checkboxes first, otherwise the show overlay might be wrong
        wv = self.currentProject[Project.FixedSerial.WorldView]
        x = OptionsManager.GetDouble("SCR_RotateIFC.coordpick1_x", 0.0)
        y = OptionsManager.GetDouble("SCR_RotateIFC.coordpick1_y", 0.0)
        z = OptionsManager.GetDouble("SCR_RotateIFC.coordpick1_z", 0.0)
        self.coordpick1.SetCoordinate(Point3D(x, y, z), self.currentProject, wv.CoordinateSystemDefinition)
        x = OptionsManager.GetDouble("SCR_RotateIFC.coordpick2_x", 0.0)
        y = OptionsManager.GetDouble("SCR_RotateIFC.coordpick2_y", 0.0)
        z = OptionsManager.GetDouble("SCR_RotateIFC.coordpick2_z", 0.0)
        self.coordpick2.SetCoordinate(Point3D(x, y, z), self.currentProject, wv.CoordinateSystemDefinition)
        x = OptionsManager.GetDouble("SCR_RotateIFC.coordpick3_x", 0.0)
        y = OptionsManager.GetDouble("SCR_RotateIFC.coordpick3_y", 0.0)
        z = OptionsManager.GetDouble("SCR_RotateIFC.coordpick3_z", 0.0)
        self.coordpick3.SetCoordinate(Point3D(x, y, z), self.currentProject, wv.CoordinateSystemDefinition)
        x = OptionsManager.GetDouble("SCR_RotateIFC.coordpick4_x", 0.0)
        y = OptionsManager.GetDouble("SCR_RotateIFC.coordpick4_y", 0.0)
        z = OptionsManager.GetDouble("SCR_RotateIFC.coordpick4_z", 0.0)
        self.coordpick4.SetCoordinate(Point3D(x, y, z), self.currentProject, wv.CoordinateSystemDefinition)
        # need to draw the overlay as last one, after resetting the coordinates
        try:
            self.showoverlay.IsChecked = OptionsManager.GetBool("SCR_RotateIFC.showoverlay", True)
        except:
            pass
    
    def SaveOptions(self):
        OptionsManager.SetValue("SCR_RotateIFC.coordpick1_x", self.computeorigin.X)
        OptionsManager.SetValue("SCR_RotateIFC.coordpick1_y", self.computeorigin.Y)
        OptionsManager.SetValue("SCR_RotateIFC.coordpick1_z", self.computeorigin.Z)
        OptionsManager.SetValue("SCR_RotateIFC.coordpick2_x", self.computex.X)
        OptionsManager.SetValue("SCR_RotateIFC.coordpick2_y", self.computex.Y)
        OptionsManager.SetValue("SCR_RotateIFC.coordpick2_z", self.computex.Z)
        OptionsManager.SetValue("SCR_RotateIFC.coordpick3_x", self.computey.X)
        OptionsManager.SetValue("SCR_RotateIFC.coordpick3_y", self.computey.Y)
        OptionsManager.SetValue("SCR_RotateIFC.coordpick3_z", self.computey.Z)
        OptionsManager.SetValue("SCR_RotateIFC.coordpick4_x", self.computez.X)
        OptionsManager.SetValue("SCR_RotateIFC.coordpick4_y", self.computez.Y)
        OptionsManager.SetValue("SCR_RotateIFC.coordpick4_z", self.computez.Z)

        OptionsManager.SetValue("SCR_RotateIFC.squarez", self.squarez.IsChecked)
        OptionsManager.SetValue("SCR_RotateIFC.independentz", self.independentz.IsChecked)
        OptionsManager.SetValue("SCR_RotateIFC.rotateaxes", self.rotateaxes.IsChecked)
        OptionsManager.SetValue("SCR_RotateIFC.showoverlay", self.showoverlay.IsChecked)
        OptionsManager.SetValue("SCR_RotateIFC.xrot", self.xrot.Angle)
        OptionsManager.SetValue("SCR_RotateIFC.yrot", self.yrot.Angle)
        OptionsManager.SetValue("SCR_RotateIFC.zrot", self.zrot.Angle)
        
    def IsValid(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, BIMEntity):
            return True
        if isinstance(o, Shell3D):
            return True

        return False

    def CoordPick1Changed(self, ctrl, e):
        self.computeorigin = self.coordpick1.Coordinate
        self.checkreadyforoverlay()
        return

    def CoordPick2Changed(self, ctrl, e):
        self.computex = self.coordpick2.Coordinate
        self.checkreadyforoverlay()
        return

    def CoordPick3Changed(self, ctrl, e):
        self.computey = self.coordpick3.Coordinate
        self.checkreadyforoverlay()
        return

    def CoordPick4Changed(self, ctrl, e):
        self.computez = self.coordpick4.Coordinate
        self.checkreadyforoverlay()
        return

    def checkreadyforoverlay(self):
        
        if self.squarez.IsChecked:
            if not any (i == Point3D.Zero for i in (self.computeorigin, self.computex, self.computey)):
                if self.computeorigin.Is3D and self.computex.Is3D and self.computey.Is3D:
                    self.drawoverlay()
        elif self.independentz.IsChecked:
            if not any (i == Point3D.Zero for i in (self.computeorigin, self.computex, self.computey, self.computez)):
                if self.computeorigin.Is3D and self.computex.Is3D and self.computey.Is3D and self.computez.Is3D:
                    self.drawoverlay()

    def transform(self, axis, rot):

        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)

        

        try:

            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                if self.objs.Count > 0 or self.savedserials.Count > 0:

                    ps1 = self.computeorigin
                    ps2 = self.computex
                    ps3 = self.computey
                    ps4 = self.computez

                    if ps1.Is3D and ps2.Is3D and ps3.Is3D:

                        xaxis = Vector3D(ps1, ps2)
                        yaxis = Vector3D(ps1, ps3)
                        zaxis = Vector3D(ps1, ps4)

                        #BuildTransformMatrix(Trimble.Vce.Geometry.Vector3D fromPoint, Trimble.Vce.Geometry.Vector3D translation, Trimble.Vce.Geometry.Spinor3D rotation, Trimble.Vce.Geometry.Vector3D scale)
                        #Spinor3D(Trimble.Vce.Geometry.BiVector3D biVector)
                        #BiVector3D(Trimble.Vce.Geometry.Vector3D rotationAxis, double rotationAngle)

                        if axis == 0: # x-axis
                            bivec = BiVector3D(xaxis, rot)
                            if self.rotateaxes.IsChecked:
                                yaxis.Rotate(bivec)
                                self.computey = self.computeorigin + yaxis
                                zaxis.Rotate(bivec)
                                self.computez = self.computeorigin + zaxis
                        elif axis == 1: # y-axis
                            bivec = BiVector3D(yaxis, rot)
                            spinor = Spinor3D(rot, yaxis)
                            if self.rotateaxes.IsChecked:
                                xaxis.Rotate(bivec)
                                self.computex = self.computeorigin + xaxis
                                zaxis.Rotate(bivec)
                                self.computez = self.computeorigin + zaxis
                        elif axis == 2: # z-axis
                            bivec = BiVector3D(zaxis, rot)
                            spinor = Spinor3D(rot, zaxis)
                            if self.rotateaxes.IsChecked:
                                xaxis.Rotate(bivec)
                                self.computex = self.computeorigin + xaxis
                                yaxis.Rotate(bivec)
                                self.computey = self.computeorigin + yaxis
                        
                        spinor = Spinor3D(bivec)
                        #BuildTransformMatrix(Vector3D fromPoint, Vector3D translation, Spinor3D rotation, Vector3D scale)

                        #matrixtozero = Matrix4D.BuildTransformMatrix(ps1, Point3D(0, 0, 0), 0, 1.0, 1.0, 1.0)
                        #matrixbacktops1 = Matrix4D.BuildTransformMatrix(Point3D(0, 0, 0), ps1,  0, 1.0, 1.0, 1.0)

                        targetmatrix = Matrix4D.BuildTransformMatrix(Vector3D(ps1), Vector3D(0, 0, 0), spinor, Vector3D(1, 1, 1))
                        #targetmatrixinverted = Matrix4D.BuildTransformMatrix(Vector3D(ps1), Vector3D(0, 0, 0), spinor.Reverse, Vector3D(1, 1, 1))
                        
                        #targetmatrix = Matrix4D.BuildTransformMatrix(Vector3D(0, 0, 0), Vector3D(0, 0, 0), spinor, Vector3D(1, 1, 1))
                        #targetmatrixinverted = Matrix4D.BuildTransformMatrix(Vector3D(0, 0, 0), Vector3D(0, 0, 0), spinor.Reverse, Vector3D(1, 1, 1))

                        #testmatrix = Matrix4D.BuildTransformMatrix(Point3D(0,0,0), ps1,  0, 1.0, 1.0, 1.0)

                        self.transformobjsn = [] # in case the user selected the BIM object in the project explorer we need to get the recursive BIMEntity

                        if self.objs.Count > 0:
                            for o in self.objs:
                                if isinstance(o, BIMEntity):
                                    self.recursivebimentities(o)
                                else:
                                    self.transformobjsn.Add(o.SerialNumber)
                        
                        elif self.objs.Count == 0 and self.savedserials.Count > 0:
                            for s in self.savedserials:
                                o = self.currentProject.Concordance.Lookup(s)
                                
                                if isinstance(o, BIMEntity):
                                    self.recursivebimentities(o)
                                else:
                                    self.transformobjsn.Add(o.SerialNumber)

                                #tt1 = TransformData(targetmatrixinverted, None)
                                #tt0 = TransformData(matrixtozero, None)
                                #tt1 = TransformData(targetmatrix, None)
                                #tt2 = TransformData(matrixbacktops1, None)

                                #for tt in o:
                                #    self.log.Add(tt.GlobalTransformation)
                                #
                                #filename = os.path.expanduser('~/Downloads/test.csv')
                                #if File.Exists(filename):
                                #    File.Delete(filename)
                                #with open(filename, 'w') as f:       
                                #
                                #    for k in range(0, self.log.Count, 2):
                                #        outputline = str(self.log[k].TranslateX) + ',' + str(self.log[k].TranslateY)+ ',' + str(self.log[k].TranslateZ)
                                #        f.write(outputline + "\n")
                                #
                                #    f.close()

                                #o.Transform(TransformData(matrixtozero, None)) # move object to zero
                                #o.Transform(TransformData(targetmatrix, None))  # apply rotation about 1 axis and origon at 0,0,0
                                #o.Transform(TransformData(targetmatrixinverted, None)) # apply the exact opposite rotation
                                #o.Transform(TransformData(matrixbacktops1, None)) # move object back to where it was
                                #for tt in o:
                                #    self.log.Add(tt.GlobalTransformation)
                                #tt9 = self.log

                        self.transformobjsn = list(set(self.transformobjsn)) # remove duplicates
                        for sn in self.transformobjsn:
                            o = self.currentProject.Concordance[sn]
                            if not isinstance(o, ShellMeshInstance):
                                o.Transform(TransformData(targetmatrix, None))
                            else:
                                pass
                                ##tt = o.GetShellMeshData()
                                ##tt1 = Matrix4D.Multiply(o.GlobalTransformation, o.Transform)
                                ##tt2 = Matrix4D.Multiply(tt1, targetmatrix)
                                ###if Vector3D(o.Transform.Translate, o.GlobalTransformation.Translate).Length == 0:
                                ##o.Transform = tt2
                                ###else:
                                ###    o.Transform = Matrix4D.Multiply(o.GlobalTransformation, targetmatrix)


                        if self.rotateaxes.IsChecked:
                            self.drawoverlay()

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

    def recursivebimentities(self, prev_e):

        for e in prev_e:
        
            if isinstance(e, BIMEntity):
                self.recursivebimentities(e)

            elif isinstance(e, ShellMeshInstance):
                self.transformobjsn.Add(e.GetSite().SerialNumber)

        return

    def Dispose(self, thisCmd, disposing):
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)

    def showoverlayChanged(self, sender, e):
        self.drawoverlay()

    def drawoverlay(self):

        ps1 = self.computeorigin
        ps2 = self.computex
        xaxis = Vector3D(ps1, ps2)
        ps3 = self.computey
        yaxis = Vector3D(ps1, ps3)
        if self.squarez.IsChecked:
            ps4 = self.computez = ps1 + Vector3D.CrossProduct(xaxis, yaxis)[0]
        else:
            ps4 = self.computez
        zaxis = Vector3D(ps1, ps4)

        # ps1 - origin, ps2 - point on x-axis, ps3 - point on y-axis
        
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)

        if self.showoverlay.IsChecked:
            self.overlayBag = OverlayBag(self.ViewOverlay) # create a new/empty overlaybag

            # make the overlay arrows a uniform length - the longest axis that has been defined by the user
            if xaxis.Length > yaxis.Length:
                yaxis.Length = xaxis.Length
                zaxis.Length = xaxis.Length
            else:
                xaxis.Length = yaxis.Length
                zaxis.Length = yaxis.Length

            #self.overlayBag.AddViewCone(ps1, ps2, double length, double radius, int color, byte opacity)
            conelength = xaxis.Length / 5   # make the cone a decent size

            #resdependent = self.overlayBag.ResolutionDependent
            #self.overlayBag.ResolutionDependent = True
            #self.overlayBag.FullCameraDependent = True
            #res = self.overlayBag.Resolution

            self.overlayBag.AddLine(ps1 + xaxis, ps1, 255, 3)
            self.overlayBag.AddViewCone(ps1 + xaxis, ps1, conelength, conelength / 5, 255, Byte.Parse("255"))

            self.overlayBag.AddLine(ps1 + yaxis, ps1, 65280, 3)
            self.overlayBag.AddViewCone(ps1 + yaxis, ps1, conelength, conelength / 5, 65280, Byte.Parse("255"))

            self.overlayBag.AddLine(ps1 + zaxis, ps1, 16711680, 3)
            self.overlayBag.AddViewCone(ps1 + zaxis, ps1, conelength, conelength / 5, 16711680, Byte.Parse("255"))

            # if it only needs to be visible in all Planview then remove the Hoops3DViewGUID
            array = Array[Guid]([DisplayWindow.Hoops3DViewGUID, DisplayWindow.HoopsPlanViewGUID])
            TrimbleOffice.TheOffice.MainWindow.AppViewManager.AddOverlayGeometry(array, self.overlayBag)

        return
        
    def restoreselection_Click(self, sender, e):

        GlobalSelection.Clear()
        GlobalSelection.Items(self.currentProject).Set(self.savedserials)


    def saveselection_Click(self, sender, e):

        self.savedserials.Clear()
        
        for o in GlobalSelection.SelectedMembers(self.currentProject):
            self.savedserials.Add(o.SerialNumber)

        self.savecount.Content = str(self.savedserials.Count)

        GlobalSelection.Clear()

    def ccw_x_Click(self, sender, e):

        self.transform(0, -1 * self.xrot.Angle)

    def cw_x_Click(self, sender, e):

        self.transform(0, self.xrot.Angle)

    def ccw_y_Click(self, sender, e):

        self.transform(1, -1 * self.yrot.Angle)

    def cw_y_Click(self, sender, e):

        self.transform(1, self.yrot.Angle)
    
    def ccw_z_Click(self, sender, e):

        self.transform(2, -1 * self.zrot.Angle)

    def cw_z_Click(self, sender, e):

        self.transform(2, self.zrot.Angle)
       

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.success.Content = ""
        self.error.Content = ""
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)



        pass


            

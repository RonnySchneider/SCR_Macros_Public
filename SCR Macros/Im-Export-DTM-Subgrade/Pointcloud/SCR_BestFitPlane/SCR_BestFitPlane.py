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
    cmdData.Key = "SCR_BestFitPlane"
    cmdData.CommandName = "SCR_BestFitPlane"
    cmdData.Caption = "_SCR_BestFitPlane"
    cmdData.UIForm = "SCR_BestFitPlane"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "Pointcloud"
        cmdData.ShortCaption = "Bestfit Plane"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.04
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "compute a Bestfit Plane into a bunch of points"
        cmdData.ToolTipTextFormatted = "compute a Bestfit Plane into a bunch of points"

    except:
        pass
    try:
        b = Bitmap(macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_BestFitPlane(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader(macroFileFolder + r"\SCR_BestFitPlane.xaml") as s:
            wpf.LoadComponent(self, s)
        self.currentProject = currentProject
        self.macroFileFolder = macroFileFolder

    def HelpClicked(self, cmd, e):
        webbrowser.open("C:\ProgramData\Trimble\MacroCommands3\SCR Macros\MacroHelp\MacroHelp.htm#" + type(self).__name__)


    def OnLoad(self, cmd, buttons, event):
        self.okBtn = buttons[0]
        #buttons[2].Content = "Help"
        #buttons[2].Visibility = Visibility.Visible
        #buttons[2].Click += self.HelpClicked
        self.Caption = cmd.Command.Caption

        self.pointcloudType = clr.GetClrType(PointCloudRegion)
        optionMenu = SelectionContextMenuHandler()
        # remove options that don't apply here
        optionMenu.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        self.objs.ButtonContextMenu = optionMenu
        self.objs.IsEntityValidCallback = self.IsValid

        try:
            self.SetDefaultOptions()
        except:
            pass


    def IsValid(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.pointcloudType):
            return True
        if isinstance(o, CoordPoint) or isinstance(o, CadPoint):
            return True
        return False

    def SetDefaultOptions(self):
        self.drawxyz.IsChecked = OptionsManager.GetBool("SCR_BestFitPlane.drawxyz", False)
        self.createplanesurface.IsChecked = OptionsManager.GetBool("SCR_BestFitPlane.createplanesurface", False)
        self.limitplanesurface.IsChecked = OptionsManager.GetBool("SCR_BestFitPlane.limitplanesurface", True)
        self.createifcface.IsChecked = OptionsManager.GetBool("SCR_BestFitPlane.createifcface", False)
    
    def SaveOptions(self):
        OptionsManager.SetValue("SCR_BestFitPlane.drawxyz", self.drawxyz.IsChecked)
        OptionsManager.SetValue("SCR_BestFitPlane.createplanesurface", self.createplanesurface.IsChecked)
        OptionsManager.SetValue("SCR_BestFitPlane.limitplanesurface", self.limitplanesurface.IsChecked)
        OptionsManager.SetValue("SCR_BestFitPlane.createifcface", self.createifcface.IsChecked)

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand()

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''
        self.success.Content = ''

        wv = self.currentProject[Project.FixedSerial.WorldView]
        wv.PauseGraphicsCache(True)

        #lc = self.currentProject[Project.FixedSerial.LayerContainer]

        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())

        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                selectionSet = GlobalSelection.Items(self.currentProject)

                rwcloudpoints = []
                cloudselectionids = []
                for o in selectionSet:   
                    if isinstance(o, CoordPoint) or isinstance(o, CadPoint):
                        rwcloudpoints.Add(RwPoint3D(o.Position.X, o.Position.Y, o.Position.Z))
                    # compile a list with the ID's of the selected cloud portions
                    if isinstance(o, self.pointcloudType):
                        cloudselectionids.Add(o.Integration.GetSelectedCloudId())
                        cloudintegration = o.Integration.PointCloudDatabase.Integration

                if cloudselectionids.Count > 0:
                    ProgressBar.TBC_ProgressBar.Title = "retrieve selected Cloud-Points"
                    #cps = o.Integration.PointCloudDatabase.Integration.GetPoints(o.Integration.GetSelectedCloudId())
                    # getselectedpoints accepts a list of CloudID's, so we just give it the ID's of the selected portions
                    cps = cloudintegration.GetSelectedPoints(cloudselectionids)
                    self.error.Content += '\nselected Cloud Point: ' + str(cps.Count)
                    ProgressBar.TBC_ProgressBar.Title = "downsampling Cloud-Points"
                    
                    if cps.Count < 15000:
                        for p in cps:
                            rwcloudpoints.Add(RwPoint3D(p.X, p.Y, p.Z))

                    else:
                        # throw that list of all the selected portions onto the sampling algorithm
                        # it returns another ID
                        cpssampleid = cloudintegration.CreateSpatiallySampledCloud(cloudselectionids, 0.025, 15000)
                        # now get the points from that specific subcloud
                        cps2 = cloudintegration.GetPoints(cpssampleid)

                        for p in cps2:
                            rwcloudpoints.Add(RwPoint3D(p.X, p.Y, p.Z))

                    self.error.Content += '\nsampled to #: ' + str(rwcloudpoints.Count)


                ProgressBar.TBC_ProgressBar.Title = "computing best fit plane from " + str(rwcloudpoints.Count) + " points"
                rwplane = RwPlane3D.FitPlaneTo3DPoints(rwcloudpoints)
                centerp = Point3D(rwplane.Point.X, rwplane.Point.Y, rwplane.Point.Z)
                v = Vector3D(rwplane.NormalVector.X, rwplane.NormalVector.Y, rwplane.NormalVector.Z)
                vx = v.Clone()
                vx.RotateAboutZ(math.pi/2)
                vx.Horizon = 0
                vy = vx.Clone()
                vy.Rotate(BiVector3D(v, math.pi/2))

                if self.createplanesurface.IsChecked:
                    # create a helper surface with full triangulation to retrieve the outer boundary rather easy
                    surface = wv.Add(clr.GetClrType(Model3D))
                    surface.Name = Model3D.GetUniqueName("Best-Fit Intermediate", None, wv) #make sure name is unique
                    surface.MaxEdgeLength = 10
                    builder = surface.GetGemBatchBuilder()
                    for p in rwcloudpoints:
                        builder.AddVertex(Point3D(p.X, p.Y, p.Z))
                    
                    builder.Construction()
                    builder.Commit()

                    if self.limitplanesurface.IsChecked:
                        # need to draw the boundary since we need it as object in order to limit the planar surface
                        l = wv.Add(clr.GetClrType(Linestring))
                        l.Append(surface.Boundary, None, False, False)
                        l.Layer = Layer.FindOrCreateLayer(self.currentProject, "Best-Fit Surface Boundary").SerialNumber

                    # now create the average slope surface
                    slope = wv.Add(clr.GetClrType(SlopingLevelSurface))
                    slope.Name = Model3D.GetUniqueName("Best-Fit Slope", None, wv) #make sure name is unique
                    if self.limitplanesurface.IsChecked:
                        slope.BoundarySerialNumber = l.SerialNumber
                    slope.ElevationSlopeType = ElevationSlopeTypes.PointAndDirection
                    slope.Point1 = centerp
                    slope.Point2 = centerp + v
                    slope.Azimuth = v.Azimuth
                    slope.LeftInclination = 0
                    slope.RightInclination = 0
                    slope.Slope = v.Horizon + math.pi/2

                if self.drawxyz.IsChecked:
                    l = wv.Add(clr.GetClrType(Linestring))
                    e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                    e.Position = centerp
                    l.AppendElement(e)
                    e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                    e.Position = centerp + v
                    l.AppendElement(e)       
                    l.Color = Color.Red
                    l.Weight = 100
                    l.Layer = Layer.FindOrCreateLayer(self.currentProject, "Best-Fit Coordinate-System").SerialNumber

                    l = wv.Add(clr.GetClrType(Linestring))
                    e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                    e.Position = centerp
                    l.AppendElement(e)
                    e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                    e.Position = centerp + vx
                    l.AppendElement(e)       
                    l.Color = Color.Blue
                    l.Weight = 100
                    l.Layer = Layer.FindOrCreateLayer(self.currentProject, "Best-Fit Coordinate-System").SerialNumber

                    l = wv.Add(clr.GetClrType(Linestring))
                    e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                    e.Position = centerp
                    l.AppendElement(e)
                    e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                    e.Position = centerp + vy
                    l.AppendElement(e)       
                    l.Color = Color.Green
                    l.Weight = 100
                    l.Layer = Layer.FindOrCreateLayer(self.currentProject, "Best-Fit Coordinate-System").SerialNumber

                failGuard.Commit()
                UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
                self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)

                if self.createplanesurface.IsChecked:
                    osite = wv.Remove(surface.SerialNumber)

        except Exception as e:
            tt = sys.exc_info()
            exc_type, exc_obj, exc_tb = sys.exc_info()
            # EndMark MUST be set no matter what
            # otherwise TBC won't work anymore and needs to be restarted
            self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
            UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
            self.error.Content += '\nan Error occurred - Result probably incomplete\n' + str(exc_type) + '\n' + str(exc_obj) + '\nLine ' + str(exc_tb.tb_lineno)
        
        self.success.Content += '\nDone'
        #ProgressBar.TBC_ProgressBar.Title = ""
        self.SaveOptions()

        wv.PauseGraphicsCache(False)
    


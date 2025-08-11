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
    cmdData.Key = "SCR_EditSideslope"
    cmdData.CommandName = "SCR_EditSideslope"
    cmdData.Caption = "_SCR_EditSideslope"
    cmdData.UIForm = "SCR_EditSideslope"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "Lines"
        cmdData.ShortCaption = "Edit Sideslope"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.07
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "edit i.e. Sideslope Chainage"
        cmdData.ToolTipTextFormatted = "edit i.e. Sideslope Chainage"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_EditSideslope(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_EditSideslope.xaml") as s:
            wpf.LoadComponent (self, s)
        self.currentProject = currentProject
        self.macroFileFolder = macroFileFolder

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


        #self.objs.IsEntityValidCallback=self.IsValid
        #optionMenu = SelectionContextMenuHandler()
        ## remove options that don't apply here
        #optionMenu.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        #self.objs.ButtonContextMenu = optionMenu
        
        self.lType = clr.GetClrType(IPolyseg)

        self.sideslopepicker.IsEntityValidCallback = self.IsValidSideslope
        self.sideslopepicker.ValueChanged += self.sideslopeChanged

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

        self.drawoverlay()

    def SetDefaultOptions(self):
        pass

    def SaveOptions(self):
        pass

    def IsValid(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        return False

    def IsValidSideslope(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, Sideslope):
            return True
        return False

    def sideslopeChanged(self, ctrl, e):
        
        sl = self.sideslopepicker.Entity
        for sle in sl:
            if isinstance(sle, CorridorTemplate):
                sl_temp = sle

        l1 = self.currentProject.Concordance.Lookup(sl.ReferenceGeometrySerial)
        if l1 != None:
            self.startstation.StationProvider = l1
            self.endstation.StationProvider = l1

            self.startstation.Distance = sl_temp.StationDistance
            self.endstation.Distance = sl.EndStation
            self.templatename.Text = sl_temp.Name
            self.nodespacing.Distance = sl_temp.MaximumNodeSpacing
            self.drawoverlay()

    def drawoverlay(self):

        wv = self.currentProject [Project.FixedSerial.WorldView]
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)
        self.overlayBag = OverlayBag(self.ViewOverlay) # create a new/empty overlaybag

        # draw a pink circle around each sideslope
        arcradius = 0.5
        for sl in wv:
            if isinstance(sl, Sideslope):
                for sle in sl:
                    if isinstance(sle, CorridorTemplate):
                        sl_temp = sle
                        sl_station = sl_temp.StationDistance
        
                
                sl_layername = self.currentProject.Concordance.Lookup(sl.Layer).Name
                l = self.currentProject.Concordance.Lookup(sl.ReferenceGeometrySerial)
                polyseg = l.ComputePolySeg()
                polyseg = polyseg.ToWorld()
                polyseg_v = l.ComputeVerticalPolySeg()
                
                p = polyseg.FindPointFromStation(sl_station)[1]
                if polyseg_v != None:
                    p.Z = polyseg_v.ComputeVerticalSlopeAndGrade(sl_station)[1]
                
                self.overlayBag.AddMarker(p, GraphicMarkerTypes.HollowCircle_IndependentColor, Color.Magenta.ToArgb(), "Layer: " + sl_layername, 0, 0, 2.0)

                #arcpoly = PolySeg.PolySeg()
                #arcpoly.Add(PolySeg.Segment.Arc(Point3D(p.X, p.Y + arcradius, p.Z), p, Point3D(p.X, p.Y - arcradius, p.Z), Side.Right))
                #arcpoly.Add(PolySeg.Segment.Arc(Point3D(p.X, p.Y - arcradius, p.Z), p, Point3D(p.X, p.Y + arcradius, p.Z), Side.Right))
                #
                #chord = arcpoly.Linearize(0.01, 0.01, 50, None, False)
                #self.overlayBag.AddPolyline(chord.ToPoint3DArray(), Color.Magenta.ToArgb(), 3)

        #sl = self.sideslopepicker.Entity
        if self.sideslopepicker.Entity:
            l1 = self.currentProject.Concordance.Lookup(self.sideslopepicker.Entity.ReferenceGeometrySerial)

            if l1 != None:
                self.overlayBag.AddPolyline(self.getpolypoints(l1), Color.Blue.ToArgb(), 5)
                
                sl = self.sideslopepicker.Entity
                for sle in sl:
                    if isinstance(sle, CorridorTemplate):
                        sl_temp = sle

                self.overlayBag.AddPolyline(self.getclippedpolypoints(l1, sl_temp.StationDistance, sl.EndStation), Color.Orange.ToArgb(), 2)

                for p in self.getarrowlocations(l1, 10): # returns list with location and perp right azimuth [Point3D, perpVector3D.Value.Azimuth]
                    self.overlayBag.AddMarker(p[0], GraphicMarkerTypes.Arrow_IndependentColor, Color.Orange.ToArgb(), "", 0, math.pi - p[1], 3.0)


        # if it only needs to be visible in all Planview then remove the Hoops3DViewGUID
        array = Array[Guid]([DisplayWindow.Hoops3DViewGUID, DisplayWindow.HoopsPlanViewGUID])
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.AddOverlayGeometry(array, self.overlayBag)

        return

    def getarrowlocations(self, l1, intervals):

        pts = []

        polyseg = l1.ComputePolySeg().Clone()
        polyseg = polyseg.ToWorld()
        polyseg_v = l1.ComputeVerticalPolySeg().Clone()
        
        interval = polyseg.ComputeStationing() / intervals
        
        computestation = interval
        while computestation < polyseg.ComputeStationing():
            outSegment = clr.StrongBox[Segment]()
            out_t = clr.StrongBox[float]()
            outPointOnCL = clr.StrongBox[Point3D]()
            perpVector3D = clr.StrongBox[Vector3D]()
            outdeflectionAngle = clr.StrongBox[float]()
            
            polyseg.FindPointFromStation(computestation, outSegment, out_t, outPointOnCL, perpVector3D, outdeflectionAngle)
            p = outPointOnCL.Value
            if polyseg_v != None:
                p.Z = polyseg_v.ComputeVerticalSlopeAndGrade(computestation)[1]

            pts.Add([p, perpVector3D.Value.Azimuth])

            computestation += interval

        return pts

    def getpolypoints(self, l):

        if l != None:
            polyseg = l.ComputePolySeg().Clone()
            polyseg = polyseg.ToWorld()
            polyseg_v = l.ComputeVerticalPolySeg().Clone()
            if not polyseg_v and not polyseg.AllPointsAre3D:
                polyseg_v = PolySeg.PolySeg()
                polyseg_v.Add(Point3D(polyseg.BeginStation,0,0))
                polyseg_v.Add(Point3D(polyseg.ComputeStationing(), 0, 0))
            chord = polyseg.Linearize(0.001, 0.001, 50, polyseg_v, False)

        return chord.ToPoint3DArray()

    # need to Clone the Polyseg upon creation
    # otherwise something weird is going on
    # and the functions above (getpoly and arrow) may return clipped sections
    # as if the ClipStationRange somehow transpires through to the original Linestring
    # although the Linestring stays unchanged on screen
    # no clue what is going on
    def getclippedpolypoints(self, l, startstation, endstation):

        if l != None:
            
            polyseg = l.ComputePolySeg().Clone()
            polyseg = polyseg.ToWorld()
            if math.isnan(endstation):
                endstation = polyseg.ComputeStationing()
            polyseg.ClipStationRange(startstation, endstation, True)
            polyseg.Trim()
            tt = polyseg.ComputeStationing()
            
            polyseg_v = l.ComputeVerticalPolySeg().Clone()
            if not polyseg_v:
                polyseg_v = PolySeg.PolySeg()
                polyseg_v.Add(Point3D(polyseg.BeginStation,0,0))
                polyseg_v.Add(Point3D(polyseg.ComputeStationing(), 0, 0))

            minmax = CommandHelper.GetMinMaxElevations(l)

            polyseg_v.Clip(Limits3D(Point3D(startstation, minmax[1]-1, 0), Point3D(endstation, minmax[2]+1, 0)), Side.Out)
            polyseg_v.Trim()

            chord = polyseg.Linearize(0.001, 0.001, 50, polyseg_v, False)

        return chord.ToPoint3DArray()

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def Dispose(self, cmd, disposing):
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)


    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''


        self.success.Content += ''

        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        
        wv = self.currentProject [Project.FixedSerial.WorldView]
        lgc = LayerGroupCollection.GetLayerGroupCollection(self.currentProject, False)
                
        wv.PauseGraphicsCache(True)

        self.success.Content=''
        # self.label_benchmark.Content = ''

        # settings = Model3DCompSettings.ProvideSettingsObject(self.currentProject)
        ProgressBar.TBC_ProgressBar.Title = "chording Lines"

        try:

            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                sl = self.sideslopepicker.Entity
                for sle in sl:
                    if isinstance(sle, CorridorTemplate):
                        sl_temp = sle

                sl_temp.StationDistance = self.startstation.Distance
                sl.EndStation = self.endstation.Distance
                sl_temp.Name = self.templatename.Text
                sl_temp.MaximumNodeSpacing = self.nodespacing.Distance

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
        
        self.success.Content += '\nDone'
        ProgressBar.TBC_ProgressBar.Title = ""
        
        wv.PauseGraphicsCache(False)

        self.SaveOptions()
        self.sideslopeChanged(None, None) # trigger the overlay update

    


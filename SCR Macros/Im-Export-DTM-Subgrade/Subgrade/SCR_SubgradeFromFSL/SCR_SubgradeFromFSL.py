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
    cmdData.Key = "SCR_SubgradeFromFSL"
    cmdData.CommandName = "SCR_SubgradeFromFSL"
    cmdData.Caption = "_SCR_SubgradeFromFSL"
    cmdData.UIForm = "SCR_SubgradeFromFSL"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "Subgrade"
        cmdData.ShortCaption = "Subgrade from dropped FSL"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.09
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Create Subgrade by droping a FSL"
        cmdData.ToolTipTextFormatted = "Create Subgrade by droping a FSL"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_SubgradeFromFSL(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_SubgradeFromFSL.xaml") as s:
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
        wv = self.currentProject [Project.FixedSerial.WorldView]
        
        #types = Array [Type] ([CadPoint]) + Array [Type] ([Point3D])    # we fill an array with TBC object types, we could combine different types
        self.linepicker1.IsEntityValidCallback=self.IsValid
        self.linepicker1.ValueChanged += self.lineChanged
        self.startstation.ValueChanged += self.lineChanged
        self.endstation.ValueChanged += self.lineChanged
        self.lType = clr.GetClrType(IPolyseg)

        types = Array [Type] (SurfaceTypeLists.AllWithCutFillMap)    # we fill an array with TBC object types, we could combine different types
        self.designsurfacepicker.FilterByEntityTypes = types    # we fill the dropdownlist by applying that types array as filter
        self.designsurfacepicker.AllowNone = False              # our list shall not show an empty field
        
        self.coordpick1.ValueChanged += self.coordpick1changed

        #self.startstation.MouseMoveInViewEvent += self.updatestation
        #self.coordpick1.ShowGdiCursor += self.DrawCustomCursor    # proof of concept for coordinates at cursor

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass
        self.unitssetup(None, None)

    #def updatestation(self, sender, e):
    #    if not isinstance(sender, clr.GetClrType(I2DProjection)):
    #        return
    #    tt = Point3D(e.Location.X, e.Location.Y, 0)
    #    tt = self.startstation.WindowToWorldSpaceTransform(tt, False, sender.PageToWorld)
    #    sta = self.startstation.FindAnchorPoint(tt)[1] # returns a tuple - 1 is the chainage/station and 2 the offset
    #    self.startstation.Distance = sta    # updating the value shown in the station edit

    #def DrawCustomCursor(self, sender, e):
    #    if not isinstance(e.MessagingView, clr.GetClrType(I2DProjection)):
    #        return
    #    tt = Point3D(e.MousePosition.X, e.MousePosition.Y, 0)
    #    tt = self.coordpick1.WindowToWorldSpaceTransform(tt, False, e.MessagingView.PageToWorld)
    #    ExploreObjectControlHelper.DrawCursorText(e.MessagingView, e.MousePosition, 10, "X: " + str(tt.X) + "\nY: " + str(tt.Y));

    def SetDefaultOptions(self):
        try:    self.designsurfacepicker.SelectIndex(OptionsManager.GetInt("SCR_SubgradeFromFSL.designsurfacepicker", 0))
        except: self.designsurfacepicker.SelectIndex(0)
        self.interval.Distance = OptionsManager.GetDouble("SCR_SubgradeFromFSL.interval", 1.000)
        self.subgradedepth.Distance = OptionsManager.GetDouble("SCR_SubgradeFromFSL.subgradedepth", 0.000)
        self.extendtodtmleft.IsChecked = OptionsManager.GetBool("SCR_SubgradeFromFSL.extendtodtmleft", True)
        self.adddtmprofileleft.IsChecked = OptionsManager.GetBool("SCR_SubgradeFromFSL.adddtmprofileleft", True)
        self.dtmprofilelengthleft.Distance = OptionsManager.GetDouble("SCR_SubgradeFromFSL.dtmprofilelengthleft", 10.000)
        self.extendtodtmright.IsChecked = OptionsManager.GetBool("SCR_SubgradeFromFSL.extendtodtmright", True)
        self.adddtmprofileright.IsChecked = OptionsManager.GetBool("SCR_SubgradeFromFSL.adddtmprofileright", True)
        self.dtmprofilelengthright.Distance = OptionsManager.GetDouble("SCR_SubgradeFromFSL.dtmprofilelengthright", 10.000)
        
        lserial = OptionsManager.GetUint("SCR_SubgradeFromFSL.layerpicker", 8)
        o = self.currentProject.Concordance.Lookup(lserial)
        if o != None:
            if isinstance(o.GetSite(), LayerCollection):    
                self.layerpicker.SetSelectedSerialNumber(lserial, InputMethod(3))
            else:                       
                self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:                       
            self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
    
    def SaveOptions(self):
        try:    # if nothing is selected it would throw an error
            OptionsManager.SetValue("SCR_SubgradeFromFSL.designsurfacepicker", self.designsurfacepicker.SelectedIndex)
        except:
            pass
        OptionsManager.SetValue("SCR_SubgradeFromFSL.interval", self.interval.Distance)
        OptionsManager.SetValue("SCR_SubgradeFromFSL.subgradedepth", self.subgradedepth.Distance)
        OptionsManager.SetValue("SCR_SubgradeFromFSL.extendtodtmleft", self.extendtodtmleft.IsChecked)
        OptionsManager.SetValue("SCR_SubgradeFromFSL.adddtmprofileleft", self.adddtmprofileleft.IsChecked)
        OptionsManager.SetValue("SCR_SubgradeFromFSL.dtmprofilelengthleft", self.dtmprofilelengthleft.Distance)
        OptionsManager.SetValue("SCR_SubgradeFromFSL.extendtodtmright", self.extendtodtmright.IsChecked)
        OptionsManager.SetValue("SCR_SubgradeFromFSL.adddtmprofileright", self.adddtmprofileright.IsChecked)
        OptionsManager.SetValue("SCR_SubgradeFromFSL.dtmprofilelengthright", self.dtmprofilelengthright.Distance)
        OptionsManager.SetValue("SCR_SubgradeFromFSL.layerpicker", self.layerpicker.SelectedSerialNumber)

    def coordpick1changed(self, ctrl, e):
        wv = self.currentProject [Project.FixedSerial.WorldView]
        if self.designsurfacepicker.SelectedSerial!=0:
            surface = wv.Lookup(self.designsurfacepicker.SelectedSerial)
            tt=surface.PickSurface(self.coordpick1.Coordinate)
            if tt[0]==True:
                self.subgradedepth.Distance = self.coordpick1.Coordinate.Z - tt[1].Z
            del surface


    def lineChanged(self, ctrl, e):
        l1 = self.linepicker1.Entity
        if l1 != None:
            self.stationframe.IsEnabled = True
            self.startstation.StationProvider = l1
            self.endstation.StationProvider = l1
            self.offsetleft.StationProvider = l1
            self.offsetright.StationProvider = l1
        else:
            self.stationframe.IsEnabled = False
        self.drawoverlay()

    def drawoverlay(self):

        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)
        self.overlayBag = OverlayBag(self.ViewOverlay) # create a new/empty overlaybag

        l1 = self.linepicker1.Entity

        if l1 != None:
            self.overlayBag.AddPolyline(self.getpolypoints(l1), Color.Green.ToArgb(), 4)

            self.overlayBag.AddPolyline(self.getclippedpolypoints(l1, self.startstation.Distance, self.endstation.Distance), Color.Orange.ToArgb(), 2)

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

    def IsValid(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        return False

    def unitssetup(self, sender, e):
        # setup everything for the unit conversions

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        self.lfp = self.lunits.Properties.Copy() # create a copy in order to set the decimals and enable/disable the suffix
        self.lfp.AddSuffix = False # disable suffix, we need to set it manually, it would always add the current projects units

        self.unitschanged(None, None)
    
    def unitschanged(self, sender, e):

        # loop through all objects of self and set the properties for all DistanceEdits
        # the code is slower than doing it manually for each single one
        # but more convenient since we don't have to worry about how many DistanceEdit Controls we have in the UI
        tt = self.__dict__.items()
        for i in self.__dict__.items():
            if i[1].GetType() == TBCWpf.DistanceEdit().GetType():
                i[1].DisplayUnit = LinearType(self.lunits.DisplayType)
                i[1].ShowControlIcon(False)
                i[1].FormatProperty.AddSuffix = ControlBoolean(1)
                #i[1].FormatProperty.NumberOfDecimals = self.lfp.NumberOfDecimals

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def Dispose(self, cmd, disposing):
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content = ''

        self.success.Content = ''

        wv = self.currentProject [Project.FixedSerial.WorldView]
        #wv.PauseGraphicsCache(True)
       
        inputok=True
        startstation = self.startstation.Distance
        endstation = self.endstation.Distance
        if endstation < startstation: startstation, endstation = endstation, startstation

        interval = abs(self.interval.Distance)
        if interval == 0.0:
            self.error.Content += '\nInterval must not be Zero'
            inputok=False

        offsetleft = self.offsetleft.Offset
        offsetright = self.offsetright.Offset
        if offsetleft > offsetright:
            offsetleft, offsetright = offsetright, offsetleft

        subgradedepth = self.subgradedepth.Distance

        dtmprofilelengthleft = self.dtmprofilelengthleft.Distance
        dtmprofilelengthright = self.dtmprofilelengthright.Distance
        
        l1=self.linepicker1.Entity
        if l1==None: 
            self.error.Content += '\nno Line 1 selected'
            inputok=False

        if inputok:
            self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
            UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
            try:
                with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                    if self.designsurfacepicker.SelectedSerial!=0:
                        profilestation = startstation
                        while profilestation <= endstation:
                            
                            self.drawprofile(profilestation, offsetleft, offsetright, subgradedepth)
                            
                            if profilestation == endstation: break
                            profilestation += interval
                            if profilestation > endstation: profilestation = endstation
    
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
        self.SaveOptions()           

        #wv.PauseGraphicsCache(False)

    
    def drawprofile(self, station, offsetleft, offsetright, depth):
        wv = self.currentProject [Project.FixedSerial.WorldView]

        layer_sn = self.layerpicker.SelectedSerialNumber
        
        surface = wv.Lookup(self.designsurfacepicker.SelectedSerial) # we get our selected surface as object

        l1=self.linepicker1.Entity
        dtmprofilelengthleft = self.dtmprofilelengthleft.Distance
        dtmprofilelengthright = self.dtmprofilelengthright.Distance

        polyseg1=l1.ComputePolySeg()
        polyseg1=polyseg1.ToWorld()
        tt=polyseg1.ComputeStationing()

        # the FindPointFromStation needs the output variables in a certain format 
        outSegment=clr.StrongBox[Segment]()
        out_t=clr.StrongBox[float]()
        outPointOnCL1=clr.StrongBox[Point3D]()
        perpVector3D=clr.StrongBox[Vector3D]()
        outdeflectionAngle=clr.StrongBox[float]()

        # compute profile 
        tt=polyseg1.FindPointFromStation(station,outSegment,out_t,outPointOnCL1,perpVector3D,outdeflectionAngle) # compute point and vector on line 1
        # we have to make sure we have a valid point on the alignment, 
        # otherwise the surface.Profile will corrupt the surface and would need a rebuilt
        if tt==True:
            vector2d=perpVector3D.Clone()   # we create a copy of the perpendicular vector
            vector2d.To2D()                 # we make it 2D
            if offsetright<0: vector2d.Rotate180()  # the vector we've got from above always points to the right, we rotate it if both offsets are negative
            vector2d.Length = abs(offsetright)      # we set the vector length to the same as the first offset
            pointright = outPointOnCL1.Value + vector2d     # we compute the point on the "right" side
            if offsetright>0: vector2d.Rotate180()      # we definetly want the vector to point left now
            vector2d.Length = abs(offsetleft - offsetright)     # we set the vectorlength to the second offset now
            pointleft = pointright + vector2d       # we compute the left side point
            seg = SegmentLine(pointright,pointleft)     # we create a segment out of those two points
            polysegstation=PolySeg.PolySeg(seg)     # and make a PolySeg out of it

            profile = surface.Profile(polysegstation,True)    # we drape the PolySeg onto the surface
            if profile!=None:       # if that worked, and wasn't i.e. outside the surface, we get a profile and we continue taking that one apart
                profile.CompressRedundantNodes(0.0001)
                profile.RepairBacktrackingSegments() # removes also 0-length segments
                l = wv.Add(clr.GetClrType(Linestring))      # we start a new string line
                l.Layer = self.layerpicker.SelectedSerialNumber
                for i in range(0, profile.NumberOfNodes):       # and add all nodes of the profile as new nodes of that string line
                    drawpoint = profile[i].Point    # we get the profile node coordinate into a temp point
                    drawpoint.Z += depth    # we apply the vertical offset
                    e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                    e.Position = drawpoint  # we draw that string line segment
                    l.AppendElement(e)
                    l.Color = Color.Black

                if self.extendtodtmright.IsChecked and profile.NumberOfNodes>1:      # if extend to DTM is checked
                    # tie right
                    tiestart = profile.FirstNode.Point  # we get us the start coordinate of the draped profile, which is still on DTM level
                    tiestart.Z += depth     # we apply the depth again
                    tievector = profile.FirstSegment.Node.Direction     # we get us the bearing of the first segment, points to the left
                    tievector.Rotate180()   # we rotate it around
                    tiepoint = clr.StrongBox[Point3D]()     # we create us the variable the ComputeTie wants for the output

                    # slope in ComputeTie is zenith angle with upwards=0
                    # Vector3D.Horizon is positive above the horizon and negative below
                    if surface.ComputeTie(tiestart, tievector, math.pi/2 + (tievector.Horizon), 100, tiepoint): # we compute the surface intersection
                        # and draw a line to it
                        l = wv.Add(clr.GetClrType(Linestring))
                        l.Layer = self.layerpicker.SelectedSerialNumber
                        l.Color = Color.Red
                        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                        e.Position = tiestart
                        l.AppendElement(e)
                        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                        e.Position = tiepoint.Value
                        l.AppendElement(e)

                        # draw the DTM profile right
                        if self.adddtmprofileright.IsChecked:  
                            v = tievector
                            v.To2D()
                            v.Length = dtmprofilelengthright
                            p4 = tiepoint.Value
                            p5 = p4 + v                   
                            seg = SegmentLine(p4, p5)
                            polysegdtm = PolySeg.PolySeg(seg)
                            addprofile = surface.Profile(polysegdtm, True)
                            l = wv.Add(clr.GetClrType(Linestring))
                            l.Layer = self.layerpicker.SelectedSerialNumber
                            l.Color = Color.Yellow
                            for i in range(0, addprofile.NumberOfNodes):
                                drawpoint = addprofile[i].Point
                                e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                e.Position = drawpoint
                                l.AppendElement(e)
                
                if self.extendtodtmleft.IsChecked and profile.NumberOfNodes>1:      # if extend to DTM is checked
                    # tie left
                    tiestart = profile.LastNode.Point
                    tiestart.Z += depth 
                    tievector = profile.LastSegment.Node.Direction
                    tiepoint = clr.StrongBox[Point3D]()

                    # slope in Computetie is zenith angle with upwards=0
                    # Vector3D.Horizon is positive above the horizon and negative below
                    if surface.ComputeTie(tiestart, tievector, math.pi/2 - (tievector.Horizon), 100, tiepoint):
                        l = wv.Add(clr.GetClrType(Linestring))
                        l.Layer = self.layerpicker.SelectedSerialNumber
                        l.Color = Color.Red
                        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                        e.Position = tiestart
                        l.AppendElement(e)
                        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                        e.Position = tiepoint.Value
                        l.AppendElement(e)

                        # draw the DTM profile left
                        if self.adddtmprofileleft.IsChecked:
                            v = tievector
                            v.To2D()
                            v.Length = dtmprofilelengthleft
                            p4 = tiepoint.Value
                            p5 = p4 + v                   
                            seg = SegmentLine(p4, p5)
                            polysegdtm=PolySeg.PolySeg(seg)
                            addprofile=surface.Profile(polysegdtm,True)
                            l = wv.Add(clr.GetClrType(Linestring))
                            l.Layer = self.layerpicker.SelectedSerialNumber
                            l.Color = Color.Yellow
                            for i in range(0, addprofile.NumberOfNodes):
                                drawpoint = addprofile[i].Point
                                e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                e.Position = drawpoint
                                l.AppendElement(e)

    



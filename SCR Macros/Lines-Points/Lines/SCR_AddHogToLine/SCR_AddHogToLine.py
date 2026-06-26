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

from System.Collections.Generic import List, IEnumerable
exec(open(r"C:\ProgramData\Trimble\MacroCommands3\SCR Macros\SCR_Imports.py").read())

_OPTIONS = {
    "hog": 0.005, "hortol": 0.0001, "vertol": 0.0001, "nodespacing": 2.0,
    "layerpicker": 8,
    "createisopachlinestring": False, "chordisopachlinestring": False,
    "keephzcurves": False, "changeexisting": False,
    "hogarc": False, "halfarc": False, "hogparabola": True, "halfparabola": False,
    "hoglinear": False, "ferguson": False, "changestart": True, "changeend": False,
    "multiline": True, "singleline": False,
    "startstation": 0.0, "endstation": 0.0,
}

def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_AddHogToLine"
    cmdData.CommandName = "SCR_AddHogToLine"
    cmdData.Caption = "_SCR_AddHogToLine"
    cmdData.UIForm = "SCR_AddHogToLine"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "Lines"
        cmdData.ShortCaption = "Weight-Hog"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.19
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "add Weight Hog to Stringlines"
        cmdData.ToolTipTextFormatted = "add Weight Hog to Stringlines"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_AddHogToLine(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_AddHogToLine.xaml") as s:
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

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        #self.lfp = self.lunits.Properties.Copy()
        self.linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation

        self.toleranceheader.Header = 'define Hog-Definition and Chording Tolerance [' + self.linearsuffix + ']'

        self.hog.DistanceType = DistanceType.Z
        #self.hog.NumberOfDecimals = 4
        #self.hortol.NumberOfDecimals = 4
        #self.vertol.NumberOfDecimals = 4
        #self.nodespacing.NumberOfDecimals = 4

        #self.fergusontol.NumberOfDecimals = 4
        #self.fergusontol.MinValue = 0.00000001
        #self.fergusontol.Value = 0.00000001

        self.lType = clr.GetClrType(IPolyseg)
        self.linepicker1.IsEntityValidCallback = self.IsValid
        self.linepicker1.ValueChanged += self.lineChanged
        self.startstation.ValueChanged += self.lineChanged
        self.endstation.ValueChanged += self.lineChanged
        self.objs.IsEntityValidCallback = self.IsValid

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def SetDefaultOptions(self):
        SCROptions.LoadMacroOptions(self, "SCR_AddHogToLine", _OPTIONS, self.currentProject)

    def SaveOptions(self):
        SCROptions.SaveMacroOptions(self, "SCR_AddHogToLine", _OPTIONS)

    def IsValid(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        return False

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def Dispose(self, cmd, disposing):
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)

    def lineChanged(self, ctrl, e):
        l1 = self.linepicker1.Entity
        if l1 != None:
            self.stationframe.IsEnabled = True
            self.startstation.StationProvider = l1
            self.endstation.StationProvider = l1
        else:
            self.stationframe.IsEnabled = False
        self.drawoverlay()

    def drawoverlay(self):

        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)
        self.overlayBag = OverlayBag(self.ViewOverlay) # create a new/empty overlaybag

        l1 = self.linepicker1.Entity

        if l1:
            self.overlayBag.AddPolyline(SCROverlayBag.getpolypoints(l1), Color.Green.ToArgb(), 2)

            self.overlayBag.AddPolyline(SCROverlayBag.getclippedpolypoints(l1, self.startstation.Distance, self.endstation.Distance), Color.Orange.ToArgb(), 4)

            for p in SCROverlayBag.getarrowlocations(l1, 10): # returns list with location and perp right azimuth [Point3D, perpVector3D.Value.Azimuth]
                self.overlayBag.AddMarker(p[0], GraphicMarkerTypes.Arrow_IndependentColor, Color.Orange.ToArgb(), "", 0, math.pi - p[1], 3.0)

        # if it only needs to be visible in all Planview then remove the Hoops3DViewGUID
        array = Array[Guid]([DisplayWindow.Hoops3DViewGUID, DisplayWindow.HoopsPlanViewGUID])
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.AddOverlayGeometry(array, self.overlayBag)

        return

    def multilineChanged(self, sender, e):
        if self.multiline.IsChecked:
            TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)
            self.keephzcurves.IsEnabled = True
            self.changestart.IsEnabled = True
            self.changeend.IsEnabled = True
            self.halfarc.IsEnabled = True
            self.halfparabola.IsEnabled = True
        elif self.singleline.IsChecked:
            self.keephzcurves.IsEnabled = False
            self.changestart.IsEnabled = False
            self.changeend.IsEnabled = False
            self.halfarc.IsEnabled = False
            self.halfparabola.IsEnabled = False
            self.drawoverlay()

    def hoglinearChanged(self, sender, e):
        if self.hoglinear.IsChecked:
            self.changestart.IsEnabled = True
            self.changeend.IsEnabled = True
            #self.fergusontol.IsEnabled = False
        else:
            self.changestart.IsEnabled = False
            self.changeend.IsEnabled = False
            #self.fergusontol.IsEnabled = False

    def fergusonChanged(self, sender, e):
        if self.ferguson.IsChecked:
            self.changestart.IsEnabled = True
            self.changeend.IsEnabled = True
            #self.fergusontol.IsEnabled = True
            #self.linearize.IsEnabled = False
        else:
            self.changestart.IsEnabled = False
            self.changeend.IsEnabled = False
            #self.fergusontol.IsEnabled = False
            #self.linearize.IsEnabled = True

    def halfarcChanged(self, sender, e):
        if (self.hogarc.IsChecked and self.halfarc.IsChecked) or (self.hogparabola.IsChecked and self.halfparabola.IsChecked):
            self.changestart.IsEnabled = True
            self.changeend.IsEnabled = True
            #self.fergusontol.IsEnabled = False
        else:
            self.changestart.IsEnabled = False
            self.changeend.IsEnabled = False
            #self.fergusontol.IsEnabled = False

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''
        self.success.Content = ''

        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        
        wv = self.currentProject [Project.FixedSerial.WorldView]
        lgc = LayerGroupCollection.GetLayerGroupCollection(self.currentProject, False)
                
        wv.PauseGraphicsCache(True)

        # self.label_benchmark.Content = ''

        # settings = Model3DCompSettings.ProvideSettingsObject(self.currentProject)

        try:

            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                
                lineobjects = []

                if self.multiline.IsChecked:
                    for l1 in self.objs:
                        if isinstance(l1, self.lType):
                            lineobjects.Add(l1)
                elif self.singleline.IsChecked:
                    if isinstance(self.linepicker1.Entity, self.lType):
                        lineobjects.Add(self.linepicker1.Entity)

                for l1 in lineobjects:

                    # get and fix line name - Trimble Access doesn't like a hyphen at the start of the name
                    # so we need to make sure we don't actively add one
                    l1name = IName.Name.__get__(l1)
                    if not l1name == '':
                        l1name += ' - '

                    # in case we change the existing line it's better to move all elevation information to the vertical profile/tab
                    if self.changeexisting.IsChecked:
                        polyseg1_v = l1.ComputeVerticalPolySeg()
                        # remove all vertical elements
                        while l1.VerticalElementCount > 0:
                            l1.RemoveVerticalElementAt(0)
                        # remove all elevation information from the horizontal tab
                        for i in range(l1.ElementCount):
                            e = l1.ElementAt(i, True)
                            epos = e.Position
                            epos.To2D()
                            e.Position = epos
                            l1.ReplaceElementAt(e, i)
                            tt2 = e
                        # replace the information on the vertical tab
                        Linestring.ConvertPolySegToLinestringVertical(l1, polyseg1_v, 0)

                    #if isinstance(l1, self.lType):
                    polyseg1 = l1.ComputePolySeg()
                    polyseg1 = polyseg1.ToWorld()
                    polyseg1_v = l1.ComputeVerticalPolySeg()
                    if not polyseg1_v and not polyseg1.AllPointsAre3D:
                        polyseg1_v = PolySeg.PolySeg()
                        polyseg1_v.Add(Point3D(polyseg1.BeginStation,0,0))
                        polyseg1_v.Add(Point3D(polyseg1.ComputeStationing(), 0, 0))
                    
                    t1 = abs(self.hortol.Distance)
                    t2 = abs(self.vertol.Distance)
                    t3 = abs(self.nodespacing.Distance)
                    fergusontol = 0.0001 # abs(self.fergusontol.Value)

                    if self.multiline.IsChecked:
                        startstation = polyseg1.BeginStation
                        endstation = polyseg1.ComputeStationing()
                    elif self.singleline.IsChecked:
                        startstation = self.startstation.Distance
                        endstation = self.endstation.Distance
                        if endstation < startstation: startstation, endstation = endstation, startstation
                    fulllength = endstation - startstation
                    halflength = (endstation - startstation) / 2

                    hog = self.hog.Distance

                    if self.hogarc.IsChecked:
                        # create a vertical polyseg
                        hogpolyseg = PolySeg.PolySeg()

                        # and add the arc geometry
                        if self.singleline.IsChecked or (self.multiline.IsChecked and not self.halfarc.IsChecked):
                            hogradius = (hog**2 + (fulllength/2)**2)/(2*hog)
                            if hog > endstation/2:
                                self.error.Content += '\nArc Solution probably incorrect\nHog is greater than the arc radius'
                                isopachname = "?? incorrect Arc ?? - " + l1name + "Isopach-Arc with " + str(hog*1000) + " mm Hog"
                                linename = "?? incorrect Arc ?? - " + l1name + "Arc with " + str(hog*1000) + " mm Hog"
                            else:
                                isopachname = l1name + "Isopach-Arc with " + str(hog*1000) + " mm Hog"
                                linename = l1name + "Arc with " + str(hog*1000) + " mm Hog"

                            hogpolyseg.Add(ArcSegment(Point3D(startstation, 0), Point3D(endstation, 0), -1*hogradius))
                        else:
                            if hog > endstation/2:
                                self.error.Content += '\nArc Solution probably incorrect\nHog is greater than the arc radius'
                                isopachname = "?? incorrect Arc ?? - " + l1name + "Isopach-Half-Arc with " + str(hog*1000) + " mm Hog"
                                linename = "?? incorrect Arc ?? - " + l1name + "Half-Arc with " + str(hog*1000) + " mm Hog"
                            else:
                                isopachname = l1name + "Isopach-Half-Arc with " + str(hog*1000) + " mm Hog"
                                linename = l1name + "Half-Arc with " + str(hog*1000) + " mm Hog"

                            hogradius = (hog**2 + (fulllength*2/2)**2)/(2*hog)
                            if self.changestart.IsChecked:
                                hogpolyseg.BeginStation = startstation - fulllength
                                hogpolyseg.Add(ArcSegment(Point3D(startstation - fulllength, 0), Point3D(endstation, 0), -1*hogradius))
                            else:
                                hogpolyseg.Add(ArcSegment(Point3D(startstation, 0), Point3D(endstation + fulllength, 0), -1*hogradius))
                                hogpolyseg.BeginStation = startstation

                    if self.hogparabola.IsChecked: # Parabola

                        # hog = a * endstation/2^2
                        # create a vertical polyseg
                        hogpolyseg = PolySeg.PolySeg()
                        # and add the parabola geometry
                        if self.singleline.IsChecked or (self.multiline.IsChecked and not self.halfparabola.IsChecked):
                            a = hog / ((fulllength/2)**2)
                            para_slope = 2 * a * (fulllength/2)
                            para_el = para_slope * (fulllength/2)

                            hogpolyseg.Add(ParabolaSegment(Point3D(startstation, 0), Point3D(startstation + halflength, para_el), Point3D(startstation + fulllength, 0)))
                            isopachname = l1name + "Isopach-Parabola with " + str(hog*1000) + " mm Hog"
                            linename = l1name + "Parabola with " + str(hog*1000) + " mm Hog"
                        else:
                            a = hog / ((fulllength*2/2)**2)
                            para_slope = 2 * a * (fulllength*2/2)
                            para_el = para_slope * (fulllength*2/2)

                            isopachname = l1name + "Isopach-Half-Parabola with " + str(hog*1000) + " mm Hog"
                            linename = l1name + "Half-Parabola with " + str(hog*1000) + " mm Hog"
                            if self.changestart.IsChecked:
                                hogpolyseg.Add(ParabolaSegment(Point3D(startstation - fulllength, 0), Point3D(startstation, para_el), Point3D(fulllength, 0)))
                            else:
                                hogpolyseg.Add(ParabolaSegment(Point3D(startstation, 0), Point3D(fulllength, para_el), Point3D(endstation + fulllength, 0)))

                    if self.hoglinear.IsChecked: # linear
                        # create a vertical polyseg
                        hogpolyseg = PolySeg.PolySeg()
                        # and add the linear geometry
                        if self.multiline.IsChecked:
                            if self.changestart.IsChecked:
                                hogpolyseg.Add(SegmentLine(Point3D(startstation, hog), Point3D(endstation, 0)))
                            else:
                                hogpolyseg.Add(SegmentLine(Point3D(startstation, 0), Point3D(endstation, hog)))
                        elif self.singleline.IsChecked:
                            hogpolyseg.Add(SegmentLine(Point3D(startstation, 0), Point3D(startstation + halflength, hog)))
                            hogpolyseg.Add(SegmentLine(Point3D(startstation + halflength, hog), Point3D(endstation, 0)))


                        isopachname = l1name + "Isopach-linear Line with " + str(hog*1000) + " mm dH"
                        linename = l1name + "linear Line with " + str(hog*1000) + " mm dH"


                    if self.ferguson.IsChecked: # Ferguson-Spline
                        # there is an issue with the discriminate function and small hog values, it won't return a point array
                        # the evaluate function does work
                        # we multiply the hog here and have to divide the resulting spline Y values later again
                        fergusonmulti = 1 # we'll iterate until the discriminate function returns us some values

                        while True:
                            if self.multiline.IsChecked and self.changestart.IsChecked:
                                tangentin = Point3D(startstation - 0.0001, hog * fergusonmulti)
                                startp = Point3D(startstation, hog * fergusonmulti)
                                endp = Point3D(endstation, 0)
                                tangentout = Point3D(endstation + 0.0001, 0)
                            elif self.multiline.IsChecked and self.changeend.IsChecked:
                                tangentin = Point3D(startstation - 0.0001, 0)
                                startp = Point3D(startstation, 0)
                                endp = Point3D(endstation, hog * fergusonmulti)
                                tangentout = Point3D(endstation + 0.0001, hog * fergusonmulti)
                            elif self.singleline.IsChecked:
                                tangentin = Point3D(0 - 0.0001, 0)
                                startp = Point3D(0, 0)
                                endp = Point3D(halflength, hog * fergusonmulti)
                                tangentout = Point3D(halflength + 0.0001, hog * fergusonmulti)

                            startend = Array[Point3D]([startp, endp])
                            ferguson = FergusonSpline()
                            ferguson.FitPoints(startend, tangentin, tangentout)

                            #tt = ferguson.Evaluate(0.3)
                            fergusonpts = ferguson.Discriminate(0, 1, fergusontol)
                            if fergusonpts.Count > 0 or fergusonmulti == 1000000:
                                break

                            fergusonmulti *= 10

                        # now we have to scale the spline back
                        if fergusonpts.Count > 0:    
                            if self.multiline.IsChecked:
                                fixedfergusonpts = Array[Point3D]([Point3D()] * fergusonpts.Count)
                                for i in range(fergusonpts.Count):
                                    pt = fergusonpts[i]
                                    fixedfergusonpts[i] = (Point3D(pt.X, pt.Y / fergusonmulti))
                            elif self.singleline.IsChecked:
                                fixedfergusonpts = Array[Point3D]([Point3D()] * (fergusonpts.Count * 2))
                                c = 0
                                for i in range(fergusonpts.Count):
                                    pt = fergusonpts[i]
                                    fixedfergusonpts[c] = (Point3D(startstation + pt.X, pt.Y / fergusonmulti))
                                    c += 1
                                for i in reversed(range(fergusonpts.Count)):
                                    pt = fergusonpts[i]
                                    fixedfergusonpts[c] = (Point3D(endstation - pt.X, pt.Y / fergusonmulti))
                                    c += 1

                            # create a vertical polyseg
                            hogpolyseg = PolySeg.PolySeg()
                            # and add the linear geometry
                            hogpolyseg.Add(fixedfergusonpts)

                            isopachname = l1name + "Isopach-Ferguson-Spline with " + str(hog*1000) + " mm dH"
                            linename = l1name + "Ferguson-Spline with " + str(hog*1000) + " mm dH"
                        else:
                            self.error.Content += '\nHog-Value seems to be too small for Ferguson-Spline'
                            hogpolyseg = None

                    if hogpolyseg != None:
                        hogpolyseg.ComputeStationing()
                        
                        # chord the hog geometry                        
                        chordedhog = hogpolyseg.Linearize(t1, t2, t3, None, False)

                        # create the new line in the worldview and draw the segments
                        if self.changeexisting.IsChecked:
                            l_with_hog = l1
                        else:
                            l_with_hog = wv.Add(clr.GetClrType(Linestring))
                            l_with_hog.Name = linename
                            l_with_hog.Layer = self.layerpicker.SelectedSerialNumber
                        

                        nodes = chordedhog.ToPoint3DArray()
                        finalhogel = Array[Point3D]([Point3D()] * nodes.Count)
                        # go through the chord nodes
                        # get the original position and elevation, add the hog value and draw it
                        
                        tt = []
                        for i in range(0, nodes.Count): # node list of linearized profile with X as chainage and Y as elevation
                            node = nodes[i]

                            p1 = polyseg1.FindPointFromStation(node.X)[1]
                            if polyseg1_v != None:
                                p1.Z = polyseg1_v.ComputeVerticalSlopeAndGrade(node.X)[1]

                            tt.Add(p1)

                            if p1.Z < 0.001:
                                tt2 = 1

                            if self.keephzcurves.IsChecked or self.singleline.IsChecked:
                                p2 = Point3D(node.X, p1.Z + node.Y, 0)
                                finalhogel[i] = p2
                            # in case of multiline or chorded output we can add to the final elements straight away
                            else:
                                p2 = Point3D(p1)
                                p2.Z += node.Y

                                e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                e.Position = p2  # we draw that string line segment
                                l_with_hog.AppendElement(e)

                        if self.keephzcurves.IsChecked or self.singleline.IsChecked:
                            
                            # create the final hog polyseg (original + hog)
                            finalhogpolyseg = PolySeg.PolySeg()
                            # and add the profile geometry
                            finalhogpolyseg.Add(finalhogel)
                            
                            if self.singleline.IsChecked:
                                # need to prepare and add the original vertical polysegs before and after
                                startseg_v = polyseg1_v.Clone()
                                tt = startseg_v.Clip(Limits3D(Point3D(0, -10000), Point3D(startstation, 10000)), Side.Out)
                                startseg_v.Trim()
                                #startseg_v.ComputeStationing()
                                endseg_v = polyseg1_v.Clone()
                                #tt = polyseg1.ComputeStationing()
                                tt = endseg_v.Clip(Limits3D(Point3D(endstation, -10000), Point3D(endseg_v.ComputeStationing(), 10000)), Side.Out)
                                endseg_v.Trim()
                                #tt = endseg_v.ComputeStationing()
                                
                                tt = startseg_v.Join(finalhogpolyseg)
                                tt = startseg_v.Join(endseg_v)
                                finalhogpolyseg = startseg_v.Clone()

                            if self.changeexisting.IsChecked:
                                # remove existing vertical information from the linestring
                                while l1.VerticalElementCount > 0:
                                    l1.RemoveVerticalElementAt(0)
                                # replace it with the hogged polyseg
                                Linestring.ConvertPolySegToLinestringVertical(l1, finalhogpolyseg, 0)
                            else:
                                l_with_hog.Append(polyseg1, finalhogpolyseg, False, False)

                        
                        # prepare and create the Isopach    
                        if self.createisopachlinestring.IsChecked:
                            
                            l_hog_iso = wv.Add(clr.GetClrType(Linestring))
                            
                            if self.singleline.IsChecked:
                                # need to add start and end to the isopach curve
                                startseg_v = PolySeg.PolySeg().Add(SegmentLine(Point3D(0, 0), Point3D(startstation, 0)))
                                tt = polyseg1.ComputeStationing()
                                endseg_v = PolySeg.PolySeg().Add(SegmentLine(Point3D(endstation, 0), Point3D(polyseg1.ComputeStationing(), 0)))
                                tt = startseg_v.Clone()
                                tt.Join(chordedhog.Clone())
                                tt.Join(endseg_v.Clone())
                                chordedhog = tt.Clone()

                                tt = startseg_v.Clone()
                                tt.Join(hogpolyseg.Clone())
                                tt.Join(endseg_v.Clone())
                                hogpolyseg = tt.Clone()
                            
                            if self.chordisopachlinestring.IsChecked:
                                l_hog_iso.Append(polyseg1, chordedhog, False, False)
                            else:
                                l_hog_iso.Append(polyseg1, hogpolyseg, False, False)
                            l_hog_iso.Name = isopachname
                            l_hog_iso.Layer = self.layerpicker.SelectedSerialNumber

                    else:
                        self.error.Content += '\nskipped invalid Objects'
                failGuard.Commit()
                self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
                UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
        
        except Exception as e:
            tt = sys.exc_info()
            exc_type, exc_obj, exc_tb = sys.exc_info()
            # EndMark MUST be set no matter what
            # otherwise TBC won't work anymore and needs to be restarted
            self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
            UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
            self.error.Content += '\nan Error occurred - Result probably incomplete\n' + str(exc_type) + '\n' + str(exc_obj) + '\nLine ' + str(exc_tb.tb_lineno)

        
        self.success.Content += '\nDone'
        
        self.drawoverlay()
        wv.PauseGraphicsCache(False)

        #Keyboard.Focus(self.linepicker1)
        Keyboard.Focus(self.objs)

        self.SaveOptions()

    


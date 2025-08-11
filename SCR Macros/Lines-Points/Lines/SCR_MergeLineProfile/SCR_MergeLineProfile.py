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
    cmdData.Key = "SCR_MergeLineProfile"
    cmdData.CommandName = "SCR_MergeLineProfile"
    cmdData.Caption = "_SCR_MergeLineProfile"
    cmdData.UIForm = "SCR_MergeLineProfile"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "Lines"
        cmdData.ShortCaption = "Merge Line+Profile"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.14
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "merge a line and a profile "
        cmdData.ToolTipTextFormatted = "merge the 2D-Compenent of Line 1 with Profile of Line 2"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_MergeLineProfile(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_MergeLineProfile.xaml") as s:
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

        self.lType = clr.GetClrType(IPolyseg)
        self.linepicker1.IsEntityValidCallback = self.IsValid
        self.linepicker2.IsEntityValidCallback = self.IsValid
        self.linepicker3.IsEntityValidCallback = self.IsValid
        self.linepicker4.IsEntityValidCallback = self.IsValid
        self.linepicker1.ValueChanged += self.lineChanged
        self.linepicker2.ValueChanged += self.lineChanged
        self.linepicker3.ValueChanged += self.lineChanged
        self.linepicker4.ValueChanged += self.lineChanged


        self.combinescalehz.NumberOfDecimals = 4
        self.combinescalehz.MinValue = 0.00000001
        self.combinescalevz.NumberOfDecimals = 4
        self.combinescalevz.MinValue = 0.00000001
        #self.combinechordtol.NumberOfDecimals = 4
        #self.combinenodespacing.NumberOfDecimals = 4
        
        self.drawscalehz.NumberOfDecimals = 4
        self.drawscalehz.MinValue = 0.00000001
        self.drawscalevz.NumberOfDecimals = 4
        self.drawscalevz.MinValue = 0.00000001
        #self.drawchordtol.NumberOfDecimals = 4
        #self.drawnodespacing.NumberOfDecimals = 4

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        #self.lfp = self.lunits.Properties.Copy()
        linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation
        #self.lfp.AddSuffix = False
        self.combinechordtollabel.Content = "chording tolerance [" + linearsuffix + "]"
        self.combinenodespacinglabel.Content = "max Node Spacing [" + linearsuffix + "]"
        self.drawchordtollabel.Content = "chording tolerance [" + linearsuffix + "]"
        self.drawnodespacinglabel.Content = "max Node Spacing [" + linearsuffix + "]"

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def SetDefaultOptions(self):
        lserial = OptionsManager.GetUint("SCR_MergeLineProfile.layerpicker", 8)
        o = self.currentProject.Concordance.Lookup(lserial) # get the object withj the saved serial number
        if o != None:   # check if we actually got an object back
            if isinstance(o.GetSite(), LayerCollection):    # check if the object is a layer
                self.layerpicker.SetSelectedSerialNumber(lserial, InputMethod(3))
            else:  # if not then set Layer zero                       
                self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:   # if not then set Layer zero               
            self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        
        self.combinemode.IsChecked = OptionsManager.GetBool("SCR_MergeLineProfile.combinemode", True)
        self.reversel2.IsChecked = OptionsManager.GetBool("SCR_MergeLineProfile.reversel2", False)
        self.l2is2d.IsChecked = OptionsManager.GetBool("SCR_MergeLineProfile.l2is2d", False)
        self.combinescalehz.Value = OptionsManager.GetDouble("SCR_MergeLineProfile.combinescalehz", 1.0)
        self.combinescalevz.Value = OptionsManager.GetDouble("SCR_MergeLineProfile.combinescalevz", 1.0)
        self.combinechordtol.Distance = OptionsManager.GetDouble("SCR_MergeLineProfile.combinechordtol", 0.0001)
        self.combinenodespacing.Distance = OptionsManager.GetDouble("SCR_MergeLineProfile.combinenodespacing", 0.2)

        self.drawmode.IsChecked = OptionsManager.GetBool("SCR_MergeLineProfile.drawmode", False)
        self.drawscalehz.Value = OptionsManager.GetDouble("SCR_MergeLineProfile.drawscalehz", 1.0)
        self.drawscalevz.Value = OptionsManager.GetDouble("SCR_MergeLineProfile.drawscalevz", 1.0)
        self.drawchordtol.Distance = OptionsManager.GetDouble("SCR_MergeLineProfile.drawchordtol", 0.0001)
        self.drawnodespacing.Distance = OptionsManager.GetDouble("SCR_MergeLineProfile.drawnodespacing", 0.2)


        self.smoothmode.IsChecked = OptionsManager.GetBool("SCR_MergeLineProfile.smoothmode", False)

        self.straightconvertmode.IsChecked = OptionsManager.GetBool("SCR_MergeLineProfile.straightconvertmode", False)
        self.straightconverttonew.IsChecked = OptionsManager.GetBool("SCR_MergeLineProfile.straightconverttonew", True)
        self.straightconvertexisting.IsChecked = OptionsManager.GetBool("SCR_MergeLineProfile.straightconvertexisting", False)

        self.smoothconvertmode.IsChecked = OptionsManager.GetBool("SCR_MergeLineProfile.smoothconvertmode", False)
        self.smoothconverttonew.IsChecked = OptionsManager.GetBool("SCR_MergeLineProfile.smoothconverttonew", True)
        self.smoothconvertexisting.IsChecked = OptionsManager.GetBool("SCR_MergeLineProfile.smoothconvertexisting", False)

        self.smoothparabolamode.IsChecked = OptionsManager.GetBool("SCR_MergeLineProfile.smoothparabolamode", True)
        self.smoothhor.IsChecked = OptionsManager.GetBool("SCR_MergeLineProfile.smoothhor", False)
        self.smoothver.IsChecked = OptionsManager.GetBool("SCR_MergeLineProfile.smoothver", False)
        self.smoothhorver.IsChecked = OptionsManager.GetBool("SCR_MergeLineProfile.smoothhorver", True)
        self.smoothchordtol.Distance = OptionsManager.GetDouble("SCR_MergeLineProfile.smoothchordtol", 0.0001)
        self.smoothnodespacing.Distance = OptionsManager.GetDouble("SCR_MergeLineProfile.smoothnodespacing", 0.2)

    def SaveOptions(self):

        OptionsManager.SetValue("SCR_MergeLineProfile.layerpicker", self.layerpicker.SelectedSerialNumber)
        OptionsManager.SetValue("SCR_MergeLineProfile.combinemode", self.combinemode.IsChecked)
        OptionsManager.SetValue("SCR_MergeLineProfile.reversel2", self.reversel2.IsChecked)
        OptionsManager.SetValue("SCR_MergeLineProfile.l2is2d", self.l2is2d.IsChecked)
        OptionsManager.SetValue("SCR_MergeLineProfile.combinescalehz", self.combinescalehz.Value)
        OptionsManager.SetValue("SCR_MergeLineProfile.combinescalevz", self.combinescalevz.Value)
        OptionsManager.SetValue("SCR_MergeLineProfile.combinechordtol", self.combinechordtol.Distance)
        OptionsManager.SetValue("SCR_MergeLineProfile.combinenodespacing", self.combinenodespacing.Distance)

        OptionsManager.SetValue("SCR_MergeLineProfile.drawmode", self.drawmode.IsChecked)
        OptionsManager.SetValue("SCR_MergeLineProfile.drawscalehz", self.drawscalehz.Value)
        OptionsManager.SetValue("SCR_MergeLineProfile.drawscalevz", self.drawscalevz.Value)
        OptionsManager.SetValue("SCR_MergeLineProfile.drawchordtol", self.drawchordtol.Distance)
        OptionsManager.SetValue("SCR_MergeLineProfile.drawnodespacing", self.drawnodespacing.Distance)

        OptionsManager.SetValue("SCR_MergeLineProfile.smoothmode", self.smoothmode.IsChecked)

        OptionsManager.SetValue("SCR_MergeLineProfile.straightconvertmode", self.straightconvertmode.IsChecked)
        OptionsManager.SetValue("SCR_MergeLineProfile.straightconverttonew", self.straightconverttonew.IsChecked)
        OptionsManager.SetValue("SCR_MergeLineProfile.straightconvertexisting", self.straightconvertexisting.IsChecked)

        OptionsManager.SetValue("SCR_MergeLineProfile.smoothconvertmode", self.smoothconvertmode.IsChecked)
        OptionsManager.SetValue("SCR_MergeLineProfile.smoothconverttonew", self.smoothconverttonew.IsChecked)
        OptionsManager.SetValue("SCR_MergeLineProfile.smoothconvertexisting", self.smoothconvertexisting.IsChecked)

        OptionsManager.SetValue("SCR_MergeLineProfile.smoothparabolamode", self.smoothparabolamode.IsChecked)
        OptionsManager.SetValue("SCR_MergeLineProfile.smoothhor", self.smoothhor.IsChecked)
        OptionsManager.SetValue("SCR_MergeLineProfile.smoothver", self.smoothver.IsChecked)
        OptionsManager.SetValue("SCR_MergeLineProfile.smoothhorver", self.smoothhorver.IsChecked)

        OptionsManager.SetValue("SCR_MergeLineProfile.smoothchordtol", self.smoothchordtol.Distance)
        OptionsManager.SetValue("SCR_MergeLineProfile.smoothnodespacing", self.smoothnodespacing.Distance)

    def IsValid(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        return False

    def lineChanged(self, ctrl, e):
        
        self.drawoverlay()

    def drawoverlay(self):

        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)
        self.overlayBag = OverlayBag(self.ViewOverlay) # create a new/empty overlaybag

        l1 = self.linepicker1.Entity
        l2 = self.linepicker2.Entity
        l3 = self.linepicker3.Entity
        l4 = self.linepicker4.Entity

        if self.combinemode.IsChecked:
            if l1 != None:
                self.overlayBag.AddPolyline(self.getpolypoints(l1), Color.Green.ToArgb(), 4)
                for p in self.getarrowlocations(l1, 10): # returns list with location and perp right azimuth [Point3D, perpVector3D.Value.Azimuth]
                    self.overlayBag.AddMarker(p[0], GraphicMarkerTypes.Arrow_IndependentColor, Color.Orange.ToArgb(), "", 0, math.pi - p[1], 3.0)

            if l2 != None:
                self.overlayBag.AddPolyline(self.getpolypoints(l2), Color.Blue.ToArgb(), 4)
                for p in self.getarrowlocations(l2, 10): # returns list with location and perp right azimuth [Point3D, perpVector3D.Value.Azimuth]
                    self.overlayBag.AddMarker(p[0], GraphicMarkerTypes.Arrow_IndependentColor, Color.Orange.ToArgb(), "", 0, math.pi - p[1], 3.0)
        elif self.drawmode.IsChecked:
            if l3 != None:
                self.overlayBag.AddPolyline(self.getpolypoints(l3), Color.Red.ToArgb(), 4)
                for p in self.getarrowlocations(l3, 10): # returns list with location and perp right azimuth [Point3D, perpVector3D.Value.Azimuth]
                    self.overlayBag.AddMarker(p[0], GraphicMarkerTypes.Arrow_IndependentColor, Color.Orange.ToArgb(), "", 0, math.pi - p[1], 3.0)
        elif self.smoothmode.IsChecked:
            if l4 != None:
                self.overlayBag.AddPolyline(self.getpolypoints(l4), Color.Magenta.ToArgb(), 4)
                for p in self.getarrowlocations(l4, 10): # returns list with location and perp right azimuth [Point3D, perpVector3D.Value.Azimuth]
                    self.overlayBag.AddMarker(p[0], GraphicMarkerTypes.Arrow_IndependentColor, Color.Orange.ToArgb(), "", 0, math.pi - p[1], 3.0)

        # if it only needs to be visible in all Planview then remove the Hoops3DViewGUID
        array = Array[Guid]([DisplayWindow.Hoops3DViewGUID, DisplayWindow.HoopsPlanViewGUID])
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.AddOverlayGeometry(array, self.overlayBag)

        return

    def getarrowlocations(self, l1, intervals):

        pts = []

        polyseg = l1.ComputePolySeg()
        polyseg = polyseg.ToWorld()
        polyseg_v = l1.ComputeVerticalPolySeg()
        
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
            polyseg = l.ComputePolySeg()
            polyseg = polyseg.ToWorld()
            polyseg_v = l.ComputeVerticalPolySeg()
            if not polyseg_v and not polyseg.AllPointsAre3D:
                polyseg_v = PolySeg.PolySeg()
                polyseg_v.Add(Point3D(polyseg.BeginStation,0,0))
                polyseg_v.Add(Point3D(polyseg.ComputeStationing(), 0, 0))
            chord = polyseg.Linearize(0.001, 0.001, 50, polyseg_v, False)

        return chord.ToPoint3DArray()

    def drawprofile(self):

        wv = self.currentProject [Project.FixedSerial.WorldView]

        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)

        try:

            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                # read the values from the GUI
                drawscalehz = self.drawscalehz.Value
                drawscalevz = self.drawscalevz.Value
                if drawscalehz > drawscalevz:
                    drawchordtol = abs(self.drawchordtol.Distance) / drawscalehz
                    drawnodespacing = abs(self.drawnodespacing.Distance) / drawscalehz
                else:
                    drawchordtol = abs(self.drawchordtol.Distance) / drawscalevz
                    drawnodespacing = abs(self.drawnodespacing.Distance) / drawscalevz


                # get the draw origin into a Point3D object
                draworigin = self.drawcoordpick.Coordinate
                draworigin.Z = 0
                
                # get the line as object
                l3 = self.linepicker3.Entity
                if l3 != None:
            
                    polyseg_v = None
                    # compute the vertical polyseg
                    # need to use clone here, if i want to run the macro multiple times
                    # I have no explanation, but without it, polyseg_v would be at the wrong place
                    # somehow the translate from the previous run would transpire through
                    # even though l3 and polyseg_v start as uninitialized
                    polyseg_v = l3.ComputeVerticalPolySeg()

                    if polyseg_v != None:
                        polyseg_v = polyseg_v.Clone()
                        newname = 'Profile - ' + IName.Name.__get__(l3) 

                        
                        if self.reversel3.IsChecked:
                            polyseg_v.Mirror() # we need to mirror it
                            polyseg_v.Reverse() # reversing the node order is not really necessary
                            polyseg_v.ShiftVerticalPolyseg(polyseg_v.BoundingBox.ptMax.X) # after mirroring the profile is left of the y-axis, need to shift it back

                        # if the vertical profile contains parabolas or spirals or the like we must chord them since they can't be converted/drawn as a horizontal polyseg
                        s = polyseg_v.FirstSegment
                        mustchord = False
                        while s is not None:
                            if s.Visible:
                                if not (isinstance(s, PolySeg.Segment.Line) or isinstance(s, PolySeg.Segment.Arc)):
                                    mustchord = True
                            s = polyseg_v.Next(s)
                            
                        # if we want to scale we need to chord the line first, otherwise we'd lose information about arcs etc.
                        if mustchord or drawscalehz != 1.0 or drawscalevz != 1.0:
                            polyseg_v = polyseg_v.Linearize(drawchordtol, drawchordtol, drawnodespacing, None, False)
                        polyseg_v.Scale(Vector3D(drawscalehz, drawscalevz, 0))

                        # get the limits at 0/0, save a difference computation later
                        limits = polyseg_v.BoundingBox.ptMax
                        
                        # shift the polyseg from 0/0 to the selected draw-origin
                        polyseg_v.Translate(polyseg_v.FirstNode, polyseg_v.LastNode, Vector3D(Point3D(0, 0, 0), draworigin))
                        
                        # draw the profile linestring
                        l = wv.Add(clr.GetClrType(Linestring))
                        l.Append(polyseg_v, None, False, False)
                        l.Name = newname
                        l.Layer = self.layerpicker.SelectedSerialNumber

                        # draw x line
                        l = wv.Add(clr.GetClrType(Linestring))
                        l.Name = 'X-Axis'
                        l.Color = Color.LemonChiffon
                        l.Layer = self.layerpicker.SelectedSerialNumber
                        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                        e.Position = draworigin
                        l.AppendElement(e)
                        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                        e.Position = draworigin + Vector3D(limits.X, 0, 0)
                        l.AppendElement(e)

                        # draw y line
                        l = wv.Add(clr.GetClrType(Linestring))
                        l.Name = 'Y-Axis'
                        l.Color = Color.LemonChiffon
                        l.Layer = self.layerpicker.SelectedSerialNumber
                        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                        e.Position = draworigin
                        l.AppendElement(e)
                        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                        e.Position = draworigin + Vector3D(0, limits.Y, 0)
                        l.AppendElement(e)

                    else:
                        self.error.Text += "\nLine doesn't seem to have valid vertical information"

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
            self.error.Text += '\nan Error occurred - Result probably incomplete\n' + str(exc_type) + '\n' + str(exc_obj) + '\nLine ' + str(exc_tb.tb_lineno)

    def combine(self):

        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        
        wv = self.currentProject [Project.FixedSerial.WorldView]
                
        #wv.PauseGraphicsCache(True)

        try:

            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                
                # get the lines as object
                l1 = self.linepicker1.Entity
                l2 = self.linepicker2.Entity

                # get the 2d profile origin as Point3D object
                combineorigin = self.combinecoordpick.Coordinate
                combineorigin.Z = 0

                # get the scale values from the GUI
                combinescalehz = self.combinescalehz.Value
                combinescalevz = self.combinescalevz.Value
                if combinescalehz > combinescalevz:
                    combinechordtol = abs(self.combinechordtol.Distance) / combinescalehz
                    combinenodespacing = abs(self.combinenodespacing.Distance) / combinescalehz
                else:
                    combinechordtol = abs(self.combinechordtol.Distance) / combinescalevz
                    combinenodespacing = abs(self.combinenodespacing.Distance) / combinescalevz

                if l1 != None and l2 != None:

                    newpolyseg = l1.ComputePolySeg()
                    newpolyseg = newpolyseg.ToWorld()

                    if self.l2is2d.IsChecked:
                        # use the horizontal information as data for the new vertical polyseg
                        newpolyseg_v = l2.ComputePolySeg()
                        newpolyseg_v = newpolyseg_v.ToWorld()
                        # shift polyseg to 0/0
                        newpolyseg_v.Translate(newpolyseg_v.FirstNode, newpolyseg_v.LastNode, Vector3D(combineorigin, Point3D(0, 0, 0)))
                    else:
                        # compute the vertical polyseg that is automatically at 0/0
                        newpolyseg_v = l2.ComputeVerticalPolySeg()

                    if newpolyseg_v != None:
                        newpolyseg_v = newpolyseg_v.Clone()
                        # if the scale value makes it necessary we need to chord the profile first
                        if combinescalehz != 1.0 or combinescalevz != 1.0:
                            newpolyseg_v = newpolyseg_v.Linearize(combinechordtol, combinechordtol, combinenodespacing, None, False)
                        # apply scale
                        newpolyseg_v.Scale(Vector3D(combinescalehz, combinescalevz, 0))

                        newname = '2D - ' + IName.Name.__get__(l1) + ' and Profile - ' + IName.Name.__get__(l2) 

                        if self.reversel2.IsChecked:
                            newpolyseg_v.Mirror()
                            newpolyseg_v.Reverse()
                            newpolyseg_v.ShiftVerticalPolyseg(newpolyseg_v.BoundingBox.ptMax.X)
                        
                        l = wv.Add(clr.GetClrType(Linestring))
                        l.Append(newpolyseg, newpolyseg_v, False, False)
                        l.Name = newname
                        l.Layer = self.layerpicker.SelectedSerialNumber
                    else:
                        self.error.Text += "\nLine 2 doesn't seem to have valid vertical information"

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
            self.error.Text += '\nan Error occurred - Result probably incomplete\n' + str(exc_type) + '\n' + str(exc_obj) + '\nLine ' + str(exc_tb.tb_lineno)

        
        self.success.Content += '\nDone'
        
        #wv.PauseGraphicsCache(False)

        Keyboard.Focus(self.linepicker1)
        #Keyboard.Focus(self.objs)

    def smoothconvertthepolyseg(self, o):

        if o:
            if self.smoothconverttonew.IsChecked:
                wv = self.currentProject [Project.FixedSerial.WorldView]
                l = wv.Add(clr.GetClrType(Linestring))
                l.Name = IName.Name.__get__(o) + ' - smoothed'
                l.Layer = self.layerpicker.SelectedSerialNumber
            
            # go through the elements of the linestring
            for i in range(0, o.ElementCount):
                ProgressBar.TBC_ProgressBar.Title = "Node " + str(i + 1) + '/' + str(o.ElementCount)
                if ProgressBar.TBC_ProgressBar.SetProgress(math.floor((i + 1) * 100 / o.ElementCount)):
                    break
                
                e = o.ElementAt(i)
                etype = e.GetType()

                if "Straight" in etype.Name:
                    if "PointId" in etype.Name:
                        enew = ElementFactory.Create(clr.GetClrType(ISmoothCurveSegment), clr.GetClrType(IPointIdLocation))
                        enew.LocationSerialNumber = e.LocationSerialNumber
                    else:
                        enew = ElementFactory.Create(clr.GetClrType(ISmoothCurveSegment), clr.GetClrType(IXYZLocation))
                        enew.Position = e.Position
                else:
                    enew = e

                if self.smoothconverttonew.IsChecked:
                    l.AppendElement(enew)
                elif self.smoothconvertexisting.IsChecked:
                    o.ReplaceElementAt(enew, i)
            self.drawoverlay()
        else:
            pass

    def straightconvertthepolyseg(self, o):

        if o:
            if self.straightconverttonew.IsChecked:
                wv = self.currentProject [Project.FixedSerial.WorldView]
                l = wv.Add(clr.GetClrType(Linestring))
                l.Name = IName.Name.__get__(o) + ' - straightened'
                l.Layer = self.layerpicker.SelectedSerialNumber
            
            # go through the elements of the linestring
            for i in range(0, o.ElementCount):
                ProgressBar.TBC_ProgressBar.Title = "Node " + str(i + 1) + '/' + str(o.ElementCount)
                if ProgressBar.TBC_ProgressBar.SetProgress(math.floor((i + 1) * 100 / o.ElementCount)):
                    break
                
                e = o.ElementAt(i)
                etype = e.GetType()

                if not "Straight" in etype.Name:
                    if "PointId" in etype.Name:
                        enew = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IPointIdLocation))
                        enew.LocationSerialNumber = e.LocationSerialNumber
                    else:
                        enew = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                        enew.Position = e.Position
                else:
                    enew = e

                if self.straightconverttonew.IsChecked:
                    l.AppendElement(enew)
                elif self.straightconvertexisting.IsChecked:
                    o.ReplaceElementAt(enew, i)
            self.drawoverlay()
        else:
            pass

    def smoothparabolathepolyseg(self, polyseg):

        if polyseg:
            # see if we can get rid of superflous nodes before we smooth it
            #polyseg = polyseg.Linearize(self.smoothchordtol.Distance, self.smoothchordtol.Distance, self.smoothnodespacing.Distance, None, False)
            # need to do it a better/cleverer way since the max nodespacing could introduce additional nodes we don't want

            ptsorg = polyseg.ToPoint3DArray()
            # need to cleanup the array - I've seen lists with duplicate entries
            pts = []
            pts.Add(ptsorg[0])
            for i in range(1, ptsorg.Count):
                if Point3D.Distance(ptsorg[i], ptsorg[i - 1])[0] > 0:
                    pts.Add(ptsorg[i])

            if pts.Count > 2:
                
                polysegsmooth = PolySeg.PolySeg()
                for i in range(pts.Count):
                    #tt = Vector3D(pts[i - 1], pts[i]).Azimuth
                    #tt2 = Vector3D(pts[i], pts[i + 1]).Azimuth

                    if i == 0:
                        polysegsmooth.Add(PolySeg.Segment.Line(pts[i], Point3D.MidPoint(pts[i], pts[i + 1])))

                    elif i == pts.Count - 1:
                        polysegsmooth.Add(PolySeg.Segment.Line(Point3D.MidPoint(pts[pts.Count - 2], pts[pts.Count - 1]), pts[pts.Count - 1]))
                        continue    # break from loop

                    elif abs(Vector3D(pts[i - 1], pts[i]).Azimuth - Vector3D(pts[i], pts[i + 1]).Azimuth) > 0.000000001:
                        tt = PolySeg.Segment.Parabola(Point3D.MidPoint(pts[i - 1], pts[i]), pts[i], Point3D.MidPoint(pts[i], pts[i + 1]))
                        polysegsmooth.Add(tt)

                    else: # if the two segements are too straight you can't compute a parabola
                        polysegsmooth.Add(PolySeg.Segment.Line(Point3D.MidPoint(pts[i - 1], pts[i]), Point3D.MidPoint(pts[i], pts[i + 1])))

                    
            else:
                # it's just a two point line, which may be a curve
                # so, keep the original shape as close as possible
                polysegsmooth = polyseg.Linearize(self.smoothchordtol.Distance, self.smoothchordtol.Distance, self.smoothnodespacing.Distance, None, False)

            return polysegsmooth

        else:
            return None

    def smoothline(self):

        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        
        wv = self.currentProject [Project.FixedSerial.WorldView]
        wv.PauseGraphicsCache(True)

        try:

            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                # get the line as object
                l4 = self.linepicker4.Entity
                if l4:
                    polyseg4 = l4.ComputePolySeg()
                    polyseg4 = polyseg4.ToWorld()
                    polyseg4_v = l4.ComputeVerticalPolySeg()

                    
                    if self.smoothparabolamode.IsChecked:
                        # smooth the horizontal polyseg                    
                        if self.smoothhor.IsChecked or self.smoothhorver.IsChecked:
                            polysegsmooth = self.smoothparabolathepolyseg(polyseg4)
                        else:
                            polysegsmooth = polyseg4

                        # smooth the vertical polyseg    
                        if self.smoothver.IsChecked or self.smoothhorver.IsChecked:                
                            polysegsmooth_v = self.smoothparabolathepolyseg(polyseg4_v)
                        else:
                            polysegsmooth_v = polyseg4_v

                        # a Linestring doesn't know Parabolas - only an Alignment does
                        # so we need to Linearize the Polyseg before adding it to the Linestring
                        polysegsmooth = polysegsmooth.Linearize(self.smoothchordtol.Distance, self.smoothchordtol.Distance, self.smoothnodespacing.Distance, None, False)
                        
                        # match the new smoothed vertical profile length to the new horizontal profile length
                        if polysegsmooth_v:
                            length1 =  polysegsmooth.ComputeStationing()
                            # for the new profile we need to work around it since we only need the "Chainage-Length"
                            # hence the use of the bounding box limits in X direction which is the same as chainage in this case
                            length2 = polysegsmooth_v.BoundingBox.ptMax.X - polysegsmooth_v.BoundingBox.ptMin.X
                            smoothscalehz = length1 / length2
                            smoothchordtol = abs(self.smoothchordtol.Distance) / smoothscalehz
                            smoothnodespacing = abs(self.smoothnodespacing.Distance) / smoothscalehz

                            polysegsmooth_v = polysegsmooth_v.Linearize(smoothchordtol, smoothchordtol, smoothnodespacing, None, False)
                            # apply scale
                            polysegsmooth_v.Scale(Vector3D(smoothscalehz, 1.0, 0))

                        newname = IName.Name.__get__(l4) + ' - smoothed'
                        l = wv.Add(clr.GetClrType(Linestring))
                        l.Append(polysegsmooth, polysegsmooth_v, False, False)
                        l.Name = newname
                        l.Layer = self.layerpicker.SelectedSerialNumber

                    elif self.smoothconvertmode.IsChecked:
                        self.smoothconvertthepolyseg(l4)

                    elif self.straightconvertmode.IsChecked:
                        self.straightconvertthepolyseg(l4)

                else:
                    self.error.Text += '\nno Line selected'

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
            self.error.Text += '\nan Error occurred - Result probably incomplete\n' + str(exc_type) + '\n' + str(exc_obj) + '\nLine ' + str(exc_tb.tb_lineno)

        
        self.success.Content += '\nDone'
        
        wv.PauseGraphicsCache(False)

        Keyboard.Focus(self.linepicker4)
        #Keyboard.Focus(self.objs)

    def computechainage_Click(self, sender, e):
        l1 = self.linepicker1.Entity
        l2 = self.linepicker2.Entity

        if l1 != None and l2 != None:
            polyseg1 = l1.ComputePolySeg()
            polyseg1 = polyseg1.ToWorld()
            polyseg2 = l2.ComputePolySeg()
            polyseg2 = polyseg2.ToWorld()

            length1 = polyseg1.ComputeStationing()
            
            # if Line2 is a 3D Line we can get the length from the "curved" horizontal component with ComputeStationing()
            # if Line2 is a Planview 2D Line we need to work around it since we only need the "Chainage-Length"
            # hence the use of the bounding box limits in X direction which is the same as chainage in this case
            if self.l2is2d.IsChecked:
                length2 = polyseg2.BoundingBox.ptMax.X - polyseg2.BoundingBox.ptMin.X
            else:    
                length2 = polyseg2.ComputeStationing()

            self.combinescalehz.Value = length1 / length2

        else:
            self.error.Text += 'select 2 lines first'

    def copyfromhzscale_Click(self, sender, e):
        self.combinescalevz.Value = self.combinescalehz.Value
            
    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def Dispose(self, cmd, disposing):
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)


    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Text=''
        self.success.Content = ''

        if self.combinemode.IsChecked:
            self.combine()
        if self.drawmode.IsChecked:
            self.drawprofile()
        if self.smoothmode.IsChecked:
            self.smoothline()


        self.SaveOptions()

    


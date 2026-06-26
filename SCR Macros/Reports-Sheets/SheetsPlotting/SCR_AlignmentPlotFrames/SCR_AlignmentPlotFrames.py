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
    "layerpicker":      8,
    "frameleft":        False,
    "frameup":          True,
    "frameright":       False,
    "framedown":        False,
    "chooserange":      False,
    "startstation":     0.0,
    "endstation":       0.0,
    "dynawidth":        400.0,
    "dynaheight":       252.0,
    "dynascale":        20.0,
    "framecount":       10.0,
    "alignclearance":   10.0,
    "createsheets":     False,
    "dynalayerpicker":  8,
    "dynax":            0.010,
    "dynay":            0.035,
}

def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_AlignmentPlotFrames"
    cmdData.CommandName = "SCR_AlignmentPlotFrames"
    cmdData.Caption = "_SCR_AlignmentPlotFrames"
    cmdData.UIForm = "SCR_AlignmentPlotFrames"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Reports"
        cmdData.DefaultTabGroupKey = "Sheets and Dynaviews"
        cmdData.ShortCaption = "Plotframes along Alignment"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.18
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Create Plotframes along an Alignment"
        cmdData.ToolTipTextFormatted = "Create Plotframes along an Alignment"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_AlignmentPlotFrames(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_AlignmentPlotFrames.xaml") as s:
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
        #types = Array [Type] ([CadPoint]) + Array [Type] ([Point3D])    # we fill an array with TBC object types, we could combine different types
        self.linepicker1.IsEntityValidCallback = self.IsValid
        self.linepicker1.ValueChanged += self.lineChanged
        
        self.lType = clr.GetClrType(IPolyseg)

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear

        #self.lfp = self.lunits.Properties.Copy()
        self.linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation
        if self.lunits.DisplayType == LinearType.Meter:
            self.linearsuffixsmall = self.lunits.Units[LinearType.Millimeter].Abbreviation
        else:
            self.linearsuffixsmall = self.lunits.Units[LinearType.Inch].Abbreviation

        self.labeldynadim.Content = 'Dynaview Dimensions on Sheet [' + self.linearsuffixsmall + ']'
        self.labelclearancedim.Content = 'Start/End Clearance [' + self.linearsuffix + ']'
        self.labelframedim.Content = 'Computed Frame Dimensions [' + self.linearsuffix + ']'
        self.labeldynax.Content = 'X [' + self.linearsuffixsmall + ']'
        self.labeldynay.Content = 'Y [' + self.linearsuffixsmall + ']'

        self.dynawidth.DisplayUnit = LinearType.DisplaySmall
        self.dynaheight.DisplayUnit = LinearType.DisplaySmall
        self.dynax.DisplayUnit = LinearType.DisplaySmall
        self.dynay.DisplayUnit = LinearType.DisplaySmall

        self.dynawidth.DistanceMin = 0
        #self.dynawidth.NumberOfDecimals = 1
        self.dynawidth.ValueChanged += self.dynavalueChanged

        self.dynaheight.DistanceMin = 0
        #self.dynaheight.NumberOfDecimals = 1
        self.dynaheight.ValueChanged += self.dynavalueChanged

        self.dynascale.MinValue = 0
        self.dynascale.NumberOfDecimals = 2
        self.dynascale.ValueChanged += self.dynavalueChanged

        #self.alignclearance.NumberOfDecimals = 3
        self.alignclearance.ValueChanged += self.dynavalueChanged

        self.framecount.MinValue = 2
        self.framecount.NumberOfDecimals = 0
        self.framecount.ValueChanged += self.dynavalueChanged

        #self.dynax.NumberOfDecimals = 1
        #self.dynay.NumberOfDecimals = 1

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass


        self.frameleft.Checked += self.dynavalueChanged
        self.frameup.Checked += self.dynavalueChanged
        self.frameright.Checked += self.dynavalueChanged
        self.framedown.Checked += self.dynavalueChanged

        self.chooserange.Checked += self.dynavalueChanged
        self.chooserange.Unchecked += self.dynavalueChanged

        self.startstation.ValueChanged += self.dynavalueChanged
        self.endstation.ValueChanged += self.dynavalueChanged
        self.startstation.ValueChanged += self.lineChanged
        self.endstation.ValueChanged += self.lineChanged
        self.chooserange.Checked += self.lineChanged
        self.chooserange.Unchecked += self.lineChanged

        vfc = self.currentProject [Project.FixedSerial.ViewFilterCollection] # getting all View Filters in one list

        self.viewfilterpicker.SearchContainer = Project.FixedSerial.ViewFilterCollection
        self.viewfilterpicker.UseSelectionEngine = False
        self.viewfilterpicker.FilterByEntityTypes = Array[Type]([clr.GetClrType(ViewFilter)])

        self.sheetsetpicker.SearchContainer = 1 # ProjectFixedSerial 1 is the project itself
        self.sheetsetpicker.SearchSubContainer = True
        self.sheetsetpicker.UseSelectionEngine = False
        self.sheetsetpicker.FilterByEntityTypes = Array[Type]([clr.GetClrType(SheetSet)])

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass


    def SetDefaultOptions(self):
        try:
            self.sheetsetpicker.SelectBySerialNumber(OptionsManager.GetUint("SCR_AlignmentPlotFrames.sheetsetpicker", 0))
        except:
            pass
        try:
            self.viewfilterpicker.SelectBySerialNumber(OptionsManager.GetUint("SCR_AlignmentPlotFrames.viewfilterpicker", 0))
        except:
            pass
        SCROptions.LoadMacroOptions(self, "SCR_AlignmentPlotFrames", _OPTIONS, self.currentProject)

    def SaveOptions(self):
        OptionsManager.SetValue("SCR_AlignmentPlotFrames.sheetsetpicker", self.sheetsetpicker.SelectedSerial)
        OptionsManager.SetValue("SCR_AlignmentPlotFrames.viewfilterpicker", self.viewfilterpicker.SelectedSerial)
        SCROptions.SaveMacroOptions(self, "SCR_AlignmentPlotFrames", _OPTIONS)

    def drawoverlay(self):

        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)
        self.overlayBag = OverlayBag(self.ViewOverlay) # create a new/empty overlaybag

        l1 = self.linepicker1.Entity

        if l1:
            self.overlayBag.AddPolyline(SCROverlayBag.getpolypoints(l1), Color.Green.ToArgb(), 4)

            if self.chooserange.IsChecked:
                self.overlayBag.AddPolyline(SCROverlayBag.getclippedpolypoints(l1, self.startstation.Distance, self.endstation.Distance), Color.Orange.ToArgb(), 2)

            for p in SCROverlayBag.getarrowlocations(l1, 10): # returns list with location and perp right azimuth [Point3D, perpVector3D.Value.Azimuth]
                self.overlayBag.AddMarker(p[0], GraphicMarkerTypes.Arrow_IndependentColor, Color.Orange.ToArgb(), "", 0, math.pi - p[1], 3.0)



        # if it only needs to be visible in all Planview then remove the Hoops3DViewGUID
        array = Array[Guid]([DisplayWindow.Hoops3DViewGUID, DisplayWindow.HoopsPlanViewGUID])
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.AddOverlayGeometry(array, self.overlayBag)

        return

    def IsValid(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        return False

    def lineChanged(self, ctrl, e):
        l1 = self.linepicker1.Entity
        if l1 != None:
            self.startstation.StationProvider = l1
            self.endstation.StationProvider = l1
            #self.dynavalueChanged()
        self.drawoverlay()

    def framecountup(self, sender, e):
        self.framecount.Value += 1

    def framecountdown(self, sender, e):
        if self.framecount.Value > 2: 
            self.framecount.Value -= 1

    def dynavalueChanged(self, ctrl, e):
        if self.dynascale.Value > 0:

            
            self.framewidth.Content = str(self.lunits.Convert(self.dynawidth.Distance * self.dynascale.Value, LinearType.Display))
            self.frameheight.Content = str(self.lunits.Convert(self.dynaheight.Distance * self.dynascale.Value, LinearType.Display))
            
            
            if self.frameup.IsChecked or self.framedown.IsChecked:
                # alignclearance can max be half the frame height
                if self.alignclearance.Distance > ((self.dynaheight.Distance * self.dynascale.Value) / 2):
                    self.alignclearance.Distance = (self.dynaheight.Distance * self.dynascale.Value) / 2
            else:
                if self.alignclearance.Distance > ((self.dynawidth.Distance * self.dynascale.Value) / 2):
                    self.alignclearance.Distance = (self.dynawidth.Distance * self.dynascale.Value) / 2


            l1 = self.linepicker1.Entity
            inputok=True

            if l1 == None:
                inputok = False
            try:
                startstation = self.startstation.Distance
            except:
                inputok = False
            try:
                endstation = self.endstation.Distance
            except:
                inputok = False

            if inputok:
                if startstation > endstation: startstation, endstation = endstation, startstation   # swap values if necessary
                polyseg1 = l1.ComputePolySeg()
                polyseg1 = polyseg1.ToWorld()

                if self.chooserange.IsChecked:
                    computestartstation = startstation
                    computeendstation = endstation
                else:
                    computestartstation = polyseg1.BeginStation
                    computeendstation = polyseg1.ComputeStationing()
                
                # account for clearance
                if self.frameup.IsChecked or self.framedown.IsChecked:
                    computestartstation -= self.alignclearance.Distance
                    computestartstation += (self.dynaheight.Distance * self.dynascale.Value) / 2
                    computeendstation +=  self.alignclearance.Distance
                    computeendstation -= (self.dynaheight.Distance * self.dynascale.Value) / 2

                    if self.framecount.Value > 1:
                        ol = (self.dynaheight.Distance * self.dynascale.Value) - ((computeendstation - computestartstation) / (self.framecount.Value - 1))
                        self.frameoverlap.Content = str(self.lunits.Convert(ol, LinearType.Display))
                    else:
                        self.frameoverlap.Content = "0"
                else:
                    computestartstation -= self.alignclearance.Distance
                    computestartstation += (self.dynawidth.Distance * self.dynascale.Value) / 2
                    computeendstation +=  self.alignclearance.Distance
                    computeendstation -= (self.dynawidth.Distance * self.dynascale.Value) / 2

                    if self.framecount.Value > 1:
                        ol = (self.dynawidth.Distance * self.dynascale.Value) - ((computeendstation - computestartstation) / (self.framecount.Value - 1))
                        self.frameoverlap.Content = str(self.lunits.Convert(ol, LinearType.Display))
                    else:
                        self.frameoverlap.Content = "0"

            else:
                self.frameoverlap.Content = "check Alignment/Ch-Range"


    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def Dispose(self, cmd, disposing):
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)


    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''
        self.success.Content = ''

        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        
        wv = self.currentProject [Project.FixedSerial.WorldView]
        #wv.PauseGraphicsCache(True)

        # get the units for linear distance
        self.aunits = self.currentProject.Units.Angular
        # we don't want the units to be included (so we make copy and turn that off). Otherwise get something like "12.50 ft"
        self.afp = self.aunits.Properties.Copy()
        #self.afp.AddSuffix = False


        inputok=True

        l1=self.linepicker1.Entity
        if l1==None: 
            self.error.Content += '\nno Line 1 selected'
            inputok = False

        try:
            startstation = self.startstation.Distance
        except:
            self.error.Content += '\nStart Chainage error'
            inputok = False

        try:
            endstation = self.endstation.Distance
        except:
            self.error.Content += '\nEnd Chainage error'
            inputok = False

        if self.chooserange.IsChecked and startstation == endstation:
            self.error.Content += '\nStart and End Chainage must not be the same'
            inputok = False

        try:

            if self.createsheets.IsChecked:
                if (self.viewfilterpicker.SelectedIndex < 0) or (self.sheetsetpicker.SelectedIndex < 0):
                    self.success.Content += '\nSelect a Sheetset/Viewfilter'
                    inputok = False
                else:
                    vfserial = self.viewfilterpicker.SelectedSerial
                    ssserial = self.sheetsetpicker.SelectedSerial

            if inputok:


                # the "with" statement will unroll any changes if something go wrong
                with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                    if startstation > endstation: startstation, endstation = endstation, startstation   # swap values if necessary

                    polyseg1 = l1.ComputePolySeg()
                    polyseg1 = polyseg1.ToWorld()

                    if self.chooserange.IsChecked:
                        computestartstation = startstation
                        computeendstation = endstation
                    else:
                        computestartstation = polyseg1.BeginStation
                        computeendstation = polyseg1.ComputeStationing()

                    # account for clearance
                    if self.frameup.IsChecked or self.framedown.IsChecked:
                        computestartstation -= self.alignclearance.Distance
                        computestartstation += (self.dynaheight.Distance * self.dynascale.Value) / 2
                        computeendstation +=  self.alignclearance.Distance
                        computeendstation -= (self.dynaheight.Distance * self.dynascale.Value) / 2
                    else:
                        computestartstation -= self.alignclearance.Distance
                        computestartstation += (self.dynawidth.Distance * self.dynascale.Value) / 2
                        computeendstation +=  self.alignclearance.Distance
                        computeendstation -= (self.dynawidth.Distance * self.dynascale.Value) / 2

                    interval = (computeendstation - computestartstation) / (self.framecount.Value - 1)

                    outSegment = clr.StrongBox[Segment]()
                    out_t = clr.StrongBox[float]()
                    outPointOnCL1 = clr.StrongBox[Point3D]()
                    perpVector3D = clr.StrongBox[Vector3D]()

                    frame_i = 0

                    while computestartstation <= computeendstation + 0.001:    # as long as we aren't at the end of the line
                        # compute a point on line 1
                        polyseg1.FindPointFromStation(computestartstation, outSegment, out_t, outPointOnCL1, perpVector3D) # compute point and vector on line 1                
                        p1 = outPointOnCL1.Value
                        
                        # take the perp vector and compute the bottom right corner of the frame
                        compvector = perpVector3D.Value
                        #account for the frame orientation setting - the main part is developed for frame-up = alignment-direction
                        if self.frameright.IsChecked:
                            compvector.Rotate90(Side.Right)
                        if self.frameleft.IsChecked:
                            compvector.Rotate90(Side.Left)
                        if self.framedown.IsChecked:
                            compvector.Rotate180()

                        frameaz = compvector.Azimuth

                        compvector.To2D()
                        compvector.Length = (self.dynawidth.Distance * self.dynascale.Value) / 2
                        p_temp = p1 + compvector
                        compvector.Rotate90(Side.Right)
                        compvector.Length = (self.dynaheight.Distance * self.dynascale.Value) / 2
                        br = p_temp + compvector
                        br.Z = 0
                        # top right
                        compvector.Rotate180()
                        compvector.Length = self.dynaheight.Distance * self.dynascale.Value
                        tr = br + compvector
                        tr.Z = 0
                        # top left
                        compvector.Rotate90(Side.Left)
                        compvector.Length = self.dynawidth.Distance * self.dynascale.Value
                        tl = tr + compvector
                        tl.Z = 0
                        # bottom left
                        compvector.Rotate90(Side.Left)
                        compvector.Length = self.dynaheight.Distance * self.dynascale.Value
                        bl = tl + compvector
                        bl.Z = 0

                        #draw some frame information
                        frame_i += 1
                        
                        t = wv.Add(clr.GetClrType(MText))
                        t.AlignmentPoint = bl
                        t.AttachPoint = AttachmentPoint.BottomLeft
                        t.Rotation = (math.pi / 2) - perpVector3D.Value.Azimuth
                        # account for frame orientation
                        if self.frameright.IsChecked: t.Rotation -= math.pi/2
                        if self.frameleft.IsChecked: t.Rotation += math.pi/2
                        if self.framedown.IsChecked: t.Rotation += math.pi
                        # text content
                        t.TextString = 'Frame ' + str(frame_i)
                        t.TextString += '\nDynaview: ' + str(self.lunits.Convert(self.dynawidth.Distance, LinearType.DisplaySmall)) + ' ' + self.linearsuffixsmall + ' / ' + str(self.lunits.Convert(self.dynaheight.Distance, LinearType.DisplaySmall)) + ' ' + self.linearsuffixsmall
                        t.TextString += '\nScale: 1 : ' + str(self.dynascale.Value)
                        t.TextString += '\nFrame: ' + str(self.lunits.Convert(self.dynawidth.Distance, LinearType.Display) * self.dynascale.Value) + ' ' + self.linearsuffix +' / ' + str(self.lunits.Convert(self.dynaheight.Distance, LinearType.Display) * self.dynascale.Value)+ ' ' + self.linearsuffix
                        t.TextString += '\noriginal Bottom Azimuth: ' + self.aunits.Format(frameaz, self.afp)
                        t.TextString += '\nfor Dynaview-Creator: ' + self.aunits.Format((2*math.pi)-(frameaz - math.pi/2), self.afp)
                        t.Height = (self.dynaheight.Distance * self.dynascale.Value) / 50
                        t.Layer = self.layerpicker.SelectedSerialNumber


                        #draw the rectangle
                        l = wv.Add(clr.GetClrType(Linestring))
                        l.Closed = True
                        l.Layer = self.layerpicker.SelectedSerialNumber
                        l.Name = 'Sheet for Frame ' + str(frame_i)
                        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                        e.Position = bl
                        l.AppendElement(e)       
                        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                        e.Position = br
                        l.AppendElement(e)       
                        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                        e.Position = tr
                        l.AppendElement(e)       
                        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                        e.Position = tl
                        l.AppendElement(e)       
                        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                        e.Position = bl
                        l.AppendElement(e)       

                        # fill the sheets
                        if self.createsheets.IsChecked:
                            # get the sheetset as object
                            sheetset = self.currentProject.Concordance.Lookup(ssserial)
                            #create a new sheet in the sheetset
                            newsheet = sheetset.Add(clr.GetClrType(BasicSheet))
                            newsheet.Name = 'Sheet for Frame ' + str(frame_i)
                            newsheet.SortRank = frame_i * 10

                            # create a dynaview on the new sheet
                            newdynaview = newsheet.Add(clr.GetClrType(DynaView))
                            newdynaview.ViewFilter = vfserial
                            newdynaview.Boundary = l.SerialNumber
                            newdynaview.Layer = self.dynalayerpicker.SelectedSerialNumber
                            newdynaview.Location = Point3D(self.dynax.Distance, self.dynay.Distance, 0)
                            newdynaview.Name = 'Frame ' + str(frame_i)
                            newdynaview.Rotation = (perpVector3D.Value.Azimuth) - math.pi/2
                            # account for frame orientation
                            if self.frameright.IsChecked: newdynaview.Rotation += math.pi/2
                            if self.frameleft.IsChecked: newdynaview.Rotation -= math.pi/2
                            if self.framedown.IsChecked: newdynaview.Rotation += math.pi

                            newdynaview.Scale = self.dynascale.Value
                        
                        
                        computestartstation += interval


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

        #wv.PauseGraphicsCache(False)

        #self.currentProject.Calculate(False)

        self.SaveOptions()

    


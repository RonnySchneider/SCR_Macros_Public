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
    cmdData.Key = "SCR_BearingPads"
    cmdData.CommandName = "SCR_BearingPads"
    cmdData.Caption = "_SCR_BearingPads"
    cmdData.UIForm = "SCR_BearingPads"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "Lines"
        cmdData.ShortCaption = "Bearing-Pads"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.07
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Create Bearing Pad Outlines"
        cmdData.ToolTipTextFormatted = "Create Bearing Pad Outlines"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_BearingPads(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_BearingPads.xaml") as s:
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

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear

        #self.lfp = self.lunits.Properties.Copy()
        self.linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation
        if self.lunits.DisplayType == LinearType.Meter:
            self.linearsuffixsmall = self.lunits.Units[LinearType.Millimeter].Abbreviation
        else:
            self.linearsuffixsmall = self.lunits.Units[LinearType.Inch].Abbreviation

        self.bearingheader.Header = 'define Bearing Dimensions [' + self.linearsuffix + ']'
        self.offsetlabel.Content = 'offset distance (positive is outwards) [' + self.linearsuffix + ']'
        self.dynalabel.Content = 'Dynaview Dimensions on Sheet [' + self.linearsuffixsmall + ']'

        self.shelfcoordpick1.ValueChanged += self.shelfcoordpick1changed
        self.shelfcoordpick2.ValueChanged += self.shelfcoordpick2changed
        self.shelfcoordpick3.ValueChanged += self.shelfcoordpick3changed

        
        self.coordpick4.ValueChanged += self.coordpick4changed
        #self.bearingwidth.NumberOfDecimals = 4
        #self.bearinglength.NumberOfDecimals = 4
        #self.bearingoffset.NumberOfDecimals = 4

        #self.lineoffset.NumberOfDecimals = 4

        #self.dynawidth.MinValue = 0
        #self.dynawidth.NumberOfDecimals = 1
        #
        #self.dynaheight.MinValue = 0
        #self.dynaheight.NumberOfDecimals = 1
        #
        self.dynawidth.DisplayUnit = LinearType.DisplaySmall
        self.dynaheight.DisplayUnit = LinearType.DisplaySmall
        self.dynascale.MinValue = 0
        self.dynascale.NumberOfDecimals = 2

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def shelfcoordpick1changed(self, ctrl, e):
        self.shelfcoordpick2.CursorStyle = CursorStyle.CrossHair | CursorStyle.RubberLine
        if self.shelfcoordpick1.ResultCoordinateSystem:
            self.shelfcoordpick2.AnchorPoint = MousePosition(self.shelfcoordpick1.ClickWindow, self.shelfcoordpick1.Coordinate, self.shelfcoordpick1.ResultCoordinateSystem)
        else:
            self.shelfcoordpick2.AnchorPoint = None

        if not self.shelfcoordpick1.Coordinate.Is3D :
            self.shelfcoordpick1.StatusMessage = "No valid coordinate defined, must be 3D"
        else:
            self.shelfcoordpick1.StatusMessage = ""

        self.drawoverlay()

    def shelfcoordpick2changed(self, ctrl, e):
        self.shelfcoordpick3.CursorStyle = CursorStyle.CrossHair | CursorStyle.RubberLine
        if self.shelfcoordpick2.ResultCoordinateSystem:
            self.shelfcoordpick3.AnchorPoint = MousePosition(self.shelfcoordpick2.ClickWindow, self.shelfcoordpick2.Coordinate, self.shelfcoordpick2.ResultCoordinateSystem)
        else:
            self.shelfcoordpick3.AnchorPoint = None

        if not self.shelfcoordpick2.Coordinate.Is3D :
            self.shelfcoordpick2.StatusMessage = "No valid coordinate defined, must be 3D"
        else:
            self.shelfcoordpick2.StatusMessage = ""

        self.drawoverlay()

    def shelfcoordpick3changed(self, ctrl, e):
        if not self.shelfcoordpick3.Coordinate.Is3D :
            self.shelfcoordpick3.StatusMessage = "No valid coordinate defined, must be 3D"
        else:
            self.shelfcoordpick3.StatusMessage = ""

        self.drawoverlay()

    def drawoverlay(self):

        wv = self.currentProject [Project.FixedSerial.WorldView]
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)
        self.overlayBag = OverlayBag(self.ViewOverlay) # create a new/empty overlaybag

        if self.shelfcoordpick1.Coordinate.Is3D and self.shelfcoordpick1.Coordinate.Is3D and self.shelfcoordpick1.Coordinate.Is3D:
       
            self.overlayBag.AddPolyline(Array[Point3D]([self.shelfcoordpick1.Coordinate, self.shelfcoordpick2.Coordinate, self.shelfcoordpick3.Coordinate, \
                                                        self.shelfcoordpick1.Coordinate]), Color.Blue.ToArgb(), 5)

            self.overlayBag.AddMarker(self.shelfcoordpick1.Coordinate, GraphicMarkerTypes.HollowCircle_IndependentColor, Color.Magenta.ToArgb(), "   V1", 0, 0, 2.0) # last 2 numbers, markercircle-rotation/scale
            self.overlayBag.AddMarker(self.shelfcoordpick2.Coordinate, GraphicMarkerTypes.HollowCircle_IndependentColor, Color.Magenta.ToArgb(), "   V2", 0, 0, 2.0) # last 2 numbers, markercircle-rotation/scale
            self.overlayBag.AddMarker(self.shelfcoordpick3.Coordinate, GraphicMarkerTypes.HollowCircle_IndependentColor, Color.Magenta.ToArgb(), "   V3", 0, 0, 2.0) # last 2 numbers, markercircle-rotation/scale
            
            # if it only needs to be visible in all Planview then remove the Hoops3DViewGUID
            array = Array[Guid]([DisplayWindow.Hoops3DViewGUID, DisplayWindow.HoopsPlanViewGUID])
            TrimbleOffice.TheOffice.MainWindow.AppViewManager.AddOverlayGeometry(array, self.overlayBag)

        return

    
    def twopadsChanged(self, sender, e):
        if self.twopads.IsChecked:
            self.bearingoffset.IsEnabled = True
        else:
            self.bearingoffset.IsEnabled = False

    def SetDefaultOptions(self):

        settingserial = OptionsManager.GetUint("SCR_BearingPads.layerpicker", 8) # 8 is FixedSerial for Layer Zero
        o = self.currentProject.Concordance.Lookup(settingserial)
        if o != None:
            if isinstance(o.GetSite(), LayerCollection):    
                self.layerpicker.SetSelectedSerialNumber(settingserial, InputMethod(3))
            else:                       
                self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:                       
            self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
    
        self.elevatetoshelf.IsChecked = OptionsManager.GetBool("SCR_BearingPads.elevatetoshelf", False)
        self.bearinglength.Distance = OptionsManager.GetDouble("SCR_BearingPads.bearinglength", 0.000)
        self.bearingwidth.Distance = OptionsManager.GetDouble("SCR_BearingPads.bearingwidth", 0.000)

        self.twopads.IsChecked = OptionsManager.GetBool("SCR_BearingPads.twopads", True)
        self.bearingoffset.Distance = OptionsManager.GetDouble("SCR_BearingPads.bearingoffset", 0.000)

        self.drawoffsetline.IsChecked = OptionsManager.GetBool("SCR_BearingPads.drawoffsetline", False)
    
        settingserial = OptionsManager.GetUint("SCR_BearingPads.layerpicker2", 8) # 8 is FixedSerial for Layer Zero
        o = self.currentProject.Concordance.Lookup(settingserial)
        if o != None:
            if isinstance(o.GetSite(), LayerCollection):    
                self.layerpicker2.SetSelectedSerialNumber(settingserial, InputMethod(3))
            else:                       
                self.layerpicker2.SetSelectedSerialNumber(8, InputMethod(3))
        else:                       
            self.layerpicker2.SetSelectedSerialNumber(8, InputMethod(3))
        
        self.lineoffset.Distance = OptionsManager.GetDouble("SCR_BearingPads.lineoffset", 0.0000)

        self.drawdynaframe.IsChecked = OptionsManager.GetBool("SCR_BearingPads.drawdynaframe", False)
        self.dynaframeonly.IsChecked = OptionsManager.GetBool("SCR_BearingPads.dynaframeonly", False)

        settingserial = OptionsManager.GetUint("SCR_BearingPads.layerpicker3", 8) # 8 is FixedSerial for Layer Zero
        o = self.currentProject.Concordance.Lookup(settingserial)
        if o != None:
            if isinstance(o.GetSite(), LayerCollection):    
                self.layerpicker3.SetSelectedSerialNumber(settingserial, InputMethod(3))
            else:                       
                self.layerpicker3.SetSelectedSerialNumber(8, InputMethod(3))
        else:                       
            self.layerpicker3.SetSelectedSerialNumber(8, InputMethod(3))

        self.dynawidth.Distance = OptionsManager.GetDouble("SCR_BearingPads.dynawidth", 0.4)
        self.dynaheight.Distance = OptionsManager.GetDouble("SCR_BearingPads.dynaheight", 0.252)
        self.dynascale.Value = OptionsManager.GetDouble("SCR_BearingPads.dynascale", 20.000)
   
    def SaveOptions(self):

        OptionsManager.SetValue("SCR_BearingPads.layerpicker", self.layerpicker.SelectedSerialNumber)

        OptionsManager.SetValue("SCR_BearingPads.elevatetoshelf", self.elevatetoshelf.IsChecked)
        OptionsManager.SetValue("SCR_BearingPads.bearinglength", self.bearinglength.Distance)
        OptionsManager.SetValue("SCR_BearingPads.bearingwidth", self.bearingwidth.Distance)

        OptionsManager.SetValue("SCR_BearingPads.twopads", self.twopads.IsChecked)

        OptionsManager.SetValue("SCR_BearingPads.bearingoffset", self.bearingoffset.Distance)

        OptionsManager.SetValue("SCR_BearingPads.drawoffsetline", self.drawoffsetline.IsChecked)
        OptionsManager.SetValue("SCR_BearingPads.layerpicker2", self.layerpicker2.SelectedSerialNumber)
        OptionsManager.SetValue("SCR_BearingPads.lineoffset", self.lineoffset.Distance)

        OptionsManager.SetValue("SCR_BearingPads.drawdynaframe", self.drawdynaframe.IsChecked)
        OptionsManager.SetValue("SCR_BearingPads.dynaframeonly", self.dynaframeonly.IsChecked)
        OptionsManager.SetValue("SCR_BearingPads.layerpicker3", self.layerpicker3.SelectedSerialNumber)
        OptionsManager.SetValue("SCR_BearingPads.dynawidth", self.dynawidth.Distance)
        OptionsManager.SetValue("SCR_BearingPads.dynaheight", self.dynaheight.Distance)
        OptionsManager.SetValue("SCR_BearingPads.dynascale", self.dynascale.Value)


    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def Dispose(self, cmd, disposing):
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)
    
    def coordpick4changed(self, ctrl, e):
        if e.Cause == InputMethod.Mouse:     
            self.OkClicked(None, None)



    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''



        self.success.Content = ""

        wv = self.currentProject [Project.FixedSerial.WorldView]
        
        # get the units for angles
        self.aunits = self.currentProject.Units.Angular
        # we don't want the units to be included (so we make copy and turn that off). Otherwise get something like "12.50 ft"
        self.afp = self.aunits.Properties.Copy()
        #self.afp.AddSuffix = False

        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        
        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                # Headstock CL
                p1 = self.coordpick1.Coordinate
                p1.Z = 0
                p2 = self.coordpick2.Coordinate
                p2.Z = 0
                # Girder CL
                p3 = self.coordpick3.Coordinate
                p3.Z = 0
                p4 = self.coordpick4.Coordinate
                p4.Z = 0

                self.p = None
                if self.shelfcoordpick1.Coordinate.Is3D and self.shelfcoordpick2.Coordinate.Is3D and self.shelfcoordpick3.Coordinate.Is3D:
                    self.p = Plane3D(self.shelfcoordpick1.Coordinate, self.shelfcoordpick2.Coordinate, self.shelfcoordpick3.Coordinate)[0] # the plane is returned as first element

                headstockseg = SegmentLine(p1, p2)
                headstockvec = Vector3D(p1, p2)
                girderseg = SegmentLine(p3, p4)
                girdervec = Vector3D(p3, p4)

                constructorvec = girdervec

                intersections = Intersections()
                
                if girderseg.Intersect(headstockseg, True, intersections):
                    ip = intersections[0].Point
                    intersections.Clear()
                    
                    # check if we only want to draw the plotframe
                    if self.drawdynaframe.IsChecked and self.dynaframeonly.IsChecked:
                        pass
                    else:
                        headstockvec.Length = self.bearingoffset.Distance

                        if self.twopads.IsChecked:
                            b1c = ip + headstockvec
                            b2c = ip - headstockvec
                        else:
                            b1c = ip
                        
                        # Pad 1
                        constructorvec.Length = self.bearinglength.Distance/2
                        b1p1 = b1c + constructorvec
                        constructorvec.Rotate90(Side.Right)
                        constructorvec.Length = self.bearingwidth.Distance/2
                        b1p1 = b1p1 + constructorvec
                        constructorvec.Rotate90(Side.Right)
                        constructorvec.Length = self.bearinglength.Distance
                        b1p2 = b1p1 + constructorvec
                        constructorvec.Rotate90(Side.Right)
                        constructorvec.Length = self.bearingwidth.Distance
                        b1p3 = b1p2 + constructorvec
                        constructorvec.Rotate90(Side.Right)
                        constructorvec.Length = self.bearinglength.Distance
                        b1p4 = b1p3 + constructorvec

                        # draw Pad 1
                        self.drawline(b1p1, b1p2, b1p3, b1p4, False)

                        if self.twopads.IsChecked:
                            # Pad 2
                            constructorvec.Length = self.bearinglength.Distance/2
                            b2p1 = b2c + constructorvec
                            constructorvec.Rotate90(Side.Right)
                            constructorvec.Length = self.bearingwidth.Distance/2
                            b2p1 = b2p1 + constructorvec
                            constructorvec.Rotate90(Side.Right)
                            constructorvec.Length = self.bearinglength.Distance
                            b2p2 = b2p1 + constructorvec
                            constructorvec.Rotate90(Side.Right)
                            constructorvec.Length = self.bearingwidth.Distance
                            b2p3 = b2p2 + constructorvec
                            constructorvec.Rotate90(Side.Right)
                            constructorvec.Length = self.bearinglength.Distance
                            b2p4 = b2p3 + constructorvec

                            # draw Pad 2
                            self.drawline(b2p1, b2p2, b2p3, b2p4, False)
                        
                    # dynaview frame
                    if self.drawdynaframe.IsChecked:
                        compvector = girdervec
                        compvector.To2D()
                        compvector.Length = (self.dynawidth.Distance * self.dynascale.Value) / 2
                        compvector.Rotate90(Side.Right)
                        p_temp = ip + compvector
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

                        # draw Dynaview-Frame
                        self.drawline(bl, br, tr, tl, True)

                        #draw some frame information
                        compvector.Rotate90(Side.Left)
                        t = wv.Add(clr.GetClrType(MText))
                        t.AlignmentPoint = bl
                        t.AttachPoint = AttachmentPoint.TopLeft
                        t.Rotation = (math.pi / 2) - compvector.Azimuth
                        # text content
                        t.TextString += '\nDynaview: ' + str(self.lunits.Convert(self.dynawidth.Distance, LinearType.DisplaySmall)) + ' ' + self.linearsuffixsmall + ' / ' + str(self.lunits.Convert(self.dynaheight.Distance, LinearType.DisplaySmall)) + ' ' + self.linearsuffixsmall
                        t.TextString += '\nScale: 1 : ' + str(self.dynascale.Value)
                        t.TextString += '\nFrame: ' + str(self.lunits.Convert(self.dynawidth.Distance, LinearType.Display) * self.dynascale.Value) + ' ' + self.linearsuffix +' / ' + str(self.lunits.Convert(self.dynaheight.Distance, LinearType.Display) * self.dynascale.Value)+ ' ' + self.linearsuffix
                        t.TextString += '\noriginal Bottom Azimuth: ' + self.aunits.Format(compvector.Azimuth, self.afp)
                        t.TextString += '\nfor Dynaview-Creator: ' + self.aunits.Format((2*math.pi)-(compvector.Azimuth - math.pi/2), self.afp)
                        t.Height = (self.dynaheight.Distance * self.dynascale.Value) / 50
                        t.Layer = self.layerpicker3.SelectedSerialNumber


                    Keyboard.Focus(self.coordpick3)
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
   
    def drawline(self, p1, p2, p3, p4, isframe):
        wv = self.currentProject [Project.FixedSerial.WorldView]

        # compute shelf elevations if ticked
        if self.elevatetoshelf.IsChecked and self.p and not isframe:
            p1 = self.planeelevation(p1)
            p2 = self.planeelevation(p2)
            p3 = self.planeelevation(p3)
            p4 = self.planeelevation(p4)

        l = wv.Add(clr.GetClrType(Linestring))
        l.Closed = True
        if isframe:
            l.Layer = self.layerpicker3.SelectedSerialNumber
        else:
            l.Layer = self.layerpicker.SelectedSerialNumber
        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
        e.Position = p1
        l.AppendElement(e)
        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
        e.Position = p2
        l.AppendElement(e) 
        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
        e.Position = p3
        l.AppendElement(e) 
        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
        e.Position = p4
        l.AppendElement(e) 
        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
        e.Position = p1
        l.AppendElement(e)


        # draw offset line if ticked and if it's not a dynaview frame
        if self.drawoffsetline.IsChecked and isframe == False:
            l = l.ComputePolySeg()
            if self.lineoffset.Distance < 0:
                side = Side.Right
            else:
                side = Side.Left

            l_offset = l.Offset(side, abs(self.lineoffset.Distance))
            
            if l_offset[0] == True:
                l = wv.Add(clr.GetClrType(Linestring))
                l.Layer=self.layerpicker2.SelectedSerialNumber
                for n in l_offset[1].Nodes():
                    try:
                        if n.Visible:
                            e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                            if self.elevatetoshelf.IsChecked and self.p:
                                e.Position = self.planeelevation(n.Point)
                            else:
                                e.Position = n.Point
                            l.AppendElement(e)
                    except:
                        break


    def planeelevation(self, pold):

        if not pold.Is3D:
            pold.Z = 0
        pnew = Plane3D.IntersectWithRay(self.p, pold, Vector3D(0, 0, 1))

        return pnew

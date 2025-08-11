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
    cmdData.Key = "SCR_AlignTextToLine"
    cmdData.CommandName = "SCR_AlignTextToLine"
    cmdData.Caption = "_SCR_AlignTextToLine"
    cmdData.UIForm = "SCR_AlignTextToLine"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Reports"
        cmdData.DefaultTabGroupKey = "Text"
        cmdData.ShortCaption = "Align Text to Line"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.15
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Align Text to Line"
        cmdData.ToolTipTextFormatted = "Align Text to Line"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_AlignTextToLine(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_AlignTextToLine.xaml") as s:
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
        self.linepicker1.IsEntityValidCallback=self.IsValidLine
        self.linepicker1.ValueChanged += self.lineChanged
        self.lType = clr.GetClrType(IPolyseg)

        self.textobjs.IsEntityValidCallback=self.IsValidText
        optionMenuText = SelectionContextMenuHandler()
        # remove options that don't apply here
        optionMenuText.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        self.textobjs.ButtonContextMenu = optionMenuText
        self.mtextType = clr.GetClrType(MText)
        self.cadtextType = clr.GetClrType(CadText)

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        #self.lfp = self.lunits.Properties.Copy()
        self.linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation

        self.labeloffset.Content = 'Offset [' + self.linearsuffix + ']'

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def SetDefaultOptions(self):


        self.singleline.IsChecked = OptionsManager.GetBool("SCR_AlignTextToLine.singleline", True)
        self.closestline.IsChecked = OptionsManager.GetBool("SCR_AlignTextToLine.closestline", False)

        # layer
        lserial = OptionsManager.GetUint("SCR_AlignTextToLine.layerpicker", 8)
        o = self.currentProject.Concordance.Lookup(lserial)
        if o != None:
            if isinstance(o.GetSite(), LayerCollection):    
                self.layerpicker.SetSelectedSerialNumber(lserial, InputMethod(3))
            else:                       
                self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:                       
            self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))

        self.adjustoffset.IsChecked = OptionsManager.GetBool("SCR_AlignTextToLine.adjustoffset", False)
        self.newoffset.Offset = OptionsManager.GetDouble("SCR_AlignTextToLine.newoffset", 0.000)

        self.textelfromlineel.IsChecked = OptionsManager.GetBool("SCR_AlignTextToLine.textelfromlineel", False)
        self.rotfromline.IsChecked = OptionsManager.GetBool("SCR_AlignTextToLine.rotfromline", True)
        self.addrot.Angle = OptionsManager.GetDouble("SCR_AlignTextToLine.addrot", 0.0)
        self.rotfixed.IsChecked = OptionsManager.GetBool("SCR_AlignTextToLine.rotfixed", False)
        self.fixedrot.Direction = OptionsManager.GetDouble("SCR_AlignTextToLine.fixedrot", 0.0)

    def SaveOptions(self):
        OptionsManager.SetValue("SCR_AlignTextToLine.singleline", self.singleline.IsChecked)
        OptionsManager.SetValue("SCR_AlignTextToLine.closestline", self.closestline.IsChecked)
        OptionsManager.SetValue("SCR_AlignTextToLine.layerpicker", self.layerpicker.SelectedSerialNumber)
        OptionsManager.SetValue("SCR_AlignTextToLine.adjustoffset", self.adjustoffset.IsChecked)
        OptionsManager.SetValue("SCR_AlignTextToLine.newoffset", self.newoffset.Offset)
        OptionsManager.SetValue("SCR_AlignTextToLine.textelfromlineel", self.textelfromlineel.IsChecked)
        OptionsManager.SetValue("SCR_AlignTextToLine.rotfromline", self.rotfromline.IsChecked)
        OptionsManager.SetValue("SCR_AlignTextToLine.addrot", self.addrot.Angle)
        OptionsManager.SetValue("SCR_AlignTextToLine.rotfixed", self.rotfixed.IsChecked)
        OptionsManager.SetValue("SCR_AlignTextToLine.fixedrot", self.fixedrot.Direction)

    def drawoverlay(self):

        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)
        self.overlayBag = OverlayBag(self.ViewOverlay) # create a new/empty overlaybag

        l1 = self.linepicker1.Entity

        if l1 != None:
            self.overlayBag.AddPolyline(self.getpolypoints(l1), Color.Green.ToArgb(), 4)

            for p in self.getarrowlocations(l1, 10): # returns list with location and perp right azimuth [Point3D, perpVector3D.Value.Azimuth]
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

    def lineChanged(self, ctrl, e):
        l1=self.linepicker1.Entity
        if l1 != None:
            self.newoffset.StationProvider = l1
            self.drawoverlay()

    def IsValidLine(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        return False

    def IsValidText(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.mtextType) or isinstance(o, self.cadtextType) or isinstance(o, CadLabel):
            return True
        return False

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def Dispose(self, cmd, disposing):
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''


        self.success.Content = ''

        wv = self.currentProject [Project.FixedSerial.WorldView]
        #wv.PauseGraphicsCache(True)

        
        inputok = True
        try:
            #if self.addrot.Text == "":
            #    addrot = 0
            #else:
            #    addrot = float(self.addrot.Text)
            #if addrot < 0: addrot += 360
            #if addrot >= 360: addrot -= 360
            #self.addrot.Text = str(addrot)
            #addrot = (addrot * math.pi)/180
            addrot = self.addrot.Angle
        except:
            self.error.Content += '\nadditional rotation error'
            inputok=False
        
        newoffset = self.newoffset.Offset
        #self.newoffset.Offset = newoffset
        if math.isnan(newoffset):
            newoffset = 0

        if self.singleline.IsChecked:
            l1 = self.linepicker1.Entity
            if l1 == None: 
                self.error.Content += '\nno Line selected'
                inputok=False

        if inputok:
            self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
            UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
            try:
                with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                    j = 0
                    ProgressBar.TBC_ProgressBar.Title = "align Textobjects"
                    for o in self.textobjs.SelectedMembers(self.currentProject): # go through all selected objects
                        j += 1
                        if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(j * 100 / self.textobjs.SelectedMembers(self.currentProject).Count)):
                            break   # function returns true if user pressed cancel

                        if isinstance(o, CadText) or isinstance(o, MText) or isinstance(o, CadLabel):  # if we have a text continue

                            if isinstance(o, CadText) or isinstance(o, MText):
                                textpoint = o.AlignmentPoint    # get the text position
                            
                            if isinstance(o, CadLabel):
                                # we need to take the initial delta into account, otherwise the transfrom would accumulate it each time
                                textpoint = o.ObjectAnchorPosition + o.Delta

                            if not textpoint.Is3D:
                                textpoint.Z = 0.0

                            if self.closestline.IsChecked:
                                l1 = self.findclosestline(textpoint)
                                if l1 == None: Continue # in case we couldn't find a line we continue with the next text object

                            polyseg1 = l1.ComputePolySeg()
                            polyseg1 = polyseg1.ToWorld()
                            polyseg1_v = l1.ComputeVerticalPolySeg()
                            tt = polyseg1.ComputeStationing()

                            # the FindPointFromPoint needs the output variables in a certain format 
                            texttestoutSegment = clr.StrongBox[Segment]()
                            out_t = clr.StrongBox[float]()
                            texttestPointOnCL1 = clr.StrongBox[Point3D]()
                            texttestStation = clr.StrongBox[float]()
                            texttestperpVector3D = clr.StrongBox[Vector3D]()
                            texttestDist = clr.StrongBox[float]()
                            texttestside = clr.StrongBox[Side]()
                            pl = Point3D()

                            # compute the text offset from the line
                            if polyseg1.FindPointFromPoint(textpoint, texttestoutSegment, out_t, texttestPointOnCL1, texttestStation, texttestperpVector3D, texttestDist, texttestside):

                                # get the elevation
                                if polyseg1_v == None:
                                    pl = texttestPointOnCL1.Value
                                else:
                                    pl.X = texttestPointOnCL1.Value.X
                                    pl.Y = texttestPointOnCL1.Value.Y
                                    pl.Z = polyseg1_v.ComputeVerticalSlopeAndGrade(texttestStation.Value)[1]


                                v = Vector3D(pl, textpoint)
                                v.To2D()
                                # text rotation works counter clockwise, with 0 in positive x direction
                                if self.rotfixed.IsChecked:
                                    newrot = self.fixedrot.Direction
                                else:
                                    newrot = math.pi/2 - (v.Azimuth + addrot)
                                    if texttestside.Value == Side.Left: newrot += math.pi
                                
                                

                                if isinstance(o, self.mtextType) or isinstance(o, CadLabel):
                                    o.Rotation = newrot
                                if isinstance(o, self.cadtextType):
                                    o.RotateAngle = newrot

                                if self.adjustoffset.IsChecked:
                                    v.Length = newoffset # * -1
                                
                                pnew = pl + v

                                if self.textelfromlineel.IsChecked:
                                    # so annoying, if it's a text or mtext with an undefined elevation we have to set it to zero first, otherwise transfrom won't work
                                    # FYI for labels; if you select to change the elevation but the labels reference point is undefined the label will stay at undefined,
                                    # even if the deltaZ shows now a value != 0, only if the point is set to zero the label will change to the new elevation
                                    if isinstance(o, self.mtextType) or isinstance(o, self.cadtextType):
                                        if not o.AlignmentPoint.Is3D:
                                            ptemp = o.AlignmentPoint
                                            ptemp.Z = 0.0
                                            o.AlignmentPoint = ptemp
                                
                                    pnew.Z = pl.Z
                                else:   
                                    pnew.Z = textpoint.Z

                                td = TransformData(Matrix4D(Vector3D(textpoint, pnew)), Matrix4D(Vector3D.Zero))
                                o.Transform(td)
                                #o.AlignmentPoint = p1 + v
                                

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
        ProgressBar.TBC_ProgressBar.Title = ""

        #wv.PauseGraphicsCache(False)

    def findclosestline(self, p2):

        layer_sn = self.layerpicker.SelectedSerialNumber
        wlc = self.currentProject[Project.FixedSerial.LayerContainer] # we get all the layers into an object, LayerCollection
        wl = wlc[layer_sn]    # we get just the source layer as an object
        layermembers = wl.Members  # we get serial number list of all the elements on that layer

        outPointOnCL = clr.StrongBox[Point3D]()
        station = clr.StrongBox[float]()

        linefound = None
        pfound = False

        for objserial in layermembers: # go through all members of the design layer
            l = self.currentProject.Concordance.Lookup(objserial)
            if isinstance(l, self.lType) and not (isinstance(l, Leader)):
                           
                polyseg = l.ComputePolySeg()
                try:
                    polyseg = polyseg.ToWorld()
                except:
                    pass
                polyseg_v = l.ComputeVerticalPolySeg()

                # try to find a perpendicular solution on that line
                if polyseg:
                    if polyseg.FindPointFromPoint(p2, outPointOnCL, station): 
                        pcompute = outPointOnCL.Value
                        if polyseg_v != None:
                            pcompute.Z = polyseg_v.ComputeVerticalSlopeAndGrade(station.Value)[1]
                        
                        if pfound == False:
                            p1 = pcompute
                            linefound = l
                            pfound = True
                        else: # if we find multiple solutions we only want the closest one
                            if Point3D.Distance2D(pcompute, p2) < Point3D.Distance2D(p1, p2):
                                p1 = pcompute
                                linefound = l
                                pfound = True
                else:
                    pass

        return linefound
    
    def ccw45_Click(self, sender, e):
        #rot = float(self.addrot.Text) - 45
        #if rot < 0: rot += 360
        #if rot >= 360: rot -= 360
        #self.addrot.Text = str(rot)
        self.addrot.Angle -= math.pi / 4
        
    def cw45_Click(self, sender, e):
        #rot = float(self.addrot.Text) + 45
        #if rot < 0: rot += 360
        #if rot >= 360: rot -= 360
        #self.addrot.Text = str(rot)
        self.addrot.Angle += math.pi / 4




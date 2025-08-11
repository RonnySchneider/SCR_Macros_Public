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
    cmdData.Key = "SCR_ImportStaOffEl"
    cmdData.CommandName = "SCR_ImportStaOffEl"
    cmdData.Caption = "_SCR_ImportStaOffEl"
    cmdData.UIForm = "SCR_ImportStaOffEl"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "Import/Export"
        cmdData.ShortCaption = "StaOffEl CSV to Points"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.09
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Import CSV with StaOffElev"
        cmdData.ToolTipTextFormatted = "compute points based on CSV with Stations, Offsets, Elevations"

    except:
        pass
    
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_ImportStaOffEl(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_ImportStaOffEl.xaml") as s:
            wpf.LoadComponent (self, s)
        self.currentProject = currentProject
        self.macroFileFolder = macroFileFolder
        if self.openfilename.Text=='': self.openfilename.Text = macroFileFolder

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
        self.linepicker1.IsEntityValidCallback=self.IsValid
        self.linepicker1.ValueChanged += self.lineChanged
        self.lType = clr.GetClrType(IPolyseg)

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass
        self.unitssetup(None, None)  

    def unitssetup(self, sender, e):
        # setup everything for the unit conversions
        self.inputunitenum = 0

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        self.lfp = self.lunits.Properties.Copy() # create a copy in order to set the decimals and enable/disable the suffix
        self.lfp.AddSuffix = False # disable suffix, we need to set it manually, it would always add the current projects units

        # fill the unitpicker
        for u in self.lunits.Units:
            item = ComboBoxItem()
            item.Content = u.Key
            item.FontSize = 1
            self.unitpicker.Items.Add(item)

        tt = self.unitpicker.Text
        self.unitpicker.SelectedIndex = 0
        if tt != "":
            self.unitpicker.Text = tt
        self.unitpicker.SelectionChanged += self.unitschanged

        self.unitschanged(None, None)
    
    def unitschanged(self, sender, e):

        # find the enum for the selected LinearType
        for e in range(0, 19):
            if LinearType(e) == self.unitpicker.SelectedItem.Content:
                self.inputunitenum = e

    def toprojectunit(self, v):
        return self.lunits.Convert(LinearType(self.inputunitenum), v, LinearType.Meter)

    def SetDefaultOptions(self):
        self.openfilename.Text = OptionsManager.GetString("SCR_ImportStaOffEl.openfilename", os.path.expanduser('~\\Downloads'))

        lserial = OptionsManager.GetUint("SCR_ImportStaOffEl.layerpicker", 8)
        o = self.currentProject.Concordance.Lookup(lserial)
        if o != None:
            if isinstance(o.GetSite(), LayerCollection):    
                self.layerpicker.SetSelectedSerialNumber(lserial, InputMethod(3))
            else:                       
                self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:                       
            self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))

        self.createdependentpoints.IsChecked = OptionsManager.GetBool("SCR_ImportStaOffEl.createdependentpoints", False)
        self.createpointonline.IsChecked = OptionsManager.GetBool("SCR_ImportStaOffEl.createpointonline", False)
        self.unitpicker.Text = OptionsManager.GetString("SCR_ImportStaOffEl.unitpicker", "Meter")
        self.textBox1.Text = OptionsManager.GetString("SCR_ImportStaOffEl.textBox1", "")
        self.textBox2.Text = OptionsManager.GetString("SCR_ImportStaOffEl.textBox2", "2")
        self.textBox3.Text = OptionsManager.GetString("SCR_ImportStaOffEl.textBox3", "3")
        self.textBox4.Text = OptionsManager.GetString("SCR_ImportStaOffEl.textBox4", "4")

        self.relativeelev.IsChecked = OptionsManager.GetBool("SCR_ImportStaOffEl.relativeelev", True)
        self.absoluteelev.IsChecked = OptionsManager.GetBool("SCR_ImportStaOffEl.absoluteelev", False)
        #self.el_checkBox.IsChecked = OptionsManager.GetBool("SCR_ImportStaOffEl.el_checkBox", True)
        
        self.textBox5.Text = OptionsManager.GetString("SCR_ImportStaOffEl.textBox5", "0")
            
    def SaveOptions(self):
        OptionsManager.SetValue("SCR_ImportStaOffEl.openfilename", self.openfilename.Text)

        OptionsManager.SetValue("SCR_ImportStaOffEl.layerpicker", self.layerpicker.SelectedSerialNumber)

        OptionsManager.SetValue("SCR_ImportStaOffEl.createdependentpoints", self.createdependentpoints.IsChecked)
        OptionsManager.SetValue("SCR_ImportStaOffEl.createpointonline", self.createpointonline.IsChecked)
        OptionsManager.SetValue("SCR_ImportStaOffEl.unitpicker", self.unitpicker.Text)
        OptionsManager.SetValue("SCR_ImportStaOffEl.textBox1", self.textBox1.Text)
        OptionsManager.SetValue("SCR_ImportStaOffEl.textBox2", self.textBox2.Text)
        OptionsManager.SetValue("SCR_ImportStaOffEl.textBox3", self.textBox3.Text)
        OptionsManager.SetValue("SCR_ImportStaOffEl.textBox4", self.textBox4.Text)

        OptionsManager.SetValue("SCR_ImportStaOffEl.relativeelev", self.relativeelev.IsChecked)
        OptionsManager.SetValue("SCR_ImportStaOffEl.absoluteelev", self.absoluteelev.IsChecked)
        #OptionsManager.SetValue("SCR_ImportStaOffEl.el_checkBox", self.el_checkBox.IsChecked)

        OptionsManager.SetValue("SCR_ImportStaOffEl.textBox5", self.textBox5.Text)

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

    def createdependentpointsChanged(self, sender, e):
        if self.createdependentpoints.IsChecked:
            self.textBox5.IsEnabled = False
            self.relativeelev.Content = "relative and dependent to Line"
        else:
            self.textBox5.IsEnabled = True
            self.relativeelev.Content = "relative to Line"
        
    def lineChanged(self, ctrl, e):
        l1=self.linepicker1.Entity
        if l1 != None:
            self.drawoverlay()

    def IsValid(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        return False

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def Dispose(self, cmd, disposing):
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)

    def openbutton_Click(self, sender, e):
        dialog=OpenFileDialog()
        dialog.InitialDirectory = self.openfilename.Text
        dialog.Filter=("CSV|*.csv")
        
        tt=dialog.ShowDialog()
        if tt==DialogResult.OK:
            self.openfilename.Text = dialog.FileName

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''

        self.errortext.Content=''
        layer_sn = self.layerpicker.SelectedSerialNumber

        if self.linepicker1.Entity==None:
            self.errortext.Content += 'no line selected\n'
        if File.Exists(self.openfilename.Text)==False:
            self.errortext.Content += 'no file selected\n'
        
        if self.linepicker1.Entity!=None and File.Exists(self.openfilename.Text):   
        
            self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
            UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())

            wv = self.currentProject [Project.FixedSerial.WorldView]
            
                
            stationlist=[]

            with open(self.openfilename.Text,'r') as csvfile: 
                reader = csv.reader(csvfile, delimiter=',', quotechar='|') 
                for row in reader:
                    stationlist.Add(row)
            

            o = self.linepicker1.Entity
            polyseg1 = o.ComputePolySeg()
            polyseg1 = polyseg1.ToWorld()
            polyseg1_v = o.ComputeVerticalPolySeg()
            
            truestation = StationEdit()
            truestation.StationProvider = o
            

            outSegment = clr.StrongBox[Segment]()
            outPointOnCL1 = clr.StrongBox[Point3D]()
            outPointOnCL2 = clr.StrongBox[Point3D]()
            perpVector3D = clr.StrongBox[Vector3D]()
            out_t = clr.StrongBox[float]()
            outdeflectionAngle = clr.StrongBox[float]()
            pnew = Point3D()
            perpVector2D = Vector2D()
            station = float()
            offset = float()
            elev = float()
            
            if str.isnumeric(self.textBox1.Text):
                pointcolumn=int(self.textBox1.Text)
            else:
                pointcolumn=0
            
            if str.isnumeric(self.textBox2.Text):
                stationcolumn=int(self.textBox2.Text)
            else:
                stationcolumn=0
            
            if str.isnumeric(self.textBox3.Text):
                offsetcolumn=int(self.textBox3.Text)
            else:
                offsetcolumn=0

            if str.isnumeric(self.textBox4.Text):
                elevcolumn=int(self.textBox4.Text)
            else:
                elevcolumn=0
            if elevcolumn==0: self.el_checkBox.IsChecked=True

            try:
                grade = float(self.textBox5.Text)
            except:
                self.textBox5.Text = "0"
                grade = 0

            try:
                with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                    for i in range(stationlist.Count):

                        if stationlist[i].Count>0:

                            ptnr = (stationlist[i][pointcolumn-1])
                            
                            truestationlist =[]

                            if stationcolumn == 0:
                                truestationlist.Add(0)
                            else:                        
                                if o.HasEquations:
                                    for j in range (1, o.StationTable.Count + 2): # stationtable count is 1 less than we have zones
                                        truestation.ClientAreaText = str(self.toprojectunit(float(stationlist[i][stationcolumn-1]))) + ":" + str(j)
                                        if not math.isnan(truestation.Distance):
                                            truestationlist.Add(truestation.Distance)
                                else:
                                    try:
                                        tt = self.toprojectunit(float(stationlist[i][stationcolumn-1]))
                                        truestationlist.Add(self.toprojectunit(float(stationlist[i][stationcolumn-1])))
                                    except: 
                                        self.errortext.Content += 'non-numeric in station column, line ' + str(i+1) + '\n'
                                        break

                            
                            if offsetcolumn==0:
                                offset=0
                            else:
                                try:
                                    offset = self.toprojectunit(float(stationlist[i][offsetcolumn-1]))
                                except:
                                    self.errortext.Content += 'non-numeric in offset column, line ' + str(i+1) + '\n'
                                    break
                            
                            if elevcolumn==0:
                                elev=0
                            else:
                                try:
                                    elev = self.toprojectunit(float(stationlist[i][elevcolumn-1]))
                                except: 
                                    self.errortext.Content += 'non-numeric in elevation column, line ' + str(i+1) + '\n'
                                    break
                            
                            if truestationlist.Count == 0:
                                self.errortext.Content += "\ncouldn't find " + stationlist[i][stationcolumn-1] + " on Alignment"

                            if truestationlist.Count > 1:
                                self.errortext.Content += "\nfound " + stationlist[i][stationcolumn-1] + " from Inputline " + str(i + 1) + " multiple times on Alignment"
                                               
                                
                            for station in truestationlist:

                                polyseg1.FindPointFromStation(station, outSegment, out_t, outPointOnCL1, perpVector3D, outdeflectionAngle)
                                
                                if polyseg1_v == None:
                                    pnew = outPointOnCL1.Value
                                else:
                                    pnew.X = outPointOnCL1.Value.X
                                    pnew.Y = outPointOnCL1.Value.Y
                                    pnew.Z = polyseg1_v.ComputeVerticalSlopeAndGrade(station)[1]

                                if self.createpointonline.IsChecked == True:
                                    
                                    if self.createdependentpoints.IsChecked:
                                        cadPoint = wv.Add(clr.GetClrType(LocationComputer.DependentPoint))
                                        cadPoint.Layer = layer_sn
                                        cadPoint.SymbolCode = 0
                                        cadPoint.LocationComputer = LocationComputer.LocationByStation(o.SerialNumber, station, 0.0) # those are directly from the file
                                        if pointcolumn > 0:
                                            cadPoint.Name = ptnr + ' - on source'
                                        cadPoint.ElevationType = LocationComputer.LocationPointElevationType.eByDeltaElevation
                                        cadPoint.ElevationDeltaElev = 0 # that's the value directly from the file
                                    else:
                                        if pointcolumn == 0:
                                            cadPoint = wv.Add(clr.GetClrType(CadPoint))
                                            cadPoint.Layer = layer_sn
                                            cadPoint.Point0 = pnew
                                        else:
                                            cadPoint = CoordPoint.CreatePoint(self.currentProject,  ptnr + ' - on source line')
                                            cadPoint.Layer = layer_sn
                                            cadPoint.AddPosition(pnew)


                                perpVector = perpVector3D.Value
                                perpVector.Z = 0
                                perpVector.Length = offset

                                pnew.X = pnew.X + perpVector.X
                                pnew.Y = pnew.Y + perpVector.Y

                                if self.relativeelev.IsChecked:
                                    if offset == 0.0:
                                        pnew.Z = pnew.Z + elev
                                    else:
                                        pnew.Z = pnew.Z + (((elev/abs(offset)) + (grade/100)) * abs(offset))
                                elif self.absoluteelev.IsChecked:
                                    if offset == 0.0:
                                        pnew.Z = elev
                                    else:
                                        pnew.Z = (((elev/abs(offset)) + (grade/100)) * abs(offset))
                                
                                if self.createdependentpoints.IsChecked:
                                    cadPoint = wv.Add(clr.GetClrType(LocationComputer.DependentPoint))
                                    cadPoint.Layer = layer_sn
                                    cadPoint.SymbolCode = 0
                                    cadPoint.LocationComputer = LocationComputer.LocationByStation(o.SerialNumber, station, offset) # those are directly from the file
                                    if pointcolumn > 0:
                                        cadPoint.Name = ptnr
                                    if self.relativeelev.IsChecked:
                                        cadPoint.ElevationType = LocationComputer.LocationPointElevationType.eByDeltaElevation
                                        cadPoint.ElevationDeltaElev = elev # that's the value directly from the file
                                    elif self.absoluteelev.IsChecked:
                                        cadPoint.ElevationType = LocationComputer.LocationPointElevationType.eElevationDefined
                                        cadPoint.Position = Point3D(cadPoint.Position.X, cadPoint.Position.Y, elev)
                                        tt = 1
                                    
                                else:
                                    if pointcolumn == 0:
                                        cadPoint = wv.Add(clr.GetClrType(CadPoint))
                                        cadPoint.Layer = layer_sn
                                        cadPoint.Point0 = pnew
                                    else:
                                        cadPoint = CoordPoint.CreatePoint(self.currentProject, ptnr)
                                        cadPoint.Layer = layer_sn
                                        cadPoint.AddPosition(pnew)



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
        
        Keyboard.Focus(self.linepicker1)
        self.SaveOptions()


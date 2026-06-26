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
    "openfilename":             os.path.expanduser('~\\Downloads'),
    "layerpicker":              8,
    "createdependentpoints":    False,
    "createpointonline":        False,
    "unitpicker":               "Meter",
    "textBox1":                 "",
    "textBox2":                 "2",
    "textBox3":                 "3",
    "textBox4":                 "4",
    "relativeelev":             True,
    "absoluteelev":             False,
    "textBox5":                 "0",
}

def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_ImportStaOffEl"
    cmdData.CommandName = "SCR_ImportStaOffEl"
    cmdData.Caption = "_SCR_ImportStaOffEl"
    cmdData.UIForm = "SCR_ImportStaOffEl"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "Import"
        cmdData.ShortCaption = "StaOffEl CSV to Points"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.15
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
        SCROptions.LoadMacroOptions(self, "SCR_ImportStaOffEl", _OPTIONS, self.currentProject)

    def SaveOptions(self):
        SCROptions.SaveMacroOptions(self, "SCR_ImportStaOffEl", _OPTIONS)

    def drawoverlay(self):

        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)
        self.overlayBag = OverlayBag(self.ViewOverlay) # create a new/empty overlaybag

        l1 = self.linepicker1.Entity

        if l1:
            self.overlayBag.AddPolyline(SCROverlayBag.getpolypoints(l1), Color.Green.ToArgb(), 4)

            for p in SCROverlayBag.getarrowlocations(l1, 10): # returns list with location and perp right azimuth [Point3D, perpVector3D.Value.Azimuth]
                self.overlayBag.AddMarker(p[0], GraphicMarkerTypes.Arrow_IndependentColor, Color.Orange.ToArgb(), "", 0, math.pi - p[1], 3.0)

        # if it only needs to be visible in all Planview then remove the Hoops3DViewGUID
        array = Array[Guid]([DisplayWindow.Hoops3DViewGUID, DisplayWindow.HoopsPlanViewGUID])
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.AddOverlayGeometry(array, self.overlayBag)

        return

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
        
            UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
            self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)

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
                pointcolumn = int(self.textBox1.Text)
            else:
                pointcolumn = 0
            
            if str.isnumeric(self.textBox2.Text):
                stationcolumn = int(self.textBox2.Text)
            else:
                stationcolumn = 0
            
            if str.isnumeric(self.textBox3.Text):
                offsetcolumn=int(self.textBox3.Text)
            else:
                offsetcolumn = 0

            if str.isnumeric(self.textBox4.Text):
                elevcolumn = int(self.textBox4.Text)
            else:
                elevcolumn = 0
            if elevcolumn == 0: self.relativeelev.IsChecked = True

            if str.isnumeric(self.textBox6.Text):
                fccolumn = int(self.textBox6.Text)
            else:
                fccolumn = 0

            try:
                grade = float(self.textBox5.Text)
            except:
                self.textBox5.Text = "0"
                grade = 0

            fm = FeatureManager.Provide(self.currentProject)
            fcm = FeatureCodeManager.Provide(self.currentProject)

            for o2 in self.currentProject:
            #find PointManager as object
                if isinstance(o2, PointManager):
                    pm = o2

            try:
                with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                    for i in range(stationlist.Count):

                        if stationlist[i].Count > 0:

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

                            
                            if offsetcolumn == 0:
                                offset = 0
                            else:
                                try:
                                    offset = self.toprojectunit(float(stationlist[i][offsetcolumn-1]))
                                except:
                                    self.errortext.Content += 'non-numeric in offset column, line ' + str(i+1) + '\n'
                                    break
                            
                            if elevcolumn == 0:
                                elev = 0
                            else:
                                try:
                                    elev = self.toprojectunit(float(stationlist[i][elevcolumn-1]))
                                except: 
                                    self.errortext.Content += 'non-numeric in elevation column, line ' + str(i+1) + '\n'
                                    break

                            if fccolumn == 0:
                                fcdic = None
                            else:
                                try:
                                    fcdic = self.fcstringtodic(stationlist[i][fccolumn-1])
                                except: 
                                    self.errortext.Content += 'error parsing FC, line ' + str(i+1) + '\n'
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
                                        namedPoint = CoordPoint.CreatePoint(self.currentProject, ptnr)
                                        namedPoint.Layer = layer_sn
                                        namedPoint.AddPosition(pnew)

                                        for setfc in fcdic:
                                            pm.SetFeatureCodeAtPoint(namedPoint.SerialNumber, setfc)
                                            features = pm.AssociatedRDFeatures(namedPoint.SerialNumber)
                                            # try to fill the feature attributes from the dictionary
                                            for f in features:
                                                for attr in f.Definition.AttributeDefinitions:
                                                    # lookup values from dictionary
                                                    attrval = fcdic.get(f.Code, {}).get(attr.Name, "")
                                                    # set attribute value
                                                    f.Add(attr.Type, attr.Name, attrval)



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
        
        Keyboard.Focus(self.linepicker1)
        self.SaveOptions()

    def fcstringtodic(self, fcstring):

        fcdic = {}

        for fc in fcstring.split("|"):

            att = fc.split(":")
            if att.Count == 3: # only try to add to dict if we have a "FC:Attribute:Value" combination

                if not att[0] in fcdic.keys(): # if the FC doesn't exist in the dict yet add it
                    fcdic.update({att[0] : {}})
                fcdic[att[0]].update({att[1] : att[2]}) # add attribute and value to FC key

        return fcdic
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
    cmdData.Key = "SCR_ExifGPS"
    cmdData.CommandName = "SCR_ExifGPS"
    cmdData.Caption = "_SCR_ExifGPS"
    cmdData.UIForm = "SCR_ExifGPS"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Expld-SNR-Relay-Prop"
        cmdData.DefaultTabGroupKey = "Properties"
        cmdData.ShortCaption = "EXIF GPS"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.10
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "retrieve EXIF GPS metadata"
        cmdData.ToolTipTextFormatted = "retrieve EXIF GPS metadata"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png") 
        cmdData.ImageSmall = b
    except:
        pass


class SCR_ExifGPS(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_ExifGPS.xaml") as s:
            wpf.LoadComponent (self, s)
        self.currentProject = currentProject
        self.macroFileFolder = macroFileFolder
        self.exiftools_exists = False

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

        self.linepicker1.IsEntityValidCallback = self.IsValid
        self.linepicker1.ValueChanged += self.lineChanged
        self.lType = clr.GetClrType(IPolyseg)
        self.coordpointType = clr.GetClrType(CoordPoint)

        self.objs.IsEntityValidCallback = self.IsValidcoordpoint
        optionMenu = SelectionContextMenuHandler()
        # remove options that don't apply here
        optionMenu.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        self.objs.ButtonContextMenu = optionMenu

        UIEvents.AfterDataProcessing += self.backgroundupdateend

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def SetDefaultOptions(self):
        self.retrievegps.IsChecked = OptionsManager.GetBool("SCR_ExifGPS.retrievegps", True)
        self.openfilename.Text = OptionsManager.GetString("SCR_ExifGPS.openfilename", os.path.expanduser('~\\Downloads'))

        lserial = OptionsManager.GetUint("SCR_ExifGPS.layerpicker", 23)
        o = self.currentProject.Concordance.Lookup(lserial) # get the object with the saved serial number
        if o != None:   # check if we actually got an object back
            if isinstance(o.GetSite(), LayerCollection):    # check if the object is a layer
                self.layerpicker.SetSelectedSerialNumber(lserial, InputMethod(3))
            else:  # if not then set Layer zero                       
                self.layerpicker.SetSelectedSerialNumber(23, InputMethod(3))
        else:   # if not then set Layer zero               
            self.layerpicker.SetSelectedSerialNumber(23, InputMethod(3))

        self.writechainagetofile.IsChecked = OptionsManager.GetBool("SCR_ExifGPS.writechainagetofile", False)

        self.copyfiles.IsChecked = OptionsManager.GetBool("SCR_ExifGPS.copyfiles", False)
        self.copytargetfolder.Text = OptionsManager.GetString("SCR_ExifGPS.copytargetfolder", os.path.expanduser('~\\Downloads'))

        lserial = OptionsManager.GetUint("SCR_ExifGPS.layerpickergood", 23)
        o = self.currentProject.Concordance.Lookup(lserial) # get the object with the saved serial number
        if o != None:   # check if we actually got an object back
            if isinstance(o.GetSite(), LayerCollection):    # check if the object is a layer
                self.layerpickergood.SetSelectedSerialNumber(lserial, InputMethod(3))
            else:  # if not then set Layer zero                       
                self.layerpickergood.SetSelectedSerialNumber(23, InputMethod(3))
        else:   # if not then set Layer zero               
            self.layerpickergood.SetSelectedSerialNumber(23, InputMethod(3))

        lserial = OptionsManager.GetUint("SCR_ExifGPS.layerpickerbad", 8)
        o = self.currentProject.Concordance.Lookup(lserial) # get the object with the saved serial number
        if o != None:   # check if we actually got an object back
            if isinstance(o.GetSite(), LayerCollection):    # check if the object is a layer
                self.layerpickerbad.SetSelectedSerialNumber(lserial, InputMethod(3))
            else:  # if not then set Layer zero                       
                self.layerpickerbad.SetSelectedSerialNumber(8, InputMethod(3))
        else:   # if not then set Layer zero               
            self.layerpickerbad.SetSelectedSerialNumber(8, InputMethod(3))

        self.backgroundupdateend(None, None)

    def SaveOptions(self):
        OptionsManager.SetValue("SCR_ExifGPS.retrievegps", self.retrievegps.IsChecked)
        OptionsManager.SetValue("SCR_ExifGPS.openfilename", self.openfilename.Text)
        OptionsManager.SetValue("SCR_ExifGPS.layerpicker", self.layerpicker.SelectedSerialNumber)
        OptionsManager.SetValue("SCR_ExifGPS.writechainagetofile", self.writechainagetofile.IsChecked)

        OptionsManager.SetValue("SCR_ExifGPS.copyfiles", self.copyfiles.IsChecked)
        OptionsManager.SetValue("SCR_ExifGPS.copytargetfolder", self.copytargetfolder.Text)
        OptionsManager.SetValue("SCR_ExifGPS.layerpickergood", self.layerpickergood.SelectedSerialNumber)
        OptionsManager.SetValue("SCR_ExifGPS.layerpickerbad", self.layerpickerbad.SelectedSerialNumber)


    def drawoverlay(self):

        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)
        self.overlayBag = OverlayBag(self.ViewOverlay) # create a new/empty overlaybag

        l1 = self.linepicker1.Entity

        if l1 != None:
            self.overlayBag.AddPolyline(self.getpolypoints(l1), Color.Green.ToArgb(), 4)

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

    def lineChanged(self, ctrl, e):
        l1 = self.linepicker1.Entity
        if l1 != None:
            self.drawoverlay()

    def IsValid(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        return False

    def IsValidcoordpoint(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.coordpointType):
            return True
        return False

    def relayerbutton_Click(self, sender, e):

        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        try:
            # the "with" statement will unroll any changes if something go wrong
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                for o in self.objs:
                    if isinstance(o, self.coordpointType):
                        
                        if sender.Name == "relayertogoodbutton":
                            o.Layer = self.layerpickergood.SelectedSerialNumber
                        elif sender.Name == "relayertobadbutton":
                            o.Layer = self.layerpickerbad.SelectedSerialNumber

                failGuard.Commit()
                UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
                self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
            
        except Exception as err:
            tt = sys.exc_info()
            exc_type, exc_obj, exc_tb = sys.exc_info()
            # EndMark MUST be set no matter what
            # otherwise TBC won't work anymore and needs to be restarted
            self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
            UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
            self.debug.Content += '\nan Error occurred - Result probably incomplete\n' + str(exc_type) + '\n' + str(exc_obj) + '\nLine ' + str(exc_tb.tb_lineno)

        Keyboard.Focus(self.objs)
        self.SaveOptions()

    def backgroundupdateend(self, sender, e):
        
        goodlayer = self.currentProject.Concordance[self.layerpickergood.SelectedSerialNumber]
        goodlayer.DefaultColor = Color.Lime
        badlayer = self.currentProject.Concordance[self.layerpickerbad.SelectedSerialNumber]
        badlayer.DefaultColor = Color.Red
        goodcount = 0
        badcount = 0
        for o in goodlayer.Members:
            if isinstance(self.currentProject.Concordance[o], self.coordpointType):
                goodcount += 1
        for o in badlayer.Members:
            if isinstance(self.currentProject.Concordance[o], self.coordpointType):
                badcount += 1
        self.goodlayercountlabel.Content = str(goodcount)
        self.badlayercountlabel.Content = str(badcount)

    def openbutton_Click(self, sender, e):
        dialog = FolderBrowserDialog()
        dialog.Reset()
        dialog.RootFolder = Environment.SpecialFolder.Desktop
        dialog.SelectedPath = self.openfilename.Text
        SendKeys.Send("{TAB}{TAB}{RIGHT}")
        tt = dialog.ShowDialog()
        if tt == DialogResult.OK:
            self.openfilename.Text = dialog.SelectedPath #.replace("\\", "/")

    def copybutton_Click(self, sender, e):
        dialog = FolderBrowserDialog()
        dialog.Reset()
        dialog.RootFolder = Environment.SpecialFolder.Desktop
        dialog.SelectedPath = self.copytargetfolder.Text
        SendKeys.Send("{TAB}{TAB}{RIGHT}")
        tt = dialog.ShowDialog()
        if tt == DialogResult.OK:
            self.copytargetfolder.Text = dialog.SelectedPath #.replace("\\", "/")

    def radiochange_Click(self, sender, e):

        if self.retrievegps.IsChecked:
            self.okBtn.Content = "read GPS tags"
        elif self.copyfiles.IsChecked:
            self.okBtn.Content = "copy selected files on disk"

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def Dispose(self, cmd, disposing):
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)
        UIEvents.AfterDataProcessing -= self.backgroundupdateend

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.success.Content=''
        self.error.Content=''

        wv = self.currentProject [Project.FixedSerial.WorldView]
        activeForm = TrimbleOffice.TheOffice.MainWindow.AppViewManager.ActiveView
        #wv.PauseGraphicsCache(True)

        # get the unit format for chainages
        self.lunits = self.currentProject.Units.Station
        # we don't want the units to be included (so we make copy and turn that off). Otherwise get something like "12.50 ft"
        self.lfp = self.lunits.Properties.Copy()
        self.lfp.AddSuffix = False

        if self.retrievegps.IsChecked:

            inputok = True

            l1 = self.linepicker1.Entity
            
            if self.writechainagetofile.IsChecked and l1 == None: 
                self.error.Content += '\nno Line selected'
                inputok = False

            if inputok:

                ProgressBar.TBC_ProgressBar.Title = "retrieve GPS-Tags"

                # get a recursive list of all files within the search folder, list contains path and file
                jpeglist = [os.path.join(root, name)
                            for root, dirs, files in os.walk(self.openfilename.Text)
                            for name in files
                            if name.endswith(tuple([".jpg", ".JPG", ".jpeg", ".JPEG"]))]

                if jpeglist.Count > 0:

                    if self.writechainagetofile.IsChecked:
                        outPointOnCL1 = clr.StrongBox[Point3D]()    
                        station1 = clr.StrongBox[float]()
                        outbool = clr.StrongBox[bool]()

                        polyseg1 = l1.ComputePolySeg()
                        polyseg1 = polyseg1.ToWorld()
                        polyseg1_v = l1.ComputeVerticalPolySeg()
                        l1name = IName.Name.__get__(l1)
                     
                    nogeotagcount = 0
                    badrtkcount = 0
                    for i in range(0, jpeglist.Count):

                        if ProgressBar.TBC_ProgressBar.SetProgress(i * 100 // jpeglist.Count):
                            break

                        try:

                            isgeotagged, latdeg, longdeg, elev, filedate, rtkflag = self.latlongelevdate_fromexif(jpeglist[i])
                            
                            if isgeotagged:
                                
                                pnew_wv = CoordPoint.CreatePoint(self.currentProject, os.path.basename(jpeglist[i]))

                                YY = int(filedate[0])
                                mm = int(filedate[1])
                                dd = int(filedate[2])
                                HH = int(filedate[3])
                                MM = int(filedate[4])
                                SS = int(filedate[5])
                                # !!! Latitude and longtude in Radians!!!
                                keyed_coord = KeyedIn(CoordSystem.eWGS84, math.pi/180 * latdeg, math.pi/180 * longdeg, CoordQuality.eSurvey, \
                                              elev, CoordQuality.eSurvey, CoordComponentType.eEllipsHeight, \
                                              System.DateTime(YY, mm, dd, HH, MM, SS))
                                OfficeEnteredCoord.AddOfficeEnteredCoord(self.currentProject, pnew_wv, keyed_coord)
                                pnew_wv.Layer = self.layerpicker.SelectedSerialNumber
                                pnew_wv.Description1 = jpeglist[i]
                                
                                if not rtkflag or int(rtkflag) < 50:
                                    standardinputlayer = self.currentProject.Concordance[self.layerpicker.SelectedSerialNumber]
                                    nortklayer = Layer.FindOrCreateLayer(self.currentProject, standardinputlayer.Name + ' - bad RTK-Flag')
                                    nortklayer.DefaultColor = Color.Blue
                                    pnew_wv.Layer = nortklayer.SerialNumber

                                    # info from Ben from Propeller - It can have the following values: 
                                    # 0 - no positioning (ex. standard accuracy)
                                    # 16 - single-point positioning mode
                                    # 34 - RTK floating solution
                                    # 50 - RTK fixed solution
                                    if not rtkflag:
                                        pnew_wv.Description2 = "no RTK flag found"
                                    elif int(rtkflag) == 0:
                                        pnew_wv.Description2 = "0 - no positioning (ex. standard accuracy)"
                                    elif int(rtkflag) == 16:
                                        pnew_wv.Description2 = "16 - single-point positioning mode"
                                    elif int(rtkflag) == 34:
                                        pnew_wv.Description2 = "34 - RTK floating solution"

                                    badrtkcount += 1
                                else:
                                        pnew_wv.Description2 = "50 - RTK fixed solution"


                            else: # no geotag
                                nogeotagcount += 1

                                nogpstaglayer = Layer.FindOrCreateLayer(self.currentProject, standardinputlayer.Name + ' - no GPS-Tag at all')
                                nogpstaglayer.DefaultColor = Color.Magenta

                                pnew_wv = CoordPoint.CreatePoint(self.currentProject, os.path.basename(jpeglist[i]))
                                pnew_wv.AddPosition(Point3D(0,0,0))

                                pnew_wv.Layer = nogpstaglayer.SerialNumber
                                pnew_wv.Description1 = jpeglist[i]



                        except Exception as e:
                            tt = sys.exc_info()
                            exc_type, exc_obj, exc_tb = sys.exc_info()
                            self.error.Content += '\nan Error occurred - Result probably incomplete\n' + str(exc_type) + '\n' + str(exc_obj) + '\nLine ' + str(exc_tb.tb_lineno)
                            continue
                        
                        if self.writechainagetofile.IsChecked:
                            if polyseg1.FindPointFromPoint(pnew_wv.AnchorPoint, outPointOnCL1, station1):
                            
                                pnew_wv.PointID = self.lunits.Format(station1.Value, self.lfp) + " - " + os.path.basename(gpslist[i][0])
                            
                if self.writechainagetofile.IsChecked:
                    subprocess.call([exiftool, '-overwrite_original',  '-csv=' + exiftoolpath + '/imagechainages.csv', self.openfilename.Text])

                self.currentProject.Calculate(False)

                if nogeotagcount > 0:
                    self.error.Content += '\nfound ' + str(nogeotagcount) + ' images without GPS-Tag'
                if badrtkcount > 0:
                    self.error.Content += '\nfound ' + str(badrtkcount) + ' images with bad RTK tag'

         
        elif self.copyfiles.IsChecked:

            i = 0
            for o in self.objs:
                
                i += 1
                ProgressBar.TBC_ProgressBar.Title = "copy image " + str(i) + '/' + str(self.objs.Count)
                if ProgressBar.TBC_ProgressBar.SetProgress(i * 100 // self.objs.Count):
                    break

                if isinstance(o, self.coordpointType):

                    try:
                        sourcefilename = o.Description1
                        shutil.copy2(o.Description1, self.copytargetfolder.Text)

                    except Exception as e:
                        tt = sys.exc_info()
                        exc_type, exc_obj, exc_tb = sys.exc_info()
                        self.error.Content += '\nan Error occurred - Result probably incomplete\n' + str(exc_type) + '\n' + str(exc_obj) + '\nLine ' + str(exc_tb.tb_lineno)



        ProgressBar.TBC_ProgressBar.Title = ""
        self.SaveOptions()

    def latlongelevdate_fromexif(self, filename):

        latdeg, longdeg, elev, filedate = None, None, None, None

        # get EXIF information from file
        #tt = JpegMetadataReader.ReadMetadata(o.Description1)
        metadata = JpegMetadataReader.ReadMetadata(filename)

        isgeotagged = False
        rtkflag = False

        for dataentry in metadata:
            if isinstance(dataentry, MetadataFormats.Exif.GpsDirectory):
                
                geoloc = dataentry.GetGeoLocation()
                latdeg = float(geoloc.Latitude)
                longdeg = float(geoloc.Longitude)
                elev = float(dataentry.GetDescription(6).replace(" metres", "")) # Altitude

                isgeotagged = True    
            
            elif isinstance(dataentry, MetadataFormats.Exif.ExifIfd0Directory):
                filedate = dataentry.GetDescription(306).replace(" ", ":").split(":")

            elif isinstance(dataentry, MetadataFormats.Xmp.XmpDirectory):

                for xmp in dataentry.XmpMeta.Properties:

                    if xmp.Path != None and str.upper("rtkflag") in str.upper(xmp.Path):

                        rtkflag = xmp.Value

        return isgeotagged, latdeg, longdeg, elev, filedate, rtkflag


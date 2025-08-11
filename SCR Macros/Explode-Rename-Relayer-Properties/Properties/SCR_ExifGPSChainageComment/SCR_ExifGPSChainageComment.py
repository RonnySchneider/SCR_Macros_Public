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
    cmdData.Key = "SCR_ExifGPSChainageComment"
    cmdData.CommandName = "SCR_ExifGPSChainageComment"
    cmdData.Caption = "_SCR_ExifGPSChainageComment"
    cmdData.UIForm = "SCR_ExifGPSChainageComment"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Expld-SNR-Relay-Prop"
        cmdData.DefaultTabGroupKey = "Properties"
        cmdData.ShortCaption = "EXIF GPS to Chainage-Comment"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.06
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "retrieve EXIF GPS and write as Chainage into Comment"
        cmdData.ToolTipTextFormatted = "retrieve EXIF GPS and write as Chainage into Comment"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png") 
        cmdData.ImageSmall = b
    except:
        pass


class SCR_ExifGPSChainageComment(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_ExifGPSChainageComment.xaml") as s:
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

        self.exiftools_exists = os.path.isfile(self.macroFileFolder + "\\exiftool.exe") 
        if not self.exiftools_exists:
            VceMessageBox.ShowError("this Macro requires 'exiftool.exe' from Phil Harvey\n\nplease download from\n\nhttps://exiftool.sourceforge.net/\n\nand copy it as 'exiftool.exe' into the macro folder\n\n" + self.macroFileFolder, "exiftool.exe missing")

    def SetDefaultOptions(self):
        self.openfilename.Text = OptionsManager.GetString("SCR_ExifGPSChainageComment.openfilename", os.path.expanduser('~\\Downloads'))

    def SaveOptions(self):
        OptionsManager.SetValue("SCR_ExifGPSChainageComment.openfilename", self.openfilename.Text)

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
        l1=self.linepicker1.Entity
        if l1 != None:
            self.drawoverlay()

    def IsValid(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        return False

    def openbutton_Click(self, sender, e):
        dialog = FolderBrowserDialog()
        dialog.Reset()
        #dialog.RootFolder = Environment.SpecialFolder.MyComputer
        dialog.SelectedPath = self.openfilename.Text
        tt = dialog.ShowDialog()
        if tt==DialogResult.OK:
            self.openfilename.Text = dialog.SelectedPath.replace("\\", "/")

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def Dispose(self, cmd, disposing):
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.success.Content=''
        self.error.Content=''

        wv = self.currentProject [Project.FixedSerial.WorldView]
        #wv.PauseGraphicsCache(True)

        # get the unit format for chainages
        self.lunits = self.currentProject.Units.Station
        # we don't want the units to be included (so we make copy and turn that off). Otherwise get something like "12.50 ft"
        self.lfp = self.lunits.Properties.Copy()
        self.lfp.AddSuffix = False

        ## # find the ProjectExplorer-"Media Files" Folder as object
        ## fpc = FilePropertiesContainer.ProvideFilePropertiesContainer(self.currentProject)
        ## # find the ProjectExplorer-MediaFilesFolder as object
        ## fpc = FilePropertiesContainer.ProvideFilePropertiesContainer(self.currentProject)

        inputok = True

        l1 = self.linepicker1.Entity
        
        if l1 == None: 
            self.error.Content += '\nno Line selected'
            inputok = False
        if not self.exiftools_exists: 
            self.error.Content += '\nexiftool.exe missing'
            inputok = False

        if inputok:

            ProgressBar.TBC_ProgressBar.Title = "retrieve GPS-Tags with ExifTool"
            if ProgressBar.TBC_ProgressBar.SetProgress(0):
                return   # function returns true if user pressed cancel

            exiftoolpath = os.path.dirname(__file__).replace("\\", "/")
            
            exiftool = exiftoolpath + "/exiftool.exe"

            #exiftool_retrieve = '"' + exiftool + '" -n -gpslongitude -gpslatitude -gpsaltitude -csv "' + self.openfilename.Text + '" > "' + exiftoolpath + '\gpstmp.csv"'
            #exiftool_retrieve = '"' + exiftool + '" -n -gpslongitude -gpslatitude -gpsaltitude -datetimeoriginal -ext jpg -csv "' + self.openfilename.Text + '" > "' + exiftoolpath + '/gpstmp.csv"'
            subprocess.call([exiftool, '-n', '-gpslongitude', '-gpslatitude', '-gpsaltitude',
                    '-datetimeoriginal', '-ext jpg', '-csv', self.openfilename.Text, '>', exiftoolpath  + '/gpstmp.csv'], shell = True)
            
            with open(exiftoolpath  + '/gpstmp.csv','r') as csvfile: 
                reader = csv.reader(csvfile, delimiter=',', quotechar='|') 
                gpslist = []
                for row in reader:
                    if row[3] == "": row[3] = "0"
                    if row[1] == "" or row[2] == "": continue
                    gpslist.Add(row)
            

            outPointOnCL1 = clr.StrongBox[Point3D]()    
            station1 = clr.StrongBox[float]()
            outbool = clr.StrongBox[bool]()

            polyseg1 = l1.ComputePolySeg()
            polyseg1 = polyseg1.ToWorld()
            polyseg1_v = l1.ComputeVerticalPolySeg()
            l1name = IName.Name.__get__(l1)
             
            imagechainages = exiftoolpath + "/imagechainages.csv"
            if File.Exists(imagechainages):
                    File.Delete(imagechainages)
            with open(imagechainages, 'w') as f:            

                f.write("SourceFile,Title\n")
                 
                ProgressBar.TBC_ProgressBar.Title = "create Points/computing Chainages"
                if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(1./3.)):
                    return   # function returns true if user pressed cancel


                for i in range(1, gpslist.Count):

                    try:
                        pnew_wv = CoordPoint.CreatePoint(self.currentProject, os.path.basename(gpslist[i][0]))

                        filedate = gpslist[i][4].replace(" ", ":").split(":")
                        YY = int(filedate[0])
                        mm = int(filedate[1])
                        dd = int(filedate[2])
                        HH = int(filedate[3])
                        MM = int(filedate[4])
                        SS = int(filedate[5])
                        # !!! Latitude and Longitude in Radians!!!
                        keyed_coord = KeyedIn(CoordSystem.eWGS84, math.pi/180 * float(gpslist[i][2]), math.pi/180 * float(gpslist[i][1]), CoordQuality.eSurvey, \
                                      float(gpslist[i][3]), CoordQuality.eSurvey, CoordComponentType.eEllipsHeight, \
                                      System.DateTime(YY, mm, dd, HH, MM, SS))
                        OfficeEnteredCoord.AddOfficeEnteredCoord(self.currentProject, pnew_wv, keyed_coord)
                        pnew_wv.Layer = self.layerpicker.SelectedSerialNumber

                    except:
                        self.success.Content += '\nsome error occurred'
                        continue

                    if polyseg1.FindPointFromPoint(pnew_wv.AnchorPoint, outPointOnCL1, station1):
                    
                        f.write(gpslist[i][0] + "," + \
                                self.lunits.Format(station1.Value, self.lfp) + " - " + l1name + "\n")

                        pnew_wv.PointID = self.lunits.Format(station1.Value, self.lfp) + " - " + os.path.basename(gpslist[i][0])
                    #fpo = fpc.FindOrCreate(os.path.basename(gpslist[i][0]), outbool)
                    #fpo.FullName = gpslist[i][0]
                    
                    #mf.Add(fpo)

            f.close()

            ProgressBar.TBC_ProgressBar.Title = "write Chainage data using ExifTool"
            if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(2./3.)):
                return   # function returns true if user pressed cancel

            #exiftool_writeback = '"' + exiftool + '" -overwrite_original -csv="' + exiftoolpath + '/imagechainages.csv" "' + self.openfilename.Text + '"'
            #exiftoolreturn = StringIO(os.popen(exiftool_writeback).read())
            
            subprocess.call([exiftool, '-overwrite_original',  '-csv=' + exiftoolpath + '/imagechainages.csv', self.openfilename.Text])

            ProgressBar.TBC_ProgressBar.Title = ""
            

        self.currentProject.Calculate(False)
        self.SaveOptions()
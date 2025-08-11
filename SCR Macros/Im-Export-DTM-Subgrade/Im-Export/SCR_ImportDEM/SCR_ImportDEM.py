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
    cmdData.Key = "SCR_ImportDEM"
    cmdData.CommandName = "SCR_ImportDEM"
    cmdData.Caption = "_SCR_ImportDEM"
    cmdData.UIForm = "SCR_ImportDEM"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "Import/Export"
        cmdData.ShortCaption = "Import DEM"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.02
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Import a DEM ASCII file"
        cmdData.ToolTipTextFormatted = "Import a DEM ASCII file"

    except:
        pass
    
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_ImportDEM(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_ImportDEM.xaml") as s:
            wpf.LoadComponent (self, s)
        self.currentProject = currentProject
        self.macroFileFolder = macroFileFolder
        if self.openfilename.Text=='': self.openfilename.Text = macroFileFolder

    def HelpClicked(self, cmd, e):
        webbrowser.open("C:\ProgramData\Trimble\MacroCommands3\SCR Macros\MacroHelp\MacroHelp.htm#" + type(self).__name__)

    def OnLoad(self, cmd, buttons, event):
        self.okBtn = buttons[0]
        buttons[2].Content = "Help"
        buttons[2].Visibility = Visibility.Visible
        buttons[2].Click += self.HelpClicked
        self.Caption = cmd.Command.Caption

        self.columncount.NumberOfDecimals = 0
        self.rowcount.NumberOfDecimals = 0

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass


    def SetDefaultOptions(self):
        self.openfilename.Text = OptionsManager.GetString("SCR_ImportDEM.openfilename", os.path.expanduser('~\\Downloads'))

        self.gridcoords.IsChecked = OptionsManager.GetBool("SCR_ImportDEM.gridcoords", True)
        self.globalcoords.IsChecked = OptionsManager.GetBool("SCR_ImportDEM.globalcoords", False)
        self.cellcorners.IsChecked = OptionsManager.GetBool("SCR_ImportDEM.cellcorners", True)
        self.cellcenter.IsChecked = OptionsManager.GetBool("SCR_ImportDEM.cellcenter", False)

            
    def SaveOptions(self):
        OptionsManager.SetValue("SCR_ImportDEM.openfilename", self.openfilename.Text)

        OptionsManager.SetValue("SCR_ImportDEM.gridcoords", self.gridcoords.IsChecked)
        OptionsManager.SetValue("SCR_ImportDEM.globalcoords", self.globalcoords.IsChecked)
        OptionsManager.SetValue("SCR_ImportDEM.cellcorners", self.cellcorners.IsChecked)
        OptionsManager.SetValue("SCR_ImportDEM.cellcenter", self.cellcenter.IsChecked)

    def IsValid(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        return False

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def openbutton_Click(self, sender, e):
        dialog = OpenFileDialog()
        dialog.InitialDirectory = self.openfilename.Text
        dialog.Filter=("DEM|*.asc; *.dem")
        
        tt = dialog.ShowDialog()
        if tt == DialogResult.OK:
            self.openfilename.Text = dialog.FileName

        self.rowlist = []
        with open(self.openfilename.Text,'r') as csvfile: 
            reader = csv.reader(csvfile, delimiter=' ', quotechar='|') 
            for row in reader:
                self.rowlist.Add(row)

        self.fileheader.Text = ''
        for i in range(0, 6):
            for all in self.rowlist[i]:
                self.fileheader.Text += all + ' '

            self.fileheader.Text += '\n'

            if "COL" in str.upper(self.rowlist[i][0]):
                self.columncount.Value = int(self.rowlist[i][len(self.rowlist[i]) - 1])
            if "ROW" in str.upper(self.rowlist[i][0]):
                tt = len(self.rowlist[i]) - 1
                self.rowcount.Value = int(self.rowlist[i][len(self.rowlist[i]) - 1])
            if "X" in str.upper(self.rowlist[i][0]) and "CORNER" in str.upper(self.rowlist[i][0]):
                self.btmleftx.NumberOfDecimals = len(self.rowlist[i][len(self.rowlist[i]) - 1]) - self.rowlist[i][len(self.rowlist[i]) - 1].find(".") - 1
                self.btmleftx.Value = float(self.rowlist[i][len(self.rowlist[i]) - 1])
            if "Y" in str.upper(self.rowlist[i][0]) and "CORNER" in str.upper(self.rowlist[i][0]):
                self.btmlefty.NumberOfDecimals = len(self.rowlist[i][len(self.rowlist[i]) - 1]) - self.rowlist[i][len(self.rowlist[i]) - 1].find(".") - 1
                self.btmlefty.Value = float(self.rowlist[i][len(self.rowlist[i]) - 1])
            if "CELL" in str.upper(self.rowlist[i][0]):
                self.cellsize.NumberOfDecimals = len(self.rowlist[i][len(self.rowlist[i]) - 1]) - self.rowlist[i][len(self.rowlist[i]) - 1].find(".") - 1
                self.cellsize.Value = float(self.rowlist[i][len(self.rowlist[i]) - 1])
            if "NO" in str.upper(self.rowlist[i][0]) and "DATA" in str.upper(self.rowlist[i][0]):
                self.nodata.Text = self.rowlist[i][len(self.rowlist[i]) - 1]

        if abs(self.btmlefty.Value) < 360 and abs(self.btmleftx.Value) < 360:
            self.globalcoords.IsChecked = True
        else:
            self.gridcoords.IsChecked = True


    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''

        if File.Exists(self.openfilename.Text) == False:
            self.error.Content += 'no file selected\n'
        
        
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())

        wv = self.currentProject [Project.FixedSerial.WorldView]
            
        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                maxx = int(self.columncount.Value)
                maxy = int(self.rowcount.Value)

                xll = self.btmleftx.Value
                yll = self.btmlefty.Value

                step = self.cellsize.Value
                nodata = float(self.nodata.Text)

                #  https://surferhelp.goldensoftware.com/subsys/subsys_ASC_Arc_Info_ASCII_Grid.htm
                #  the position is given for the bottom left, and the elevations in the matrix follow standard x/y
                #  so bottom left corner is in the first column and last row

                ptslist = []
                if self.gridcoords.IsChecked:
                    stepadjust = -0.001 # for X/Y
                else:
                    stepadjust = -0.00000001 # for lat/long

                ProgressBar.TBC_ProgressBar.Title = "computing Cell Values"
                for y in range(maxy):
                    if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(y * 100 / maxy)):
                        break   # function returns true if user pressed cancel
                
                    for x in range(maxx):
                
                        if float(self.rowlist[6 + y][x]) - nodata != 0:
                            tt = self.rowlist[6 + y][x]
                            if self.cellcenter.IsChecked:
                                # cell center
                                pll = Point3D()
                                pll.Y = yll + ((maxy - y-1) * step) + step/2 # latitude
                                pll.X = xll + (x * step) + step/2   # longitude
                                pll.Z = float(self.rowlist[6 + y][x])
                                ptslist.Add(pll)

                            else:
                                # cell lower left
                                pll = Point3D()
                                #Sheet2.Cells(output, 1) = output
                                pll.Y = yll + ((maxy - y-1) * step) # latitude
                                pll.X = xll + (x * step)    # longitude
                                pll.Z = float(self.rowlist[6 + y][x])
                                ptslist.Add(pll)
                                #output = output + 1
                                
                                # cell upper left
                                pul = Point3D()
                                #Sheet2.Cells(output, 1) = output
                                pul.Y = yll + ((maxy - y-1) * step) + step + stepadjust    # latitude
                                pul.X = xll + (x * step)    # longitude
                                pul.Z = float(self.rowlist[6 + y][x])
                                ptslist.Add(pul)
                                #output = output + 1
                                
                                # cell lower right
                                plr = Point3D()
                                #Sheet2.Cells(output, 1) = output
                                plr.Y = yll + ((maxy - y-1) * step)   # latitude
                                plr.X = xll + (x * step) + step + stepadjust    # longitude
                                plr.Z = float(self.rowlist[6 + y][x])
                                ptslist.Add(plr)
                                #output = output + 1
                                #
                                # cell upper right
                                pur = Point3D()
                                #Sheet2.Cells(output, 1) = output
                                pur.Y = yll + ((maxy - y-1) * step) + step + stepadjust    # latitude
                                pur.X = xll + (x * step) + step + stepadjust    # longitude
                                pur.Z = float(self.rowlist[6 + y][x])
                                ptslist.Add(pur)
                                #output = output + 1
                
                ProgressBar.TBC_ProgressBar.Title = "creating new surface"
                newSurface = wv.Add(clr.GetClrType(Model3D))
                
                #tt = os.path.splitext(os.path.basename(self.openfilename.Text))[0]
                tt = os.path.basename(self.openfilename.Text)
                if self.cellcenter.IsChecked:
                    tt += " - interpolated cell centers"
                else:
                    tt += " - stepped cells"
                newSurface.Name = Model3D.GetUniqueName(tt, None, wv) #make sure name is unique
                if self.gridcoords.IsChecked:
                    newSurface.MaxEdgeLength = step * 2

                builder = newSurface.GetGemBatchBuilder()
                
                output = 1
                pnew_wv = CoordPoint.CreatePoint(self.currentProject, "tmp")
                for p in ptslist:

                    if self.gridcoords.IsChecked:
                        v = builder.AddVertex(p)

                    else:

                        # !!! Latitude and Longitude in Radians!!!
                        keyed_coord = KeyedIn(CoordSystem.eWGS84, math.pi/180 * p.Y, math.pi/180 * p.X, CoordQuality.eSurvey, \
                                      p.Z, CoordQuality.eSurvey, CoordComponentType.eElevation, \
                                      System.DateTime.UtcNow)

                        OfficeEnteredCoord.AddOfficeEnteredCoord(self.currentProject, pnew_wv, keyed_coord)
                        
                        v = builder.AddVertex(pnew_wv.AnchorPoint)

                pnew_wv.GetSite().Remove(pnew_wv.SerialNumber)
                    
                    
                    #cadPoint = wv.Add(clr.GetClrType(CadPoint))
                    #cadPoint.Point0 = pnew_wv.AnchorPoint

                    #output += 1


                builder.Construction()
                builder.Commit()

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


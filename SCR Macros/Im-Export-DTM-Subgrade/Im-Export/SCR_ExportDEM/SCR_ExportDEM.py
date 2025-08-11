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
    cmdData.Key = "SCR_ExportDEM"
    cmdData.CommandName = "SCR_ExportDEM"
    cmdData.Caption = "_SCR_ExportDEM"
    cmdData.UIForm = "SCR_ExportDEM"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "Import/Export"
        cmdData.ShortCaption = "Export DEM"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.04
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Export a surface as DEM Grid ASCII file"
        cmdData.ToolTipTextFormatted = "Export a surface as DEM Grid ASCII file"

    except:
        pass
    
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_ExportDEM(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_ExportDEM.xaml") as s:
            wpf.LoadComponent (self, s)
        self.currentProject = currentProject
        self.macroFileFolder = macroFileFolder
        #if self.openfilename.Text=='': self.openfilename.Text = macroFileFolder

    def HelpClicked(self, cmd, e):
        webbrowser.open("C:\ProgramData\Trimble\MacroCommands3\SCR Macros\MacroHelp\MacroHelp.htm#" + type(self).__name__)

    def OnLoad(self, cmd, buttons, event):
        self.okBtn = buttons[0]
        buttons[2].Content = "Help"
        buttons[2].Visibility = Visibility.Visible
        buttons[2].Click += self.HelpClicked
        self.Caption = cmd.Command.Caption

        types = Array[Type](SurfaceTypeLists.AllWithCutFillMap) #+Array[Type]([clr.GetClrType(ProjectedSurface)])
        self.surfacepicker.FilterByEntityTypes = types
        self.surfacepicker.AllowNone = False
        self.cellsize.ValueChanged += self.cellsizeChanged
        self.cellsizedeg.ValueChanged += self.cellsizedegChanged

        self.error.Content += '\ncurrently hard coded output to'
        self.error.Content += '\n' + os.path.expanduser('~/Downloads/DEM-Test.asc')

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass


    def SetDefaultOptions(self):
        #self.openfilename.Text = OptionsManager.GetString("SCR_ExportDEM.openfilename", os.path.expanduser('~\\Downloads'))
        # Select surface
        try:    self.surfacepicker.SelectIndex(OptionsManager.GetInt("SCR_ExportDEM.surfacepicker", 0))
        except: self.surfacepicker.SelectIndex(0)
        self.cellsize.Distance = OptionsManager.GetDouble("SCR_ExportDEM.cellsize", 0.1)

            
    def SaveOptions(self):
        #OptionsManager.SetValue("SCR_ExportDEM.openfilename", self.openfilename.Text)
        OptionsManager.SetValue("SCR_ExportDEM.surfacepicker", self.surfacepicker.SelectedIndex)
        OptionsManager.SetValue("SCR_ExportDEM.cellsize", self.cellsize.Distance)

    def IsValid(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        return False

    def cellsizeChanged(self, ctrl, e):

        self.cellsizedeg.Angle = math.atan(self.cellsize.Distance / 6380000)

    def cellsizedegChanged(self, ctrl, e):

        self.cellsize.Distance = math.tan(self.cellsizedeg.Angle) * 6380000

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def openbutton_Click(self, sender, e):
        dialog=OpenFileDialog()
        dialog.InitialDirectory = self.openfilename.Text
        dialog.Filter=("DEM|*.asc; *.dem")
        
        tt=dialog.ShowDialog()
        if tt==DialogResult.OK:
            self.openfilename.Text = dialog.FileName

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''

        #if File.Exists(self.openfilename.Text) == False:
        #    self.error.Content += 'no file selected\n'
        
        
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())

        wv = self.currentProject [Project.FixedSerial.WorldView]
            
        if self.gridoutput.IsChecked:
            cellsize = self.cellsize.Distance
        else:
            
            keyed_coord1 = KeyedIn(CoordSystem.eLocal, self.coordpick1.Coordinate.X, self.coordpick1.Coordinate.Y, CoordQuality.eSurvey, System.DateTime.UtcNow)

            cellsize = self.cellsizedeg.Angle

        rowcount = math.ceil((self.coordpick2.Coordinate.Y - self.coordpick1.Coordinate.Y) / cellsize)
        colcount = math.ceil((self.coordpick2.Coordinate.X - self.coordpick1.Coordinate.X) / cellsize)

        rowlist = [[[None for k in range(4)] for j in range(colcount)] for i in range(rowcount)]
        # filling this array needs more code, but should be fast on large surfaces
        # no need to compute the corner elevations all over again

        #with open(self.openfilename.Text,'r') as csvfile: 
        #    reader = csv.reader(csvfile, delimiter=' ', quotechar='|') 
        #    for row in reader:
        #        rowlist.Add(row)
        #
        #tt = 1
        
        surfaceoutPoint = clr.StrongBox[Point3D]()

        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                ProgressBar.TBC_ProgressBar.Title = "computing Cell Values"
                surface = wv.Lookup(self.surfacepicker.SelectedSerial)
                for y in range(0, rowcount, 1): # fill bottom to top, but later write top to bottom
                    if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(y * 100 / rowcount)):
                        break   # function returns true if user pressed cancel

                    for x in range(0, colcount, 1):

                        #     |1 2|
                        #cell |0 3|


                        # compute the corner elevations for the current cell, if necessary
                        #Cell - BL
                        if rowlist[y][x][0] == None:
                            p = Point3D(self.coordpick1.Coordinate.X + x * cellsize, self.coordpick1.Coordinate.Y + y * cellsize)
                            if surface.PickSurface(p, surfaceoutPoint):
                                rowlist[y][x][0] = surfaceoutPoint.Value.Z
                            else:
                                rowlist[y][x][0] = -9999

                        #Cell - TL
                        if rowlist[y][x][1] == None:
                            p = Point3D(self.coordpick1.Coordinate.X + x * cellsize, self.coordpick1.Coordinate.Y + y * cellsize + cellsize)
                            if surface.PickSurface(p, surfaceoutPoint):
                                rowlist[y][x][1] = surfaceoutPoint.Value.Z
                            else:
                                rowlist[y][x][1] = -9999

                        #Cell - TR
                        if rowlist[y][x][2] == None:
                            p = Point3D(self.coordpick1.Coordinate.X + x * cellsize + cellsize, self.coordpick1.Coordinate.Y + y * cellsize + cellsize)
                            if surface.PickSurface(p, surfaceoutPoint):
                                rowlist[y][x][2] = surfaceoutPoint.Value.Z
                            else:
                                rowlist[y][x][2] = -9999

                        #Cell - BR
                        if rowlist[y][x][3] == None:
                            p = Point3D(self.coordpick1.Coordinate.X + x * cellsize + cellsize, self.coordpick1.Coordinate.Y + y * cellsize)
                            if surface.PickSurface(p, surfaceoutPoint):
                                rowlist[y][x][3] = surfaceoutPoint.Value.Z
                            else:
                                rowlist[y][x][3] = -9999
                        
                        # copy the corner values to the neighbour cells if possible, reducing the DTM calculations
                        if y < rowcount - 1:
                                rowlist[y + 1][x][0] = rowlist[y][x][1]
                                rowlist[y + 1][x][3] = rowlist[y][x][2]
                        if x < colcount - 1:
                                rowlist[y][x + 1][0] = rowlist[y][x][3]
                                rowlist[y][x + 1][1] = rowlist[y][x][2]
                        if y < rowcount - 1 and x < colcount - 1:
                                rowlist[y + 1][x + 1][0] = rowlist[y][x][2]

                ProgressBar.TBC_ProgressBar.Title = "writing File"
                filename = os.path.expanduser('~/Downloads/DEM-Test.asc')
                if File.Exists(filename):
                    File.Delete(filename)
                with open(filename, 'w') as f:         

                    f.write("ncols "+ str(colcount) + "\n")
                    f.write("nrows "+ str(rowcount) + "\n")
                    f.write("xllcorner "+ str("{:.{}f}".format(self.coordpick1.Coordinate.X, 3)) + "\n")
                    f.write("yllcorner "+ str("{:.{}f}".format(self.coordpick1.Coordinate.Y, 3)) + "\n")
                    f.write("cellsize "+ str("{:.{}f}".format(cellsize, 3)) + "\n")
                    f.write("NODATA_value "+ str("-9999") + "\n")

                    for row in reversed(rowlist):
                        outputline = ''
                        i = 0
                        for cell in row:
                            i += 1
                            if cell[0] == -9999 or cell[1] == -9999 or cell[2] == -9999 or cell[3] == -9999:
                                outputline += str(-9999)
                            else:
                                el = ((cell[0] + cell[1] + cell[2] + cell[3]) / 4)
                                outputline += str("{:.{}f}".format(el, 3))
                            
                            if i < row.Count:
                                outputline += " "
                            else:
                                outputline += "\n"

                        f.write(outputline)
                    f.close()

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
        
        ProgressBar.TBC_ProgressBar.Title = ""

        self.SaveOptions()


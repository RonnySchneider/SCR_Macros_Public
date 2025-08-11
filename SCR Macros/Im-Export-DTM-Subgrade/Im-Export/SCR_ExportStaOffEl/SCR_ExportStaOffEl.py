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
    cmdData.Key = "SCR_ExportStaOffEl"
    cmdData.CommandName = "SCR_ExportStaOffEl"
    cmdData.Caption = "_SCR_ExportStaOffEl"
    cmdData.UIForm = "SCR_ExportStaOffEl"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "Import/Export"
        cmdData.ShortCaption = "Points to StaOffEl CSV"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.06
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Export CSV with StaOffElev"
        cmdData.ToolTipTextFormatted = "compute the Stations, Offsets, Elevations of Points and Export into CSV file"

    except:
        pass
    
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_ExportStaOffEl(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_ExportStaOffEl.xaml") as s:
            wpf.LoadComponent (self, s)
        self.currentProject = currentProject
        self.macroFileFolder = macroFileFolder
        #if self.openfilename.Text=='': self.openfilename.Text = macroFileFolder

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
        self.outputunitenum = 0

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
                self.outputunitenum = e

    #def toprojectunit(self, v):
    #    return self.lunits.Convert(LinearType(self.outputunitenum), v, LinearType.Meter)
    
    def tooutputunit(self, v):
    
        #self.lfp.AddSuffix = self.addunitsuffix.IsChecked
        return self.lunits.Format(LinearType.Meter, v, self.lfp, LinearType(self.outputunitenum))

    def SetDefaultOptions(self):
        #self.openfilename.Text = OptionsManager.GetString("SCR_ExportStaOffEl.openfilename", os.path.expanduser('~\\Downloads'))

        #lserial = OptionsManager.GetUint("SCR_ExportStaOffEl.layerpicker", 8)
        #o = self.currentProject.Concordance.Lookup(lserial)
        #if o != None:
        #    if isinstance(o.GetSite(), LayerCollection):    
        #        self.layerpicker.SetSelectedSerialNumber(lserial, InputMethod(3))
        #    else:                       
        #        self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        #else:                       
        #    self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))

        #self.createpointonline.IsChecked = OptionsManager.GetBool("SCR_ExportStaOffEl.createpointonline", True)
        self.unitpicker.Text = OptionsManager.GetString("SCR_ExportStaOffEl.unitpicker", "Meter")
        #self.textBox1.Text = OptionsManager.GetString("SCR_ExportStaOffEl.textBox1", "")
        #self.textBox2.Text = OptionsManager.GetString("SCR_ExportStaOffEl.textBox2", "2")
        #self.textBox3.Text = OptionsManager.GetString("SCR_ExportStaOffEl.textBox3", "3")
        #self.textBox4.Text = OptionsManager.GetString("SCR_ExportStaOffEl.textBox4", "4")

        #self.el_checkBox.IsChecked = OptionsManager.GetBool("SCR_ExportStaOffEl.el_checkBox", True)
        
        #self.textBox5.Text = OptionsManager.GetString("SCR_ExportStaOffEl.textBox5", "0")
            
    def SaveOptions(self):
        #OptionsManager.SetValue("SCR_ExportStaOffEl.openfilename", self.openfilename.Text)

        #OptionsManager.SetValue("SCR_ExportStaOffEl.layerpicker", self.layerpicker.SelectedSerialNumber)

        #OptionsManager.SetValue("SCR_ExportStaOffEl.createpointonline", self.createpointonline.IsChecked)
        OptionsManager.SetValue("SCR_ExportStaOffEl.unitpicker", self.unitpicker.Text)
        #OptionsManager.SetValue("SCR_ExportStaOffEl.textBox1", self.textBox1.Text)
        #OptionsManager.SetValue("SCR_ExportStaOffEl.textBox2", self.textBox2.Text)
        #OptionsManager.SetValue("SCR_ExportStaOffEl.textBox3", self.textBox3.Text)
        #OptionsManager.SetValue("SCR_ExportStaOffEl.textBox4", self.textBox4.Text)
        #
        #OptionsManager.SetValue("SCR_ExportStaOffEl.el_checkBox", self.el_checkBox.IsChecked)
        #
        #OptionsManager.SetValue("SCR_ExportStaOffEl.textBox5", self.textBox5.Text)

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
            self.drawoverlay()

    def IsValid(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        return False

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def Dispose(self, cmd, disposing):
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)

    #def openbutton_Click(self, sender, e):
    #    dialog=OpenFileDialog()
    #    dialog.InitialDirectory = self.openfilename.Text
    #    dialog.Filter=("CSV|*.csv")
    #    
    #    tt=dialog.ShowDialog()
    #    if tt==DialogResult.OK:
    #        self.openfilename.Text = dialog.FileName

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''

        self.errortext.Content=''
        #layer_sn = self.layerpicker.SelectedSerialNumber

        inputok = True
        l1 = self.linepicker1.Entity
        if l1 == None: 
            self.error.Content += '\nno Line selected'
            inputok = False
        
        wv = self.currentProject [Project.FixedSerial.WorldView]

        if inputok:

            outSegment1 = clr.StrongBox[Segment]()
            out_t = clr.StrongBox[float]()
            outPointOnCL1 = clr.StrongBox[Point3D]()
            station1 = clr.StrongBox[float]()
            outaz = clr.StrongBox[Vector3D]()
            out_offset = clr.StrongBox[float]()
            out_side = clr.StrongBox[Side]()
        
            filename = os.path.expanduser('~/Downloads/Line-Offsets.csv')
            if File.Exists(filename):
                File.Delete(filename)
            with open(filename, 'w') as f:         

                polyseg1 = l1.ComputePolySeg()
                polyseg1 = polyseg1.ToWorld()
                polyseg1_v = l1.ComputeVerticalPolySeg()

                for o in self.objs:
                    if isinstance(o, CoordPoint) or isinstance(o, CadPoint):
                        
                        polyseg1.FindPointFromPoint(o.Position, outSegment1, out_t, outPointOnCL1, station1, outaz, out_offset, out_side)
                        el = 0
                        if polyseg1_v != None:
                            el = polyseg1_v.ComputeVerticalSlopeAndGrade(station1.Value)[1]

                        el = o.Position.Z - el
                        
                        if out_side.Value == Side.Right:
                            offset = out_offset.Value
                        else:
                            offset = -1 * out_offset.Value

                        outputline = ''
                        if isinstance(o, CoordPoint):
                            outputline += o.AnchorName
                        else:
                            outputline += o.Description
                        outputline += ',' + self.tooutputunit(station1.Value)
                        outputline += ',' + self.tooutputunit(offset)
                        outputline += ',' + self.tooutputunit(el)

                        f.write(outputline + "\n")

                f.close()
        
                self.success.Content += '\nDone'
        Keyboard.Focus(self.linepicker1)
        self.SaveOptions()


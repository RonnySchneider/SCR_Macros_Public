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
    cmdData.Key = "SCR_PointSurfaceSlopeToCSV"
    cmdData.CommandName = "SCR_PointSurfaceSlopeToCSV"
    cmdData.Caption = "_SCR_PointSurfaceSlopeToCSV"
    cmdData.UIForm = "SCR_PointSurfaceSlopeToCSV"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "Points"
        cmdData.ShortCaption = "Surface-Slope at Point"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.05
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Surface Slope at Point Locations to CSV"
        cmdData.ToolTipTextFormatted = "Surface Slope at Point Locations to CSV"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png") 
        cmdData.ImageSmall = b
    except:
        pass


class SCR_PointSurfaceSlopeToCSV(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_PointSurfaceSlopeToCSV.xaml") as s:
            wpf.LoadComponent (self, s)
        self.currentProject = currentProject
        self.macroFileFolder = macroFileFolder

    def HelpClicked(self, cmd, e):
        webbrowser.open("C:\ProgramData\Trimble\MacroCommands3\SCR Macros\MacroHelp\MacroHelp.htm#" + type(self).__name__)


    def OnLoad(self, cmd, buttons, event):
        self.okBtn = buttons[0]
        buttons[2].Content = "Help"
        buttons[2].Visibility = Visibility.Visible
        buttons[2].Click += self.HelpClicked
        self.Caption = cmd.Command.Caption

        optionMenu = SelectionContextMenuHandler()
        # remove options that don't apply here
        optionMenu.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        self.objs.ButtonContextMenu = optionMenu
        self.objs.IsEntityValidCallback = self.IsValid

        types = Array [Type] (SurfaceTypeLists.AllWithCutFillMap)    # we fill an array with TBC object types, we could combine different types
        self.designsurfacepicker.FilterByEntityTypes = types    # we fill the dropdownlist by applying that types array as filter
        self.designsurfacepicker.AllowNone = False              # our list shall not show an empty field

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass
        self.unitssetup(None, None)

    def SetDefaultOptions(self):
        try:    self.designsurfacepicker.SelectIndex(OptionsManager.GetInt("SCR_PointSurfaceSlopeToCSV.designsurfacepicker", 0))
        except: self.designsurfacepicker.SelectIndex(0)

        self.unitpicker.Text = OptionsManager.GetString("SCR_PointSurfaceSlopeToCSV.unitpicker", "Meter")
        self.addunitsuffix.IsChecked = OptionsManager.GetBool("SCR_PointSurfaceSlopeToCSV.addunitsuffix", False)
        self.textdecimals.Value = OptionsManager.GetDouble("SCR_PointSurfaceSlopeToCSV.textdecimals", 3)

        self.createtext.IsChecked = OptionsManager.GetBool("SCR_PointSurfaceSlopeToCSV.createtext", False)
        self.textheight.Distance = OptionsManager.GetDouble("SCR_PointSurfaceSlopeToCSV.textheight", 0.2)

        settingserial = OptionsManager.GetUint("SCR_PointSurfaceSlopeToCSV.textlayerpicker", 8) # 8 is FixedSerial for Layer Zero
        o = self.currentProject.Concordance.Lookup(settingserial)
        if o != None:
            if isinstance(o.GetSite(), LayerCollection):    
                self.textlayerpicker.SetSelectedSerialNumber(settingserial, InputMethod(3))
            else:                       
                self.textlayerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:                       
            self.textlayerpicker.SetSelectedSerialNumber(8, InputMethod(3))


    def SaveOptions(self):
        try:    # if nothing is selected it would throw an error
            OptionsManager.SetValue("SCR_PointSurfaceSlopeToCSV.designsurfacepicker", self.designsurfacepicker.SelectedIndex)
        except:
            pass
        OptionsManager.SetValue("SCR_PointSurfaceSlopeToCSV.unitpicker", self.unitpicker.Text)
        OptionsManager.SetValue("SCR_PointSurfaceSlopeToCSV.addunitsuffix", self.addunitsuffix.IsChecked)
        OptionsManager.SetValue("SCR_PointSurfaceSlopeToCSV.textdecimals", self.textdecimals.Value)
        
        OptionsManager.SetValue("SCR_PointSurfaceSlopeToCSV.createtext", self.createtext.IsChecked)
        OptionsManager.SetValue("SCR_PointSurfaceSlopeToCSV.textheight", self.textheight.Distance)
        
        OptionsManager.SetValue("SCR_PointSurfaceSlopeToCSV.textlayerpicker", self.textlayerpicker.SelectedSerialNumber)

    def IsValid(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, CoordPoint):
            return True
        if isinstance(o, CadPoint):
            return True
        return False
        
    def unitssetup(self, sender, e):
        # setup everything for the unit conversions
        self.outputunitenum = 0
        self.textdecimals.NumberOfDecimals = 0

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
        self.textdecimals.MinValue = 0.0
        self.textdecimals.ValueChanged += self.unitschanged

        self.lfp.NumberOfDecimals = int(self.textdecimals.Value)
        self.unitschanged(None, None)
    
    def unitschanged(self, sender, e):

        # find the enum for the selected LinearType
        for e in range(0, 19):
            if LinearType(e) == self.unitpicker.SelectedItem.Content:
                self.outputunitenum = e
        
        # loop through all objects of self and set the properties for all DistanceEdits
        # the code is slower than doing it manually for each single one
        # but more convenient since we don't have to worry about how many DistanceEdit Controls we have in the UI
        tt = self.__dict__.items()
        for i in self.__dict__.items():
            if i[1].GetType() == TBCWpf.DistanceEdit().GetType():
                i[1].DisplayUnit = LinearType(self.outputunitenum)
                i[1].ShowControlIcon(False)
                i[1].FormatProperty.AddSuffix = ControlBoolean(1)
                i[1].FormatProperty.NumberOfDecimals = int(self.textdecimals.Value)

    def decdecimals_Click(self, sender, e):
        if not self.textdecimals.Value == 0:
            self.textdecimals.Value -= 1
             # setup the linear format properties
            self.lfp.NumberOfDecimals = int(self.textdecimals.Value)
            self.unitschanged(None, None)

    def incdecimals_Click(self, sender, e):
        self.textdecimals.Value += 1
        # setup the linear format properties
        self.lfp.NumberOfDecimals = int(self.textdecimals.Value)
        self.unitschanged(None, None)

    def tooutputunit(self, v):
        
        self.lfp.AddSuffix = self.addunitsuffix.IsChecked
        return self.lunits.Format(LinearType.Meter, v, self.lfp, LinearType(self.outputunitenum))

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''

        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())

        self.success.Content = ''

        wv = self.currentProject [Project.FixedSerial.WorldView]
        
        surface = wv.Lookup(self.designsurfacepicker.SelectedSerial)
        if isinstance(surface,ProjectedSurface):
            projected=True
        else:
            projected=False
        
        outPoint = clr.StrongBox[Point3D]()
        outPrimitive = clr.StrongBox[Primitive]()
        outInt = clr.StrongBox[Int32]()
        outByte = clr.StrongBox[Byte]()

        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                filename = os.path.expanduser('~/Downloads/Point-Surface-Slopes.csv')
                if File.Exists(filename):
                        File.Delete(filename)
                with open(filename, 'w') as f:            

                    f.write('Pointname,Easting,Northing,Elevation,Slope [%],Surface Name' + "\n")

                    for o in self.objs:
                        if (isinstance(o, CoordPoint) or isinstance(o, CadPoint)) and projected == False:

                            if surface.PickSurface(o.Position, outPrimitive, outInt, outByte, outPoint):
                                
                                verticelist = []
                                outstring = ''

                                i = outInt.Value
                                if surface.GetTriangleMaterial(i) != surface.NullMaterialIndex():
                                    verticelist.Add(surface.GetVertexPoint(surface.GetTriangleIVertex(i,0)))
                                    verticelist.Add(surface.GetVertexPoint(surface.GetTriangleIVertex(i,1)))
                                    verticelist.Add(surface.GetVertexPoint(surface.GetTriangleIVertex(i,2)))
                                
                                if verticelist.Count != 3:
                                    continue
                                p = Plane3D(verticelist[0], verticelist[1], verticelist[2])[0]

                                if self.createtext.IsChecked:
                                    t = wv.Add(clr.GetClrType(MText))
                                    t.AlignmentPoint = o.Position
                                    t.TextString = str("{:.{}f}".format(p.Slope, int(abs(self.textdecimals.Value))))
                                    if self.addunitsuffix.IsChecked:
                                        t.TextString += ' %'
                                    try:    
                                        t.Height = self.textheight.Distance
                                    except:
                                        t.Height = 1

                                    t.Layer = self.textlayerpicker.SelectedSerialNumber


                                if (isinstance(o, CoordPoint)):
                                    outstring += o.Name + ','
                                    outstring += self.tooutputunit(o.Position.X) + ','
                                    outstring += self.tooutputunit(o.Position.Y) + ','
                                    outstring += self.tooutputunit(o.Position.Z) + ','
                                    outstring += str("{:.{}f}".format(p.Slope, int(abs(self.textdecimals.Value)))) + ','
                                    outstring += surface.Name
                                    f.write(outstring + "\n")
                                
                                if (isinstance(o, CadPoint)):
                                    outstring += 'Noname CADPoint,'
                                    outstring += self.tooutputunit(o.Position.X) + ','
                                    outstring += self.tooutputunit(o.Position.Y) + ','
                                    outstring += self.tooutputunit(o.Position.Z) + ','
                                    outstring += str("{:.{}f}".format(p.Slope, int(abs(self.textdecimals.Value)))) + ','
                                    outstring += surface.Name
                                    f.write(outstring + "\n")


                            else:
                                self.success.Content += '\nPoints seem to be outside of selected Surface'


                    #f.close() not necessary - with statement closes it automatically

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

        self.success.Content += '\nData written to\n' + os.path.expanduser('~/Downloads/Point-Surface-Slopes.csv')
        self.SaveOptions()   

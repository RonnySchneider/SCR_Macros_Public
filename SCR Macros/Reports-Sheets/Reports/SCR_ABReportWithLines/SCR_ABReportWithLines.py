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
    cmdData.Key = "SCR_ABReportWithLines"
    cmdData.CommandName = "SCR_ABReportWithLines"
    cmdData.Caption = "_SCR_ABReportWithLines"
    cmdData.UIForm = "SCR_ABReportWithLines"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Reports"
        cmdData.DefaultTabGroupKey = "Reports"
        cmdData.ShortCaption = "relative to Linework"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.23
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "AB Report relative to Linework"
        cmdData.ToolTipTextFormatted = "AB Report relative to Linework, lines need to be exploded before"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_ABReportWithLines(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_ABReportWithLines.xaml") as s:
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
        #types = Array [Type] ([CadPoint]) + Array [Type] ([Point3D])    # we fill an array with TBC object types, we could combine different types

        self.blockpicker.SearchContainer = Project.FixedSerial.BlockCollection    # put the content of the BlockCollection into the Dropdownbox
        self.blockpicker.UseSelectionEngine = False
        self.blockpicker.SetEntityType(clr.GetClrType(BlockView), self.currentProject)    # set the object type that we want to see in the list
        self.blockpicker.SelectIndex(0)

        self.objs.IsEntityValidCallback = self.IsValidPoints
        self.polyType = clr.GetClrType(IPolyseg)
        
        self.linepicker1.IsEntityValidCallback=self.IsValidLine
        
        optionMenu = SelectionContextMenuHandler()
        # remove options that don't apply here
        optionMenu.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        self.objs.ButtonContextMenu = optionMenu
        self.point3dType = clr.GetClrType(Point3D)
        self.coordpointType = clr.GetClrType(CoordPoint)
        self.cadpointType = clr.GetClrType(CadPoint)

        types = Array [Type] (SurfaceTypeLists.AllWithCutFillMap)    # we fill an array with TBC object types, we could combine different types
        self.designsurfacepicker.FilterByEntityTypes = types    # we fill the dropdownlist by applying that types array as filter
        self.designsurfacepicker.AllowNone = False              # our list shall not show an empty field

        self.coordpick.ShowElevationIf3D = True
        self.coordpick.ValueChanged += self.CoordPickChanged
        self.coordpick.AutoTab = False

        self.lType = clr.GetClrType(IPolyseg)
        #self.designthicknessentity.IsEntityValidCallback = self.IsValidThicknessEntity
        #self.asbuiltthicknessentity.IsEntityValidCallback = self.IsValidThicknessEntity

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass
        self.unitssetup(None, None)

    def SetDefaultOptions(self):
        settings = ConstructionCommandsSettings.ProvideObject(self.currentProject)

        try:    self.blockpicker.SelectIndex(settings.GetUInt32("SCR_ABReportWithLines.blockpicker", 0))
        except: self.blockpicker.SelectIndex(0)
        self.blockscale.Value = settings.GetDouble("SCR_ABReportWithLines.blockscale", 1.000)
        self.arrowstmrspec.IsChecked = settings.GetBoolean("SCR_ABReportWithLines.arrowstmrspec", True)

        self.unitpicker.Text = settings.GetString("SCR_ABReportWithLines.unitpicker", "Meter")
        self.addunitsuffix.IsChecked = settings.GetBoolean("SCR_ABReportWithLines.addunitsuffix", False)
        self.textdecimals.Value = settings.GetDouble("SCR_ABReportWithLines.textdecimals", 3)
        self.textheight.Distance = settings.GetDouble("SCR_ABReportWithLines.textheight", 0.2)
        
        self.thresh1.Distance = settings.GetDouble("SCR_ABReportWithLines.thresh1", 0.005)
        self.thresh2.Distance = settings.GetDouble("SCR_ABReportWithLines.thresh2", 0.010)
        self.thresh3.Distance = settings.GetDouble("SCR_ABReportWithLines.thresh3", 0.015)
        
        c = settings.GetInt32("SCR_ABReportWithLines.thresh1colorpicker")
        if  c == 0:
            self.thresh1colorpicker.SelectedColor = Color.Green
        else:
             self.thresh1colorpicker.SelectedColor = Color.FromArgb(c)

        c = settings.GetInt32("SCR_ABReportWithLines.thresh2colorpicker")
        if  c == 0:
            self.thresh2colorpicker.SelectedColor = Color.Yellow
        else:
             self.thresh2colorpicker.SelectedColor = Color.FromArgb(c)

        c = settings.GetInt32("SCR_ABReportWithLines.thresh3colorpicker")
        if  c == 0:
            self.thresh3colorpicker.SelectedColor = Color.Magenta
        else:
             self.thresh3colorpicker.SelectedColor = Color.FromArgb(c)

        c = settings.GetInt32("SCR_ABReportWithLines.thresh4colorpicker")
        if  c == 0:
            self.thresh4colorpicker.SelectedColor = Color.Red
        else:
             self.thresh4colorpicker.SelectedColor = Color.FromArgb(c)

        self.searchradius.Distance = settings.GetDouble("SCR_ABReportWithLines.searchradius", 0.25)
        self.horoffset.Distance = settings.GetDouble("SCR_ABReportWithLines.horoffset", 0.000)
        self.veroffset.Distance = settings.GetDouble("SCR_ABReportWithLines.veroffset", 0.000)

        self.drawoffset.IsChecked = settings.GetBoolean("SCR_ABReportWithLines.drawoffset", True)
        self.search3d.IsChecked = settings.GetBoolean("SCR_ABReportWithLines.search3d", True)
        self.usedtm.IsChecked = settings.GetBoolean("SCR_ABReportWithLines.usedtm", False)

        try:    self.designsurfacepicker.SelectIndex(settings.GetUInt32("SCR_ABReportWithLines.designsurfacepicker", 0))
        except: self.designsurfacepicker.SelectIndex(0)

        settingserial = settings.GetUInt32("SCR_ABReportWithLines.designlayerpicker", 8) # 8 is FixedSerial for Layer Zero
        o = self.currentProject.Concordance.Lookup(settingserial)
        if o != None:
            if isinstance(o.GetSite(), LayerCollection):    
                self.designlayerpicker.SetSelectedSerialNumber(settingserial, InputMethod(3))
            else:                       
                self.designlayerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:                       
            self.designlayerpicker.SetSelectedSerialNumber(8, InputMethod(3))

        self.specificline.IsChecked = settings.GetBoolean("SCR_ABReportWithLines.specificline", False)

        self.drawleader.IsChecked = settings.GetBoolean("SCR_ABReportWithLines.drawleader", False)
        self.leaderarrowsize.Distance = settings.GetDouble("SCR_ABReportWithLines.leaderarrowsize", 0.1)

        #self.computethickness.IsChecked = settings.GetBoolean("SCR_ABReportWithLines.computethickness", False)
        #self.showdesignthicknessRL.IsChecked = settings.GetBoolean("SCR_ABReportWithLines.showdesignthicknessRL", True)
        #self.textdesignthicknessRL.Text = settings.GetString("SCR_ABReportWithLines.textdesignthicknessRL", "DesignUS-RL=")
        #self.showdesignthickness.IsChecked = settings.GetBoolean("SCR_ABReportWithLines.showdesignthickness", True)
        #self.textdesignthickness.Text = settings.GetString("SCR_ABReportWithLines.textdesignthickness", "DesignThickness=")
        #self.showasbuiltthicknessRL.IsChecked = settings.GetBoolean("SCR_ABReportWithLines.showasbuiltthicknessRL", True)
        #self.textasbuiltthicknessRL.Text = settings.GetString("SCR_ABReportWithLines.textasbuiltthicknessRL", "AsbuiltBlindingRL=")
        #self.showasbuiltthickness.IsChecked = settings.GetBoolean("SCR_ABReportWithLines.showasbuiltthickness", True)
        #self.textasbuiltthickness.Text = settings.GetString("SCR_ABReportWithLines.textasbuiltthickness", "AsbuiltThickness=")
        #self.showthicknessdeviation.IsChecked = settings.GetBoolean("SCR_ABReportWithLines.showthicknessdeviation", True)
        #self.textthicknessdeviation.Text = settings.GetString("SCR_ABReportWithLines.textthicknessdeviation", "ThicknessDeviation=")

        settingserial = settings.GetUInt32("SCR_ABReportWithLines.ablayerpicker", 8) # 8 is FixedSerial for Layer Zero
        o = self.currentProject.Concordance.Lookup(settingserial)
        if o != None:
            if isinstance(o.GetSite(), LayerCollection):    
                self.ablayerpicker.SetSelectedSerialNumber(settingserial, InputMethod(3))
            else:                       
                self.ablayerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:                       
            self.ablayerpicker.SetSelectedSerialNumber(8, InputMethod(3))

        self.multipoint.IsChecked = settings.GetBoolean("SCR_ABReportWithLines.multipoint", True)
        self.singlepoint.IsChecked = settings.GetBoolean("SCR_ABReportWithLines.singlepoint", False)

    def SaveOptions(self):
        settings = ConstructionCommandsSettings.ProvideObject(self.currentProject)

        settings.SetUInt32("SCR_ABReportWithLines.blockpicker", self.blockpicker.SelectedIndex)
        settings.SetDouble("SCR_ABReportWithLines.blockscale", self.blockscale.Value)
        settings.SetBoolean("SCR_ABReportWithLines.arrowstmrspec", self.arrowstmrspec.IsChecked)

        settings.SetString("SCR_ABReportWithLines.unitpicker", self.unitpicker.Text)
        settings.SetBoolean("SCR_ABReportWithLines.addunitsuffix", self.addunitsuffix.IsChecked)
        settings.SetDouble("SCR_ABReportWithLines.textdecimals", self.textdecimals.Value)
        settings.SetDouble("SCR_ABReportWithLines.textheight", self.textheight.Distance)

        settings.SetDouble("SCR_ABReportWithLines.thresh1", self.thresh1.Distance)
        settings.SetDouble("SCR_ABReportWithLines.thresh2", self.thresh2.Distance)
        settings.SetDouble("SCR_ABReportWithLines.thresh3", self.thresh3.Distance)

        settings.SetInt32("SCR_ABReportWithLines.thresh1colorpicker", self.thresh1colorpicker.SelectedColor.ToArgb())
        settings.SetInt32("SCR_ABReportWithLines.thresh2colorpicker", self.thresh2colorpicker.SelectedColor.ToArgb())
        settings.SetInt32("SCR_ABReportWithLines.thresh3colorpicker", self.thresh3colorpicker.SelectedColor.ToArgb())
        settings.SetInt32("SCR_ABReportWithLines.thresh4colorpicker", self.thresh4colorpicker.SelectedColor.ToArgb())


        settings.SetDouble("SCR_ABReportWithLines.searchradius", self.searchradius.Distance)
        settings.SetDouble("SCR_ABReportWithLines.horoffset", self.horoffset.Distance)
        settings.SetDouble("SCR_ABReportWithLines.veroffset", self.veroffset.Distance)

        settings.SetBoolean("SCR_ABReportWithLines.drawoffset", self.drawoffset.IsChecked)
        settings.SetBoolean("SCR_ABReportWithLines.search3d", self.search3d.IsChecked)
        settings.SetBoolean("SCR_ABReportWithLines.usedtm", self.usedtm.IsChecked)

        try:    # if nothing is selected it would throw an error
            settings.SetUInt32("SCR_ABReportWithLines.designsurfacepicker", self.designsurfacepicker.SelectedIndex)
        except:
            pass
        settings.SetUInt32("SCR_ABReportWithLines.designlayerpicker", self.designlayerpicker.SelectedSerialNumber)
        settings.SetBoolean("SCR_ABReportWithLines.specificline", self.specificline.IsChecked)

        settings.SetBoolean("SCR_ABReportWithLines.drawleader", self.drawleader.IsChecked)
        settings.SetDouble("SCR_ABReportWithLines.leaderarrowsize", self.leaderarrowsize.Distance)

        #OptionsManager.SetValue("SCR_ABReportWithLines.computethickness", self.computethickness.IsChecked)
        #OptionsManager.SetValue("SCR_ABReportWithLines.showdesignthicknessRL", self.showdesignthicknessRL.IsChecked)
        #OptionsManager.SetValue("SCR_ABReportWithLines.textdesignthicknessRL", self.textdesignthicknessRL.Text)
        #OptionsManager.SetValue("SCR_ABReportWithLines.showdesignthickness", self.showdesignthickness.IsChecked)
        #OptionsManager.SetValue("SCR_ABReportWithLines.textdesignthickness", self.textdesignthickness.Text)
        #OptionsManager.SetValue("SCR_ABReportWithLines.showasbuiltthicknessRL", self.showasbuiltthicknessRL.IsChecked)
        #OptionsManager.SetValue("SCR_ABReportWithLines.textasbuiltthicknessRL", self.textasbuiltthicknessRL.Text)
        #OptionsManager.SetValue("SCR_ABReportWithLines.showasbuiltthickness", self.showasbuiltthickness.IsChecked)
        #OptionsManager.SetValue("SCR_ABReportWithLines.textasbuiltthickness", self.textasbuiltthickness.Text)
        #OptionsManager.SetValue("SCR_ABReportWithLines.showthicknessdeviation", self.showthicknessdeviation.IsChecked)
        #OptionsManager.SetValue("SCR_ABReportWithLines.textthicknessdeviation", self.textthicknessdeviation.Text)

        settings.SetUInt32("SCR_ABReportWithLines.ablayerpicker", self.ablayerpicker.SelectedSerialNumber)
        settings.SetBoolean("SCR_ABReportWithLines.multipoint", self.multipoint.IsChecked)
        settings.SetBoolean("SCR_ABReportWithLines.singlepoint", self.singlepoint.IsChecked)

    #def IsValidThicknessEntity(self, serial):
    #    o = self.currentProject.Concordance.Lookup(serial)
    #    if isinstance(o, self.lType):
    #        return True
    #    if isinstance(o, Model3D) and not isinstance(o, clr.GetClrType(ProjectedSurface)):
    #        return True
    #    return False

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

    def IsValidLine(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        return False

    def CoordPickChanged(self, ctrl, e):
        if e.Cause == InputMethod.Mouse:     
            self.OkClicked(None, None)

    def search3dChanged(self, sender, e):
        if self.search3d.IsChecked==True:
            self.dtmborder.IsEnabled=True
        else:
            self.dtmborder.IsEnabled=False

    def usedtmChanged(self, sender, e):
        if self.usedtm.IsChecked==True:
            self.designsurfacepicker.IsEnabled=True
        else:
            self.designsurfacepicker.IsEnabled=False
        
    def IsValidPoints(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.coordpointType):
            return True
        if isinstance(o, self.cadpointType):
            return True
        return False
    
    def specificlineChanged(self, sender, e):
        if self.specificline.IsChecked:
            self.designlayerlabel.IsEnabled = False
            self.designlayerpicker.IsEnabled = False
            self.specificlinelabel.IsEnabled = True
            self.linepicker1.IsEnabled = True
        else:
            self.designlayerlabel.IsEnabled = True
            self.designlayerpicker.IsEnabled = True
            self.specificlinelabel.IsEnabled = False
            self.linepicker1.IsEnabled = False

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''

        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())

        wv = self.currentProject [Project.FixedSerial.WorldView]
        bc = self.currentProject [Project.FixedSerial.BlockCollection]    # getting all blocks as collection

        design_layer_sn = self.designlayerpicker.SelectedSerialNumber
        wlc = self.currentProject[Project.FixedSerial.LayerContainer] # we get all the layers into an object, LayerCollection
        wl = wlc[design_layer_sn]    # we get just the source layer as an object
        if self.specificline.IsChecked:
            designmembers = []
            l1 = self.linepicker1.Entity
            if l1!=None:
                designmembers.Add(l1.SerialNumber)
        else:
            designmembers = wl.Members  # we get serial number list of all the elements on that layer
        
        if self.search3d.IsChecked and self.usedtm.IsChecked:
            designsurface = wv.Lookup (self.designsurfacepicker.SelectedSerial)    # we get our selected surface as object

        pihalf = math.pi/2
        twopi = 2*math.pi

        #check thresholds
        thresholds=True
        try:
            blockscale = self.blockscale.Value
            thresh1 = self.thresh1.Distance
            thresh2 = self.thresh2.Distance
            thresh3 = self.thresh3.Distance
            searchradius = self.searchradius.Distance
            horoffset = self.horoffset.Distance
            veroffset = self.veroffset.Distance
            textheight = self.textheight.Distance
            textdecimals = int(self.textdecimals.Value)
        except:
            thresholds = False
        
        # p1 - Design Point
        # p2 - AB Point
        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
            
            
                outSegment=clr.StrongBox[Segment]()
                out_t=clr.StrongBox[float]()
                outPointOnCL = clr.StrongBox[Point3D]()
                station = clr.StrongBox[float]()
                perpVector3D=clr.StrongBox[Vector3D]()
                testDist=clr.StrongBox[float]()
                testside=clr.StrongBox[Side]()
                
                pfound = Point3D()
                pointlist = []

                if self.multipoint.IsChecked and self.objs.SelectedMembers(self.currentProject).Count > 0:
                    for p2 in self.objs.SelectedMembers(self.currentProject): # go through all selected AB points
                        if p2.Layer != self.designlayerpicker.SelectedSerialNumber: # we don't want to report on points in the design layer
                            if isinstance(p2, CoordPoint) or isinstance(p2, self.cadpointType):
                                pointlist.Add(p2.Position)
                        else:
                            self.error.Content += '\nPoints on Design-Layer have been ignored'

                if self.singlepoint.IsChecked:
                    pointlist.Add(self.coordpick.Coordinate)
                
                if pointlist.Count > 0 and designmembers.Count  >0 and thresholds==True:
                    for p in pointlist: # go through all selected AB points
                        if p:
                            p2 = p  # I changed the code later and was too lazy too adjust the columns
                           
                            pfound = False
                            for designserial in designmembers: # go through all members of the design layer
                                l=self.currentProject.Concordance.Lookup(designserial)
                                if isinstance(l, self.lType):
                                    polyseg=l.ComputePolySeg()
                                    polyseg=polyseg.ToWorld()
                                    polyseg_v=l.ComputeVerticalPolySeg()

                                    # try to find a perpendicular solution on that line
                                    if polyseg.FindPointFromPoint(p2, outSegment, out_t, outPointOnCL, station, perpVector3D, testDist, testside): 
                                        pcompute=outPointOnCL.Value
                                        if polyseg_v != None:
                                            pcompute.Z = polyseg_v.ComputeVerticalSlopeAndGrade(station.Value)[1]
                                        
                                        tt = pcompute.Distance2D(p2)
                                        tt2 = Point3D.Distance(pcompute, p2)
                                        if self.search3d.IsChecked and self.usedtm.IsChecked==False:
                                            if Point3D.Distance(pcompute, p2)[0] <= searchradius: # we check if it's actually within our search radius
                                                if pfound == False:
                                                    p1 = pcompute
                                                    p1side = testside.Value
                                                    designsegment = outSegment.Value
                                                    pfound = True
                                                else: # if we find multiple solutions we only want the closest one
                                                    if Point3D.Distance(pcompute, p2)[0] < Point3D.Distance(p1, p2)[0]:
                                                        p1 = pcompute
                                                        p1side = testside.Value
                                                        designsegment = outSegment.Value
                                                        pfound = True
                                        else:
                                            if pcompute.Distance2D(p2) <= searchradius: # we check if it's actually within our search radius
                                                if pfound == False:
                                                    p1 = pcompute
                                                    p1side = testside.Value
                                                    designsegment = outSegment.Value
                                                    pfound = True
                                                else: # if we find multiple solutions we only want the closest one
                                                    if pcompute.Distance2D(p2) < p1.Distance2D(p2):
                                                        p1 = pcompute
                                                        p1side = testside.Value
                                                        designsegment = outSegment.Value
                                                        pfound = True
                              
                            # p1 - Design Point
                            # p2 - AB Point                        
                            if pfound:
                                # compute deltas and draw block etc.    
                                deltavector = Vector3D(p1, p2)
                                if self.search3d.IsChecked==False and deltavector.Length2D > searchradius: continue  # ignore points that are too far away
                                if self.search3d.IsChecked and deltavector.Length > searchradius: continue  # ignore points that are too far away

                                # compute delta offset value
                                deltaoff = p2.Distance2D(p1)

                                # get it right for very small offsets, arrow might otherwise point to North 
                                if abs(deltaoff) < 0.0001:
                                    azo = (2*math.pi) - Vector3D(designsegment.BeginPoint, designsegment.EndPoint).Azimuth - math.pi/2
                                else:
                                    azo = (2*math.pi) - deltavector.Azimuth

                                # apply horizontal offset and flip arrow if necessary
                                if horoffset != 0.0:
                                    if ((deltaoff < 0) == (deltaoff + horoffset < 0)): # check if sign is changing when offset is applied
                                        deltaoff += horoffset
                                    else:
                                        deltaoff += horoffset
                                        azo += math.pi



                                #deltaoffstring = "dO=" + self.lunits.Format(self.lunits.Convert(abs(deltaoff), LinearType.Display), self.lfp)
                                #https://mkaz.blog/code/python-string-format-cookbook/
                                #deltaoffstring = "dO=" + str("{:.{}f}".format(abs(deltaoff), textdecimals))
                                deltaoffstring = "dO=" + self.tooutputunit(abs(deltaoff))

                                # compute the z value if necessary and change the textcolor
                                if self.search3d.IsChecked:
                                    if self.usedtm.IsChecked:
                                        tt=designsurface.PickSurface(p2)
                                        if tt[0]==True:
                                            deltaz = (p2.Z - tt[1].Z) + veroffset
                                            ## test purpose draw p1
                                            #cadPoint = wv.Add(clr.GetClrType(CadPoint))
                                            #cadPoint.Point0 = tt[1]
                                            #cadPoint.Layer = self.ablayerpicker.SelectedSerialNumber
                                        else:
                                            deltaz = 9999.9
                                    else:   # only elevation comparison between p1 and p2
                                        deltaz = (p2.Z-p1.Z) + veroffset
                                    #deltazstring = "dz=" + self.lunits.Format(self.lunits.Convert(deltaz, LinearType.Display), self.lfp)
                                    #deltazstring = "dz=" + str("{:.{}f}".format(deltaz, textdecimals))
                                    deltazstring = "dz=" + self.tooutputunit(deltaz)

                                else:
                                    deltazstring = ''
                                
                                deltahigh=0
                                # find the highest delta
                                if self.drawoffset.IsChecked:
                                    deltahigh=abs(deltaoff)
                                if self.search3d.IsChecked and abs(deltaz)>deltahigh: deltahigh=abs(deltaz)

                                # find the right color for the text
                                if deltahigh <= thresh1: abcolor = self.thresh1colorpicker.SelectedColor
                                if deltahigh > thresh1: abcolor = self.thresh2colorpicker.SelectedColor
                                if deltahigh > thresh2: abcolor = self.thresh3colorpicker.SelectedColor
                                if deltahigh > thresh3: abcolor = self.thresh4colorpicker.SelectedColor

                                # color for the arrow
                                if abs(deltaoff) <= thresh1: offcolor = self.thresh1colorpicker.SelectedColor
                                if abs(deltaoff) > thresh1: offcolor = self.thresh2colorpicker.SelectedColor
                                if abs(deltaoff) > thresh2: offcolor = self.thresh3colorpicker.SelectedColor
                                if abs(deltaoff) > thresh3: offcolor = self.thresh4colorpicker.SelectedColor

                                # we draw the block
                                if self.drawoffset.IsChecked:
                                    b = wv.Add(clr.GetClrType(BlockReference))
                                    b.BlockSerial=self.blockpicker.SelectedSerial     # we set the serial of our new block to the serial we've selected in the dropdown box
                                    b.InsertionPoint=p2     # we set the insertion point to the one from our coordinate picker
                                    if self.arrowstmrspec.IsChecked:     #original code was to point away from design
                                        b.Rotation=azo + math.pi
                                    else:
                                        b.Rotation=azo       
                                    b.Layer = self.ablayerpicker.SelectedSerialNumber
                                    b.Color=offcolor
                                    b.XScaleFactor = blockscale
                                    b.YScaleFactor = blockscale
                                    b.ZScaleFactor = blockscale

                                # draw the text
                                t = wv.Add(clr.GetClrType(MText))
                                textvector = deltavector.Clone()
                                textvector.To2D()
                                textvector.Length = textheight/2
                                textvector.Z = 0
                                
                                if p1side == Side.Left:
                                    #textvector.RotateAboutZ(-pihalf/2)
                                    pt = p2 + textvector
                                    #pt.Z = p2.Z
                                    t.AlignmentPoint = pt
                                    t.AttachPoint = AttachmentPoint.BottomRight
                                    t.Rotation = azo - pihalf
                                else:
                                    #textvector.RotateAboutZ(pihalf/2)
                                    pt = p2 + textvector
                                    #pt.Z = p2.Z
                                    t.AlignmentPoint = pt
                                    t.AttachPoint = AttachmentPoint.BottomLeft
                                    t.Rotation = azo + pihalf
                                t.TextString=''
                                if self.drawoffset.IsChecked: 
                                    if len(t.TextString)>0:
                                        t.TextString += '\\P'+ deltaoffstring
                                    else:
                                        t.TextString += deltaoffstring
                                if self.search3d.IsChecked: 
                                    if len(t.TextString)>0:
                                        t.TextString += '\\P' + deltazstring
                                    else:
                                        t.TextString += deltazstring


                                t.Height = textheight
                                t.Layer = self.ablayerpicker.SelectedSerialNumber
                                t.Color = abcolor

                                if self.drawleader.IsChecked:
                                    textvector.Length *= 5
                                    t.AlignmentPoint = t.AlignmentPoint + textvector
                                    leaderpoints = List[Point3D]()
                                    leaderpoints.Add(p2)
                                    leaderpoints.Add(t.AlignmentPoint)
                                    
                                    l = wv.Add(clr.GetClrType(Leader))
                                    l.AnnotationSerial = t.SerialNumber
                                    l.ScaleFactor = 1
                                    l.Points = leaderpoints
                                    l.Layer = self.ablayerpicker.SelectedSerialNumber
                                    l.LeaderType = LeaderType.LineWithArrow
                                    l.ArrowheadType = DimArrowheadType.ArrowDefault
                                    l.ArrowheadSize = abs(self.leaderarrowsize.Distance)
                                    l.TextGap = 0.1
                                
                            
                        else:
                            pass
                else:
                    pass
                
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
        
        if self.multipoint.IsChecked:
            Keyboard.Focus(self.objs)
        if self.singlepoint.IsChecked:
            Keyboard.Focus(self.coordpick)
        self.SaveOptions()           

        
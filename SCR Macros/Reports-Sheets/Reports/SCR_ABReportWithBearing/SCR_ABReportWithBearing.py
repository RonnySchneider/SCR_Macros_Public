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
    cmdData.Key = "SCR_ABReportWithBearing"
    cmdData.CommandName = "SCR_ABReportWithBearing"
    cmdData.Caption = "_SCR_ABReportWithBearing"
    cmdData.UIForm = "SCR_ABReportWithBearing"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Reports"
        cmdData.DefaultTabGroupKey = "Reports"
        cmdData.ShortCaption = "relative to Bearing"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.16
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "AB Report point to point, relative to Bearing"
        cmdData.ToolTipTextFormatted = "AB Report point to point, relative to Bearing"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_ABReportWithBearing(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_ABReportWithBearing.xaml") as s:
            wpf.LoadComponent (self, s)
        self.currentProject = currentProject
        self.macroFileFolder = macroFileFolder
        self.outputunitenum = 0

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

        self.objs.IsEntityValidCallback = self.IsValid1
               
        optionMenu1 = SelectionContextMenuHandler()
        # remove options that don't apply here
        optionMenu1.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        self.objs.ButtonContextMenu = optionMenu1
        
        self.cadpointType = clr.GetClrType(CadPoint)
        self.coordpointType = clr.GetClrType(CoordPoint)
        self.lType = clr.GetClrType(IPolyseg)

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
        settings = ConstructionCommandsSettings.ProvideObject(self.currentProject)

        try:    self.blockpicker.SelectIndex(settings.GetUInt32("SCR_ABReportWithBearing.blockpicker", 0))
        except: self.blockpicker.SelectIndex(0)
        self.blockscale.Value = settings.GetDouble("SCR_ABReportWithBearing.blockscale", 1.000)
        self.arrowstmrspec.IsChecked = settings.GetBoolean("SCR_ABReportWithBearing.arrowstmrspec", True)

        self.unitpicker.Text = settings.GetString("SCR_ABReportWithBearing.unitpicker", "Meter")
        self.addunitsuffix.IsChecked = settings.GetBoolean("SCR_ABReportWithBearing.addunitsuffix", False)
        self.textdecimals.Value = settings.GetDouble("SCR_ABReportWithBearing.textdecimals", 3)
        self.textheight.Distance = settings.GetDouble("SCR_ABReportWithBearing.textheight", 0.2)

        self.thresh1.Distance = settings.GetDouble("SCR_ABReportWithBearing.thresh1", 0.005)
        self.thresh2.Distance = settings.GetDouble("SCR_ABReportWithBearing.thresh2", 0.010)
        self.thresh3.Distance = settings.GetDouble("SCR_ABReportWithBearing.thresh3", 0.015)
        
        c = settings.GetInt32("SCR_ABReportWithBearing.thresh1colorpicker")
        if  c == 0:
            self.thresh1colorpicker.SelectedColor = Color.Green
        else:
             self.thresh1colorpicker.SelectedColor = Color.FromArgb(c)

        c = settings.GetInt32("SCR_ABReportWithBearing.thresh2colorpicker")
        if  c == 0:
            self.thresh2colorpicker.SelectedColor = Color.Yellow
        else:
             self.thresh2colorpicker.SelectedColor = Color.FromArgb(c)

        c = settings.GetInt32("SCR_ABReportWithBearing.thresh3colorpicker")
        if  c == 0:
            self.thresh3colorpicker.SelectedColor = Color.Magenta
        else:
             self.thresh3colorpicker.SelectedColor = Color.FromArgb(c)

        c = settings.GetInt32("SCR_ABReportWithBearing.thresh4colorpicker")
        if  c == 0:
            self.thresh4colorpicker.SelectedColor = Color.Red
        else:
             self.thresh4colorpicker.SelectedColor = Color.FromArgb(c)

        self.searchradius.Distance = settings.GetDouble("SCR_ABReportWithBearing.searchradius", 0.25)

        self.drawx.IsChecked = settings.GetBoolean("SCR_ABReportWithBearing.drawx", True)
        self.drawy.IsChecked = settings.GetBoolean("SCR_ABReportWithBearing.drawy", True)
        self.search3d.IsChecked = settings.GetBoolean("SCR_ABReportWithBearing.search3d", True)
        self.usedtm.IsChecked = settings.GetBoolean("SCR_ABReportWithBearing.usedtm", False)

        try:    self.designsurfacepicker.SelectIndex(settings.GetUInt32("SCR_ABReportWithBearing.designsurfacepicker", 0))
        except: self.designsurfacepicker.SelectIndex(0)

        settingserial = settings.GetUInt32("SCR_ABReportWithBearing.designlayerpicker", 8) # 8 is FixedSerial for Layer Zero
        o = self.currentProject.Concordance.Lookup(settingserial)
        if o != None:
            if isinstance(o.GetSite(), LayerCollection):    
                self.designlayerpicker.SetSelectedSerialNumber(settingserial, InputMethod(3))
            else:                       
                self.designlayerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:                       
            self.designlayerpicker.SetSelectedSerialNumber(8, InputMethod(3))

        settingserial = settings.GetUInt32("SCR_ABReportWithBearing.ablayerpicker", 8) # 8 is FixedSerial for Layer Zero
        o = self.currentProject.Concordance.Lookup(settingserial)
        if o != None:
            if isinstance(o.GetSite(), LayerCollection):    
                self.ablayerpicker.SetSelectedSerialNumber(settingserial, InputMethod(3))
            else:                       
                self.ablayerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:                       
            self.ablayerpicker.SetSelectedSerialNumber(8, InputMethod(3))

        self.bearingpicker.Direction = settings.GetDouble("SCR_ABReportWithBearing.bearingpicker", 0.000)

    def SaveOptions(self):
        settings = ConstructionCommandsSettings.ProvideObject(self.currentProject)

        settings.SetUInt32("SCR_ABReportWithBearing.blockpicker", self.blockpicker.SelectedIndex)
        settings.SetDouble("SCR_ABReportWithBearing.blockscale", self.blockscale.Value)
        settings.SetBoolean("SCR_ABReportWithBearing.arrowstmrspec", self.arrowstmrspec.IsChecked)

        settings.SetString("SCR_ABReportWithBearing.unitpicker", self.unitpicker.Text)
        settings.SetBoolean("SCR_ABReportWithBearing.addunitsuffix", self.addunitsuffix.IsChecked)
        settings.SetDouble("SCR_ABReportWithBearing.textdecimals", self.textdecimals.Value)
        settings.SetDouble("SCR_ABReportWithBearing.textheight", self.textheight.Distance)

        settings.SetDouble("SCR_ABReportWithBearing.thresh1", self.thresh1.Distance)
        settings.SetDouble("SCR_ABReportWithBearing.thresh2", self.thresh2.Distance)
        settings.SetDouble("SCR_ABReportWithBearing.thresh3", self.thresh3.Distance)

        settings.SetInt32("SCR_ABReportWithBearing.thresh1colorpicker", self.thresh1colorpicker.SelectedColor.ToArgb())
        settings.SetInt32("SCR_ABReportWithBearing.thresh2colorpicker", self.thresh2colorpicker.SelectedColor.ToArgb())
        settings.SetInt32("SCR_ABReportWithBearing.thresh3colorpicker", self.thresh3colorpicker.SelectedColor.ToArgb())
        settings.SetInt32("SCR_ABReportWithBearing.thresh4colorpicker", self.thresh4colorpicker.SelectedColor.ToArgb())

        settings.SetDouble("SCR_ABReportWithBearing.searchradius", self.searchradius.Distance)

        settings.SetBoolean("SCR_ABReportWithBearing.drawx", self.drawx.IsChecked)
        settings.SetBoolean("SCR_ABReportWithBearing.drawy", self.drawy.IsChecked)
        settings.SetBoolean("SCR_ABReportWithBearing.search3d", self.search3d.IsChecked)
        settings.SetBoolean("SCR_ABReportWithBearing.usedtm", self.usedtm.IsChecked)

        try:    # if nothing is selected it would throw an error
            settings.SetUInt32("SCR_ABReportWithBearing.designsurfacepicker", self.designsurfacepicker.SelectedIndex)
        except:
            pass
        settings.SetUInt32("SCR_ABReportWithBearing.designlayerpicker", self.designlayerpicker.SelectedSerialNumber)
        settings.SetUInt32("SCR_ABReportWithBearing.ablayerpicker", self.ablayerpicker.SelectedSerialNumber)
        
        settings.SetDouble("SCR_ABReportWithBearing.bearingpicker", self.bearingpicker.Direction)

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
        
    def IsValid1(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.coordpointType):
            return True
        if isinstance(o, self.cadpointType):
            return True
        return False

       
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
        wl=wlc[design_layer_sn]    # we get just the source layer as an object
        designmembers=wl.Members  # we get serial number list of all the elements on that layer

        if self.search3d.IsChecked and self.usedtm.IsChecked:
            designsurface = wv.Lookup (self.designsurfacepicker.SelectedSerial)    # we get our selected surface as object

        pihalf=math.pi/2
        twopi=2*math.pi

        #check thresholds
        thresholds = True
        try:
            blockscale = self.blockscale.Value
            thresh1 = self.thresh1.Distance
            thresh2 = self.thresh2.Distance
            thresh3 = self.thresh3.Distance
            searchradius = self.searchradius.Distance
            textheight = self.textheight.Distance
            textdecimals = int(self.textdecimals.Value)
        except:
            thresholds=False

        # p1 - Design Point
        # p2 - AB Point
        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                if self.objs.SelectedMembers(self.currentProject).Count>0 and thresholds==True:
                    for p2 in self.objs.SelectedMembers(self.currentProject): # go through all selected AB points
                        if (isinstance(p2, CoordPoint) or isinstance(p2, CadPoint)) and designmembers.Count>0 and p2.Layer!=design_layer_sn:
                            
                            #find closest design point
                            foundclosest = False
                            if self.search3d.IsChecked and self.usedtm.IsChecked==False:
                                c=0
                                for i in designmembers:             # go through all design points
                                    p1=self.currentProject.Concordance.Lookup(i)
                                    if (isinstance(p1, CoordPoint) or isinstance(p1, CadPoint)) and p1.SerialNumber!=p2.SerialNumber:
                                        if c==0:
                                            shortestdist=Vector3D(p1.Position,p2.Position).Length
                                            closestp1=p1
                                            c+=1
                                            foundclosest = True
                                        else:
                                            if Vector3D(p1.Position,p2.Position).Length<shortestdist:
                                                shortestdist=Vector3D(p1.Position,p2.Position).Length
                                                closestp1=p1
                                                foundclosest = True
                            else:
                                c=0
                                for i in designmembers:             # go through all design points
                                    p1=self.currentProject.Concordance.Lookup(i)
                                    if (isinstance(p1, CoordPoint) or isinstance(p1, CadPoint)) and p1.SerialNumber!=p2.SerialNumber:
                                        if c==0:
                                            shortestdist=Vector3D(p1.Position,p2.Position).Length2D
                                            closestp1=p1
                                            foundclosest = True
                                            c+=1
                                        else:
                                            if Vector3D(p1.Position,p2.Position).Length2D<shortestdist:
                                                shortestdist=Vector3D(p1.Position,p2.Position).Length2D
                                                closestp1=p1
                                                foundclosest = True
                            
                            if foundclosest == True:
                                
                                p1 = closestp1

                                # compute deltas and draw block etc.    
                                deltavector=Vector3D(p1.Position,p2.Position)
                                if self.search3d.IsChecked==False and deltavector.Length2D > searchradius: continue  # ignore points that are too far away
                                if self.search3d.IsChecked and deltavector.Length > searchradius: continue  # ignore points that are too far away

                                az1=deltavector.Azimuth
                                az2=self.bearingpicker.Direction   # bearing picker value is counter clockwise and starts in x direction

                                az2comp=(2*math.pi)-self.bearingpicker.Direction+pihalf  
                                if az2comp>=twopi: az2=az2-twopi
                                if az2comp<0: az2=az2+twopi
                                
                                # compute delta x/y values
                                deltax = math.cos(az2comp-az1)*deltavector.Length2D
                                deltay = math.sin(az2comp-az1)*deltavector.Length2D

                                # create the bearings for the x and y arrow block
                                if deltax>=0: azx = az2 - pihalf
                                else: azx = az2 + pihalf
                                if deltay>=0: azy = az2
                                else: azy = az2 + math.pi

                                #deltaxstring = "dx=" + String.Format("{0:0.000}", abs(deltax))
                                #deltaystring = "dy=" + String.Format("{0:0.000}", abs(deltay))
                                # https://mkaz.blog/code/python-string-format-cookbook/
                                deltaxstring = "dx=" + self.tooutputunit(abs(deltax))
                                deltaystring = "dy=" + self.tooutputunit(abs(deltay))

                                # compute the z value if necessary and change the textcolor
                                if self.search3d.IsChecked:
                                    if self.usedtm.IsChecked:
                                        tt=designsurface.PickSurface(p2.Position)
                                        if tt[0]==True:
                                            deltaz = p2.Position.Z - tt[1].Z
                                        else:
                                            deltaz = 9999
                                    else:   # only elevation comparison between p1 and p2
                                        deltaz = round(p2.Position.Z-p1.Position.Z, 3)
                                    deltazstring = "dz=" + self.tooutputunit(deltaz)
                                else:
                                    deltazstring = ''
                                
                                deltahigh=0
                                # find the highest delta
                                if self.drawx.IsChecked and self.drawy.IsChecked:
                                    deltahigh=abs(deltax)
                                    if abs(deltay)>deltahigh: deltahigh=abs(deltay)
                                if self.drawx.IsChecked and self.drawy.IsChecked==False:
                                    deltahigh=abs(deltax)
                                if self.drawx.IsChecked==False and self.drawy.IsChecked:
                                    deltahigh=abs(deltay)
                                if self.search3d.IsChecked and abs(deltaz)>deltahigh: deltahigh=abs(deltaz)

                                # find the right color
                                if deltahigh <= thresh1: abcolor = self.thresh1colorpicker.SelectedColor
                                if deltahigh > thresh1: abcolor = self.thresh2colorpicker.SelectedColor
                                if deltahigh > thresh2: abcolor = self.thresh3colorpicker.SelectedColor
                                if deltahigh > thresh3: abcolor = self.thresh4colorpicker.SelectedColor

                                if abs(deltax) <= thresh1: xcolor = self.thresh1colorpicker.SelectedColor
                                if abs(deltax) > thresh1: xcolor = self.thresh2colorpicker.SelectedColor
                                if abs(deltax) > thresh2: xcolor = self.thresh3colorpicker.SelectedColor
                                if abs(deltax) > thresh3: xcolor = self.thresh4colorpicker.SelectedColor

                                if abs(deltay) <= thresh1: ycolor = self.thresh1colorpicker.SelectedColor
                                if abs(deltay) > thresh1: ycolor = self.thresh2colorpicker.SelectedColor
                                if abs(deltay) > thresh2: ycolor = self.thresh3colorpicker.SelectedColor
                                if abs(deltay) > thresh3: ycolor = self.thresh4colorpicker.SelectedColor
                                
                                # we draw the block
                                if self.drawx.IsChecked:
                                    b = wv.Add(clr.GetClrType(BlockReference))
                                    b.BlockSerial = self.blockpicker.SelectedSerial     # we set the serial of our new block to the serial we've selected in the dropdown box
                                    b.InsertionPoint = p2.Position     # we set the insertion point to the one from our coordinate picker
                                    if self.arrowstmrspec.IsChecked:     #original code was to point away from design
                                        b.Rotation = azx + math.pi       # we set the block rotation to the one from our bearing picker
                                    else:
                                        b.Rotation=azx       # we set the block rotation to the one from our bearing picker
                                    b.Layer = self.ablayerpicker.SelectedSerialNumber
                                    b.Color=xcolor
                                    b.XScaleFactor = blockscale
                                    b.YScaleFactor = blockscale
                                    b.ZScaleFactor = blockscale

                                if self.drawy.IsChecked:
                                    b = wv.Add(clr.GetClrType(BlockReference))
                                    b.BlockSerial=self.blockpicker.SelectedSerial     # we set the serial of our new block to the serial we've selected in the dropdown box
                                    b.InsertionPoint=p2.Position     # we set the insertion point to the one from our coordinate picker
                                    if self.arrowstmrspec.IsChecked:     #original code was to point away from design
                                        b.Rotation = azy + math.pi    # we set the block rotation to the one from our bearing picker
                                    else:
                                        b.Rotation = azy     # we set the block rotation to the one from our bearing picker
                                    b.Layer = self.ablayerpicker.SelectedSerialNumber
                                    b.Color=ycolor
                                    b.XScaleFactor = blockscale
                                    b.YScaleFactor = blockscale
                                    b.ZScaleFactor = blockscale


                                t = wv.Add(clr.GetClrType(MText))
                                t.AlignmentPoint=p2.Position
                                t.TextString=''
                                if self.drawx.IsChecked:
                                    if len(t.TextString)>0:
                                        t.TextString += '\\P' + deltaxstring
                                    else:
                                        t.TextString += deltaxstring
                                if self.drawy.IsChecked:
                                    if len(t.TextString)>0:
                                        t.TextString += '\\P' + deltaystring
                                    else:
                                        t.TextString += deltaystring
                                if self.search3d.IsChecked:
                                    if len(t.TextString)>0:
                                        t.TextString += '\\P' + deltazstring
                                    else:
                                        t.TextString += deltazstring
                                t.Height = textheight
                                t.Layer = self.ablayerpicker.SelectedSerialNumber
                                t.Color = abcolor
                                t.Rotation = self.bearingpicker.Direction
                            
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

        Keyboard.Focus(self.objs)
        GlobalSelection.Clear()
        self.SaveOptions()           

        
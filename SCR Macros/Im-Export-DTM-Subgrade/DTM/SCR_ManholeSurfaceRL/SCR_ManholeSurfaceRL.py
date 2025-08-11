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
    cmdData.Key = "SCR_ManholeSurfaceRL"
    cmdData.CommandName = "SCR_ManholeSurfaceRL"
    cmdData.Caption = "_SCR_ManholeSurfaceRL"
    cmdData.UIForm = "SCR_ManholeSurfaceRL"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "DTM"
        cmdData.ShortCaption = "Manhole RL"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.25
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "set Manhole-RL to Surface-RL"
        cmdData.ToolTipTextFormatted = "change Utility-Manhole-RLs to Surface-RL"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png") 
        cmdData.ImageSmall = b
    except:
        pass

class SCR_ManholeSurfaceRL(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_ManholeSurfaceRL.xaml") as s:
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

        types = Array [Type] (SurfaceTypeLists.AllWithCutFillMap)    # we fill an array with TBC object types, we could combine different types
        self.surfacepicker.FilterByEntityTypes = types    # we fill the dropdownlist by applying that types array as filter
        self.surfacepicker.AllowNone = False              # our list shall not show an empty field
        self.addsurfacepicker.FilterByEntityTypes = types    # we fill the dropdownlist by applying that types array as filter
        self.addsurfacepicker.AllowNone = False              # our list shall not show an empty field

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        self.lfp = self.lunits.Properties.Copy()
        linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation
        #self.lfp.AddSuffix = False
        self.verticaloffsetlabel.Content = "Vertical Offset [" + linearsuffix + "]"
        self.hz1offsetlabel.Content = "Horizontal Offset [" + linearsuffix + "]"
        self.hz2offsetlabel.Content = "Horizontal Offset [" + linearsuffix + "]"
        self.vz2offsetlabel.Content = "Vertical Offset [" + linearsuffix + "]"

        self.verticaloffset.NumberOfDecimals = self.lunits.Units[self.lunits.DisplayType].NumberOfDecimals
        self.hz1offset.NumberOfDecimals = self.lunits.Units[self.lunits.DisplayType].NumberOfDecimals
        self.hz2offset.NumberOfDecimals = self.lunits.Units[self.lunits.DisplayType].NumberOfDecimals
        self.vz2offset.NumberOfDecimals = self.lunits.Units[self.lunits.DisplayType].NumberOfDecimals

        self.objs.IsEntityValidCallback=self.IsValid

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def IsValid(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, UtilityNode):
            return True
        return False
        
    def SetDefaultOptions(self):
        try:    self.surfacepicker.SelectIndex(OptionsManager.GetInt("SCR_ManholeSurfaceRL.surfacepicker", 0))
        except: self.surfacepicker.SelectIndex(0)
        self.adjustrim.IsChecked = OptionsManager.GetBool("SCR_ManholeSurfaceRL.adjustrim", True)
        self.adjustfloor.IsChecked = OptionsManager.GetBool("SCR_ManholeSurfaceRL.adjustfloor", False)
        self.verticaloffset.Value = OptionsManager.GetDouble("SCR_ManholeSurfaceRL.verticaloffset", 0.0)
        self.createlines.IsChecked = OptionsManager.GetBool("SCR_ManholeSurfaceRL.createlines", False)
        self.createoutline.IsChecked = OptionsManager.GetBool("SCR_ManholeSurfaceRL.createoutline", False)
        self.createhz1offset.IsChecked = OptionsManager.GetBool("SCR_ManholeSurfaceRL.createhz1offset", False)
        self.hz1offset.Value = OptionsManager.GetDouble("SCR_ManholeSurfaceRL.hz1offset", 0.0)
        self.createhz2vzoffset.IsChecked = OptionsManager.GetBool("SCR_ManholeSurfaceRL.createhz2vzoffset", False)
        self.hz2offset.Value = OptionsManager.GetDouble("SCR_ManholeSurfaceRL.hz2offset", 0.0)
        self.vz2offset.Value = OptionsManager.GetDouble("SCR_ManholeSurfaceRL.vz2offset", 0.0)
        
        settingserial = OptionsManager.GetUint("SCR_ManholeSurfaceRL.layerpicker", 8) # 8 is FixedSerial for Layer Zero
        o = self.currentProject.Concordance.Lookup(settingserial)
        if o != None:
            if isinstance(o.GetSite(), LayerCollection):    
                self.layerpicker.SetSelectedSerialNumber(settingserial, InputMethod(3))
            else:                       
                self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:                       
            self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))

        self.addtosurface.IsChecked = OptionsManager.GetBool("SCR_ManholeSurfaceRL.addtosurface", False)

        self.addtosurface.IsChecked = OptionsManager.GetBool("SCR_ManholeSurfaceRL.addtosurface", False)
        try:    self.addsurfacepicker.SelectIndex(OptionsManager.GetInt("SCR_ManholeSurfaceRL.addsurfacepicker", 0))
        except: self.addsurfacepicker.SelectIndex(0)

    def SaveOptions(self):
        OptionsManager.SetValue("SCR_ManholeSurfaceRL.adjustrim", self.adjustrim.IsChecked)
        OptionsManager.SetValue("SCR_ManholeSurfaceRL.adjustfloor", self.adjustfloor.IsChecked)
        try:    # if nothing is selected it would throw an error
            OptionsManager.SetValue("SCR_ManholeSurfaceRL.surfacepicker", self.surfacepicker.SelectedIndex)
        except:
            pass
        OptionsManager.SetValue("SCR_ManholeSurfaceRL.verticaloffset", self.verticaloffset.Value)
        OptionsManager.SetValue("SCR_ManholeSurfaceRL.createlines", self.createlines.IsChecked)
        OptionsManager.SetValue("SCR_ManholeSurfaceRL.createoutline", self.createoutline.IsChecked)
        OptionsManager.SetValue("SCR_ManholeSurfaceRL.createhz1offset", self.createhz1offset.IsChecked)
        OptionsManager.SetValue("SCR_ManholeSurfaceRL.hz1offset", self.hz1offset.Value)
        OptionsManager.SetValue("SCR_ManholeSurfaceRL.createhz2vzoffset", self.createhz2vzoffset.IsChecked)
        OptionsManager.SetValue("SCR_ManholeSurfaceRL.hz2offset", self.hz2offset.Value)
        OptionsManager.SetValue("SCR_ManholeSurfaceRL.vz2offset", self.vz2offset.Value)
        OptionsManager.SetValue("SCR_ManholeSurfaceRL.layerpicker", self.layerpicker.SelectedSerialNumber)
        OptionsManager.SetValue("SCR_ManholeSurfaceRL.addtosurface", self.addtosurface.IsChecked)
        try:    # if nothing is selected it would throw an error
            OptionsManager.SetValue("SCR_ManholeSurfaceRL.addsurfacepicker", self.addsurfacepicker.SelectedIndex)
        except:
            pass
            
    def createlinesChanged(self, sender, e):
        if self.createlines.IsChecked:
            self.surfacepicker.IsEnabled = False
            self.verticaloffset.IsEnabled = False
            self.hz1offset.IsEnabled = True
            self.layerpicker.IsEnabled = True
        else:
            self.surfacepicker.IsEnabled = True
            self.verticaloffset.IsEnabled = True
            self.hz1offset.IsEnabled = False
            self.layerpicker.IsEnabled = False

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''
        self.success.Content=''


        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())

        wv = self.currentProject [Project.FixedSerial.WorldView]

        surface = wv.Lookup(self.surfacepicker.SelectedSerial)

        try:
            # the "with" statement will unroll any changes if something go wrong
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                # TBC internally computes in meter
                verticaloffset_meter = self.lunits.Convert(self.lunits.DisplayType, self.verticaloffset.Value, self.lunits.InternalType)

                if self.objs.Count == 0:
                    self.error.Content += '\nnothing selected'

                addtosurfaceserials = [] # will now actually be the whole object

                for o in self.objs:
                    if isinstance(o, UtilityNode):
                        
                        if not self.createlines.IsChecked: # compute DTM elevation and change node values
                            p = o.Point0

                            surfp = surface.PickSurface(p)
                            if surfp[0]:

                                if self.adjustrim.IsChecked:

                                    o.RimElevation = surfp[1].Z + verticaloffset_meter
                
                                else:

                                    p.Z = surfp[1].Z + verticaloffset_meter
                                    o.Point0 = p

                            else:
                                self.error.Content += "\ncouldn't compute a valid surface elevation"

                        else: # create outlines

                            # get the elevation
                            if self.adjustrim.IsChecked: 
                                el = o.RimElevation
                            else:
                                el = o.FloorElevation

                            if self.createhz1offset.IsChecked:
                                hz1offset_meter = self.lunits.Convert(self.lunits.DisplayType, self.hz1offset.Value, self.lunits.InternalType)
                            else:
                                hz1offset_meter = 0.0
                            hz2offset_meter = self.lunits.Convert(self.lunits.DisplayType, self.hz2offset.Value, self.lunits.InternalType)
                            vz2offset_meter = self.lunits.Convert(self.lunits.DisplayType, self.vz2offset.Value, self.lunits.InternalType)

                            # Manhole = 6; Eccentric Manhole = 7; Junctionbox = 12
                            if o.NodeType == UtilityNodeType(6) or o.NodeType == UtilityNodeType(7):
                                cp = o.Point0
                                cp.Z = el
                                
                                if self.createoutline.IsChecked:
                                    circle = wv.Add(clr.GetClrType(CadCircle))
                                    circle.CenterPoint = cp
                                    circle.Radius = (o.SiteImprovement.Diameter / 2) + o.SiteImprovement.Thickness
                                    circle.Layer = self.layerpicker.SelectedSerialNumber

                                    addtosurfaceserials.Add(circle)#.SerialNumber)

                                if self.createhz1offset.IsChecked and hz1offset_meter != 0.0:
                                    os_circle = wv.Add(clr.GetClrType(CadCircle))
                                    os_circle.CenterPoint = cp
                                    os_circle.Radius = (o.SiteImprovement.Diameter / 2) + \
                                                        o.SiteImprovement.Thickness + hz1offset_meter
                                    os_circle.Layer = self.layerpicker.SelectedSerialNumber

                                    addtosurfaceserials.Add(os_circle)#.SerialNumber)

                                if self.createhz2vzoffset.IsChecked:
                                    cp.Z = el + vz2offset_meter
                                    os_circle = wv.Add(clr.GetClrType(CadCircle))
                                    os_circle.CenterPoint = cp
                                    os_circle.Radius = (o.SiteImprovement.Diameter / 2) + \
                                                        o.SiteImprovement.Thickness + hz1offset_meter + hz2offset_meter
                                    os_circle.Layer = self.layerpicker.SelectedSerialNumber
                                    
                                    addtosurfaceserials.Add(os_circle)#.SerialNumber)

                            elif o.NodeType == UtilityNodeType(12):
                                boxlength = o.SiteImprovement.Length + (2 * o.SiteImprovement.Thickness)
                                boxwidth  = o.SiteImprovement.Width + (2 * o.SiteImprovement.Thickness)

                                boxvec = Vector3D(Vector3D(o.Point0), 1.0)
                                boxvec.Azimuth = 2 * math.pi - o.Orientation + math.pi/2
                                boxvec.Horizon = 0.0 # for some reason the vector was not exactly horizontal
                                                     # what lead to slight elevation errors at the corners
                               
                                # compute the 4 outline corners
                                corners = List[Point3D]()
                                # get to the first corner, adding vectors and rotating them
                                tmpp = o.Point0
                                tmpp.Z = el
                                boxvec.Length = boxlength / 2
                                tmpp += boxvec
                                boxvec.Rotate90(Side.Left)
                                boxvec.Length = boxwidth / 2
                                tmpp += boxvec
                                corners.Add(Point3D(tmpp))
                                # go to the 2nd corner
                                boxvec.Rotate90(Side.Left)
                                boxvec.Length = boxlength
                                tmpp += boxvec
                                corners.Add(Point3D(tmpp))
                                # go to the 3rd corner
                                boxvec.Rotate90(Side.Left)
                                boxvec.Length = boxwidth
                                tmpp += boxvec
                                corners.Add(Point3D(tmpp))
                                # go to the 4th corner
                                boxvec.Rotate90(Side.Left)
                                boxvec.Length = boxlength
                                tmpp += boxvec
                                corners.Add(Point3D(tmpp))
                                #close the line
                                corners.Add(Point3D(corners[0]))
                                
                                outlinepolyseg = PolySeg.PolySeg()
                                outlinepolyseg.Add(corners)

                                if self.createoutline.IsChecked:
                                    l = wv.Add(clr.GetClrType(Linestring))
                                    l.Layer = self.layerpicker.SelectedSerialNumber
                                    l.Append(outlinepolyseg, None, False, False)

                                    addtosurfaceserials.Add(l)#.SerialNumber)

                                if self.createhz1offset.IsChecked and hz1offset_meter != 0.0:
                                    # computing and creating the offset line
                                    poly_offset = outlinepolyseg.Offset(Side.Right, hz1offset_meter)
                                    
                                    if poly_offset[0] == True: # if the offset was computed
                                        l = wv.Add(clr.GetClrType(Linestring))
                                        l.Layer = self.layerpicker.SelectedSerialNumber
                                        l.Append(poly_offset[1], None, False, False)
                                        
                                        addtosurfaceserials.Add(l)#.SerialNumber)

                                if self.createhz2vzoffset.IsChecked:
                                    # computing and creating the offset line
                                    # re-fill the corner array with new elevations
                                    for i in range(0, corners.Count):
                                        p = Point3D(corners[i])
                                        p.Z = el + vz2offset_meter
                                        corners[i] = Point3D(p)
                                    
                                    # compute a new outline  with adjusted height
                                    outlinepolyseg = PolySeg.PolySeg()
                                    outlinepolyseg.Add(corners)
                                    poly_offset = outlinepolyseg.Offset(Side.Right, hz1offset_meter + hz2offset_meter)
                                    
                                    if poly_offset[0] == True: # if the offset was computed
                                        l = wv.Add(clr.GetClrType(Linestring))
                                        l.Layer = self.layerpicker.SelectedSerialNumber
                                        l.Append(poly_offset[1], None, False, False)
                                        
                                        addtosurfaceserials.Add(l)#.SerialNumber)

                            else:
                                self.error.Content += '\nfound an unsupported NodeType for line creation'
                            
                if self.addtosurface.IsChecked and addtosurfaceserials.Count > 0:
                    surface = wv.Lookup(self.addsurfacepicker.SelectedSerial)
                    if surface:
                        # this method wants a list with the actual objects, not just the serial numbers
                        surface.AddInfluences(addtosurfaceserials)
                    else:
                        self.error.Content += '\nno surface selected'
                    
                    #builder = surface.GetGemBatchBuilder()
                    #b = DTMSharpness.eSoft
                    #for sn in addtosurfaceserials:
                    #    o = wv.Lookup(sn)
                    #    polyseg = o.ComputePolySeg()
                    #    polyseg = polyseg.ToWorld() # just in case
                    #    if isinstance(o, CadCircle):
                    #        polyseg = polyseg.Linearize(0.0001, 0.0001, 0.1, None, False)
                    #
                    #    previndex = None
                    #
                    #    tt = polyseg.ToPoint3DArray()
                    #    tt2 = surfacegem.AddBreaklines(True, Byte(b), sn, tt, tt.Count)
                    #    tt3 = tt2

                        #for n in polyseg.Point3Ds():
                        #    curindex = builder.AddVertex(n) # that results the index in the new surface
                        #    if not previndex == None:
                        #        builder.AddBreakline(Byte(b), previndex[0], curindex[0])
                        #        previndex = curindex
                        #    else:
                        #        previndex = curindex
                    
                    #builder.Construction()
                    #builder.Commit()          

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

        self.success.Content += '\nDone'
        self.SaveOptions()  
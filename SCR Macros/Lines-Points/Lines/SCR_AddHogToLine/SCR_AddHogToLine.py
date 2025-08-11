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
    cmdData.Key = "SCR_AddHogToLine"
    cmdData.CommandName = "SCR_AddHogToLine"
    cmdData.Caption = "_SCR_AddHogToLine"
    cmdData.UIForm = "SCR_AddHogToLine"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "Lines"
        cmdData.ShortCaption = "Weight-Hog"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.13
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "add Weight Hog to Stringlines"
        cmdData.ToolTipTextFormatted = "add Weight Hog to Stringlines"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_AddHogToLine(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_AddHogToLine.xaml") as s:
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

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        #self.lfp = self.lunits.Properties.Copy()
        self.linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation

        self.toleranceheader.Header = 'define Hog-Definition and Chording Tolerance [' + self.linearsuffix + ']'

        self.hog.DistanceType = DistanceType.Z
        #self.hog.NumberOfDecimals = 4
        #self.hortol.NumberOfDecimals = 4
        #self.vertol.NumberOfDecimals = 4
        #self.nodespacing.NumberOfDecimals = 4

        #self.fergusontol.NumberOfDecimals = 4
        #self.fergusontol.MinValue = 0.00000001
        #self.fergusontol.Value = 0.00000001

        self.lType = clr.GetClrType(IPolyseg)
        #self.linepicker1.IsEntityValidCallback = self.IsValid
        self.objs.IsEntityValidCallback=self.IsValid

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def SetDefaultOptions(self):

        self.hog.Distance = OptionsManager.GetDouble("SCR_AddHogToLine.hog", 0.005)
        self.hortol.Distance = OptionsManager.GetDouble("SCR_AddHogToLine.hortol", 0.0001)
        self.vertol.Distance = OptionsManager.GetDouble("SCR_AddHogToLine.vertol", 0.0001)
        self.nodespacing.Distance = OptionsManager.GetDouble("SCR_AddHogToLine.nodespacing", 2)

        lserial = OptionsManager.GetUint("SCR_AddHogToLine.layerpicker", 8)
        o = self.currentProject.Concordance.Lookup(lserial) # get the object with the saved serial number
        if o != None:   # check if we actually got an object back
            if isinstance(o.GetSite(), LayerCollection):    # check if the object is a layer
                self.layerpicker.SetSelectedSerialNumber(lserial, InputMethod(3))
            else:  # if not then set Layer zero                       
                self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:   # if not then set Layer zero               
            self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))


        self.createisopachlinestring.IsChecked = OptionsManager.GetBool("SCR_AddHogToLine.createisopachlinestring", False)
        self.chordisopachlinestring.IsChecked = OptionsManager.GetBool("SCR_AddHogToLine.chordisopachlinestring", False)
        self.keephzcurves.IsChecked = OptionsManager.GetBool("SCR_AddHogToLine.keephzcurves", False)
        self.hogarc.IsChecked = OptionsManager.GetBool("SCR_AddHogToLine.hogarc", False)
        self.halfarc.IsChecked = OptionsManager.GetBool("SCR_AddHogToLine.halfarc", False)
        self.hogparabola.IsChecked = OptionsManager.GetBool("SCR_AddHogToLine.hogparabola", True)
        self.halfparabola.IsChecked = OptionsManager.GetBool("SCR_AddHogToLine.halfparabola", False)

        self.hoglinear.IsChecked = OptionsManager.GetBool("SCR_AddHogToLine.hoglinear", False)
        self.ferguson.IsChecked = OptionsManager.GetBool("SCR_AddHogToLine.ferguson", False)
        self.changestart.IsChecked = OptionsManager.GetBool("SCR_AddHogToLine.changestart", True)
        self.changeend.IsChecked = OptionsManager.GetBool("SCR_AddHogToLine.changeend", False)
        #self.fergusontol.Value = OptionsManager.GetDouble("SCR_AddHogToLine.fergusontol", 0.0001)

    def SaveOptions(self):

        OptionsManager.SetValue("SCR_AddHogToLine.hog", self.hog.Distance)
        OptionsManager.SetValue("SCR_AddHogToLine.hortol", self.hortol.Distance)
        OptionsManager.SetValue("SCR_AddHogToLine.vertol", self.vertol.Distance)
        OptionsManager.SetValue("SCR_AddHogToLine.nodespacing", self.nodespacing.Distance)
        
        OptionsManager.SetValue("SCR_AddHogToLine.layerpicker", self.layerpicker.SelectedSerialNumber)

        OptionsManager.SetValue("SCR_AddHogToLine.createisopachlinestring", self.createisopachlinestring.IsChecked)    
        OptionsManager.SetValue("SCR_AddHogToLine.chordisopachlinestring", self.chordisopachlinestring.IsChecked)    
        OptionsManager.SetValue("SCR_AddHogToLine.keephzcurves", self.keephzcurves.IsChecked)    
        OptionsManager.SetValue("SCR_AddHogToLine.hogarc", self.hogarc.IsChecked)    
        OptionsManager.SetValue("SCR_AddHogToLine.halfarc", self.halfarc.IsChecked)    
        OptionsManager.SetValue("SCR_AddHogToLine.hogparabola", self.hogparabola.IsChecked)    
        OptionsManager.SetValue("SCR_AddHogToLine.halfparabola", self.halfparabola.IsChecked)    

        OptionsManager.SetValue("SCR_AddHogToLine.hoglinear", self.hoglinear.IsChecked)    
        OptionsManager.SetValue("SCR_AddHogToLine.ferguson", self.ferguson.IsChecked)    
        OptionsManager.SetValue("SCR_AddHogToLine.changestart", self.changestart.IsChecked)    
        OptionsManager.SetValue("SCR_AddHogToLine.changeend", self.changeend.IsChecked)    
        #OptionsManager.SetValue("SCR_AddHogToLine.fergusontol", self.fergusontol.Value)

    def IsValid(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        return False

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def hoglinearChanged(self, sender, e):
        if self.hoglinear.IsChecked:
            self.changestart.IsEnabled = True
            self.changeend.IsEnabled = True
            #self.fergusontol.IsEnabled = False
        else:
            self.changestart.IsEnabled = False
            self.changeend.IsEnabled = False
            #self.fergusontol.IsEnabled = False

    def fergusonChanged(self, sender, e):
        if self.ferguson.IsChecked:
            self.changestart.IsEnabled = True
            self.changeend.IsEnabled = True
            #self.fergusontol.IsEnabled = True
            self.linearize.IsEnabled = False
        else:
            self.changestart.IsEnabled = False
            self.changeend.IsEnabled = False
            #self.fergusontol.IsEnabled = False
            self.linearize.IsEnabled = True

    def halfarcChanged(self, sender, e):
        if (self.hogarc.IsChecked and self.halfarc.IsChecked) or (self.hogparabola.IsChecked and self.halfparabola.IsChecked):
            self.changestart.IsEnabled = True
            self.changeend.IsEnabled = True
            #self.fergusontol.IsEnabled = False
        else:
            self.changestart.IsEnabled = False
            self.changeend.IsEnabled = False
            #self.fergusontol.IsEnabled = False

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''
        self.success.Content = ''

        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        
        wv = self.currentProject [Project.FixedSerial.WorldView]
        lgc = LayerGroupCollection.GetLayerGroupCollection(self.currentProject, False)
                
        wv.PauseGraphicsCache(True)

        # self.label_benchmark.Content = ''

        # settings = Model3DCompSettings.ProvideSettingsObject(self.currentProject)

        try:

            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                #l1 = self.linepicker1.Entity
                for l1 in self.objs:

                    # get and fix line name - Trimble Access doesn't like a hyphen at the start of the name
                    # so we need to make sure we don't actively add one
                    l1name = IName.Name.__get__(l1)
                    if not l1name == '':
                        l1name += ' - '

                    if isinstance(l1, self.lType):
                        polyseg1 = l1.ComputePolySeg()
                        polyseg1 = polyseg1.ToWorld()
                        polyseg1_v = l1.ComputeVerticalPolySeg()
                        
                        t1 = abs(self.hortol.Distance)
                        t2 = abs(self.vertol.Distance)
                        t3 = abs(self.nodespacing.Distance)
                        fergusontol = 0.0001 # abs(self.fergusontol.Value)

                        startstation = polyseg1.BeginStation
                        endstation = polyseg1.ComputeStationing()
                        fulllength = endstation - startstation
                        halflength = (endstation - startstation) / 2

                        hog = self.hog.Distance

                        if self.hogarc.IsChecked:
                            # create a vertical polyseg
                            hogpolyseg = PolySeg.PolySeg()

                            # and add the arc geometry
                            if not self.halfarc.IsChecked:
                                hogradius = (hog**2 + (fulllength/2)**2)/(2*hog)
                                if hog > endstation/2:
                                    self.error.Content += '\nArc Solution probably incorrect\nHog is greater than the arc radius'
                                    isopachname = "?? incorrect Arc ?? - " + l1name + "Isopach-Arc with " + str(hog*1000) + " mm Hog"
                                    linename = "?? incorrect Arc ?? - " + l1name + "Arc with " + str(hog*1000) + " mm Hog"
                                else:
                                    isopachname = l1name + "Isopach-Arc with " + str(hog*1000) + " mm Hog"
                                    linename = l1name + "Arc with " + str(hog*1000) + " mm Hog"

                                hogpolyseg.Add(ArcSegment(Point3D(startstation, 0), Point3D(endstation, 0), -1*hogradius))
                            else:
                                if hog > endstation/2:
                                    self.error.Content += '\nArc Solution probably incorrect\nHog is greater than the arc radius'
                                    isopachname = "?? incorrect Arc ?? - " + l1name + "Isopach-Half-Arc with " + str(hog*1000) + " mm Hog"
                                    linename = "?? incorrect Arc ?? - " + l1name + "Half-Arc with " + str(hog*1000) + " mm Hog"
                                else:
                                    isopachname = l1name + "Isopach-Half-Arc with " + str(hog*1000) + " mm Hog"
                                    linename = l1name + "Half-Arc with " + str(hog*1000) + " mm Hog"

                                hogradius = (hog**2 + (fulllength*2/2)**2)/(2*hog)
                                if self.changestart.IsChecked:
                                    hogpolyseg.BeginStation = startstation - fulllength
                                    hogpolyseg.Add(ArcSegment(Point3D(startstation - fulllength, 0), Point3D(endstation, 0), -1*hogradius))
                                else:
                                    hogpolyseg.Add(ArcSegment(Point3D(startstation, 0), Point3D(endstation + fulllength, 0), -1*hogradius))
                                    hogpolyseg.BeginStation = startstation

                        if self.hogparabola.IsChecked: # Parabola

                            # hog = a * endstation/2^2
                            # create a vertical polyseg
                            hogpolyseg = PolySeg.PolySeg()
                            # and add the parabola geometry
                            if not self.halfparabola.IsChecked:
                                a = hog / ((fulllength/2)**2)
                                para_slope = 2 * a * (fulllength/2)
                                para_el = para_slope * (fulllength/2)

                                hogpolyseg.Add(ParabolaSegment(Point3D(startstation, 0), Point3D(halflength, para_el), Point3D(fulllength, 0)))
                                isopachname = l1name + "Isopach-Parabola with " + str(hog*1000) + " mm Hog"
                                linename = l1name + "Parabola with " + str(hog*1000) + " mm Hog"
                            else:
                                a = hog / ((fulllength*2/2)**2)
                                para_slope = 2 * a * (fulllength*2/2)
                                para_el = para_slope * (fulllength*2/2)

                                isopachname = l1name + "Isopach-Half-Parabola with " + str(hog*1000) + " mm Hog"
                                linename = l1name + "Half-Parabola with " + str(hog*1000) + " mm Hog"
                                if self.changestart.IsChecked:
                                    hogpolyseg.Add(ParabolaSegment(Point3D(startstation - fulllength, 0), Point3D(startstation, para_el), Point3D(fulllength, 0)))
                                else:
                                    hogpolyseg.Add(ParabolaSegment(Point3D(startstation, 0), Point3D(fulllength, para_el), Point3D(endstation + fulllength, 0)))

                        if self.hoglinear.IsChecked: # linear
                            # create a vertical polyseg
                            hogpolyseg = PolySeg.PolySeg()
                            # and add the linear geometry
                            if self.changestart.IsChecked:
                                hogpolyseg.Add(SegmentLine(Point3D(startstation, hog), Point3D(endstation, 0)))
                            else:
                                hogpolyseg.Add(SegmentLine(Point3D(startstation, 0), Point3D(endstation, hog)))

                            isopachname = l1name + "Isopach-linear Line with " + str(hog*1000) + " mm dH"
                            linename = l1name + "linear Line with " + str(hog*1000) + " mm dH"


                        if self.ferguson.IsChecked: # Ferguson-Spline
                            # there is an issue with the discriminate function and small hog values, it won't return a point array
                            # the evaluate function does work
                            # we multiply the hog here and have to divide the resulting spline Y values later again
                            fergusonmulti = 1 # we'll iterate until the discriminate function returns us some values

                            while True:
                                if self.changestart.IsChecked:
                                    tangentin = Point3D(startstation - 0.0001, hog * fergusonmulti)
                                    startp = Point3D(startstation, hog * fergusonmulti)
                                    endp = Point3D(endstation, 0)
                                    tangentout = Point3D(endstation + 0.0001, 0)
                                else:
                                    tangentin = Point3D(startstation - 0.0001, 0)
                                    startp = Point3D(startstation, 0)
                                    endp = Point3D(endstation, hog * fergusonmulti)
                                    tangentout = Point3D(endstation + 0.0001, hog * fergusonmulti)

                                startend = Array[Point3D]([startp, endp])
                                ferguson = FergusonSpline()
                                ferguson.FitPoints(startend, tangentin, tangentout)

                                #tt = ferguson.Evaluate(0.3)
                                fergusonpts = ferguson.Discriminate(0, 1, fergusontol)
                                if fergusonpts.Count > 0 or fergusonmulti == 1000000:
                                    break

                                fergusonmulti *= 10

                            # now we have to scale the spline back
                            if fergusonpts.Count > 0:    
                                fixedfergusonpts = Array[Point3D]([Point3D()] * fergusonpts.Count)
                                for i in range(fergusonpts.Count):
                                    pt = fergusonpts[i]
                                    fixedfergusonpts[i] = (Point3D(pt.X, pt.Y / fergusonmulti))

                                # create a vertical polyseg
                                hogpolyseg = PolySeg.PolySeg()
                                # and add the linear geometry
                                hogpolyseg.Add(fixedfergusonpts)

                                isopachname = l1name + "Isopach-Ferguson-Spline with " + str(hog*1000) + " mm dH"
                                linename = l1name + "Ferguson-Spline with " + str(hog*1000) + " mm dH"
                            else:
                                self.error.Content += '\nHog-Value seems to be too small for Ferguson-Spline'
                                hogpolyseg = None

                        if hogpolyseg != None:
                            hogpolyseg.ComputeStationing()
                            
                            # chord the hog geometry                        
                            chordedhog = hogpolyseg.Linearize(t1, t2, t3, None, False)
                            
                            if self.createisopachlinestring.IsChecked:
                                
                                l_hog_iso = wv.Add(clr.GetClrType(Linestring))
                                if self.chordisopachlinestring.IsChecked:
                                    l_hog_iso.Append(polyseg1, chordedhog, False, False)
                                else:
                                    l_hog_iso.Append(polyseg1, hogpolyseg, False, False)
                                l_hog_iso.Name = isopachname
                                l_hog_iso.Layer = self.layerpicker.SelectedSerialNumber


                            # create the new line in the worldview and draw the segments
                            l_with_hog = wv.Add(clr.GetClrType(Linestring))
                            l_with_hog.Name = linename
                            l_with_hog.Layer = self.layerpicker.SelectedSerialNumber

                            nodes = chordedhog.ToPoint3DArray()
                            finalhogel = Array[Point3D]([Point3D()] * nodes.Count)
                            # go through the chord nodes
                            # get the original position and elevation, add the hog value and draw it
                            
                            for i in range(0, nodes.Count): # node list of linearized profile with X as chainage and Y as elevation
                                node = nodes[i]
                                p1 = polyseg1.FindPointFromStation(node.X)[1]
                                if polyseg1_v != None:
                                    p1.Z = polyseg1_v.ComputeVerticalSlopeAndGrade(node.X)[1]

                                if self.keephzcurves.IsChecked:
                                    p2 = Point3D(node.X, p1.Z + node.Y,0)
                                    finalhogel[i] = p2
                                else:
                                    p2 = Point3D(p1)
                                    p2.Z += node.Y

                                    e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                    e.Position = p2  # we draw that string line segment
                                    l_with_hog.AppendElement(e)

                            if self.keephzcurves.IsChecked:
                                # create the final hog polyseg (original + hog)
                                finalhogpolyseg = PolySeg.PolySeg()
                                # and add the profile geometry
                                finalhogpolyseg.Add(finalhogel)

                                l_with_hog.Append(polyseg1, finalhogpolyseg, False, False)

                    else:
                        self.error.Content += '\nskipped invalid Objects'
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
        
        wv.PauseGraphicsCache(False)

        #Keyboard.Focus(self.linepicker1)
        Keyboard.Focus(self.objs)

        self.SaveOptions()

    


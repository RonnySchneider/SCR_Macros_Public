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


from System.Collections.Generic import List, IEnumerable
exec(open(r"C:\ProgramData\Trimble\MacroCommands3\SCR Macros\SCR_Imports.py").read())

_OPTIONS = {
    "unitpicker":              "Meter",
    "addunitsuffix":           False,
    "textdecimals":            3.0,
    "textheight":              0.2,
    "thresh1":                 0.005,
    "thresh2":                 0.010,
    "thresh3":                 0.015,
    "thresh1colorpicker":      Color.Green,
    "thresh2colorpicker":      Color.Yellow,
    "thresh3colorpicker":      Color.Magenta,
    "thresh4colorpicker":      Color.Red,
    "showcompareRL":           False,
    "textcompareRL":           "DesRL=",
    "showasbuiltRL":           False,
    "textasbuiltRL":           "ABTOC=",
    "showRLdifference":        False,
    "textRLdifference":        "dZ=",
    "drawleader":              False,
    "leaderarrow":             False,
    "leaderarrowsize":         0.1,
    "usedtm":                  False,
    "comparesurfacepicker":    0,
    "usepoint":                False,
    "useline":                 False,
    "usemanualinput":          False,
    "manualelevation":         0.0,
    "uselayercontent":         False,
    "designlayerpicker":       8,
    "searchradius":            0.1,
    "computethickness":        False,
    "showdesignthicknessRL":   True,
    "textdesignthicknessRL":   "DesUS-RL=",
    "showdesignthickness":     True,
    "textdesignthickness":     "DesThick=",
    "showasbuiltthicknessRL":  True,
    "textasbuiltthicknessRL":  "ABBlndRL=",
    "showasbuiltthickness":    True,
    "textasbuiltthickness":    "ABThick=",
    "showthicknessdeviation":  True,
    "textthicknessdeviation":  "dThick=",
    "ablayerpicker":           8,
}

def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_ABReportElevationDifference"
    cmdData.CommandName = "SCR_ABReportElevationDifference"
    cmdData.Caption = "_SCR_ABReportElevationDifference"
    cmdData.UIForm = "SCR_ABReportElevationDifference"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = ""
    cmdData.HelpTopic = ""

    try:
        cmdData.DefaultTabKey = "SCR Reports"
        cmdData.DefaultTabGroupKey = "Reports"
        cmdData.ShortCaption = "Elevation Difference"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.30
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "AB Report Elevation Difference"
        cmdData.ToolTipTextFormatted = "AB Report Elevation Difference"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_ABReportElevationDifference(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_ABReportElevationDifference.xaml") as s:
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

        #buttons[2].Content = "Help"
        #buttons[2].Visibility = Visibility.Visible
        #buttons[2].Click += self.HelpClicked

        self.cadpointType = clr.GetClrType(CadPoint)
        self.coordpointType = clr.GetClrType(CoordPoint)
        self.polyType = clr.GetClrType(IPolyseg)

        self.comparepointpicker.IsEntityValidCallback = self.IsValidPoint
        self.abpointpicker.IsEntityValidCallback = self.IsValidPoint
        self.linepicker.IsEntityValidCallback = self.IsValidLine

        self.lType = clr.GetClrType(IPolyseg)
        self.designthicknessentity.IsEntityValidCallback = self.IsValidThicknessEntity
        self.asbuiltthicknessentity.IsEntityValidCallback = self.IsValidThicknessEntity

        types = Array [Type] (SurfaceTypeLists.AllWithCutFillMap)    # we fill an array with TBC object types, we could combine different types
        self.comparesurfacepicker.FilterByEntityTypes = types    # we fill the dropdownlist by applying that types array as filter
        self.comparesurfacepicker.AllowNone = False              # our list shall not show an empty field

        #self.manualelevation.NumberOfDecimals = 4

        self.abpointpicker.AutoTab = False
        self.abpointpicker.IsEntityValidCallback = self.IsValid

        SCRExpanders.wire_pairs([
            (self.expander_computethickness, self.computethickness),
        ])
		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass
        self.unitssetup(None, None)

    def IsValidThicknessEntity(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        if isinstance(o, Model3D) and not isinstance(o, clr.GetClrType(ProjectedSurface)):
            return True
        return False

    def IsValid(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.coordpointType):
            return True
        if isinstance(o, self.cadpointType):
            return True
        return False

    def SetDefaultOptions(self):
        SCROptions.LoadProjectOptions(self, "SCR_ABReportElevationDifference", _OPTIONS, self.currentProject)

    def SaveOptions(self):
        SCROptions.SaveProjectOptions(self, "SCR_ABReportElevationDifference", _OPTIONS, self.currentProject)

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
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.polyType):
            return True
        return False

    def IsValidPoint(self, serial):
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
        self.error.Content = ''
        self.success.Content = ''

        wv = self.currentProject [Project.FixedSerial.WorldView]

        wlc = self.currentProject[Project.FixedSerial.LayerContainer] # we get all the layers into an object, LayerCollection
        wl = wlc[self.designlayerpicker.SelectedSerialNumber]    # we get just the source layer as an object
        designmembers = wl.Members

        #check thresholds
        inputok=True
        try:

            thresh1 = self.thresh1.Distance
            thresh2 = self.thresh2.Distance
            thresh3 = self.thresh3.Distance
            searchradius = self.searchradius.Distance
            textheight = self.textheight.Distance
            textdecimals = int(self.textdecimals.Value)
        except:
            inputok=False
        
        # p1 - Design Point
        # p2 - AB Point
        if inputok:
            UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
            self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)

            pointlist = []
            for p2 in self.abpointpicker.SelectedMembers(self.currentProject): # go through all selected AB points
                #if p2.Layer != self.designlayerpicker.SelectedSerialNumber: # we don't want to report on points in the design layer
                if isinstance(p2, CoordPoint) or isinstance(p2, self.cadpointType): # and \p2.Layer != self.designlayerpicker.SelectedSerialNumber:
                    pointlist.Add(p2.Position)
            try:
                with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                    
                    p1 = Point3D(0,0,0)
                    if pointlist.Count > 0:
                        for p2 in pointlist: # go through all selected AB points
                            p1p2ok = True
                            # depending on what is selected we get the design point/elevation in different ways
                            # always getting a point p1 in order to keep the text creation simple
                            
                            if self.showasbuiltRL.IsChecked and self.showRLdifference.IsChecked == False and self.showcompareRL.IsChecked == False:
                                # we only want to show the Asbuilt RL, so we don't need to find a P1
                                # we need to disable the other Checkboxes
                                p1 = Point3D(p2)
                                self.usepoint.IsChecked = False
                                self.usedtm.IsChecked = False
                                self.useline.IsChecked = False
                                self.usemanualinput.IsChecked = False
                                self.computethickness.IsChecked = False

                            if self.usepoint.IsChecked: 
                                if self.comparepointpicker.SerialNumber != 0: #checking if a point was selected
                                    p1 = self.currentProject.Concordance.Lookup(self.comparepointpicker.SerialNumber).Position
                                else:
                                    self.error.Content += "\nno Design Point selected"    # if not show an error message
                                    p1p2ok = False                                          # set the ok check to false
                            
                            if self.usedtm.IsChecked:
                                comparesurface = wv.Lookup(self.comparesurfacepicker.SelectedSerial)    # we get our selected surface as object
                                tt = comparesurface.PickSurface(p2)
                                if tt[0]==True:
                                    p1 = tt[1]
                                else:
                                    self.error.Content += "\nAsbuilt Point is outside of the Surface"     # if not show an error message
                                    p1p2ok = False                                                          # set the ok check to false
                            
                            if self.useline.IsChecked:
                                if self.linepicker.Entity != None:
                                    p1 = self.perplinepoint(p2, self.linepicker.Entity)

                                    if p1 == None:
                                        self.error.Content += "\ncouldn't find a solution on the line"     # if not show an error message
                                        p1p2ok = False                                                       # set the ok check to false
                                else:
                                    self.error.Content += "\nno design line selected"

                            if  self.usemanualinput.IsChecked:
                                p1 = Point3D(p2)        # create 
                                p1.Z = self.manualelevation.Distance

                            if self.uselayercontent.IsChecked:
                                if designmembers.Count>0:
                                    pfound = False
                                    for designserial in designmembers: # go through all members of the design layer
                                        l = self.currentProject.Concordance.Lookup(designserial)
                                        if isinstance(l, self.lType):
                                            pcompute = self.perplinepoint(p2, l)
                                            if pcompute:
                                                if Point3D.Distance(pcompute, p2)[0] <= searchradius: # we check if it's actually within our search radius
                                                    if pfound == False:
                                                        p1 = pcompute
                                                        pfound = True
                                                    else: # if we find multiple solutions we only want the closest one
                                                        if Point3D.Distance(pcompute, p2)[0] < Point3D.Distance(p1, p2)[0]:
                                                            p1 = pcompute
                                                            pfound = True
                                else:
                                    self.error.Content += "\nDesign-Layer is empty"
                                if not pfound: p1p2ok = False

                            if p1p2ok and math.isnan(p1.Z):
                                self.error.Content += "\nDesign Point without Elevation"
                                p1p2ok = False
                            if p1p2ok and math.isnan(p2.Z):
                                self.error.Content += "\nAsbuilt Point without Elevation"
                                p1p2ok = False

                            if p1p2ok:

                                #p1.Z = round(p1.Z, abs(int(self.textdecimals.Text)))
                                #p2.Z = round(p2.Z, abs(int(self.textdecimals.Text)))

                                deltaz = p2.Z - p1.Z

                                thicknessok = False
                                # in case we also need to compute the thickness
                                if (self.computethickness.IsChecked and
                                   self.designthicknessentity.Entity != None and
                                   self.asbuiltthicknessentity.Entity != None):

                                    thicknessok = True
                                    # compute point on thickness design entity
                                    if isinstance(self.designthicknessentity.Entity, self.lType):
                                        p3 = self.perplinepoint(p1, self.designthicknessentity.Entity)

                                        if p3 == None:
                                            self.error.Content += "\ncouldn't find Thickness Design Elevation"     # if not show an error message
                                            thicknessok = False                                                       # set the ok check to false
                                    else:
                                        tt = self.designthicknessentity.Entity.PickSurface(p1)
                                        if tt[0]==True:
                                            p3 = tt[1]
                                        else:
                                            self.error.Content += "\ncouldn't find Thickness Design Elevation"     # if not show an error message
                                            thicknessok = False
                                    # compute point on asbuilt design entity
                                    if isinstance(self.asbuiltthicknessentity.Entity, self.lType):
                                        p4 = self.perplinepoint(p2, self.asbuiltthicknessentity.Entity)

                                        if p4 == None:
                                            self.error.Content += "\ncouldn't find Thickness Asbuilt Elevation"     # if not show an error message
                                            thicknessok = False                                                       # set the ok check to false
                                    else:
                                        tt = self.asbuiltthicknessentity.Entity.PickSurface(p2)
                                        if tt[0]==True:
                                            p4 = tt[1]
                                        else:
                                            self.error.Content += "\ncouldn't find Thickness Asbuilt Elevation"     # if not show an error message
                                            thicknessok = False
                                    
                                    if thicknessok:
                                        #p3.Z = round(p3.Z, abs(int(self.textdecimals.Text)))
                                        #p4.Z = round(p4.Z, abs(int(self.textdecimals.Text)))
                                        designthickness = p1.Z - p3.Z
                                        asbuiltthickness = p2.Z - p4.Z
                                        thicknessdeviation = asbuiltthickness - designthickness
                                        thicknessdesignRLstring = self.textdesignthicknessRL.Text + self.tooutputunit(p3.Z)
                                        thicknessdesignstring = self.textdesignthickness.Text + self.tooutputunit(designthickness)
                                        thicknessasbuiltRLstring = self.textasbuiltthicknessRL.Text + self.tooutputunit(p4.Z)
                                        thicknessasbuiltstring = self.textasbuiltthickness.Text + self.tooutputunit(asbuiltthickness)
                                        thicknessdeviationstring = self.textthicknessdeviation.Text + self.tooutputunit(thicknessdeviation)


                                # https://mkaz.blog/code/python-string-format-cookbook/
                                #compareRLstring = self.textcompareRL.Text + str("{:.{}f}".format(p1.Z, textdecimals))
                                #asbuiltRLstring = self.textasbuiltRL.Text + str("{:.{}f}".format(p2.Z, textdecimals))
                                #deltazstring = self.textRLdifference.Text + str("{:.{}f}".format(deltaz, textdecimals))
                                compareRLstring = self.textcompareRL.Text + self.tooutputunit(p1.Z)
                                asbuiltRLstring = self.textasbuiltRL.Text + self.tooutputunit(p2.Z)
                                deltazstring = self.textRLdifference.Text + self.tooutputunit(deltaz)

                                # find the right color
                                if deltaz <= thresh1: abcolor = self.thresh1colorpicker.SelectedColor
                                if deltaz >  thresh1: abcolor = self.thresh2colorpicker.SelectedColor
                                if deltaz >  thresh2: abcolor = self.thresh3colorpicker.SelectedColor
                                if deltaz >  thresh3: abcolor = self.thresh4colorpicker.SelectedColor

                                if self.showcompareRL.IsChecked or self.showasbuiltRL.IsChecked or self.showRLdifference.IsChecked:
                                    t = wv.Add(clr.GetClrType(MText))
                                    t.AlignmentPoint = p2
                                    t.TextString = ''

                                    if self.showcompareRL.IsChecked: t.TextString = self.addtotextstring(t.TextString, compareRLstring)
                                    if self.showasbuiltRL.IsChecked: t.TextString = self.addtotextstring(t.TextString, asbuiltRLstring)
                                    if self.showRLdifference.IsChecked: t.TextString = self.addtotextstring(t.TextString, deltazstring)
                                    
                                    if thicknessok:
                                        if self.showdesignthicknessRL.IsChecked: t.TextString = self.addtotextstring(t.TextString, thicknessdesignRLstring)
                                        if self.showdesignthickness.IsChecked: t.TextString = self.addtotextstring(t.TextString, thicknessdesignstring)
                                        if self.showasbuiltthicknessRL.IsChecked: t.TextString = self.addtotextstring(t.TextString, thicknessasbuiltRLstring)
                                        if self.showasbuiltthickness.IsChecked: t.TextString = self.addtotextstring(t.TextString, thicknessasbuiltstring)
                                        if self.showthicknessdeviation.IsChecked: t.TextString = self.addtotextstring(t.TextString, thicknessdeviationstring)

                                    t.Height = textheight
                                    t.Layer = self.ablayerpicker.SelectedSerialNumber
                                    t.Color = abcolor


                                    if self.drawleader.IsChecked:
                                        leaderpoints = List[Point3D]()
                                        leaderpoints.Add(t.AlignmentPoint)
                                        t.AlignmentPoint = Point3D(t.AlignmentPoint.X + 0.1, t.AlignmentPoint.Y - 0.1, t.AlignmentPoint.Z)
                                        leaderpoints.Add(t.AlignmentPoint)
                                        
                                        l = wv.Add(clr.GetClrType(Leader))
                                        l.AnnotationSerial = t.SerialNumber
                                        l.ScaleFactor = 1
                                        l.Points = leaderpoints
                                        l.Layer = self.ablayerpicker.SelectedSerialNumber
                                        if self.leaderarrow.IsChecked:
                                            l.LeaderType = LeaderType.LineWithArrow
                                        else:
                                            l.LeaderType = LeaderType.LineNoArrow
                                        l.ArrowheadType = DimArrowheadType.ArrowDefault
                                        l.ArrowheadSize = abs(self.leaderarrowsize.Distance)
                                        l.TextGap = 0.1
                                        l.Color = Color.DarkGray
                                        #l.AnnotationTargetDelta = t.AlignmentPoint

                    else:
                        self.error.Content += "\nno Asbuilt Point selected"
                        
                    failGuard.Commit()
                    self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
                    UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
            
            except Exception as e:
                tt = sys.exc_info()
                exc_type, exc_obj, exc_tb = sys.exc_info()
                # EndMark MUST be set no matter what
                # otherwise TBC won't work anymore and needs to be restarted
                self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
                UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
                self.error.Content += '\nan Error occurred - Result probably incomplete\n' + str(exc_type) + '\n' + str(exc_obj) + '\nLine ' + str(exc_tb.tb_lineno)
        
        
                
        
        if self.usedtm.IsChecked  or self.useline.IsChecked or self.usemanualinput.IsChecked or self.computethickness.IsChecked:
            Keyboard.Focus(self.abpointpicker)
        else:
            Keyboard.Focus(self.comparepointpicker)

        if self.showasbuiltRL.IsChecked and self.showRLdifference.IsChecked == False and self.showcompareRL.IsChecked == False:
            Keyboard.Focus(self.abpointpicker)

        self.SaveOptions()           

    def addtotextstring(self, t, addstring):
        if len(t) > 0:
            t += '\\P' + addstring
        else:
            t += addstring

        return t
        
    def perplinepoint(self, p2, l):

        outSegment=clr.StrongBox[Segment]()
        out_t=clr.StrongBox[float]()
        outPointOnCL = clr.StrongBox[Point3D]()
        station = clr.StrongBox[float]()
        perpVector3D=clr.StrongBox[Vector3D]()
        testDist=clr.StrongBox[float]()
        testside=clr.StrongBox[Side]()
        
        polyseg = l.ComputePolySeg()
        try:
            polyseg = polyseg.ToWorld()
        except:
            pass
        polyseg_v = l.ComputeVerticalPolySeg()

        if polyseg:
            # try to find a perpendicular solution on that line
            if polyseg.FindPointFromPoint(p2, outSegment, out_t, outPointOnCL, station, perpVector3D, testDist, testside): 
                p1 = outPointOnCL.Value
                if polyseg_v != None:
                    p1.Z = polyseg_v.ComputeVerticalSlopeAndGrade(station.Value)[1]
            
                return p1
            else:
                return None
        else:
            return None

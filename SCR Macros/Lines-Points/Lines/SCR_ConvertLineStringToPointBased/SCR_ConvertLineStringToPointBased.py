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
    cmdData.Key = "SCR_ConvertLinestringToPointBased"
    cmdData.CommandName = "SCR_ConvertLinestringToPointBased"
    cmdData.Caption = "_SCR_ConvertLinestringToPointBased"
    cmdData.UIForm = "SCR_ConvertLinestringToPointBased"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "Lines"
        cmdData.ShortCaption = "Linestring XY to P#"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.11
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Convert Linestring-Nodes from XYZ to Point-based"
        cmdData.ToolTipTextFormatted = "Convert Linestring-Nodes from XYZ to Point-based, including a global coordinate for transformation"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_ConvertLinestringToPointBased(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_ConvertLinestringToPointBased.xaml") as s:
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

        self.objs.IsEntityValidCallback = self.IsValid
        self.lType = clr.GetClrType(IPolyseg)
        
		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def SetDefaultOptions(self):
        #settingserial = OptionsManager.GetUint("SCR_ConvertLinestringToPointBased.layerpicker", 8) # 8 is FixedSerial for Layer Zero
        #o = self.currentProject.Concordance.Lookup(settingserial)
        #if o != None:
        #    if isinstance(o.GetSite(), LayerCollection):    
        #        self.layerpicker.SetSelectedSerialNumber(settingserial, InputMethod(3))
        #    else:                       
        #        self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        #else:                       
        #    self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))

        self.newpoints.IsChecked = OptionsManager.GetBool("SCR_ConvertLinestringToPointBased.newpoints", False)
        self.useexisting.IsChecked = OptionsManager.GetBool("SCR_ConvertLinestringToPointBased.useexisting", True)
        self.matchFC.IsChecked = OptionsManager.GetBool("SCR_ConvertLinestringToPointBased.matchFC", False)
        self.relayerpoints.IsChecked = OptionsManager.GetBool("SCR_ConvertLinestringToPointBased.relayerpoints", False)
        self.drawcircles.IsChecked = OptionsManager.GetBool("SCR_ConvertLinestringToPointBased.drawcircles", True)
        self.createnew.IsChecked = OptionsManager.GetBool("SCR_ConvertLinestringToPointBased.createnew", False)

    def SaveOptions(self):
        #OptionsManager.SetValue("SCR_ConvertLinestringToPointBased.layerpicker", self.layerpicker.SelectedSerialNumber)
        OptionsManager.SetValue("SCR_ConvertLinestringToPointBased.newpoints", self.newpoints.IsChecked)
        OptionsManager.SetValue("SCR_ConvertLinestringToPointBased.useexisting", self.useexisting.IsChecked)
        OptionsManager.SetValue("SCR_ConvertLinestringToPointBased.matchFC", self.matchFC.IsChecked)
        OptionsManager.SetValue("SCR_ConvertLinestringToPointBased.relayerpoints", self.relayerpoints.IsChecked)
        OptionsManager.SetValue("SCR_ConvertLinestringToPointBased.drawcircles", self.drawcircles.IsChecked)
        OptionsManager.SetValue("SCR_ConvertLinestringToPointBased.createnew", self.createnew.IsChecked)


    def IsValid(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        return False

    def MyFilterCriteria(self, filter, filterCriteria):
        return True

    def preparepointtables(self):

        # in large datasets we need a table to quickly find closeby points instead of going through all of them
        
        ProgressBar.TBC_ProgressBar.Title = "preparing lookup tables"
        wv = self.currentProject[Project.FixedSerial.WorldView]

        # find RawDataContainer as object
        for o in wv:
            if isinstance(o, RawDataContainer):
                rdc = o
                break
        
        self.lookupx = {}
        self.lookupy = {}

        for pserial in rdc.AllPoints:

            #pserial = rdc.AllPoints[i]   # go through all the point serials in the RawDataContainer
            rdc_point = self.currentProject.Concordance[pserial] # get the point object via its serial number        

            # https://stackoverflow.com/questions/20585920/how-to-add-multiple-values-to-a-dictionary-key

            keyx = math.trunc(rdc_point.Position.X)
            self.lookupx.setdefault(keyx, []).append(pserial)

            keyy = math.trunc(rdc_point.Position.Y)
            self.lookupy.setdefault(keyy, []).append(pserial)

        return

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''
        self.success.Content=''

        wv = self.currentProject [Project.FixedSerial.WorldView]
        wv.PauseGraphicsCache(True)

        self.preparepointtables()

        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        try:

            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                p1 = 0

                surfaceedits = []

                for o in self.objs:
                    #ProgressBar.TBC_ProgressBar.Title = "Convert Line " + str(p1) + '/' + str(self.objs.Count)
                    #if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(p1 * 100 / self.objs.Count)):
                    #    break
                    p1 += 1

                    if isinstance(o, Linestring):

                        pointidmissing = 0
                        
                        # get the line feature code
                        for observes in self.currentProject.Concordance.GetIsObservedBy(o.SerialNumber):
                            if isinstance(observes, LineFeature):
                                linecode = observes.Definition.Code
                        # if line is part of surfaces then set them to manual rebuild if necessary
                        for observedby in self.currentProject.Concordance.GetObserversOf(o.SerialNumber):
                            if isinstance(observedby, Model3D) and observedby.RebuildMethod == SurfaceRebuildMethod.Auto:
                                surfaceedits.Add(observedby)
                                observedby.RebuildMethod = SurfaceRebuildMethod.ByUser

                        polyseg = o.ComputePolySeg()
                        polyseg = polyseg.ToWorld()
                        polyseg_v = o.ComputeVerticalPolySeg()

                        outPointOnCL = clr.StrongBox[Point3D]()
                        station = clr.StrongBox[float]()

                        j = 0
                        createnew = self.createnew.IsChecked
                        inputlayer = self.currentProject.Concordance.Lookup(o.Layer)

                        if createnew:
                            l = wv.Add(clr.GetClrType(Linestring))
                            outputlayer = Layer.FindOrCreateLayer(self.currentProject, inputlayer.Name + ' - Point-based')
                            l.Layer = outputlayer.SerialNumber
                            outputlayer.LayerGroupSerial = inputlayer.LayerGroupSerial
                            l.Name = IName.Name.__get__(o)
                            l.Color = o.Color
                            l.LineStyle = o.LineStyle
                            l.LineTypeScale = o.LineTypeScale
                            l.Weight = o.Weight

                        # go through the elements of the linestring
                        for i in range(0, o.ElementCount):
                            ProgressBar.TBC_ProgressBar.Title = "Line " + str(p1) + '/' + str(self.objs.Count) + ' - Node ' + str(i + 1) + '/' + str(o.ElementCount)
                            if ProgressBar.TBC_ProgressBar.SetProgress(math.floor((i + 1) * 100 / o.ElementCount)):
                                break
                            
                            e = o.ElementAt(i)
                            etype = e.GetType()
                            #einterfaces = etype.GetInterfaces()
                            #newi = einterfaces[einterfaces.Count - 1]

                            if "XYZ" in etype.Name:

                                if not self.useexisting.IsChecked:
                                    # find free point number
                                    while PointHelper.Helpers.Find(self.currentProject, self.nameinput.SelectedName).Count != 0:
                                        self.nameinput.SelectedName = PointHelper.Helpers.Increment(self.nameinput.SelectedName, None, True)

                                    # compute the elevation
                                    polyseg.FindPointFromPoint(e.Position, outPointOnCL, station) # only to get the station value
                                    p_position = e.Position
                                    if p_position.Is2D: # in case the elevation is undefined
                                        p_position.Z = 0.0

                                    if polyseg_v != None:
                                        p_position.Z = polyseg_v.ComputeVerticalSlopeAndGrade(station.Value)[1]

                                    # create the point
                                    pnew_wv = CoordPoint.CreatePoint(self.currentProject, self.nameinput.SelectedName)
                                    # set the correct layer
                                    if createnew:
                                        pnew_wv.Layer = l.Layer
                                    else:
                                        pnew_wv.Layer = o.Layer
                                    # add the location to the point
                                    pnew_wv.AddPosition(p_position)
                                    # retrieve lat/long/elel 
                                    pnew_wv_glob = pnew_wv.GetPosition(CoordType.LLHGlobalElev) # result is a Trimble.Vce.Coordinates.ICoordinate
                                    # returns a double array containing the position values
                                    pnew_wv_glob = pnew_wv_glob.GetPosition(Array[CoordComponentType]([CoordComponentType.eHorizontal, CoordComponentType.eEllipsHeight]))
                                    
                                    # create a office entered WGS84 coordinate for the point
                                    # !!! Latitude and Longitude in Radians!!!
                                    keyed_coord = KeyedIn(CoordSystem.eWGS84, pnew_wv_glob[0], pnew_wv_glob[1], CoordQuality.eSurvey, \
                                                  pnew_wv_glob[2], CoordQuality.eSurvey, CoordComponentType.eEllipsHeight, \
                                                  System.DateTime.Now)
                                    # add it to the just created point
                                    OfficeEnteredCoord.AddOfficeEnteredCoord(self.currentProject, pnew_wv, keyed_coord)

                                else: # use existing Points
                                    foundpoint = False
                                    
                                    # new algorithm in order to reduce search time
                                    # uses lookup tables
                                    # if it doesn't find the dictionary key, hence if it doesn't find even a close match
                                    # the whole list serial/combination will fail with an exception
                                    try:
                                        closebyserials = list(set(self.lookupx[math.trunc(e.Position.X)]) & set(self.lookupy[math.trunc(e.Position.Y)]))
                                    except:
                                        closebyserials = []

                                    for sn in closebyserials:

                                        rdc_point = self.currentProject.Concordance[sn]

                                        #tt = Point3D.IsDuplicate3D(rdc_point.AnchorPoint, e.Position, 0.000001)
                                        if rdc_point.AnchorPoint.Is3D and e.Position.Is3D and Point3D.IsDuplicate3D(rdc_point.AnchorPoint, e.Position, 0.000001)[0]:
                                            if self.matchFC.IsChecked:
                                                matchFC = True
                                                # long point codes could potentially have parts of the line code somewhere in it
                                                # so we can't just test is linecode in pointcode
                                                # test if the start of the code matches
                                                for lci in range(0, len(linecode)):
                                                    if linecode[lci] != rdc_point.FeatureCode[lci]:
                                                        matchFC = False
                                                # in case the point code is longer than the line code
                                                # the only way it is a match is when the next character is a number
                                                # if the user added numbers in the field if he had multiple lines with the same code
                                                if matchFC and len(linecode) != len(rdc_point.FeatureCode):
                                                    tt = rdc_point.FeatureCode[len(linecode)]
                                                    if not str.isnumeric(rdc_point.FeatureCode[len(linecode)]):
                                                        matchFC = False
                                                if matchFC:
                                                    pnew_wv = rdc_point
                                                    foundpoint = True
                                                    break
                                            else:
                                                pnew_wv = rdc_point
                                                foundpoint = True
                                                break


                                        #tt = 1

                                    # relayer point if ticked
                                    if foundpoint and self.relayerpoints.IsChecked:
                                        rdc_point.Layer = o.Layer

                                # create new elements
                                if etype.Name == "XYZStraightElement":
                                    enew = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IPointIdLocation))
                                if etype.Name == "XYZSmoothCurveElement":
                                    enew = ElementFactory.Create(clr.GetClrType(ISmoothCurveSegment), clr.GetClrType(IPointIdLocation))
                                if etype.Name == "XYZBestFitArcElement":
                                    enew = ElementFactory.Create(clr.GetClrType(IBestFitArcSegment), clr.GetClrType(IPointIdLocation))
                                if etype.Name == "XYZArcElement":
                                    enew = ElementFactory.Create(clr.GetClrType(IArcSegment), clr.GetClrType(IPointIdLocation))
                                    enew.LargeSmall = e.LargeSmall
                                    enew.LeftRight = e.LeftRight
                                    enew.Radius = e.Radius
                                if etype.Name == "XYZPIArcElement":
                                    enew = ElementFactory.Create(clr.GetClrType(IPIArcSegment), clr.GetClrType(IPointIdLocation))
                                    enew.Radius = e.Radius
                                if etype.Name == "XYZTangentArcElement":
                                    enew = ElementFactory.Create(clr.GetClrType(ITangentArcSegment), clr.GetClrType(IPointIdLocation))
                                    enew.TangentType = e.TangentType

                                if self.useexisting.IsChecked and not foundpoint:
                                    enew = e
                                    pointidmissing += 1

                                    if self.drawcircles.IsChecked:
                                        # draw circle on warning layer
                                        # create the layer first if necessary
                                        linelayer = self.currentProject.Concordance.Lookup(o.Layer)
                                        outputlayer = Layer.FindOrCreateLayer(self.currentProject, linelayer.Name + ' - PointID Warnings')
                                        outputlayer.DefaultColor = Color.Red
                                        outputlayer.LayerGroupSerial = linelayer.LayerGroupSerial

                                        # in case we have an undefined elevation (question mark interpolation)
                                        # we still want to draw the circle at the correct elevation
                                        newcircle = wv.Add(clr.GetClrType(CadCircle))
                                        newcircle.Layer = outputlayer.SerialNumber
                                        newcircle.Radius = 0.2
                                        
                                        # compute the elevation
                                        if not e.Position.Is3D:
                                            polyseg.FindPointFromPoint(e.Position, outPointOnCL, station) # only to get the station value
                                            p_position = e.Position
                                            if polyseg_v != None:
                                                p_position.Z = polyseg_v.ComputeVerticalSlopeAndGrade(station.Value)[1]

                                            newcircle.CenterPoint = p_position
                                        else:
                                            newcircle.CenterPoint = e.Position

                                else:
                                    # that's the same for all of them
                                    enew.LocationSerialNumber = pnew_wv.SerialNumber


                                if createnew:
                                    l.AppendElement(enew)
                                else:
                                    o.ReplaceElementAt(enew, i)

                            else: # is already using PointID
                                if createnew:
                                    l.AppendElement(e)
                                else:
                                    pass # do nothing
                            
                        if createnew:
                            Linestring.ConvertPolySegToLinestringVertical(l, polyseg_v, 0.0)

                        if self.useexisting.IsChecked and pointidmissing > 0:
                            self.error.Content += '\nnot fully Point based:\nLine: ' + IName.Name.__get__(o) + '\non Layer: ' + inputlayer.Name + '\nNodes without ID: ' + str(pointidmissing) + '\n'

                    else:
                        self.error.Content += '\nfound unsupported Linetype\n    Name - ' + IName.Name.__get__(o) + "\n    Type - " + o.GetType().ToString()

                # set the changed surfaces back to auto rebuild
                if surfaceedits.Count != 0:
                    for s in surfaceedits:
                        s.RebuildMethod = SurfaceRebuildMethod.Auto

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

        ProgressBar.TBC_ProgressBar.Title = ''
        self.SaveOptions()
        wv.PauseGraphicsCache(False)


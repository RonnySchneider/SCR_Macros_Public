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
    cmdData.Key = "SCR_AutoConnectPoints"
    cmdData.CommandName = "SCR_AutoConnectPoints"
    cmdData.Caption = "_SCR_AutoConnectPoints"
    cmdData.UIForm = "SCR_AutoConnectPoints"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "Points"
        cmdData.ShortCaption = "AutoConnect Points"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.22
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Autoconnect Points"
        cmdData.ToolTipTextFormatted = "Start at a source point and always connect to the nearest one"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_AutoConnectPoints(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_AutoConnectPoints.xaml") as s:
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
        self.objs.IsEntityValidCallback = self.IsValid
        optionMenu = SelectionContextMenuHandler()
        # remove options that don't apply here
        optionMenu.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        self.objs.ButtonContextMenu = optionMenu
        
        self.cadpointType = clr.GetClrType(CadPoint)
        self.point3dType = clr.GetClrType(Point3D)
        self.coordpointType = clr.GetClrType(CoordPoint)
        self.pointcloudType = clr.GetClrType(PointCloudRegion)
        self.coordCtl1.ValueChanged += self.Coord1Changed

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def SetDefaultOptions(self):
        lserial = OptionsManager.GetUint("SCR_AutoConnectPoints.layerpicker", 8)
        o = self.currentProject.Concordance.Lookup(lserial) # get the object withj the saved serial number
        if o != None:   # check if we actually got an object back
            if isinstance(o.GetSite(), LayerCollection):    # check if the object is a layer
                self.layerpicker.SetSelectedSerialNumber(lserial, InputMethod(3))
            else:  # if not then set Layer zero                       
                self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:   # if not then set Layer zero               
            self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))

        self.usepointid.IsChecked = OptionsManager.GetBool("SCR_AutoConnectPoints.usepointid", True)
        self.boltspecial.IsChecked = OptionsManager.GetBool("SCR_AutoConnectPoints.boltspecial", False)

    def SaveOptions(self):
        OptionsManager.SetValue("SCR_AutoConnectPoints.layerpicker", self.layerpicker.SelectedSerialNumber)
        OptionsManager.SetValue("SCR_AutoConnectPoints.usepointid", self.usepointid.IsChecked)
        OptionsManager.SetValue("SCR_AutoConnectPoints.boltspecial", self.boltspecial.IsChecked)


    def IsValid(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.cadpointType):
            return True
        if isinstance(o, self.point3dType):
            return True
        if isinstance(o, self.coordpointType):
            return True
        if isinstance(o, self.pointcloudType):
            return True
        return False
        

    def selectionChanged(self, sender, e):        # in case we select a new surface from the list we update the min/max
        pass

        
    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def Coord1Changed(self, ctrl, e):
        # set keyboard focus if change was due to mouse pick
        if e.Cause == InputMethod.Mouse:     
            self.OkClicked(None, None)


    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''


        # self.label_benchmark.Content = ''

        # start_t = timer ()
        wv = self.currentProject [Project.FixedSerial.WorldView]


        pointlist=[]

        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())

        try:

            # the "with" statement will unroll any changes if something go wrong
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                for o in self.objs.SelectedMembers(self.currentProject):
                    if isinstance(o, self.cadpointType):
                        pointlist.Add([o.Point0, 0])
                    elif isinstance(o, self.point3dType):
                        tt=o
                    elif isinstance(o, self.coordpointType):
                        pointlist.Add([o.Position, o.SerialNumber])
                    elif isinstance(o, self.pointcloudType):
                        integration = o.Integration  # = SdePointCloudRegionIntegration
                        selectedid = integration.GetSelectedCloudId() # it seems the selected points form a sub-cloud

                        regiondb = integration.PointCloudDatabase # PointCloudDatabase
                        sdedb = regiondb.Integration # SdePointCloudDatabaseIntegration
                        scanpointlist = sdedb.GetPoints(selectedid)
                        for p in scanpointlist:
                            pointlist.Add([p, 0])
                
                if not pointlist.Count == 0:
                    
                    j = 0
                    jj = pointlist.Count
                    ProgressBar.TBC_ProgressBar.Title = "connect Points"
                    if self.boltspecial.IsChecked:  # always just connect the two closest points
                        
                        while pointlist.Count>1: # as long as we have more than 1 point in the list
                            i1 = 0
                            i2 = 1
                            if pointlist.Count>2:
                                for i in range(2, pointlist.Count): # we find the closest second point
                                    if Vector3D(pointlist[i1][0], pointlist[i][0]).Length2D < Vector3D(pointlist[i1][0], pointlist[i2][0]).Length2D:
                                        i2 = i
                            # now that we have two close points we draw the line and delete them from the list
                            l = wv.Add(clr.GetClrType(Linestring))
                            l.Layer = self.layerpicker.SelectedSerialNumber
                            if self.usepointid.IsChecked and pointlist[i1][1] > 0:
                                e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IPointIdLocation))
                                e.LocationSerialNumber = pointlist[i1][1] 
                            else:
                                e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                e.Position = pointlist[i1][0]
                            l.AppendElement(e)       
                            
                            if self.usepointid.IsChecked and pointlist[i2][1] > 0:
                                e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IPointIdLocation))
                                e.LocationSerialNumber = pointlist[i2][1] 
                            else:                            
                                e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                e.Position = pointlist[i2][0]
                            l.AppendElement(e)       

                            pointlist.RemoveAt(i2) # we have to remove the higher index first
                            pointlist.RemoveAt(i1)

                            # update the progress bar
                            j += 1
                            if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(j * 100 / jj)):
                                break   # function returns true if user pressed cancel

                    else:   # standard connect of all points
                        # create the line and add the start point as first element
                        l = wv.Add(clr.GetClrType(Linestring))
                        l.Layer = self.layerpicker.SelectedSerialNumber

                        if math.isnan(self.coordCtl1.Coordinate.X): # if the coordinate picker is empty use the first selected point
                            if self.usepointid.IsChecked and pointlist[0][1] > 0:
                                e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IPointIdLocation))
                                e.LocationSerialNumber = pointlist[0][1]
                            else:
                                e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                e.Position = pointlist[0][0]

                            startpoint = pointlist[0][0]

                        else: # we have to see if we snapped to a selected point
                            found = -1
                            for i in range(0, pointlist.Count): 
                                if Point3D.Distance(self.coordCtl1.Coordinate, pointlist[i][0])[0] == 0:
                                    found = i
                                    break
                            if found > -1:
                                if self.usepointid.IsChecked and pointlist[found][1] > 0:
                                    e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IPointIdLocation))
                                    e.LocationSerialNumber = pointlist[found][1]
                                    startpoint = pointlist[found][0]
                                else:
                                    e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                    e.Position = pointlist[found][0]
                                    startpoint = pointlist[found][0]
                                pointlist.RemoveAt(found)
                            else:        
                                e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                e.Position = self.coordCtl1.Coordinate  # use the value from the coordinate picker
                                startpoint = self.coordCtl1.Coordinate

                        l.AppendElement(e)       

                        linejuststarted = True
                        # as long as we have points that haven't been connected go on
                        while pointlist.Count > 0:
                            
                            # always reset the compare values
                            dist = 0
                            shortest_dist = 0
                            smallest_defl = 0
                            forsearchjuststarted = True
                            closest_p = None

                            # go through the rest of the points and find the next closest one
                            for i in reversed(range(0, pointlist.Count)):

                                if pointlist[i][1] != -1:

                                    # if we check for deflection angles must make the second coordinate input the next best point
                                    # no matter what
                                    if self.deflectionlimit.IsChecked and linejuststarted:
                                        for k in range(pointlist.Count):
                                            if Point3D.IsDuplicate3D(pointlist[k][0], self.coordCtl2.Coordinate, 0.000001)[0]:
                                                coord2index = k

                                        closest_p = pointlist[coord2index]
                                        shortest_dist = Point3D.Distance(startpoint, pointlist[coord2index][0])[0]
                                        pointlist.RemoveAt(coord2index)
                                        linejuststarted = False
                                        break # from for loop
                                                
                                    dist = Point3D.Distance(startpoint, pointlist[i][0])[0]
                                    
                                    if dist == 0:  # if the distance is 0, the point is the same as we just connected we remove it from the list and go on
                                        pointlist.RemoveAt(i)
                                        continue

                                    if forsearchjuststarted and not self.deflectionlimit.IsChecked:    # if we started going through the list we save it as closest solution, which it probably isn't
                                        closest_p = pointlist[i]
                                        shortest_dist = dist
                                        forsearchjuststarted = False
                                        continue

                                    if forsearchjuststarted and self.deflectionlimit.IsChecked and not linejuststarted:
                                        az1 = Vector3D(prevpoint, startpoint).Azimuth
                                        az2 = Vector3D(startpoint, pointlist[i][0]).Azimuth
                                        if abs(az1 - az2) < abs(self.deflectionlimitangle.Angle) and dist < 20:
                                            closest_p = pointlist[i]
                                            smallest_defl = abs(az1 - az2)
                                            shortest_dist = dist
                                            forsearchjuststarted = False
                                            continue

                                    # if we find a closer solution than we already have we save the coordinate
                                    if dist < shortest_dist:

                                        if self.deflectionlimit.IsChecked:
                                            az1 = Vector3D(prevpoint, startpoint).Azimuth
                                            az2 = Vector3D(startpoint, pointlist[i][0]).Azimuth
                                            if abs(az1 - az2) < abs(self.deflectionlimitangle.Angle) and dist < 20:
                                                closest_p = pointlist[i]
                                                smallest_defl = abs(az1 - az2)
                                                shortest_dist = dist
                                        else:
                                            closest_p = pointlist[i]
                                            shortest_dist = dist

                            if closest_p:
                                if self.usepointid.IsChecked and closest_p[1] > 0:
                                    e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IPointIdLocation))
                                    e.LocationSerialNumber = closest_p[1]
                                else:
                                    e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                    e.Position = closest_p[0]
                                l.AppendElement(e)
                            
                                prevpoint = startpoint
                                startpoint = closest_p[0]
                            else:
                                break
                            # pointlist.RemoveAt(shortest_i)
                            # update the progress bar
                            j += 1
                            if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(j * 100 / jj)):
                                break   # function returns true if user pressed cancel

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
        Keyboard.Focus(self.objs)
        self.SaveOptions()

           


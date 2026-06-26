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
    "selectcadpoints": True,
}

def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_SelectPtsOffLine"
    cmdData.CommandName = "SCR_SelectPtsOffLine"
    cmdData.Caption = "_SCR_SelectPtsOffLine"
    cmdData.UIForm = "SCR_SelectPtsOffLine"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "Points"
        cmdData.ShortCaption = "Select Points Off Line"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.16
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Select Points Off Line"
        cmdData.ToolTipTextFormatted = "Select Points Off Line"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_SelectPtsOffLine(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_SelectPtsOffLine.xaml") as s:
            wpf.LoadComponent (self, s)
        self.currentProject = currentProject
        self.macroFileFolder = macroFileFolder

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
        #types = Array [Type] ([CadPoint]) + Array [Type] ([Point3D])    # we fill an array with TBC object types, we could combine different types
        #self.linepicker1.IsEntityValidCallback = self.IsValid

        self.objs.IsEntityValidCallback = self.IsValid
        optionMenu = SelectionContextMenuHandler()
        # remove options that don't apply here
        optionMenu.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        self.objs.ButtonContextMenu = optionMenu
        
        self.lType = clr.GetClrType(IPolyseg)

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

        #self.linepicker1.ValueChanged += self.lineChanged
        #self.linepicker1.AutoTab = False
        #self.objs.ValueChanged += self.lineChanged
        self.objs.AutoTab = False

    def lineChanged(self, ctrl, e):
        self.OkClicked(None, None)
        
    def drawoverlay(self):

        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)
        self.overlayBag = OverlayBag(self.ViewOverlay) # create a new/empty overlaybag

        
        for p in self.verticewithnearbycoordpoint:
            self.overlayBag.AddMarker(p, GraphicMarkerTypes.HollowCircle_IndependentColor, Color.Green.ToArgb(), "", 0, 0, 4.0)

        for p in self.verticewithnearbycadpoint:
            self.overlayBag.AddMarker(p, GraphicMarkerTypes.HollowCircle_IndependentColor, Color.Blue.ToArgb(), "", 0, 0, 2.0)

        for p in self.verticewithnonearbypoint:
            self.overlayBag.AddMarker(p, GraphicMarkerTypes.HollowCircle_IndependentColor, Color.Red.ToArgb(), "", 0, 0, 4.0)

        for p in self.vpi:
            self.overlayBag.AddMarker(p, GraphicMarkerTypes.HollowCircle_IndependentColor, Color.Magenta.ToArgb(), "", 0, 0, 3.0)

        for sn in self.linenotfullypointbased:
            o = self.currentProject.Concordance[sn]
            if isinstance(o, self.lType):
                self.overlayBag.AddPolyline(SCROverlayBag.getpolypoints(o), Color.Red.ToArgb(), 4)

        for sn in self.linefullypointbased:
            o = self.currentProject.Concordance[sn]
            if isinstance(o, self.lType):
                self.overlayBag.AddPolyline(SCROverlayBag.getpolypoints(o), Color.Lime.ToArgb(), 4)

        for sn in self.wronglinetype:
            o = self.currentProject.Concordance[sn]
            if isinstance(o, self.lType):
                self.overlayBag.AddPolyline(SCROverlayBag.getpolypoints(o), Color.Brown.ToArgb(), 4)

        # if it only needs to be visible in all Planview then remove the Hoops3DViewGUID
        array = Array[Guid]([DisplayWindow.Hoops3DViewGUID, DisplayWindow.HoopsPlanViewGUID])
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.AddOverlayGeometry(array, self.overlayBag)

        return

    def SetDefaultOptions(self):
        SCROptions.LoadMacroOptions(self, "SCR_SelectPtsOffLine", _OPTIONS)

    def SaveOptions(self):
        SCROptions.SaveMacroOptions(self, "SCR_SelectPtsOffLine", _OPTIONS)

    def IsValid(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        return False

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def Dispose(self, cmd, disposing):
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)

    def preparepointtables(self):
        ProgressBar.TBC_ProgressBar.Title = "preparing lookup tables"

        items = []
        rdc = RawDataContainer.RawDataContainerIfFound(self.currentProject)
        if rdc is not None:
            for sn in rdc.AllPoints:
                p = self.currentProject.Concordance[sn]
                items.append((p.Position.X, p.Position.Y, p.Position.Z, sn))

        if self.selectcadpoints.IsChecked:
            wv = self.currentProject[Project.FixedSerial.WorldView]
            for o in wv:
                if isinstance(o, CadPoint) or isinstance(o, LocationComputer.DependentPoint):
                    items.append((o.Position.X, o.Position.Y, o.Position.Z, o.SerialNumber))

        self.point_tree = SCROctree()
        self.point_tree.build(items)

    def OkClicked(self, cmd, e):
        
        Keyboard.Focus(self.okBtn)
        self.errorcoordpointnearby.Content = ''
        self.errorcadpointnearby.Content = ''
        self.errornoverticenearby.Content = ''
        self.errorvpi.Content = ''
        self.errorcadline.Content = ''
        self.error.Content = ''
        self.success.Content = ''

        self.verticewithnearbycoordpoint = []
        self.verticewithnearbycadpoint = []
        self.verticewithnonearbypoint = []
        self.vpi = []
        self.linefullypointbased = []
        self.linenotfullypointbased = []
        self.wronglinetype = []

        #UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        #self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        wv = self.currentProject [Project.FixedSerial.WorldView]
        #wv.PauseGraphicsCache(True)

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        # we don't want the units to be included (so we make copy and turn that off). Otherwise get something like "12.50 ft"
        self.lfp = self.lunits.Properties.Copy()
        self.lfp.AddSuffix = False

        #inputok = True
        #l1 = self.linepicker1.Entity
        #if l1==None: 
        #    self.success.Content += '\nno Line 1 selected'
        #    inputok=False
        #if not isinstance(l1, Linestring):
        #    self.success.Content += '\nwrong Linetype'
        #    inputok=False
        
        #if inputok:
   
        if self.objs.Count > 0:
            
            self.preparepointtables()
            newselection = []

            #find PointManager as object
            for o in self.currentProject:
                if isinstance(o, PointManager):
                    pm = o
            rdc = RawDataContainer.RawDataContainerIfFound(self.currentProject)

            filename = os.path.expanduser('~/Downloads/Node-Output.csv')
            if File.Exists(filename):
                File.Delete(filename)
            with open(filename, 'w') as f:            

                for o in self.objs:

                    if isinstance(o, Linestring):

                        # the "with" statement will unroll any changes if something go wrong
                        #with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:


                        linenodes = self.currentProject.Concordance[o.SerialNumber].GetElements()
                        vnodes = self.currentProject.Concordance[o.SerialNumber].GetVerticalElements()

                        ProgressBar.TBC_ProgressBar.Title = "checking " + str(linenodes.Count) + " line nodes against " + str(rdc.AllPoints.Count if rdc is not None else 0) + " database points"
                        i = 0
                        verticeswithpointid = []
                        time1 = datetime.now()
                        for node in linenodes:

                            i += 1
                            # reducing the number of times the progressbar is updated, could lead to serious slowdown
                            # 10 % increments                            
                            #if (i * 100 // linenodes.Count) % 20 == 0:
                            #tt = datetime.now() - time1 # delivers .days .microseconds . seconds
                            if (datetime.now() - time1).seconds > 1:    
                                if ProgressBar.TBC_ProgressBar.SetProgress(i * 100 // linenodes.Count):
                                    break   # function returns true if user pressed cancel
                            time1 = datetime.now()

                            outputline = ''
                            try:
                                node_location = self.currentProject.Concordance[node.LocationSerialNumber]
                                newselection.Add(node.LocationSerialNumber)
                                verticeswithpointid.Add(node.LocationSerialNumber)
                                #if isinstance(node_location, LocationComputer.DependentPoint):
                                outputline = node_location.Name
                                #else:
                                #    outputline = node_location.AnchorName


                                outputline += ',' + self.lunits.Format(self.lunits.Convert(node_location.Position.X, LinearType.Display), self.lfp)
                                outputline += ',' + self.lunits.Format(self.lunits.Convert(node_location.Position.Y, LinearType.Display), self.lfp)
                                outputline += ',' + self.lunits.Format(self.lunits.Convert(node_location.Position.Z, LinearType.Display), self.lfp)
                                if pm.AssociatedRDFeatures(node.LocationSerialNumber).Count > 0:
                                    attr_list = pm.AssociatedRDFeatures(node.LocationSerialNumber)[0].Attributes
                                    if attr_list.Count > 0:
                                        for att in attr_list:
                                            outputline += ',' + str(att.Name) + ":" + str(att.Value)
                            
                            except: # if it's a node that is just a coordinate

                                foundpoint = False

                                if node.Position.Is3D:
                                    closebyserials = self.point_tree.find_within(node.Position.X, node.Position.Y, node.Position.Z, 0.001)
                                else:
                                    closebyserials = self.point_tree.find_within_2d(node.Position.X, node.Position.Y, 0.001)
                                
                                for pserial in closebyserials:

                                    db_point = self.currentProject.Concordance[pserial]

                                    if isinstance(db_point, CoordPoint):
                                        if Vector3D(db_point.Position, node.Position).Length == 0:
                                            newselection.Add(pserial)


                                            outputline = "??? " + db_point.Name
                                            outputline += ',' + self.lunits.Format(self.lunits.Convert(db_point.Position.X, LinearType.Display), self.lfp)
                                            outputline += ',' + self.lunits.Format(self.lunits.Convert(db_point.Position.Y, LinearType.Display), self.lfp)
                                            outputline += ',' + self.lunits.Format(self.lunits.Convert(db_point.Position.Z, LinearType.Display), self.lfp)
                                            if pm.AssociatedRDFeatures(pserial).Count > 0:
                                                attr_list = pm.AssociatedRDFeatures(pserial)[0].Attributes
                                                if attr_list.Count > 0:
                                                    for att in attr_list:
                                                        outputline += ',' + str(att.Name) + ":" + str(att.Value)
                                            outputline += ' ???'

                                            foundpoint = True
                                            self.verticewithnearbycoordpoint.Add(db_point.Position)
                                            break

                                    elif self.selectcadpoints.IsChecked and (isinstance(db_point, CadPoint) or isinstance(db_point, LocationComputer.DependentPoint)):
                                        if Vector3D(db_point.Position, node.Position).Length == 0:
                                            newselection.Add(pserial)

                                            outputline = "??? CadPoint"
                                            outputline += ',' + self.lunits.Format(self.lunits.Convert(db_point.Position.X, LinearType.Display), self.lfp)
                                            outputline += ',' + self.lunits.Format(self.lunits.Convert(db_point.Position.Y, LinearType.Display), self.lfp)
                                            outputline += ',' + self.lunits.Format(self.lunits.Convert(db_point.Position.Z, LinearType.Display), self.lfp)
                                            outputline += ' ???'

                                            foundpoint = True
                                            self.verticewithnearbycadpoint.Add(db_point.Position)
                                            break

                                if foundpoint == False:
                                    outputline += '!!! Linevertex without DB-Point'
                                    
                                    outputline += ',' + self.lunits.Format(self.lunits.Convert(node.Position.X, LinearType.Display), self.lfp)
                                    outputline += ',' + self.lunits.Format(self.lunits.Convert(node.Position.Y, LinearType.Display), self.lfp)
                                    outputline += ',' + self.lunits.Format(self.lunits.Convert(node.Position.Z, LinearType.Display), self.lfp)
                                    outputline += ' !!!'

                                    if node.Position.Is3D:
                                        self.verticewithnonearbypoint.Add(node.Position)
                                    else:
                                        tt = self.computepointwithelevation(node.Position, o)
                                        self.verticewithnonearbypoint.Add(self.computepointwithelevation(node.Position, o))

                        for node in vnodes: # vertical nodes are never pointbased

                            polyseg = o.ComputePolySeg()
                            polyseg = polyseg.ToWorld()
                            polyseg_v = o.ComputeVerticalPolySeg()

                            p_new = polyseg.FindPointFromStation(node.StationDistance)[1]
                            p_new.Z = node.Elevation

                            self.vpi.Add(p_new)

                            #self.success.Content += outputline + "\n"
                            f.write(outputline + "\n")
                            
                        
                        if verticeswithpointid.Count == linenodes.Count:
                            self.linefullypointbased.Add(o.SerialNumber)
                        else: 
                            self.linenotfullypointbased.Add(o.SerialNumber)
                            if self.verticewithnearbycoordpoint.Count > 0:
                                self.errorcoordpointnearby.Content = 'Vertice without PointID - but ID-Point nearby'
                            if self.verticewithnearbycadpoint.Count > 0:
                                self.errorcadpointnearby.Content = 'Vertice without PointID - simple CAD-Point nearby'
                            if self.verticewithnonearbypoint.Count > 0:
                                self.errornoverticenearby.Content = 'Vertice without PointID'

                        if vnodes.Count > 0:
                            self.errorvpi.Content = 'VPI are never point-based'
                    
                    else: # something else than Linestring, i.e. CadLine

                        self.errorcadline.Content = 'non-Linestrings are never point-based'
                        self.wronglinetype.Add(o.SerialNumber)

                f.close()
                GlobalSelection.Items(self.currentProject).Set(newselection)

            self.drawoverlay()

            #failGuard.Commit()
            
            #self.success.Content += '\n\nResult also written to:\n'
            #self.success.Content += str(filename)


        #self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
        #UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())

        self.success.Content += '\n\nDone'
        self.SaveOptions()
        ProgressBar.TBC_ProgressBar.Title = ""
        Keyboard.Focus(self.objs)

    def computepointwithelevation(self, p, l):

        polyseg = l.ComputePolySeg()
        polyseg = polyseg.ToWorld()
        polyseg_v = l.ComputeVerticalPolySeg()

        outPointOnCL = clr.StrongBox[Point3D]()
        station = clr.StrongBox[float]()

        if polyseg.FindPointFromPoint(p, outPointOnCL, station):
            p_new = outPointOnCL.Value
            if polyseg_v != None:  
                p_new.Z = polyseg_v.ComputeVerticalSlopeAndGrade(station.Value)[1]

        return p_new



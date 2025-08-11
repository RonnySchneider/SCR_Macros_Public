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
    cmdData.Key = "SCR_DepthToFC"
    cmdData.CommandName = "SCR_DepthToFC"
    cmdData.Caption = "_SCR_DepthToFC"
    cmdData.UIForm = "SCR_DepthToFC"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Expld-SNR-Relay-Prop"
        cmdData.DefaultTabGroupKey = "Properties"
        cmdData.ShortCaption = "Depth to FC"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.04
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "apply (average) distance from surface to points/lines to FC"
        cmdData.ToolTipTextFormatted = "apply (average) distance from surface to points/lines to FC"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_DepthToFC(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_DepthToFC.xaml") as s:
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

        self.objs.IsEntityValidCallback = self.IsValid
        optionMenu = SelectionContextMenuHandler()
        # remove options that don't apply here
        optionMenu.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        self.objs.ButtonContextMenu = optionMenu

        self.lType = clr.GetClrType(IPolyseg)
        self.coordpointType = clr.GetClrType(CoordPoint)
        self.cadpointType = clr.GetClrType(CadPoint)

        types = Array[Type](SurfaceTypeLists.AllWithCutFillMap)
        self.surfacepicker.FilterByEntityTypes = types
        self.surfacepicker.AllowNone = False

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        self.lfp = self.lunits.Properties.Copy()
        # self.linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation

        # get the units for chainage distance
        self.chunits = self.currentProject.Units.Station
        self.chfp = self.chunits.Properties.Copy()

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass


    def SetDefaultOptions(self):
        # Select surface
        try:    self.surfacepicker.SelectIndex(OptionsManager.GetInt("SCR_DepthToFC.surfacepicker", 0))
        except: self.surfacepicker.SelectIndex(0)

        self.changed12d.IsChecked = OptionsManager.GetBool("SCR_DepthToFC.changed12d", True)
        self.changefc.IsChecked = OptionsManager.GetBool("SCR_DepthToFC.changefc", False)
        self.fcnameline.SelectedName = OptionsManager.GetString("SCR_DepthToFC.fcnameline", "Average Depth (m)")
        self.fcnamepoint.SelectedName = OptionsManager.GetString("SCR_DepthToFC.fcnamepoint", "Depth (m)")

    def SaveOptions(self):
        OptionsManager.SetValue("SCR_DepthToFC.surfacepicker", self.surfacepicker.SelectedIndex)
        OptionsManager.SetValue("SCR_DepthToFC.changed12d", self.changed12d.IsChecked)
        OptionsManager.SetValue("SCR_DepthToFC.changefc", self.changefc.IsChecked)
        OptionsManager.SetValue("SCR_DepthToFC.fcnameline", self.fcnameline.SelectedName)
        OptionsManager.SetValue("SCR_DepthToFC.fcnamepoint", self.fcnamepoint.SelectedName)
    
    def IsValid(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        if isinstance(o, self.coordpointType):
            return True
        if isinstance(o, self.cadpointType):
            return True
        return False

    def drawoverlay(self):

        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)
        self.overlayBag = OverlayBag(self.ViewOverlay) # create a new/empty overlaybag

        
        for p in self.nodepth:
            self.overlayBag.AddMarker(p, GraphicMarkerTypes.HollowCircle_IndependentColor, Color.Blue.ToArgb(), "", 0, 0, 4.0)

        for sn in self.lineswithno12d:
            o = self.currentProject.Concordance[sn]
            if isinstance(o, self.lType):
                self.overlayBag.AddPolyline(self.getpolypoints(o), Color.Magenta.ToArgb(), 4)
            elif isinstance(o, self.coordpointType) or isinstance(o, self.cadpointType):
                self.overlayBag.AddMarker(o.Position, GraphicMarkerTypes.HollowCircle_IndependentColor, Color.Magenta.ToArgb(), "", 0, 0, 3.0)

        for sn in self.lineswithnofc:
            o = self.currentProject.Concordance[sn]
            if isinstance(o, self.lType):
                self.overlayBag.AddPolyline(self.getpolypoints(o), Color.Green.ToArgb(), 4)
            elif isinstance(o, self.coordpointType) or isinstance(o, self.cadpointType):
                self.overlayBag.AddMarker(o.Position, GraphicMarkerTypes.HollowCircle_IndependentColor, Color.Green.ToArgb(), "", 0, 0, 3.0)
        

        # if it only needs to be visible in all Planview then remove the Hoops3DViewGUID
        array = Array[Guid]([DisplayWindow.Hoops3DViewGUID, DisplayWindow.HoopsPlanViewGUID])
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.AddOverlayGeometry(array, self.overlayBag)

        return

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


    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def Dispose(self, cmd, disposing):
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.errordepth.Content = ''
        self.errorfc.Content = ''
        self.error12d.Content = ''

        self.error.Content = ''
        self.success.Content = ''

        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)

        uda = self.currentProject.Lookup(27) # User Defined Attributes are always Serial# 27

        wv = self.currentProject [Project.FixedSerial.WorldView]

        self.success.Content=''

        try:

            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                
                pc = PointCollection.ProvidePointCollection(self.currentProject)
                surface = self.currentProject.Concordance.Lookup(self.surfacepicker.SelectedSerial)

                self.nodepth = []
                self.lineswithnofc = []
                self.lineswithno12d = []

                for o in self.objs:

                    depthaveragestr = ''

                    if isinstance(o, self.lType):

                        # get the line data as polyseg, in world coordinates
                        polyseg = o.ComputePolySeg()
                        polyseg = polyseg.ToWorld()
                        polyseg_v = o.ComputeVerticalPolySeg()

                        polysegcombined = polyseg.ComputePolySeg3D(polyseg_v)

                        depthlist = []
                        for p in polysegcombined.Point3Ds():    # go through all the nodes in the linestring

                            outPoint = clr.StrongBox[Point3D]()
                            if surface.PickSurface(p, outPoint):
                                depthlist.Add(p.Z - outPoint.Value.Z)
                            else:
                                self.nodepth.Add(p)
                        
                        depthaveragefloat = sum(depthlist) / depthlist.Count
                        depthaveragestr = "{:.4f}".format(depthaveragefloat)

                    if isinstance(o, self.coordpointType) or isinstance(o, self.cadpointType):

                        outPoint = clr.StrongBox[Point3D]()
                        if surface.PickSurface(o.Position, outPoint):
                            depthaveragefloat = o.Position.Z - outPoint.Value.Z
                            depthaveragestr = "{:.4f}".format(depthaveragefloat)
                        else:
                            self.nodepth.Add(o.Position)

                    if self.changed12d.IsChecked:
                        attrfound = False
                        try:
                            dic = json.loads(uda.GetAssociatedAttributes(o.SerialNumber)['Sitech12dAttributes'])
                            for i in range(0, dic.Count):
                                if isinstance(o, self.lType):
                                    if dic[i]["Name"] == self.fcnameline.SelectedName:
                                        dic[i]["ValueForOutput"] = depthaveragestr
                                        dic[i]["Value"] = depthaveragestr
                                        attrfound = True
                                else:
                                    if dic[i]["Name"] == self.fcnamepoint.SelectedName:
                                        dic[i]["ValueForOutput"] = depthaveragestr
                                        dic[i]["Value"] = depthaveragestr
                                        attrfound = True

                            oatt = uda.GetAssociatedAttributes(o.SerialNumber)
                            oatt['Sitech12dAttributes'] = json.dumps(dic)
                        except:
                            self.lineswithno12d.Add(o.SerialNumber)

                        if not attrfound:            
                            self.lineswithno12d.Add(o.SerialNumber)

                    elif self.changefc.IsChecked:
                        
                        attrfound = False
                        for observes in self.currentProject.Concordance.GetIsObservedBy(o.SerialNumber):
                            try:
                                for attr in observes.Attributes:
                                    if isinstance(o, self.lType):
                                        if attr.Name == self.fcnameline.SelectedName:
                                            attr.Value = depthaveragefloat
                                            attrfound = True
                                    else:
                                        if attr.Name == self.fcnamepoint.SelectedName:
                                            attr.Value = depthaveragefloat
                                            attrfound = True
                            except:
                                pass

                        if not attrfound:            
                            self.lineswithnofc.Add(o.SerialNumber)


                self.lineswithnofc = list(set(self.lineswithnofc)) # remove duplicates
                self.lineswithno12d = list(set(self.lineswithno12d)) # remove duplicates

                if self.nodepth.Count > 0:
                    self.errordepth.Content += 'could not compute depth for all nodes'
                if self.lineswithnofc.Count > 0:
                    self.errorfc.Content += 'found Object without the requested FC-Attribute'
                if self.lineswithno12d.Count > 0:
                    self.error12d.Content += 'found Object without the requested 12d Attribute'

                self.drawoverlay()

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
        self.SaveOptions()


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
    cmdData.Key = "SCR_CopyXSectionsToPlanview"
    cmdData.CommandName = "SCR_CopyXSectionsToPlanview"
    cmdData.Caption = "_SCR_CopyXSectionsToPlanview"
    cmdData.UIForm = "SCR_CopyXSectionsToPlanview"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Reports"
        cmdData.DefaultTabGroupKey = "Sheets and Dynaviews"
        cmdData.ShortCaption = "Copy XSectionsToPlanview"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.04
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "copy XSection content to Planview"
        cmdData.ToolTipTextFormatted = "copy XSection content to Planview"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png") # we have to include a icon revision, otherwise TBC might not show the new one
        cmdData.ImageSmall = b
    except:
        pass


class SCR_CopyXSectionsToPlanview(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_CopyXSectionsToPlanview.xaml") as s:
            wpf.LoadComponent (self, s)
        self.currentProject = currentProject

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

        self.coordpick2.CoordinateType = Array[CoordinateType]([CoordinateType.SheetX, CoordinateType.SheetY, CoordinateType.SheetZ])
        self.coordpick3.CoordinateType = Array[CoordinateType]([CoordinateType.SheetX, CoordinateType.SheetY, CoordinateType.SheetZ])

        try:
            self.SetDefaultOptions()
        except:
            pass

    def SetDefaultOptions(self):

        self.drawup.IsChecked = OptionsManager.GetBool("SCR_CopyXSectionsToPlanview.drawup", True)
        self.drawdown.IsChecked = OptionsManager.GetBool("SCR_CopyXSectionsToPlanview.drawdown", False)

        wv = self.currentProject[Project.FixedSerial.WorldView]
        x = OptionsManager.GetDouble("SCR_CopyXSectionsToPlanview.coordpick1_x", 0.0)
        y = OptionsManager.GetDouble("SCR_CopyXSectionsToPlanview.coordpick1_y", 0.0)
        z = OptionsManager.GetDouble("SCR_CopyXSectionsToPlanview.coordpick1_z", 0.0)
        self.coordpick1.SetCoordinate(Point3D(x, y, z), self.currentProject, wv.CoordinateSystemDefinition)
        
        for o in self.currentProject:
            if isinstance(o, PlanSetSheetViews):
                csd = o.CoordinateSystemDefinition
        x = OptionsManager.GetDouble("SCR_CopyXSectionsToPlanview.coordpick2_x", 0.010)
        y = OptionsManager.GetDouble("SCR_CopyXSectionsToPlanview.coordpick2_y", 0.035)
        z = OptionsManager.GetDouble("SCR_CopyXSectionsToPlanview.coordpick2_z", 0.0)
        self.coordpick2.SetCoordinate(Point3D(x, y, z), self.currentProject, csd)
        x = OptionsManager.GetDouble("SCR_CopyXSectionsToPlanview.coordpick3_x", 0.410)
        y = OptionsManager.GetDouble("SCR_CopyXSectionsToPlanview.coordpick3_y", 0.287)
        z = OptionsManager.GetDouble("SCR_CopyXSectionsToPlanview.coordpick3_z", 0.0)
        self.coordpick3.SetCoordinate(Point3D(x, y, z), self.currentProject, csd)

    def SaveOptions(self):
        OptionsManager.SetValue("SCR_CopyXSectionsToPlanview.drawup", self.drawup.IsChecked)    
        OptionsManager.SetValue("SCR_CopyXSectionsToPlanview.drawdown", self.drawdown.IsChecked)    

        OptionsManager.SetValue("SCR_CopyXSectionsToPlanview.coordpick1_x", self.coordpick1.Coordinate.X)
        OptionsManager.SetValue("SCR_CopyXSectionsToPlanview.coordpick1_y", self.coordpick1.Coordinate.Y)
        OptionsManager.SetValue("SCR_CopyXSectionsToPlanview.coordpick1_z", self.coordpick1.Coordinate.Z)
        OptionsManager.SetValue("SCR_CopyXSectionsToPlanview.coordpick2_x", self.coordpick2.Coordinate.X)
        OptionsManager.SetValue("SCR_CopyXSectionsToPlanview.coordpick2_y", self.coordpick2.Coordinate.Y)
        OptionsManager.SetValue("SCR_CopyXSectionsToPlanview.coordpick2_z", self.coordpick2.Coordinate.Z)
        OptionsManager.SetValue("SCR_CopyXSectionsToPlanview.coordpick3_x", self.coordpick3.Coordinate.X)
        OptionsManager.SetValue("SCR_CopyXSectionsToPlanview.coordpick3_y", self.coordpick3.Coordinate.Y)
        OptionsManager.SetValue("SCR_CopyXSectionsToPlanview.coordpick3_z", self.coordpick3.Coordinate.Z)

    def OkClicked(self, cmd, e):
        
        #prog = Action[int]
        #test = Array[float](3*[0.000]) + Array[Point3D](4 * [Point3D()])

        self.success.Content = ""
        self.error.Content = ""

        wv = self.currentProject [Project.FixedSerial.WorldView]
        dp = self.currentProject.CreateDuplicator()

        inputok = True

        if not (self.coordpick3.Coordinate.X > self.coordpick2.Coordinate.X) and \
           not (self.coordpick3.Coordinate.Y > self.coordpick2.Coordinate.Y):
            inputok = False
            self.error.Content='check Frame definition'
        
        activeForm = TrimbleOffice.TheOffice.MainWindow.AppViewManager.ActiveView

        if not isinstance(activeForm, HoopsSheetView):
            inputok = False
            self.error.Content='activate Sheetview'

        if not activeForm.View.LeftMouseOperation == LeftMouseModeType.PolygonSelection:
            self.error.Content = "\nyou must be in 'Polygon Select Mode'"
            return

        if inputok:
            #UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
            self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
            
            try:
            
                with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:


                    # need to save the serials since we are selecting items with the fake mouse gesture
                    # and change the global selection all the time
                    savedserials = []
                    for o in GlobalSelection.SelectedMembers(self.currentProject):
                        # add if it's a Sheetset of to correct type
                        if isinstance(o, XSSheetSet) or isinstance(o, ProfileSheetSet):
                            savedserials.Add(o.SerialNumber)
                        # but also if it's at least a single Sheet of this kind
                        if isinstance(o, BasicSheet):
                            if isinstance(o.GetSite(), XSSheetSet) or isinstance(o.GetSite(), ProfileSheetSet):
                                savedserials.Add(o.SerialNumber)

                    scale = 1
                    column = 0
                    p0 = Point3D(0,0,0) # needs to stay at 0,0,0
                    
                    if self.drawup.IsChecked:
                        dir = 1
                    else:
                        dir = -1

                    p1 = self.coordpick1.Coordinate
                    if p1.Is2D: p1.Z = 0 

                    # now go through the selected serials and do the magic
                    for s in savedserials:
                    
                        sheetobj = self.currentProject.Concordance.Lookup(s)
                        setname = ""

                        sheetserials = []
                        if isinstance(sheetobj, BasicSheet):
                            sheetserials.Add(s)
                            setname = sheetobj.GetSite().Name
                            if isinstance(sheetobj.GetSite(), XSSheetSet):
                                scale = 1 / sheetobj.GetSite().XsSheetSettings.HorizontalScale
                            if isinstance(sheetobj.GetSite(), ProfileSheetSet):
                                scale = 1 / sheetobj.GetSite().ProfileSheetSettings.HorizontalScale
                            
                        elif isinstance(sheetobj, XSSheetSet):
                            setname = sheetobj.Name
                            scale = 1 / sheetobj.XsSheetSettings.HorizontalScale
                            for s in sheetobj.ContainedSerials:
                                # we don't want any text that's sitting on the SheetSet level
                                if isinstance(self.currentProject.Concordance.Lookup(s), BasicSheet):
                                    sheetserials.Add(s)
                        elif isinstance(sheetobj, ProfileSheetSet):
                            setname = sheetobj.Name
                            scale = 1 / sheetobj.ProfileSheetSettings.HorizontalScale
                            for s in sheetobj.ContainedSerials:
                                # we don't want any text that's sitting on the SheetSet level
                                if isinstance(self.currentProject.Concordance.Lookup(s), BasicSheet):
                                    sheetserials.Add(s)

                        column += 1
                        framewidth = self.coordpick3.Coordinate.X - self.coordpick2.Coordinate.X
                        frameheight = self.coordpick3.Coordinate.Y - self.coordpick2.Coordinate.Y
                        xoff = self.coordpick2.Coordinate.X
                        yoff = self.coordpick2.Coordinate.Y

                        
                        # draw the SheetSet name
                        t = wv.Add(clr.GetClrType(MText))
                        if self.drawdown.IsChecked:
                            # for MText you have to use Trimble.Vce.ForeignCad.AttachmentPoint
                            # 1 2 3
                            # 4 5 6
                            # 7 8 9
                            t.AttachPoint = AttachmentPoint(7)
                        else:
                            t.AttachPoint = AttachmentPoint(1)
                        t.AlignmentPoint = p1 + Vector3D(0, -1 * dir * 0.005 * scale, 0)
                        t.TextString = setname
                        t.Height = 0.01 * scale

                        for i in range(0, sheetserials.Count):

                            xshift = -(xoff * scale)
                            if self.drawup.IsChecked:
                                yshift = -(yoff * scale) + (dir * i * frameheight * scale)
                            else:
                                yshift = -(yoff * scale) + (dir * (i+1) * frameheight * scale)
                            p3 = p1 + Vector3D(xshift, yshift,0)
                            td = TransformData(Matrix4D.BuildTransformMatrix(p0, p3, 0, scale, scale, scale), Matrix4D(Vector3D.Zero))

                            # coordinates for selection polygon
                            # needs to be here since the fake mouse empties the list every time
                            ptsList = List[Point3D]() # we need a generic list, not an array
                            ptsList.Add(Point3D(xoff, yoff, 0))
                            ptsList.Add(Point3D(xoff, yoff + frameheight, 0))
                            ptsList.Add(Point3D(xoff + framewidth, yoff + frameheight, 0))
                            ptsList.Add(Point3D(xoff + framewidth, yoff, 0))
                            ptsList.Add(Point3D(xoff, yoff, 0))

                            #for sheetobserves in self.currentProject.Concordance.GetIsObservedBy(sheetserials[i]):
                            #    tt = sheetobserves


                            #testlimit = Limits3D(Point3D(0.010,0.035,0), Point3D(0.410,0.287,0))

                            framepoly = PolySeg.PolySeg()
                            framepoly.Add(ptsList)

                            l = wv.Add(clr.GetClrType(Linestring))
                            l.Append(framepoly, None, False, False)
                            l.Transform(td)

                            activeForm.SetSheetSerial(sheetserials[i])
                            #activeForm.UpdateDisplay()
                            #time.sleep(0.5)

                            op =  activeForm.View.CurrentOperator # should be PolygonSelectionOperator
                            op.PolygonStarted = True
                            op.PolygonPoints = ptsList
                            activeForm.View.PolygonSelectionStarted = True # need to fake a mouse gesture
                            tt = op.PerformSelection()
                            #activeForm.View.PolygonSelectionStarted = False # need to fake a mouse gesture
                            
                            selectionSet = GlobalSelection.Items(self.currentProject)


                            for o in selectionSet:

                                # create a copy of the selected object
                                
                                o_new = wv.Add(clr.GetClrType(o.GetType()))
                                o_new.CopyBody(self.currentProject.Concordance, self.currentProject.TransactionManager, o, dp)                        

                                #oldsite = o.GetSite()
                                #o_new.SetSite(wv)
                                #self.currentProject.Concordance.AddObserver(o_new.SerialNumber, 2)
                                #self.currentProject.Concordance.RemoveObserver(o_new.SerialNumber, oldsite.SerialNumber)#RemoveObserver(uint target, uint source) source is the observer

                                o_new.Transform(td)

                                tt2 = 1
                        
                        p1 = p1 + Vector3D(framewidth * scale, 0,0)

                    failGuard.Commit()
                    #UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
                    self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
            
            except Exception as e:
                tt = sys.exc_info()
                exc_type, exc_obj, exc_tb = sys.exc_info()
                # EndMark MUST be set no matter what
                # otherwise TBC won't work anymore and needs to be restarted
                self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
                #UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
                self.error.Content += '\nan Error occurred - Result probably incomplete\n' + str(exc_type) + '\n' + str(exc_obj) + '\nLine ' + str(exc_tb.tb_lineno)

        self.SaveOptions()

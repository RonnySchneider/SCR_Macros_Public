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
    cmdData.Key = "SCR_LocateDynaviewSource"
    cmdData.CommandName = "SCR_LocateDynaviewSource"
    cmdData.Caption = "_SCR_LocateDynaviewSource"
    cmdData.UIForm = "SCR_LocateDynaviewSource"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Reports"
        cmdData.DefaultTabGroupKey = "Sheets and Dynaviews"
        cmdData.ShortCaption = "Locate DynaView Source"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.03
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "locate the source outline of a DynaView"
        cmdData.ToolTipTextFormatted = "locate the source outline of a DynaView"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png") # we have to include a icon revision, otherwise TBC might not show the new one
        cmdData.ImageSmall = b
    except:
        pass


class SCR_LocateDynaviewSource(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_LocateDynaviewSource.xaml") as s:
            wpf.LoadComponent (self, s)
        self.currentProject = currentProject


    def OnLoad(self, cmd, buttons, event):
        self.okBtn = buttons[0]
        self.Caption = cmd.Command.Caption

        self.dynapicker.IsEntityValidCallback = self.IsValid
        self.dynapicker.ValueChanged += self.dynapickerChanged

        try:
            self.SetDefaultOptions()
        except:
            pass

    def SetDefaultOptions(self):
        self.jumptosheet.IsChecked = OptionsManager.GetBool("SCR_LocateDynaviewSource.jumptosheet", False)

    def SaveOptions(self):
        OptionsManager.SetValue("SCR_LocateDynaviewSource.jumptosheet", self.jumptosheet.IsChecked)


    def IsValid(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, DynaView):
            return True
        return False

    def dynapickerChanged(self, ctrl, e):
        
        self.OkClicked(None, None)


    def OkClicked(self, cmd, e):
        
        #prog = Action[int]
        #test = Array[float](3*[0.000]) + Array[Point3D](4 * [Point3D()])

        self.success.Content = ""
        self.error.Content = ""

        o = self.dynapicker.Entity

        if isinstance(o, DynaView):

            dynaviewwindow = TrimbleOffice.TheOffice.MainWindow.AppViewManager.ActiveView
            GlobalSelection.Clear()
            GlobalSelection.Items(self.currentProject).Set([o.Boundary])
            boundary = self.currentProject.Concordance.Lookup(o.Boundary)
            boundarylayer = self.currentProject.Concordance.Lookup(boundary.Layer)
            if boundarylayer.LayerGroupSerial == 0:
                boundarylayergroupname = "<None> -> Layers"
            else:
                boundarylayergroupname = self.currentProject.Concordance.Lookup(boundarylayer.LayerGroupSerial).Name

            tt = boundary.GetSite()
            boundarysite = boundary.GetSite()
            try:
                boundarysitename = boundary.GetSite().Name
            except:
                boundarysitename = ""

            boundarytype = boundary.GetType().FullName
            self.success.Content += "\nBoundary is located in:"
            self.success.Content += "\n" + boundarysite.Description + " - " + boundarysitename
            self.success.Content += "\n\nin Layer-Group:"
            self.success.Content += "\n" + boundarylayergroupname
            self.success.Content += "\n\non Layer:"
            self.success.Content += "\n" + boundarylayer.Name
            self.success.Content += "\n\nis of Type:"
            self.success.Content += "\n" + boundarytype
            self.error.Content += "\nBoundary is now selected"
            self.error.Content += "\n\ndon't forget to turn on CAD-Group !!!"
            
            polyseg = boundary.ComputePolySeg()
            try:
                polyseg = polyseg.ToWorld()
            except:
                pass

            loopsfound = clr.StrongBox[bool]()
            centroid = clr.StrongBox[Point3D]()
            polyside = clr.StrongBox[Side]()
            region = clr.StrongBox[RegionBuilder]()

            polyseg.AreaWithLoops(loopsfound, centroid, polyside, region)

            cgrav = centroid.Value

            if math.isnan(cgrav.Z): cgrav.Z = 0

            
            if isinstance(boundarysite, PlanSetSheetViews) or \
                isinstance(boundarysite, PlanSetSheetView) or \
                isinstance(boundarysite, SheetSet) or \
                isinstance(boundarysite, BasicSheet):
                
                if self.jumptosheet.IsChecked:
                    dynaviewwindow.Activate() # just in case another view is active, the setsheetserial would fail
                    dynaviewwindow.SetSheetSerial(boundarysite.SerialNumber)
                    dynaviewwindow.CenterView(cgrav)

                    self.error.Content += "\n\nSheet selected and centered"

            else:

                for v in TrimbleOffice.TheOffice.MainWindow.AppViewManager.Views:
                    i = 0
                    if isinstance(v, clr.GetClrType(Hoops2dView)) or \
                       isinstance(v, clr.GetClrType(Hoops3dView)) and not \
                       isinstance(v, clr.GetClrType(HoopsSheetView)):

                        i += 1
                        v.CenterView(cgrav)
                        
                        #if isinstance(v, clr.GetClrType(Hoops2dView)):
                        #    #campos, camtarget, camup, camwidth, camheight, strporj = v.Viewing.GetCamera()
                        #    #campos = cgrav + Point3D(0,0,100)
                        #    #camdir = Vector3D(campos, cgrav)
                        #    #camup = Vector3D(camup)
                        #    #v.RenderedModel.SetCameraToArea(campos, camdir, camup, 50, 50)
                        #    v.RenderedModel.SetCameraToArea(366400, 8117400, 366650, 8117600)
                        #
                        #tt = 1

                if i > 0:
                    self.error.Content += "\n\n2D/3D Views centered"

        self.SaveOptions()


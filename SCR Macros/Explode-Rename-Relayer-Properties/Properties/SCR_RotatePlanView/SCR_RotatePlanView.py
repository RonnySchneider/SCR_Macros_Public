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
    cmdData.Key = "SCR_RotatePlanView"
    cmdData.CommandName = "SCR_RotatePlanView"
    cmdData.Caption = "_SCR_RotatePlanView"
    cmdData.UIForm = "SCR_RotatePlanView"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Expld-SNR-Relay-Prop"
        cmdData.DefaultTabGroupKey = "Properties"
        cmdData.ShortCaption = "Rotate Plan View"
        cmdData.DefaultRibbonToolSize = 0 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.04
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Rotate Plan View"
        cmdData.ToolTipTextFormatted = "Rotate Plan View"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_RotatePlanView(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_RotatePlanView.xaml") as s:
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
        self.linepicker1.IsEntityValidCallback=self.IsValid
        self.linepicker1.ValueChanged += self.lineChanged
        self.station.ValueChanged += self.stationChanged
        self.station.AutoTab = False
        
        self.bearing.ValueChanged += self.bearingChanged
        self.bearing.AutoTab = False

        self.lType = clr.GetClrType(IPolyseg)

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def SetDefaultOptions(self):
        settings = ConstructionCommandsSettings.ProvideObject(self.currentProject)
        #self.bearing.Direction = settings.GetDouble("SCR_RotatePlanView.bearing", 0)
        self.usealignment.IsChecked = settings.GetBoolean("SCR_RotatePlanView.usealignment", False)
        #serial = settings.GetUInt32("SCR_RotatePlanView.l1",0)
        #o = self.currentProject.Concordance.Lookup(serial)
        #if isinstance(o, self.lType):
        #    self.linepicker1.SerialNumber = serial
        #    self.station.Distance = settings.GetDouble("SCR_RotatePlanView.station", 0)

        self.save1_bearing.Direction = settings.GetDouble("SCR_RotatePlanView.save1_bearing", 0)
        self.save1_text.Text = settings.GetString("SCR_RotatePlanView.save1_text", "Save 1")
        self.save2_bearing.Direction = settings.GetDouble("SCR_RotatePlanView.save2_bearing", 0)
        self.save2_text.Text = settings.GetString("SCR_RotatePlanView.save2_text", "Save 2")
        self.save3_bearing.Direction = settings.GetDouble("SCR_RotatePlanView.save3_bearing", 0)
        self.save3_text.Text = settings.GetString("SCR_RotatePlanView.save3_text", "Save 3")
        self.save4_bearing.Direction = settings.GetDouble("SCR_RotatePlanView.save4_bearing", 0)
        self.save4_text.Text = settings.GetString("SCR_RotatePlanView.save4_text", "Save 4")
        self.save5_bearing.Direction = settings.GetDouble("SCR_RotatePlanView.save5_bearing", 0)
        self.save5_text.Text = settings.GetString("SCR_RotatePlanView.save5_text", "Save 5")

    def SaveOptions(self):
        settings = ConstructionCommandsSettings.ProvideObject(self.currentProject)
        #settings.SetDouble("SCR_RotatePlanView.bearing", self.bearing.Direction)
        settings.SetBoolean("SCR_RotatePlanView.usealignment", self.usealignment.IsChecked)
        #try:
        #    settings.SetUInt32("SCR_RotatePlanView.l1", self.linepicker1.Entity.SerialNumber)
        #    settings.SetDouble("SCR_RotatePlanView.station", self.station.Distance)
        #except:
        #    pass
        settings.SetDouble("SCR_RotatePlanView.save1_bearing", self.save1_bearing.Direction)
        settings.SetString("SCR_RotatePlanView.save1_text", self.save1_text.Text)
        settings.SetDouble("SCR_RotatePlanView.save2_bearing", self.save2_bearing.Direction)
        settings.SetString("SCR_RotatePlanView.save2_text", self.save2_text.Text)
        settings.SetDouble("SCR_RotatePlanView.save3_bearing", self.save3_bearing.Direction)
        settings.SetString("SCR_RotatePlanView.save3_text", self.save3_text.Text)
        settings.SetDouble("SCR_RotatePlanView.save4_bearing", self.save4_bearing.Direction)
        settings.SetString("SCR_RotatePlanView.save4_text", self.save4_text.Text)
        settings.SetDouble("SCR_RotatePlanView.save5_bearing", self.save5_bearing.Direction)
        settings.SetString("SCR_RotatePlanView.save5_text", self.save5_text.Text)

    def lineChanged(self, ctrl, e):
        l1 = self.linepicker1.Entity
        if l1 != None:
            self.station.StationProvider = l1

    def stationChanged(self, ctrl, e):
        self.OkClicked(None, None)
        Keyboard.Focus(self.station)

    def bearingChanged(self, ctrl, e):
        self.OkClicked(None, None)

    def save_Click(self, sender, e):
        activeForm = TrimbleOffice.TheOffice.MainWindow.AppViewManager.ActiveView
        if not isinstance(activeForm, clr.GetClrType(Hoops2dView)):
            return
        else:
            activeFormState = Hoops2dView.Hoops2dViewState(activeForm)
            cam_az = Vector3D(activeFormState.cameraUpVector).Azimuth

            if self.radio1.IsChecked: self.save1_bearing.Direction = math.pi/2 - cam_az
            if self.radio2.IsChecked: self.save2_bearing.Direction = math.pi/2 - cam_az
            if self.radio3.IsChecked: self.save3_bearing.Direction = math.pi/2 - cam_az
            if self.radio4.IsChecked: self.save4_bearing.Direction = math.pi/2 - cam_az
            if self.radio5.IsChecked: self.save5_bearing.Direction = math.pi/2 - cam_az

        self.SaveOptions()

    def ccw90_Click(self, sender, e):
        activeForm = TrimbleOffice.TheOffice.MainWindow.AppViewManager.ActiveView
        if not isinstance(activeForm, clr.GetClrType(Hoops2dView)):
            return
        else:
            activeFormState = Hoops2dView.Hoops2dViewState(activeForm)
            alignvector = Vector3D(activeFormState.cameraUpVector)

            alignvector.RotateAboutZ(-math.pi/2)

            activeForm.SetCameraUpVector(alignvector)
            #activeForm.CenterView(anchor)
            activeForm.ToggleGrids(True)    # toggle grid twice to view it correctly
            activeForm.ToggleGrids(True)

    def cw90_Click(self, sender, e):
        activeForm = TrimbleOffice.TheOffice.MainWindow.AppViewManager.ActiveView
        if not isinstance(activeForm, clr.GetClrType(Hoops2dView)):
            return
        else:
            activeFormState = Hoops2dView.Hoops2dViewState(activeForm)
            alignvector = Vector3D(activeFormState.cameraUpVector)

            alignvector.RotateAboutZ(math.pi/2)

            activeForm.SetCameraUpVector(alignvector)
            #activeForm.CenterView(anchor)
            activeForm.ToggleGrids(True)    # toggle grid twice to view it correctly
            activeForm.ToggleGrids(True)

    def restore_Click(self, sender, e):
        self.success.Content = ''
        activeForm = TrimbleOffice.TheOffice.MainWindow.AppViewManager.ActiveView
        if not isinstance(activeForm, clr.GetClrType(Hoops2dView)):
            self.success.Content += '\nplease activate 2D-Planview-Window first'
            return
        else:
            self.usealignment.IsChecked = False
            if self.radio1.IsChecked: self.bearing.Direction = self.save1_bearing.Direction
            if self.radio2.IsChecked: self.bearing.Direction = self.save2_bearing.Direction
            if self.radio3.IsChecked: self.bearing.Direction = self.save3_bearing.Direction
            if self.radio4.IsChecked: self.bearing.Direction = self.save4_bearing.Direction
            if self.radio5.IsChecked: self.bearing.Direction = self.save5_bearing.Direction
            self.OkClicked(None, None)

    def reset_Click(self, sender, e):
        self.success.Content = ''
        activeForm = TrimbleOffice.TheOffice.MainWindow.AppViewManager.ActiveView
        if not isinstance(activeForm, clr.GetClrType(Hoops2dView)):
            self.success.Content += '\nplease activate 2D-Planview-Window first'
            return
        else:
            self.usealignment.IsChecked = False
            self.bearing.Direction = math.pi/2
            self.OkClicked(None, None)

    def usealignmentChanged(self, sender, e):
        if self.usealignment.IsChecked:
            self.alignmentbox.IsEnabled = True
            self.bearing.IsEnabled = False
        else:
            self.alignmentbox.IsEnabled = False
            self.bearing.IsEnabled = True


    def IsValid(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        return False
        
    def CancelClicked(self, cmd, args):
        
        cmd.CloseUICommand ()


    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''


        # self.label_benchmark.Content = ''
        self.success.Content = ''
        # start_t = timer ()
        wv = self.currentProject [Project.FixedSerial.WorldView]
        inputok = True
        activeForm = TrimbleOffice.TheOffice.MainWindow.AppViewManager.ActiveView

        if not isinstance(activeForm, clr.GetClrType(Hoops2dView)):
            self.success.Content += '\nplease activate 2D-Planview-Window first'
            return
        else:
            activeFormState = Hoops2dView.Hoops2dViewState(activeForm)
            cam_az = Vector3D(activeFormState.cameraUpVector).Azimuth
            self.SaveOptions()
            if self.bearing.IsEnabled:
                
                alignvector = Vector3D(1,0,0)
                alignvector.RotateAboutZ(self.bearing.Direction)

                activeForm.SetCameraUpVector(alignvector)
                #activeForm.CenterView(anchor)
                activeForm.ToggleGrids(True)    # toggle grid twice to view it correctly
                activeForm.ToggleGrids(True)

            if self.usealignment.IsChecked:
                l1=self.linepicker1.Entity
                if l1==None: 
                    self.success.Content += '\nno Line 1 selected'
                    inputok=False

                try:
                    startstation = self.station.Distance
                except:
                    self.success.Content += '\nStart Chainage error'
                    inputok=False

                if inputok:

                    outstation = clr.StrongBox[float]()
                    outoffset = clr.StrongBox[float]()
                    
                    anchor = self.station.FindAnchorPoint(self.station.ClickLocation, outstation, outoffset)
                    if anchor != None:
                        alignvector = Vector3D(anchor, self.station.ClickLocation)

                        activeForm.SetCameraUpVector(alignvector)
                        activeForm.CenterView(anchor)
                        activeForm.ToggleGrids(True)    # toggle grid twice to view it correctly
                        activeForm.ToggleGrids(True)

                        self.bearing.Direction = math.pi/2 - cam_az
                    # Keyboard.Focus(self.linepicker1)



            

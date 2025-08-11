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
    cmdData.Key = "SCR_PointFromVectorOffsetStation"
    cmdData.CommandName = "SCR_PointFromVectorOffsetStation"
    cmdData.Caption = "_SCR_PointFromVectorOffsetStation"
    cmdData.UIForm = "SCR_PointFromVectorOffsetStation"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "Points"
        cmdData.ShortCaption = "Vector Station/Offset"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.05
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "create Points along a 2-point-line"
        cmdData.ToolTipTextFormatted = "create Points along a 2-point-line"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_PointFromVectorOffsetStation(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_PointFromVectorOffsetStation.xaml") as s:
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
        wv = self.currentProject [Project.FixedSerial.WorldView]

        #self.station.NumberOfDecimals = 4
        #self.interval.NumberOfDecimals = 4
        #self.offset.NumberOfDecimals = 4
        #self.voffset.NumberOfDecimals = 4
 
        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        self.lfp = self.lunits.Properties.Copy()
        self.linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation

        self.labelstation.Content = 'Station (used on OK-Click) [' + self.linearsuffix + ']'
        self.enableinterval.Content = 'enable Interval [' + self.linearsuffix + ']'
        self.labelhzoffset.Content = 'horizontal Offset [' + self.linearsuffix + ']'
        self.labelvzoffset.Content = 'vertical Offset [' + self.linearsuffix + ']'

 
        self.coordpick1.ShowElevationIf3D = True
        self.coordpick2.ShowElevationIf3D = True

        self.pointnamestart.NumberOfDecimals = 0
        self.pointnameincrement.NumberOfDecimals = 0

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

        #self.coordpick1.SetCoordinate(Point3D(100,100,100), self.currentProject, )

    def SetDefaultOptions(self):
        lserial = OptionsManager.GetUint("SCR_PointFromVectorOffsetStation.layerpicker", 8)
        o = self.currentProject.Concordance.Lookup(lserial)
        if o != None:
            if isinstance(o.GetSite(), LayerCollection):    
                self.layerpicker.SetSelectedSerialNumber(lserial, InputMethod(3))
            else:                       
                self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:                       
            self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))

        self.stationis2d.IsChecked = OptionsManager.GetBool("SCR_PointFromVectorOffsetStation.stationis2d", True)
        self.stationis3d.IsChecked = OptionsManager.GetBool("SCR_PointFromVectorOffsetStation.stationis3d", False)
        self.station.Distance = OptionsManager.GetDouble("SCR_PointFromVectorOffsetStation.station", 0.000)
        self.enableinterval.IsChecked = OptionsManager.GetBool("SCR_PointFromVectorOffsetStation.enableinterval", False)
        self.interval.Distance = OptionsManager.GetDouble("SCR_PointFromVectorOffsetStation.interval", 0.000)

        self.offset.Distance = OptionsManager.GetDouble("SCR_PointFromVectorOffsetStation.offset", 0.000)
        self.dhisplumb.IsChecked = OptionsManager.GetBool("SCR_PointFromVectorOffsetStation.dhisplumb", True)
        self.dhisperp.IsChecked = OptionsManager.GetBool("SCR_PointFromVectorOffsetStation.dhisperp", False)
        self.voffset.Distance = OptionsManager.GetDouble("SCR_PointFromVectorOffsetStation.voffset", 0.000)

        self.createnamedpoint.IsChecked = OptionsManager.GetBool("SCR_PointFromVectorOffsetStation.createnamedpoint", False)
        self.pointnametext.Text = OptionsManager.GetString("SCR_PointFromVectorOffsetStation.pointnametext", "")
        self.pointnamestart.Value = OptionsManager.GetDouble("SCR_PointFromVectorOffsetStation.pointnamestart", 0)
        self.pointnameincrement.Value = OptionsManager.GetDouble("SCR_PointFromVectorOffsetStation.pointnameincrement", 0)

    def SaveOptions(self):
        OptionsManager.SetValue("SCR_PointFromVectorOffsetStation.layerpicker", self.layerpicker.SelectedSerialNumber)
        
        OptionsManager.SetValue("SCR_PointFromVectorOffsetStation.stationis2d", self.stationis2d.IsChecked)
        OptionsManager.SetValue("SCR_PointFromVectorOffsetStation.stationis3d", self.stationis3d.IsChecked)
        
        OptionsManager.SetValue("SCR_PointFromVectorOffsetStation.station", self.station.Distance)
        OptionsManager.SetValue("SCR_PointFromVectorOffsetStation.enableinterval", self.enableinterval.IsChecked)
        OptionsManager.SetValue("SCR_PointFromVectorOffsetStation.interval", self.interval.Distance)

        OptionsManager.SetValue("SCR_PointFromVectorOffsetStation.offset", self.offset.Distance)
        OptionsManager.SetValue("SCR_PointFromVectorOffsetStation.dhisplumb", self.dhisplumb.IsChecked)
        OptionsManager.SetValue("SCR_PointFromVectorOffsetStation.dhisperp", self.dhisperp.IsChecked)
        OptionsManager.SetValue("SCR_PointFromVectorOffsetStation.voffset", self.voffset.Distance)

        OptionsManager.SetValue("SCR_PointFromVectorOffsetStation.createnamedpoint", self.createnamedpoint.IsChecked)
        OptionsManager.SetValue("SCR_PointFromVectorOffsetStation.pointnametext", self.pointnametext.Text)
        OptionsManager.SetValue("SCR_PointFromVectorOffsetStation.pointnamestart", round(self.pointnamestart.Value, 0))
        OptionsManager.SetValue("SCR_PointFromVectorOffsetStation.pointnameincrement", round(self.pointnameincrement.Value, 0))

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def enableintervalChanged(self, sender, e):
        if self.enableinterval.IsChecked==True:
            self.interval.IsEnabled=True
        else:
            self.interval.IsEnabled=False

    def createnamedpointChanged(self, sender, e):
        if self.createnamedpoint.IsChecked:
            self.createnamedpointpanel.IsEnabled = True
        else:
            self.createnamedpointpanel.IsEnabled = False


    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''

        self.success.Content = ''

        wv = self.currentProject [Project.FixedSerial.WorldView]
        #wv.PauseGraphicsCache(True)

        inputok=True

        station = self.station.Distance
        offset = self.offset.Distance
        voffset = self.voffset.Distance
        self.pointnamestart.Value = round(self.pointnamestart.Value, 0)
        self.pointnameincrement.Value = round(self.pointnameincrement.Value, 0)
        #try:
        #    station = float(self.station.Text)
        #except:
        #    self.success.Content += '\nStation error'
        #    inputok=False
        #
        #try:
        #    offset = float(self.offset.Text)
        #except:
        #    self.success.Content += '\nOffset error'
        #    inputok=False


        if inputok:
            self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
            UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
            try:
                with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                    
                    p1 = self.coordpick1.Coordinate
                    p2 = self.coordpick2.Coordinate
                    
                    orgvector = Vector3D(p1, p2)
                    
                    # create 2D Vector to the side
                    offsetvector = orgvector.Clone()
                    #offsetvector.To2D()
                    offsetvector.Z = 0.0
                    offsetvector.Rotate90(Side.Right)
                    
                    # create perp-dh vector
                    perpdh = orgvector.Clone()
                    perpdh.Rotate(BiVector3D(offsetvector, math.pi / 2))
                    perpdh.Length = voffset

                    # set the vector length only now - in case of offset zero the perpdh would fail
                    offsetvector.Length = offset

                    # create 2 points parallel to the original ones
                    p1_2 = p1 + offsetvector
                    p2_2 = p2 + offsetvector
                    
                    # apply dh
                    if self.dhisplumb.IsChecked:
                        p1_2.Z = p1.Z + voffset
                        p2_2.Z = p2.Z + voffset
                    else:
                        p1_2 += perpdh
                        p2_2 += perpdh

                    # create a vector that is parallel to the original one, including the dh 
                    parallelvector = Vector3D(p1_2, p2_2)
                    # Vector3D.Horizon is positive above the horizon and negative below
                    if self.stationis2d.IsChecked:
                        # adjust the horizontal station to 3d length
                        parallelvector.Length = station / math.cos(orgvector.Horizon)
                    else:
                        parallelvector.Length = station

                    pnew = p1_2 + parallelvector
                    
                    if self.createnamedpoint.IsChecked:
                        pnew_wv = CoordPoint.CreatePoint(self.currentProject, self.pointnametext.Text + '{0:g}'.format(round(self.pointnamestart.Value, 0)))
                        pnew_wv.AddPosition(pnew)
                        pnew_wv.Layer = self.layerpicker.SelectedSerialNumber
                        self.pointnamestart.Value += self.pointnameincrement.Value
                    else:
                        cadPoint = wv.Add(clr.GetClrType(CadPoint))
                        cadPoint.Layer = self.layerpicker.SelectedSerialNumber
                        cadPoint.Point0 = pnew

                    if self.enableinterval.IsChecked:
                        self.station.Distance += self.interval.Distance

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

        #wv.PauseGraphicsCache(False)

    

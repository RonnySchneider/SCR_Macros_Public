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
    cmdData.Key = "SCR_P1_XY_P2_Z"
    cmdData.CommandName = "SCR_P1_XY_P2_Z"
    cmdData.Caption = "_SCR_P1_XY_P2_Z"
    cmdData.UIForm = "SCR_P1_XY_P2_Z"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "Points"
        cmdData.ShortCaption = "P = P1-XY + P2-Z"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.15
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Use XY from P1 and Z from P2"
        cmdData.ToolTipTextFormatted = "Use XY from P1 and Z from P2"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_P1_XY_P2_Z(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_P1_XY_P2_Z.xaml") as s:
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
        
        self.coordpick1.ShowElevationIf3D = True
        self.coordpick2.ShowElevationIf3D = True

        self.coordpick2.ValueChanged += self.coordpick2changed

        self.cadpointType = clr.GetClrType(CadPoint)
        self.coordpointType = clr.GetClrType(CoordPoint)

        self.objs.IsEntityValidCallback = self.IsValid
               
        optionMenu = SelectionContextMenuHandler()
        # remove options that don't apply here
        optionMenu.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        self.objs.ButtonContextMenu = optionMenu

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        #self.lfp = self.lunits.Properties.Copy()
        self.linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation

        self.labelsearchradius.Content = '3D search radius [' + self.linearsuffix + ']'

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def IsValid(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.coordpointType):
            return True
        return False

    def createnamedpointChanged(self, sender, e):
        if self.createnamedpoint.IsChecked:
            self.namedpointbox.IsEnabled = True
        else:
            self.namedpointbox.IsEnabled = False

    def SetDefaultOptions(self):

        settingserial = OptionsManager.GetUint("SCR_P1_XY_P2_Z.layerpicker", 8) # 8 is FixedSerial for Layer Zero
        o = self.currentProject.Concordance.Lookup(settingserial)
        if o != None:
            if isinstance(o.GetSite(), LayerCollection):    
                self.layerpicker.SetSelectedSerialNumber(settingserial, InputMethod(3))
            else:                       
                self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:                       
            self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        
        self.createnamedpoint.IsChecked = OptionsManager.GetBool("SCR_P1_XY_P2_Z.createnamedpoint", True)

        self.manualinput.IsChecked = OptionsManager.GetBool("SCR_P1_XY_P2_Z.manualinput", True)
        self.addtoname.Text = OptionsManager.GetString("SCR_P1_XY_P2_Z.addtoname", "_DERIVED")
        self.newcode.Text = OptionsManager.GetString("SCR_P1_XY_P2_Z.newcode", "")
    
        self.autoincrementslider.IsChecked = OptionsManager.GetBool("SCR_P1_XY_P2_Z.autoincrementslider", False)

        self.autoinput.IsChecked = OptionsManager.GetBool("SCR_P1_XY_P2_Z.autoinput", False)
        self.p1code.Text = OptionsManager.GetString("SCR_P1_XY_P2_Z.p1code", "")
        self.p2code.Text = OptionsManager.GetString("SCR_P1_XY_P2_Z.p2code", "")
        self.searchradius.Distance = OptionsManager.GetDouble("SCR_P1_XY_P2_Z.searchradius", 0.2)
    
   
    def SaveOptions(self):

        OptionsManager.SetValue("SCR_P1_XY_P2_Z.layerpicker", self.layerpicker.SelectedSerialNumber)

        OptionsManager.SetValue("SCR_P1_XY_P2_Z.createnamedpoint", self.createnamedpoint.IsChecked)

        OptionsManager.SetValue("SCR_P1_XY_P2_Z.manualinput", self.manualinput.IsChecked)
        OptionsManager.SetValue("SCR_P1_XY_P2_Z.addtoname", self.addtoname.Text)
        OptionsManager.SetValue("SCR_P1_XY_P2_Z.newcode", self.newcode.Text)
        
        OptionsManager.SetValue("SCR_P1_XY_P2_Z.autoincrementslider", self.autoincrementslider.IsChecked)

        OptionsManager.SetValue("SCR_P1_XY_P2_Z.autoinput", self.autoinput.IsChecked)
        OptionsManager.SetValue("SCR_P1_XY_P2_Z.p1code", self.p1code.Text)
        OptionsManager.SetValue("SCR_P1_XY_P2_Z.p2code", self.p2code.Text)
        OptionsManager.SetValue("SCR_P1_XY_P2_Z.searchradius", self.searchradius.Distance)
       
    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()
    
    def coordpick2changed(self, ctrl, e):

        self.success.Content = ""
        self.error.Content = ""

        wv = self.currentProject [Project.FixedSerial.WorldView]

        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())

        addtoname = self.addtoname.Text
        newcode = self.newcode.Text

        p1 = self.coordpick1.Coordinate
        p2 = self.coordpick2.Coordinate
        
        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                p1found=False
                # go through all elements and find the point clicked
                if self.createnamedpoint.IsChecked:
                    for o in wv:
                        if isinstance(o, PointCollection):
                            for p in o:
                                if Point3D.Distance(p1, p.Position)[0] == 0:
                                    p1name = p.Name
                                    p1code = p.FeatureCode
                                    p1found = True
                
                if p1found == False: self.success.Content += "\nno Number found for P1"
                if math.isnan(p2.Z): self.success.Content += "\nP2 has no elevation"

                if self.createnamedpoint.IsChecked and p1found and not math.isnan(p2.Z):
                    #find PointManager as object
                    for o in self.currentProject:
                        if isinstance(o, PointManager):
                            pm = o

                    pnew = Point3D()
                    pnew.X = p1.X
                    pnew.Y = p1.Y
                    pnew.Z = p2.Z

                    pnew_wv = CoordPoint.CreatePoint(self.currentProject, p1name + addtoname)
                    #change FeatureCode by using the PointManager
                    if newcode=="":
                        pm.SetFeatureCodeAtPoint(pnew_wv.SerialNumber, p1code)
                    else:    
                        pm.SetFeatureCodeAtPoint(pnew_wv.SerialNumber, newcode)
                    
                    pnew_wv.AddPosition(pnew)
                    pnew_wv.Layer = self.layerpicker.SelectedSerialNumber
                
                if (self.createnamedpoint.IsChecked == False or p1found == False ) and math.isnan(p2.Z)==False:
                    pnew = Point3D()
                    pnew.X = p1.X
                    pnew.Y = p1.Y
                    pnew.Z = p2.Z
                    
                    cadPoint = wv.Add(clr.GetClrType(CadPoint))
                    cadPoint.Point0 = pnew
                    cadPoint.Layer = self.layerpicker.SelectedSerialNumber

                failGuard.Commit()
                UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
                self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
        
        except:
            # EndMark MUST be set no matter what
            # otherwise TBC won't work anymore and needs to be restarted
            self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
            UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
            self.error.Content += '\nan Error occurred - Result probably incomplete'

        
        if self.autoincrementslider.IsChecked:
            activeForm = TrimbleOffice.TheOffice.MainWindow.AppViewManager.ActiveView
            if isinstance(activeForm, clr.GetClrType(LimitSliceWindow)):
                slider = activeForm.SliderData
                for st in slider.ListedStations:
                    if st > slider.CurrentStation:
                        slider.CurrentStation = st
                        break
            else:
                self.error.Content += '\nAuto-Slider is ticked'
                self.error.Content += '\nbut no Cutting Plane View is activated'

        Keyboard.Focus(self.coordpick1)
        self.SaveOptions() 

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''


        
        self.success.Content = ""

        wv = self.currentProject [Project.FixedSerial.WorldView]

        if self.autoinput.IsChecked:

            #find PointManager as object
            for o in self.currentProject:
                if isinstance(o, PointManager):
                    pm = o

            addtoname = self.addtoname.Text
            newcode = self.newcode.Text
            p1code = self.p1code.Text
            p2code = self.p2code.Text

            # check searchradius
            inputok = True
            try:
                searchradius = self.searchradius.Distance
            except:
                inputok = False
       
            if inputok:
                self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
                UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
                try:
                    with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                        for p1 in self.objs: # go through all the selected points
                            if isinstance(p1, self.coordpointType):    # check if it is a point with feature code

                                if p1code in p1.FeatureCode: # check if the search feature code matches

                                    i = 0
                                    closestp2 = None
                                    for p2 in self.objs: # go through all the selected points to find a P2
                                        if p1 != p2 and isinstance(p2, self.coordpointType):    # check if it is a point with feature code

                                            if p2code in p2.FeatureCode: # check if the search feature code matches

                                                if i == 0:
                                                    if Vector3D(p1.Position, p2.Position).Length <= searchradius:
                                                        closestp2 = p2
                                                        i += 1
                                                else:
                                                    if Vector3D(p1.Position, p2.Position).Length < Vector3D(p1.Position, closestp2.Position).Length:
                                                        closestp2 = p2
                                    
                                    # now that we have a matching P2 we can draw the new point
                                    if closestp2 != None and math.isnan(closestp2.Position.Z) == False:

                                        if self.createnamedpoint.IsChecked:

                                            pnew = Point3D()
                                            pnew.X = p1.Position.X
                                            pnew.Y = p1.Position.Y
                                            pnew.Z = closestp2.Position.Z

                                            pnew_wv = CoordPoint.CreatePoint(self.currentProject, p1.Name + addtoname)
                                            #change FeatureCode by using the PointManager
                                            if newcode == "":
                                                pm.SetFeatureCodeAtPoint(pnew_wv.SerialNumber, p1.FeatureCode)
                                            else:    
                                                pm.SetFeatureCodeAtPoint(pnew_wv.SerialNumber, newcode)
                                            
                                            pnew_wv.AddPosition(pnew)
                                            pnew_wv.Layer = self.layerpicker.SelectedSerialNumber
                                        
                                        if self.createnamedpoint.IsChecked == False:
                                            pnew = Point3D()
                                            pnew.X = p1.Position.X
                                            pnew.Y = p1.Position.Y
                                            pnew.Z = closestp2.Position.Z
                                            
                                            cadPoint = wv.Add(clr.GetClrType(CadPoint))
                                            cadPoint.Point0 = pnew
                                            cadPoint.Layer = self.layerpicker.SelectedSerialNumber

                                    closestp2 = None

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

        self.SaveOptions()           

        
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
    cmdData.Key = "SCR_ElevateSHPLines"
    cmdData.CommandName = "SCR_ElevateSHPLines"
    cmdData.Caption = "_SCR_ElevateSHPLines"
    cmdData.UIForm = "SCR_ElevateSHPLines"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Expld-SNR-Relay-Prop"
        cmdData.DefaultTabGroupKey = "Properties"
        cmdData.ShortCaption = "elevate SHP-Lines"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.04
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "elevate lines based on their feature attribute"
        cmdData.ToolTipTextFormatted = "elevate lines based on their feature attribute"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png") # we have to include a icon revision, otherwise TBC might not show the new one
        cmdData.ImageSmall = b
    except:
        pass


class SCR_ElevateSHPLines(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_ElevateSHPLines.xaml") as s:
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
        self.objs.IsEntityValidCallback = self.IsValid

        self.lType = clr.GetClrType(IPolyseg)
        #self.compType = clr.GetClrType(IPolyseg)

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear

        #self.lfp = self.lunits.Properties.Copy()
        self.linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation

        self.SetDefaultOptions()

    def SetDefaultOptions(self):
        self.featureattr.Text = OptionsManager.GetString("SCR_ElevateSHPLines.featureattr", "Elevation")

        lserial = OptionsManager.GetUint("SCR_ElevateSHPLines.layerpicker", 8)
        o = self.currentProject.Concordance.Lookup(lserial) # get the object withj the saved serial number
        if o != None:   # check if we actually got an object back
            if isinstance(o.GetSite(), LayerCollection):    # check if the object is a layer
                self.layerpicker.SetSelectedSerialNumber(lserial, InputMethod(3))
            else:  # if not then set Layer zero                       
                self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:   # if not then set Layer zero               
            self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))

    def SaveOptions(self):
        OptionsManager.SetValue("SCR_ElevateSHPLines.featureattr", self.featureattr.Text)
        OptionsManager.SetValue("SCR_ElevateSHPLines.layerpicker", self.layerpicker.SelectedSerialNumber)

    def IsValid(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        return False

    def OkClicked(self, cmd, e):

        Keyboard.Focus(self.okBtn)  # a trick to evaluate all input fields before execution, otherwise you'd have to click in another field first
        
        self.success.Content = ""
        self.error.Content = ""


        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())


        try:
            # the "with" statement will unroll any changes if something go wrong
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:


                for o in self.currentProject:
                #find FeatureManager as object
                    if isinstance(o, FeatureManager):
                        fm = o

                for o in self.objs:
                    
                    if isinstance(o, ICompositeGeometry):
                        foundelev = False

                        # need to jump throughs hoops since the featurecode is protected in composite geometry
                        # need to do it in reverse, go through all featurecodes and look if it's referring to the current object
                        for f in fm:
                            for e in self.currentProject.Concordance.GetObserversOf(f.SerialNumber):

                                if e == o:
                            
                                    fc = f
                                    break
                        if fc:
                            for attr in fc.Attributes:
                                if attr.Name == self.featureattr.Text:
                                    elev = self.lunits.Convert(self.lunits.DisplayType, float(attr.Value), self.lunits.InternalType)
                                    foundelev = True
                                    break

                            if foundelev:

                                for e in o:
                                    if isinstance(e.SnapIn, self.lType):
                                        self.drawline(e.SnapIn, elev)

                            else:
                                o.Color = Color.Red

                        
                    elif isinstance(o, self.lType):

                        # in case the user has selected a combination of linestrings and composites
                        # we'd color linestrings red if we wouldn't check if they are part of a composite
                        # only the composite contains the attribute
                        partofcomposite = False
                        for observedby in self.currentProject.Concordance.GetObserversOf(o.SerialNumber):
                            if isinstance(observedby, ICompositeGeometry):
                                partofcomposite = True
                                break

                        if not partofcomposite:
                            foundelev = False
                            # get the line feature code
                            for observes in self.currentProject.Concordance.GetIsObservedBy(o.SerialNumber):
                                if observes and isinstance(observes, LineFeature):
                                    for attr in observes.Attributes:
                                        if attr.Name == self.featureattr.Text:
                                            elev = self.lunits.Convert(self.lunits.DisplayType, float(attr.Value), self.lunits.InternalType)
                                            foundelev = True
                                            break

                            if foundelev:

                                self.drawline(o, elev)

                            else:
                                o.Color = Color.Red

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


                    #self.success.Content += '\n' + o.GetType().Name + ' - SN#: ' + str(o.SerialNumber)
                    #tt = o.GetSite()
                    #tt2 = o

        self.SaveOptions()

    def drawline(self, l, elev):

        wv = self.currentProject [Project.FixedSerial.WorldView]

        polyseg = l.ComputePolySeg()
        polyseg = polyseg.ToWorld()
        polyseg_v = PolySeg.PolySeg()
        polyseg_v.Add(Point3D(0, elev, 0))
        
        ls = wv.Add(clr.GetClrType(Linestring))
        ls.Append(polyseg, polyseg_v, False, False)
        ls.Layer = self.layerpicker.SelectedSerialNumber

        return


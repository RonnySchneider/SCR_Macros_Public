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
    cmdData.Key = "SCR_VolumeContours"
    cmdData.CommandName = "SCR_VolumeContours"
    cmdData.Caption = "_SCR_VolumeContours"
    cmdData.UIForm = "SCR_VolumeContours"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "DTM"
        cmdData.ShortCaption = "Volume Contour"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.07
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Volume Contour"
        cmdData.ToolTipTextFormatted = "creates Contours at Volume-Intervals"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png") 
        cmdData.ImageSmall = b
    except:
        pass


class SCR_VolumeContours(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_VolumeContours.xaml") as s:
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

        types = Array [Type] (SurfaceTypeLists.AllWithCutFillMap)    # we fill an array with TBC object types, we could combine different types
        self.surfacepicker.FilterByEntityTypes = types    # we fill the dropdownlist by applying that types array as filter
        self.surfacepicker.AllowNone = False              # our list shall not show an empty field

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        self.lfp = self.lunits.Properties.Copy()
        linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation
        #self.lfp.AddSuffix = False
        self.sliceintervallabel.Content = "Slice Interval [" + linearsuffix + "]" + \
                                          "\n(minimum 0.0001 m / 0.000329 ft\notherwise the Earthworks Report returns only Totals)"

        # get the units for volumes
        self.vunits = self.currentProject.Units.Volume
        self.vfp = self.vunits.Properties.Copy()
        #self.vfp.AddSuffix = False
        self.targetvolumelabel.Content = "Target Volume [" + self.vunits.Units[self.vunits.DisplayType].Abbreviation + "]"
        self.volumeintervallabel.Content = "Volume Interval [" + self.vunits.Units[self.vunits.DisplayType].Abbreviation + "]"
        
        # slice interval minimum is 0.0001 metres, below that the volume computation only returns totals
        self.sliceinterval.MinValue = self.lunits.Convert(self.lunits.InternalType, 0.0001, self.lunits.DisplayType)
        self.sliceinterval.NumberOfDecimals = self.lunits.Units[self.lunits.DisplayType].NumberOfDecimals
        self.targetvolume.MinValue = 0
        self.targetvolume.NumberOfDecimals = self.vunits.Units[self.vunits.DisplayType].NumberOfDecimals
        self.volumeinterval.MinValue = 0.0001 # otherwise the volume computation fails
        self.volumeinterval.NumberOfDecimals = self.vunits.Units[self.vunits.DisplayType].NumberOfDecimals

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def SetDefaultOptions(self):
        try:    self.surfacepicker.SelectIndex(OptionsManager.GetInt("SCR_VolumeContours.surfacepicker", 0))
        except: self.surfacepicker.SelectIndex(0)
        self.sliceinterval.Value = OptionsManager.GetDouble("SCR_VolumeContours.sliceinterval", 0.0001)
        self.startaddingfromsurfacemin.IsChecked = OptionsManager.GetBool("SCR_VolumeContours.startaddingfromsurfacemin", True)
        self.startaddingfromelevation.IsChecked = OptionsManager.GetBool("SCR_VolumeContours.startaddingfromelevation", False)
        self.elstart.Elevation = OptionsManager.GetDouble("SCR_VolumeContours.elstart", 1.00)
        self.singlesolution.IsChecked = OptionsManager.GetBool("SCR_VolumeContours.singlesolution", True)
        self.targetvolume.Value = OptionsManager.GetDouble("SCR_VolumeContours.targetvolume", 1.00)
        self.multisolution.IsChecked = OptionsManager.GetBool("SCR_VolumeContours.multisolution", False)
        self.volumeinterval.Value = OptionsManager.GetDouble("SCR_VolumeContours.volumeinterval", 1.00)

        self.createcontour.IsChecked = OptionsManager.GetBool("SCR_VolumeContours.createcontour", True)

    def SaveOptions(self):
        try:    # if nothing is selected it would throw an error
            OptionsManager.SetValue("SCR_VolumeContours.surfacepicker", self.surfacepicker.SelectedIndex)
        except:
            pass
        OptionsManager.SetValue("SCR_VolumeContours.sliceinterval", self.sliceinterval.Value)
        OptionsManager.SetValue("SCR_VolumeContours.startaddingfromsurfacemin", self.startaddingfromsurfacemin.IsChecked)
        OptionsManager.SetValue("SCR_VolumeContours.startaddingfromelevation", self.startaddingfromelevation.IsChecked)
        OptionsManager.SetValue("SCR_VolumeContours.elstart", self.elstart.Elevation)
        OptionsManager.SetValue("SCR_VolumeContours.singlesolution", self.singlesolution.IsChecked)
        OptionsManager.SetValue("SCR_VolumeContours.targetvolume", self.targetvolume.Value)
        OptionsManager.SetValue("SCR_VolumeContours.multisolution", self.multisolution.IsChecked)
        OptionsManager.SetValue("SCR_VolumeContours.volumeinterval", self.volumeinterval.Value)

        OptionsManager.SetValue("SCR_VolumeContours.createcontour", self.createcontour.IsChecked)

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''


        self.success.Content = ""

        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())

        wv = self.currentProject [Project.FixedSerial.WorldView]

        try:
            # the "with" statement will unroll any changes if something go wrong
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                # TBC internally computes in meter
                sliceinterval_meter = self.lunits.Convert(self.lunits.DisplayType, self.sliceinterval.Value, self.lunits.InternalType)
                targetvolume_meter = self.vunits.Convert(self.vunits.DisplayType, self.targetvolume.Value, self.vunits.InternalType)
                volinterval = self.vunits.Convert(self.vunits.DisplayType, self.volumeinterval.Value, self.vunits.InternalType)
                # get the surface as object
                surface = wv.Lookup(self.surfacepicker.SelectedSerial)
                
                # set the computation settings
                comp_settings = DtmVolumes.Settings()
                comp_settings.method = DtmVolumes.Method.DatumToDtm
                comp_settings.summary = DtmVolumes.Summary.ByElevation
                comp_settings.zDatum = surface.MaxElevation
                comp_settings.zIncrement = sliceinterval_meter
                if self.startaddingfromsurfacemin.IsChecked:
                    comp_settings.zIndex = surface.MinElevation
                else:
                    comp_settings.zIndex = self.elstart.Elevation # it's an elevation picker, it takes care of the unit conversion itself

                
                # prepare the output variables
                out_comp_results = clr.StrongBox[DtmVolumes.Results]()
                out_isopach = clr.StrongBox[Gem]()
    
                no_gem = clr.StrongBox[Gem](None)
                ProgressBar.TBC_ProgressBar.Title = "computing Earthworks report with slices"
                # create the isopach, since we compute DatumToDtm we have to pass it the no_gem Strongbox, it won't accept None
                if DtmVolumes.ComputeVolumeIsopach(comp_settings, no_gem, surface.Gem, out_isopach):
                    # compute the volumeslices - here it must be the Strongbox.Value, it won't accept None
                    if DtmVolumes.ComputeVolumeResults(comp_settings, no_gem.Value, out_isopach.Value, out_comp_results):
    
                        volumesCut = out_comp_results.Value.volumesCut
                        volumesFill = out_comp_results.Value.volumesFill
                        levels = out_comp_results.Value.zLevels
                        
                        # sum up the slices until we hit our target
                        chkinterval = volinterval
                        volumesum_meter = 0
                        solutions_meter = [] # list of elevations where we print a result and maybe draw a contour
                        # in case we start at an arbitrary elevation we also draw a contour there
                        if self.startaddingfromelevation.IsChecked and (self.elstart.Elevation > surface.MinElevation):
                            solutions_meter.Add([self.elstart.Elevation, 0])
                        
                        ProgressBar.TBC_ProgressBar.Title = "adding Slices and draw Contours"
                        # go through all the slices
                        for i in range(1, volumesCut.Count):
                            if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(i * 100 / volumesCut.Count)):
                                break   # function returns true if user pressed cancel
                            # either we add all
                            if self.startaddingfromsurfacemin.IsChecked or \
                               (self.startaddingfromelevation.IsChecked and (levels[i] >= self.elstart.Elevation)):
                                # or we start adding once we are past the start elevation
                                
                                volumesum_meter += volumesCut[i] # add to volume sum
                                
                                # check if single solution was requested and if we are past the target volumes
                                if self.singlesolution.IsChecked and volumesum_meter >= targetvolume_meter:
                                    solutions_meter.Add([levels[i], volumesum_meter])
                                    break
                                # in case of multi request check if we are past the volumeinterval
                                if self.multisolution.IsChecked and volumesum_meter >= chkinterval:
                                    solutions_meter.Add([levels[i], volumesum_meter])
                                    chkinterval += volinterval

    
                        # check if we found a solution
                        if solutions_meter.Count > 0:
                            for solmeter in solutions_meter:
                                # create the result strings - back in project dimensions
                                solutionelevation = self.lunits.Format(solmeter[0], self.lfp)
                                solutionvolume = self.vunits.Format(solmeter[1], self.vfp)
                                solstring = ""
                                if self.startaddingfromsurfacemin.IsChecked:
                                    solstring += "Base " +  str(self.lunits.Format(surface.MinElevation, self.lfp))
                                else:
                                    solstring += "Base " + str(self.lunits.Format(self.elstart.Elevation, self.lfp))
                                
                                solstring += " -> RL " + solutionelevation + " -> " + solutionvolume

                                # write the result to the macro window
                                self.success.Content += "\n" + solstring

                                # create a contour if asked for
                                if self.createcontour.IsChecked:
                                    cb = Model3DContoursBuilder()
                                    cb.SetSurface(surface)
                                    cb.SetName(solstring)
                                    cb.SetContourType(Model3DQuickContours.SurfaceContourType.eSingleElevation)
                                    cb.SetSingleElevation(solmeter[0])
                                    cb.SetSingleContourColor(Color.Red)
                                    cb.Commit()

                        else:
                            self.success.Content += "\nNo Solution found"
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
        self.SaveOptions()  
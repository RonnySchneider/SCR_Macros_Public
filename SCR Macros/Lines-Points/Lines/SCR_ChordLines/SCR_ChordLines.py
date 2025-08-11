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
    cmdData.Key = "SCR_ChordLines"
    cmdData.CommandName = "SCR_ChordLines"
    cmdData.Caption = "_SCR_ChordLines"
    cmdData.UIForm = "SCR_ChordLines"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "Lines"
        cmdData.ShortCaption = "Chord Lines"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.08
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "convert curved Lines to chorded Polylines"
        cmdData.ToolTipTextFormatted = "convert curved Lines to chorded Polylines"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_ChordLines(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_ChordLines.xaml") as s:
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

        #self.hortol.NumberOfDecimals = 4
        #self.vertol.NumberOfDecimals = 4
        #self.nodespacing.NumberOfDecimals = 4

        self.objs.IsEntityValidCallback=self.IsValid
        optionMenu = SelectionContextMenuHandler()
        # remove options that don't apply here
        optionMenu.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        self.objs.ButtonContextMenu = optionMenu
        
        self.lType = clr.GetClrType(IPolyseg)

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        #self.lfp = self.lunits.Properties.Copy()
        self.linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation

        self.toleranceheader.Header = 'define Computation Tolerance [' + self.linearsuffix + ']'

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def SetDefaultOptions(self):

        self.hortol.Distance = OptionsManager.GetDouble("SCR_ChordLines.hortol", 0.0001)
        self.vertol.Distance = OptionsManager.GetDouble("SCR_ChordLines.vertol", 0.0001)
        self.nodespacing.Distance = OptionsManager.GetDouble("SCR_ChordLines.nodespacing", 2)

        self.sourcelayer.IsChecked = OptionsManager.GetBool("SCR_ChordLines.sourcelayer", True)

        self.setlayer.IsChecked = OptionsManager.GetBool("SCR_ChordLines.setlayer", False)

        lserial = OptionsManager.GetUint("SCR_ChordLines.layerpicker", 8)
        o = self.currentProject.Concordance.Lookup(lserial) # get the object withj the saved serial number
        if o != None:   # check if we actually got an object back
            if isinstance(o.GetSite(), LayerCollection):    # check if the object is a layer
                self.layerpicker.SetSelectedSerialNumber(lserial, InputMethod(3))
            else:  # if not then set Layer zero                       
                self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:   # if not then set Layer zero               
            self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))

        self.setprefixsuffix.IsChecked = OptionsManager.GetBool("SCR_ChordLines.setprefixsuffix", False)
        self.prefix.Text = OptionsManager.GetString("SCR_ChordLines.prefix", "")
        self.suffix.Text = OptionsManager.GetString("SCR_ChordLines.suffix", "")

        self.deletesource.IsChecked = OptionsManager.GetBool("SCR_ChordLines.deletesource", False)

    def SaveOptions(self):

        OptionsManager.SetValue("SCR_ChordLines.hortol", self.hortol.Distance)
        OptionsManager.SetValue("SCR_ChordLines.vertol", self.vertol.Distance)
        OptionsManager.SetValue("SCR_ChordLines.nodespacing", self.nodespacing.Distance)

        OptionsManager.SetValue("SCR_ChordLines.sourcelayer", self.sourcelayer.IsChecked)    

        OptionsManager.SetValue("SCR_ChordLines.setlayer", self.setlayer.IsChecked)    
        OptionsManager.SetValue("SCR_ChordLines.layerpicker", self.layerpicker.SelectedSerialNumber)

        OptionsManager.SetValue("SCR_ChordLines.setprefixsuffix", self.setprefixsuffix.IsChecked)    
        OptionsManager.SetValue("SCR_ChordLines.prefix", self.prefix.Text)
        OptionsManager.SetValue("SCR_ChordLines.suffix", self.suffix.Text)

        OptionsManager.SetValue("SCR_ChordLines.deletesource", self.deletesource.IsChecked)    

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


        self.success.Content += ''

        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        
        wv = self.currentProject [Project.FixedSerial.WorldView]
        lgc = LayerGroupCollection.GetLayerGroupCollection(self.currentProject, False)
                
        wv.PauseGraphicsCache(True)

        self.success.Content=''
        # self.label_benchmark.Content = ''

        # settings = Model3DCompSettings.ProvideSettingsObject(self.currentProject)
        ProgressBar.TBC_ProgressBar.Title = "chording Lines"

        try:

            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                j = 0
                objs_count = self.objs.Count # if delete source is ticked the self.objs.Count would go down to Zero and produce a divison by zero
                for o in self.objs:
                    if isinstance(o, self.lType):
                        polyseg1 = o.ComputePolySeg()
                        polyseg1 = polyseg1.ToWorld()
                        polyseg1_v = o.ComputeVerticalPolySeg()
                        
                        t1 = abs(self.hortol.Distance)
                        t2 = abs(self.vertol.Distance)
                        t3 = abs(self.nodespacing.Distance)

                        polyseg_new = polyseg1.Linearize(abs(self.hortol.Distance), abs(self.vertol.Distance),
                            abs(self.nodespacing.Distance), polyseg1_v, False)
                        
                        if polyseg_new != None:       # if that worked
                            l = wv.Add(clr.GetClrType(Linestring))      # we start a new string line
                            oname = IName.Name.__get__(o)
                            if oname == '':
                                l.Name = l.Name + "chorded"
                            else:
                                l.Name = oname + " - chorded"

                            SnapInAttributeExtension.CopyUserAttributes(o, l)
                            if self.setlayer.IsChecked:
                                l.Layer = self.layerpicker.SelectedSerialNumber
                                try:
                                    l.Color = o.Color
                                    l.Weight = o.Weight
                                except: pass
                            
                            if self.sourcelayer.IsChecked:
                                l.Layer = o.Layer
                                try:
                                    l.Color = o.Color
                                    l.Weight = o.Weight
                                except: pass

                            if self.setprefixsuffix.IsChecked:
                                inputlayer = self.currentProject.Concordance.Lookup(o.Layer)
                                inputlayergroup = self.currentProject.Concordance.Lookup(inputlayer.LayerGroupSerial)
                                outputlayer = Layer.FindOrCreateLayer(self.currentProject, self.prefix.Text + inputlayer.Name + self.suffix.Text)
                                if inputlayergroup: # if the source layer is in a layer group
                                    # we check if the group exists, otherwise it is created
                                    outputlayergroup = lgc.FindOrCreateLayerGroup(self.prefix.Text + inputlayergroup.Name + self.suffix.Text)
                                    # we set the outputlayer group the the one we might just have created
                                    outputlayer.LayerGroupSerial = outputlayergroup.SerialNumber
                                # setting the values for the layer itself
                                outputlayer.DefaultColor = inputlayer.DefaultColor
                                outputlayer.LineStyle = inputlayer.LineStyle
                                outputlayer.LineWeight = inputlayer.LineWeight
                                # in case the line settings are not ByLayer we have to set them as well
                                l.Layer = outputlayer.SerialNumber
                                try:
                                    l.Color = o.Color
                                    l.LineStyle = o.LineStyle
                                    l.Weight = o.Weight
                                except: pass

                            
                            for i in range(0, polyseg_new.NumberOfNodes):       # and add all nodes of the profile as new nodes of that string line
                                e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                e.Position = polyseg_new[i].Point  # we draw that string line segment
                                l.AppendElement(e)
                                l.Color = o.Color

                        if self.deletesource.IsChecked:
                            osite = o.GetSite()    # we find out in which container the serial number reside
                            osite.Remove(o.SerialNumber)   # we delete the object from that container

                        j += 1
                        if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(j * 100 / objs_count)):
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
        
        self.success.Content += '\nDone'
        ProgressBar.TBC_ProgressBar.Title = ""
        
        wv.PauseGraphicsCache(False)

        self.SaveOptions()

    


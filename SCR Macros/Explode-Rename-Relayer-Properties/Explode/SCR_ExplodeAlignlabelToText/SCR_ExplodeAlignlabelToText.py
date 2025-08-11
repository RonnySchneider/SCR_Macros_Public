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
    cmdData.Key = "SCR_ExplodeAlignlabelToText"
    cmdData.CommandName = "SCR_ExplodeAlignlabelToText"
    cmdData.Caption = "_SCR_ExplodeAlignlabelToText"
    cmdData.UIForm = "SCR_ExplodeAlignlabelToText"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Expld-SNR-Relay-Prop"
        cmdData.DefaultTabGroupKey = "Explode"
        cmdData.ShortCaption = "Alignlabels to Text"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.02
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Explode Alignment Labels"
        cmdData.ToolTipTextFormatted = "detaches Labels from the Alignment and stores them as text on a layer"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_ExplodeAlignlabelToText(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_ExplodeAlignlabelToText.xaml") as s:
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
        self.objs.IsEntityValidCallback=self.IsValid
        optionMenu = SelectionContextMenuHandler()
        # remove options that don't apply here
        optionMenu.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        self.objs.ButtonContextMenu = optionMenu
        #self.poly3dType = clr.GetClrType(Poly3D)
        #self.polylineType = clr.GetClrType(PolyLine)
        #self.linestringType = clr.GetClrType(Linestring)
        #self.circleType = clr.GetClrType(CadCircle)
        self.alignmentlabelType = clr.GetClrType(AlignmentLabel)

    def IsValid(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        #if isinstance(o, self.poly3dType):
        #    return True
        if isinstance(o, self.alignmentlabelType):
            return True
        return False

    def CheckBoxChanged(self, sender, e):
        if self.checkBox.IsChecked:
            self.selectionpicker.Visibility=Visibility.Hidden
            self.objs.Visibility=Visibility.Hidden
        else:
            self.selectionpicker.Visibility=Visibility.Visible
            self.objs.Visibility=Visibility.Visible
        

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()


    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''



        self.success.Content=''


        wv = self.currentProject [Project.FixedSerial.WorldView]
        wl = self.currentProject[Project.FixedSerial.LayerContainer]
        
        layer_sn = self.layerpicker.SelectedSerialNumber



        for o in wv: # we go through the elements of the worldview
            convert=False
            if self.checkBox.IsChecked: convert=True # always convert if the checkbox is checked
            if self.checkBox.IsChecked==False:
                if self.objs.SelectedMembers(self.currentProject).Count>0:
                    for o2 in self.objs.SelectedMembers(self.currentProject): # only convert if it was selected
                        if o.SerialNumber==o2.SerialNumber: 
                            convert=True
                else:
                    self.success.Content='nothing selected'
                    break

            if convert==True:

                if isinstance(o,AlignmentLabel): # if they have the right type then continue
                    
                    labelseriallist=o.ContainedSerials # gets us a list with all the text serial numbers 
                    for i in range(0,o.Count):         # we count through the serial numbers
                        labelserial=labelseriallist[i]  # we get us one serial number
                        #tt=o2[labelserial]
                        if isinstance(o[labelserial],CadText): #the label container can also contain lines, so we have to make sure we work on a text
                            oldtext=o[labelserial]
                            newtext = wv.Add(clr.GetClrType(CadText))
                            newtext.AddWhiteOut = oldtext.AddWhiteOut
                            newtext.Alignment = oldtext.Alignment
                            newtext.AlignmentPoint = oldtext.AlignmentPoint
                            newtext.AutoFlip = oldtext.AutoFlip
                            # is read only - newtext.BackWard = oldtext.BackWard
                            # is read only - newtext.CanTransform = oldtext.CanTransform
                            newtext.Color = oldtext.Color
                            # is read only - newtext.Description = oldtext.Description
                            # is read only - newtext.Dirty = oldtext.Dirty
                            newtext.FixedOrientation = oldtext.FixedOrientation
                            newtext.Height = oldtext.Height
                            # is read only - newtext.Italic = oldtext.Italic
                            newtext.Layer = layer_sn        # we change the layer
                            # is read only - newtext.MaxElevation = oldtext.MaxElevation
                            # is read only - newtext.MinElevation = oldtext.MinElevation
                            newtext.Normal = oldtext.Normal
                            newtext.ObliqueAngle = oldtext.ObliqueAngle
                            # is read only - newtext.OriginOfUcs = oldtext.OriginOfUcs
                            newtext.RotateAngle = oldtext.RotateAngle
                            newtext.StyleName = oldtext.StyleName
                            newtext.TextString = oldtext.TextString
                            newtext.TextStyleSerial = oldtext.TextStyleSerial
                            newtext.Thickness = oldtext.Thickness
                            newtext.Visible = oldtext.Visible
                            newtext.Weight = oldtext.Weight
                            newtext.WidthFactor = oldtext.WidthFactor

    
            self.success.Content='Success'

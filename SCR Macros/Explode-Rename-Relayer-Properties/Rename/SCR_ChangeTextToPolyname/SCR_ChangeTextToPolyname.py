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
    cmdData.Key = "SCR_ChangeTextToPolyname"
    cmdData.CommandName = "SCR_ChangeTextToPolyname"
    cmdData.Caption = "_SCR_ChangeTextToPolyname"
    cmdData.UIForm = "SCR_ChangeTextToPolyname"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Expld-SNR-Relay-Prop"
        cmdData.DefaultTabGroupKey = "Renaming"
        cmdData.ShortCaption = "Text from Polyname"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.02
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Texts get Polylinename"
        cmdData.ToolTipTextFormatted = "all Texts inside a Polyline get changed to the name of the Polyline"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_ChangeTextToPolyname(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_ChangeTextToPolyname.xaml") as s:
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
        self.polylineType = clr.GetClrType(PolyLine)
        self.linestringType = clr.GetClrType(Linestring)
        self.circleType = clr.GetClrType(CadCircle)


    def IsValid(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        #if isinstance(o, self.poly3dType):
        #    return True
        if isinstance(o, self.polylineType):
            return True
        if isinstance(o, self.linestringType):
            return True
        if isinstance(o, self.circleType):
            return True
        return False
        

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()


    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''



        self.success.Content=''

        wv = self.currentProject [Project.FixedSerial.WorldView]
        wl = self.currentProject[Project.FixedSerial.LayerContainer]
            

        for o in self.objs.SelectedMembers(self.currentProject): # we go through all the selected Lines

            if isinstance(o,Linestring) or isinstance(o,PolyLine) or isinstance(o,CadCircle): # if they have the right type then continue
                
                if self.checkBox_uselayername.IsChecked==True: # in case the tickbox is ticked we get the layername
                    for o3 in wl:
                        if o3.SerialNumber==o.Layer:    #probably not the best way to get the layer name
                            linename=o3.Name
                else:                               # otherwise we use the line name
                    if isinstance(o,Linestring):    # depending on the Line Type we have to get the name differently
                        linename=o.AnchorName
                    else: 
                        linename=o.Name
                
                polyseg=o.ComputePolySeg()      # we convert the line object to a more unified polyseg, for which we have nifty functions available
                
                if polyseg.IsClosed:            # if the line is closed then continue
                    
                    for o2 in wv:               # we go through all drawing objects
                        
                        if isinstance(o2,CadText) or isinstance(o2,MText):      # in case it is some kind of text we continue
                            if polyseg.PointInPolyseg(o2.AlignmentPoint)==Side.In:      # we compute if the text anchor point is inside the polyseg
                                o2.TextString = self.textBox1.Text + linename + self.textBox2.Text      # if it is, we change the text
                                if self.checkBox_relayer.IsChecked==True: # if the checkbox is checked we also change the layer to that of the line
                                    o2.Layer = o.Layer
                        
                        if isinstance(o2,AlignmentLabel):      # in case of aligment label text we have to unwrap it first
                            labelseriallist=o2.ContainedSerials # gets us a list with all the text serial numbers 
                            for i in range(0,o2.Count):         # we count through the serial numbers
                                labelserial=labelseriallist[i]  # we get us one serial number
                                #tt=o2[labelserial]
                                if isinstance(o2[labelserial],CadText): #the label container can also contain lines, so we have to make sure we work on a text
                                    if polyseg.PointInPolyseg(o2[labelserial].AlignmentPoint)==Side.In:     # we compute if the text anchor point is inside the polyseg
                                        o2[labelserial].TextString = self.textBox1.Text + linename + self.textBox2.Text     # we change the text
                                        if self.checkBox_relayer.IsChecked==True: # if the checkbox is checked we also change the layer to that of the line
                                            o2[labelserial].Layer = o.Layer

        if self.objs.SelectedMembers(self.currentProject).Count==0:
            self.success.Content+='nothing selected'
        else:
            self.success.Content='Success'


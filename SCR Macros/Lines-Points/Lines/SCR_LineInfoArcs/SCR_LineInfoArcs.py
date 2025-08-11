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
    cmdData.Key = "SCR_LineInfoArcs"
    cmdData.CommandName = "SCR_LineInfoArcs"
    cmdData.Caption = "_SCR_LineInfoArcs"
    cmdData.UIForm = "SCR_LineInfoArcs"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "Lines"
        cmdData.ShortCaption = "Line-Info Arcs"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.11
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "creates points at the Arc-Centers and PI locations"
        cmdData.ToolTipTextFormatted = "creates points at the Arc-Centers and PI locations"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_LineInfoArcs(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_LineInfoArcs.xaml") as s:
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

        self.objs.IsEntityValidCallback = self.IsValid
        optionMenu = SelectionContextMenuHandler()
        # remove options that don't apply here
        optionMenu.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        self.objs.ButtonContextMenu = optionMenu

        self.lType = clr.GetClrType(IPolyseg)

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        self.lfp = self.lunits.Properties.Copy()
        # self.linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation

        # get the units for chainage distance
        self.chunits = self.currentProject.Units.Station
        self.chfp = self.chunits.Properties.Copy()

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass


    def SetDefaultOptions(self):
        self.selectlayer.IsChecked = OptionsManager.GetBool("SCR_LineInfoArcs.selectlayer", False)

        lserial = OptionsManager.GetUint("SCR_LineInfoArcs.layerpicker", 8)
        o = self.currentProject.Concordance.Lookup(lserial) # get the object withj the saved serial number
        if o != None:   # check if we actually got an object back
            if isinstance(o.GetSite(), LayerCollection):    # check if the object is a layer
                self.layerpicker.SetSelectedSerialNumber(lserial, InputMethod(3))
            else:  # if not then set Layer zero                       
                self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:   # if not then set Layer zero               
            self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))

        self.manualname.IsChecked = OptionsManager.GetBool("SCR_LineInfoArcs.manualname", False)
        self.nameinput.SelectedName = OptionsManager.GetString("SCR_LineInfoArcs.nameinput", "")
        self.createarcp.IsChecked = OptionsManager.GetBool("SCR_LineInfoArcs.createarcp", True)
        self.createpis.IsChecked = OptionsManager.GetBool("SCR_LineInfoArcs.createpis", True)
        self.drawarclines.IsChecked = OptionsManager.GetBool("SCR_LineInfoArcs.drawarclines", True)
        self.drawpilines.IsChecked = OptionsManager.GetBool("SCR_LineInfoArcs.drawpilines", True)

    def SaveOptions(self):
        OptionsManager.SetValue("SCR_LineInfoArcs.selectlayer", self.selectlayer.IsChecked)
        OptionsManager.SetValue("SCR_LineInfoArcs.layerpicker", self.layerpicker.SelectedSerialNumber)
        OptionsManager.SetValue("SCR_LineInfoArcs.manualname", self.manualname.IsChecked)
        OptionsManager.SetValue("SCR_LineInfoArcs.nameinput", self.nameinput.SelectedName)
        OptionsManager.SetValue("SCR_LineInfoArcs.createarcp", self.createarcp.IsChecked)
        OptionsManager.SetValue("SCR_LineInfoArcs.createpis", self.createpis.IsChecked)
        OptionsManager.SetValue("SCR_LineInfoArcs.drawarclines", self.drawarclines.IsChecked)
        OptionsManager.SetValue("SCR_LineInfoArcs.drawpilines", self.drawpilines.IsChecked)
        
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

        self.success.Content=''

        try:

            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                
                pc = PointCollection.ProvidePointCollection(self.currentProject)

                for o in self.objs:
                    if isinstance(o, self.lType):
                        if self.selectlayer.IsChecked:
                            self.layerserial = self.layerpicker.SelectedSerialNumber
                        else:
                            self.layerserial = o.Layer

                        # get the line data as polyseg, in world coordinates
                        polyseg1 = o.ComputePolySeg()
                        polyseg1 = polyseg1.ToWorld()

                        for n in polyseg1.Nodes():    # go through all the nodes in the linestring

                            # if ArcNode
                            if n and n.Visible and n.Type == PolySeg.Node.Type(3):
                                if self.createarcp.IsChecked or self.createpis.IsChecked: # only increment if we need to create points
                                    if self.manualname.IsChecked:
                                        pointname = self.nameinput.SelectedName
                                        # increment the name
                                        self.nameinput.SelectedName = PointHelper.Helpers.Increment(self.nameinput.SelectedName, None, True)
                                    else:
                                        # compile the pointname from the line name and start/end chainage
                                        pointname = IName.Name.__get__(o) + " - CH: " + \
                                                    str(self.chunits.Format(n.Station, self.chfp)) + " - " + \
                                                    str(self.chunits.Format(n.Station + n.Length, self.chfp))
                                
                                if self.createarcp.IsChecked:
                                    # create the point
                                    pnew_wv = CoordPoint.CreatePoint(self.currentProject, pointname + " - Arc-Center")

                                    # set the correct layer
                                    pnew_wv.Layer = self.layerserial

                                    # the node gives us all the values, no need to compute anything
                                    pnew_wv.AddPosition(n.RadiusPoint)
                                    pnew_wv.Description1 = "Radius - " + str(self.lunits.Format(n.Radius, self.lfp))
                                    pnew_wv.Description2 = "Length - " + str(self.lunits.Format(n.Length, self.lfp))

                                if self.createpis.IsChecked:
                                    # do the same for the PI
                                    # create the point
                                    pnew_wv = CoordPoint.CreatePoint(self.currentProject, pointname + " - PI")
                                    # set the correct layer
                                    pnew_wv.Layer = self.layerserial
                                    ## the node gives us all the values, no need to compute anything
                                    #pnew_wv.AddPosition(n.RadiusPoint)
                                    #pnew_wv.Description1 = "Radius - " + str(self.lunits.Format(n.Radius, self.lfp))
                                    #pnew_wv.Description2 = "Length - " + str(self.lunits.Format(n.Length, self.lfp))

                                    # the node gives us all the values, no need to compute anything
                                    pnew_wv.AddPosition(n.PointOfIntersection)

                                if self.drawarclines.IsChecked:
                                    # draw the lines
                                    # center to radius start/end
                                    self.drawline(n.RadiusPoint, n.Point)
                                    self.drawline(n.RadiusPoint, n.EndPoint)
                                if self.drawpilines.IsChecked:
                                    # start/end to PI
                                    self.drawline(n.PointOfIntersection, n.EndPoint)
                                    self.drawline(n.Point, n.PointOfIntersection)

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

        Keyboard.Focus(self.objs)
        self.SaveOptions()

    def drawline(self, p1, p2):
        
        wv = self.currentProject [Project.FixedSerial.WorldView]
        l = wv.Add(clr.GetClrType(Linestring))
        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
        e.Position = p1 
        l.AppendElement(e)
        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
        e.Position = p2
        l.AppendElement(e)
        l.Layer = self.layerserial
        #l.Color = self.outcolorpicker.SelectedColor

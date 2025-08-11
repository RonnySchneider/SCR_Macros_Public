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
    cmdData.Key = "SCR_LineInfoVPI"
    cmdData.CommandName = "SCR_LineInfoVPI"
    cmdData.Caption = "_SCR_LineInfoVPI"
    cmdData.UIForm = "SCR_LineInfoVPI"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "Lines"
        cmdData.ShortCaption = "Line-Info VPI's"
        cmdData.DefaultRibbonToolSize = 0 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.13
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Create Points at Line VPI's"
        cmdData.ToolTipTextFormatted = "Create Points at Line VPI's"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_LineInfoVPI(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_LineInfoVPI.xaml") as s:
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
        self.lType = clr.GetClrType(IPolyseg)
        self.objs.IsEntityValidCallback=self.IsValid
        optionMenu = SelectionContextMenuHandler()
        # remove options that don't apply here
        optionMenu.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        self.objs.ButtonContextMenu = optionMenu
		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def SetDefaultOptions(self):
        self.createpoints.IsChecked = OptionsManager.GetBool("SCR_LineInfoVPI.createpoints", True)
        self.vpionly.IsChecked = OptionsManager.GetBool("SCR_LineInfoVPI.vpionly", False)
        self.createvpis.IsChecked = OptionsManager.GetBool("SCR_LineInfoVPI.createvpis", False)
        self.deletesource.IsChecked = OptionsManager.GetBool("SCR_LineInfoVPI.deletesource", False)
        self.setlayer.IsChecked = OptionsManager.GetBool("SCR_LineInfoVPI.setlayer", False)

        lserial = OptionsManager.GetUint("SCR_LineInfoVPI.layerpicker", 8)
        o = self.currentProject.Concordance.Lookup(lserial) # get the object withj the saved serial number
        if o != None:   # check if we actually got an object back
            if isinstance(o.GetSite(), LayerCollection):    # check if the object is a layer
                self.layerpicker.SetSelectedSerialNumber(lserial, InputMethod(3))
            else:  # if not then set Layer zero                       
                self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:   # if not then set Layer zero               
            self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))

    def SaveOptions(self):
        OptionsManager.SetValue("SCR_LineInfoVPI.createpoints", self.createpoints.IsChecked)    
        OptionsManager.SetValue("SCR_LineInfoVPI.vpionly", self.vpionly.IsChecked)    
        OptionsManager.SetValue("SCR_LineInfoVPI.createvpis", self.createvpis.IsChecked)    
        OptionsManager.SetValue("SCR_LineInfoVPI.deletesource", self.deletesource.IsChecked)    
        OptionsManager.SetValue("SCR_LineInfoVPI.setlayer", self.setlayer.IsChecked)    
        OptionsManager.SetValue("SCR_LineInfoVPI.layerpicker", self.layerpicker.SelectedSerialNumber)

    def IsValid(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        return False

    def setlayerChanged(self, sender, e):
        if self.setlayer.IsChecked:
            self.layerpicker.IsEnabled = True
        else:
            self.layerpicker.IsEnabled = False
        
    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''

        wv = self.currentProject [Project.FixedSerial.WorldView]

        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())

        outPointOnCL = clr.StrongBox[Point3D]()

        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                
                if self.createpoints.IsChecked:
                    
                    if self.vpionly.IsChecked:
                        vpionly = True
                    else:
                        vpionly = False


                    for o in self.objs:
                        if isinstance(o, self.lType):
                            # get the line data as polyseg, in world coordinates
                            polyseg1 = o.ComputePolySeg()
                            polyseg1 = polyseg1.ToWorld()
                            polyseg1_v = o.ComputeVerticalPolySeg()

                            for n in polyseg1_v.Nodes():    # go through all the nodes in the vertical profile
                                try:
                                    # we are working with a profile, so X is the Chainage and Y is the Elevation
                                    # unfortunately the naming convention is different for all the node types
                                    
                                    # LineNode or PointNode
                                    if n.Visible and vpionly == False and \
                                       (n.Type == PolySeg.Node.Type(2) or n.Type == PolySeg.Node.Type(1)): 
                                        polyseg1.FindPointFromStation(n.Point.X, outPointOnCL)  # compute the world XY from the Chainage
                                        self.drawpoint(outPointOnCL.Value, n.Point.Y, o)

                                    # ArcNode
                                    if n.Visible and n.Type == PolySeg.Node.Type(3): 
                                        if vpionly == False:
                                            #Start of Arc
                                            polyseg1.FindPointFromStation(n.Point.X, outPointOnCL)  # compute the world XY from the Chainage
                                            self.drawpoint(outPointOnCL.Value, n.Point.Y, o)

                                        # VPI of Arc
                                        polyseg1.FindPointFromStation(n.PointOfIntersection.X, outPointOnCL) # compute the world XY from the Chainage
                                        self.drawpoint(outPointOnCL.Value, n.PointOfIntersection.Y, o)

                                    # ParabolaNode
                                    if n.Visible and n.Type == PolySeg.Node.Type(4): 
                                        if vpionly == False:
                                            #Start of Parabola
                                            polyseg1.FindPointFromStation(n.Point.X, outPointOnCL)  # compute the world XY from the Chainage
                                            self.drawpoint(outPointOnCL.Value, n.Point.Y, o)
                                        
                                        # VPI of Parabola
                                        polyseg1.FindPointFromStation(n.VPI.X, outPointOnCL)    # compute the world XY from the Chainage
                                        self.drawpoint(outPointOnCL.Value, n.VPI.Y, o)

                                except:
                                    break


                else: # move elevations to VPI tab
                    for o in self.objs:
                        if isinstance(o, self.lType):
                            polyseg1 = o.ComputePolySeg()
                            polyseg1 = polyseg1.ToWorld()
                            polyseg1_v = o.ComputeVerticalPolySeg()

                            if not polyseg1_v:
                                polyseg1_v = PolySeg.PolySeg()
                            polyseg1_v.ReadOnly = False
                            
                            outPointOnCL1 = clr.StrongBox[Point3D]()
                            station1 = clr.StrongBox[float]()

                            nodes = polyseg1.ToPoint3DArray()
                            
                            # go through the chord nodes
                            # and compute the chainage
                            for node in nodes:
                                # getting the chainage of the horizontal node
                                polyseg1.FindPointFromPoint(node, outPointOnCL1, station1)

                                # getting the nodes of the vertical profile
                                nodesv = polyseg1_v.ToPoint3DArray()
                                addvnode = True
                                for nodev in nodesv: # check if that station already exists
                                    if nodev.X == station1.Value: addvnode = False
                                
                                if addvnode:
                                    # add a point with chainage and elevation to the vertical 
                                    polyseg1_v.Add(Point3D(station1.Value, node.Z))

                            l = wv.Add(clr.GetClrType(Linestring))
                            l.Append(polyseg1, polyseg1_v, False, False)
                            SnapInAttributeExtension.CopyUserAttributes(o, l)
                            
                            l.Color = o.Color
                            l.Name = IName.Name.__get__(o)
                            
                            if self.setlayer.IsChecked:
                                l.Layer = self.layerpicker.SelectedSerialNumber
                            else:
                                l.Layer = o.Layer

                            if self.deletesource.IsChecked:
                                osite = o.GetSite()
                                tt2 = osite.Remove(o.SerialNumber)

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

    def drawpoint(self, p_xy, p_z, o):
        wv = self.currentProject[Project.FixedSerial.WorldView]
        vpi = p_xy
        vpi.Z = p_z
        cadPoint = wv.Add(clr.GetClrType(CadPoint))
        cadPoint.Point0 = vpi
        if self.setlayer.IsChecked:
            cadPoint.Layer = self.layerpicker.SelectedSerialNumber
        else:
            cadPoint.Layer = o.Layer


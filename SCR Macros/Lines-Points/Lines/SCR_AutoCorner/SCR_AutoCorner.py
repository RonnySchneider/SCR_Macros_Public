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
    cmdData.Key = "SCR_AutoCorner"
    cmdData.CommandName = "SCR_AutoCorner"
    cmdData.Caption = "_SCR_AutoCorner"
    cmdData.UIForm = "SCR_AutoCorner"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "Lines"
        cmdData.ShortCaption = "Autocorner"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.02
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "extend a line around an asbuilt corner"
        cmdData.ToolTipTextFormatted = "will extend Line 1 around a corner to connect to Line 2"

    except:
        pass
    try:
        b = Bitmap(macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_AutoCorner(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader(macroFileFolder + r"\SCR_AutoCorner.xaml") as s:
            wpf.LoadComponent(self, s)
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

        self.linepicker1.IsEntityValidCallback = self.IsValid
        
        self.linepicker2.IsEntityValidCallback = self.IsValid
        self.linepicker2.ValueChanged += self.line2Changed

        self.lType = clr.GetClrType(IPolyseg)

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass
        Keyboard.Focus(self.linepicker1)

    def SetDefaultOptions(self):
        self.extendl1grade.IsChecked = OptionsManager.GetBool("SCR_AutoCorner.extendl1grade", True)
        self.cleanzero.IsChecked = OptionsManager.GetBool("SCR_AutoCorner.cleanzero", True)


    def SaveOptions(self):
        OptionsManager.SetValue("SCR_AutoCorner.extendl1grade", self.extendl1grade.IsChecked)
        OptionsManager.SetValue("SCR_AutoCorner.cleanzero", self.cleanzero.IsChecked)


    def line2Changed(self, ctrl, e):
        if e.Cause == InputMethod.Mouse:
            l2 = self.linepicker2.Entity
            if l2 != None:
                self.OkClicked(None, None)
                Keyboard.Focus(self.linepicker1)

    def IsValid(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType) and not isinstance(o, ProfileView):
            return True
        return False

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand()

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''
        self.success.Content = ''

        wv = self.currentProject[Project.FixedSerial.WorldView]

        inputok = True
        intersections = Intersections()

        l1 = self.linepicker1.Entity
        l2 = self.linepicker2.Entity
        
        if l1 == None: 
            self.error.Content += '\nno Line 1 selected'
            inputok = False
        if l2 == None: 
            self.error.Content += '\nno Line 2 selected'
            inputok = False
        
        if l1 == l2:
            self.error.Content += '\nLine 1 and 2 are the same'
            inputok = False

        if inputok:


            self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
            UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())

                
            try:
                # the "with" statement will unroll any changes if something go
                # wrong
                with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                    l1_elements = l1.GetElements()
                    l2_elements = l2.GetElements()

                    if self.cleanzero.IsChecked:
                        rem_l1 = 0
                        rem_l2 = 0
                        # cleanup zero length elements
                        for i in reversed(range(1, l1_elements.Count)):
                            if Vector3D(l1_elements[i].Position, l1_elements[i-1].Position).Length2D == 0.000:
                                l1.RemoveElementAt(i)
                                rem_l1 += 1
                                self.success.Content += '\nLine 1 - removed ' + str(rem_l1) + ' Segments'

                        for i in reversed(range(1, l2_elements.Count)):
                            if Vector3D(l2_elements[i].Position, l2_elements[i-1].Position).Length2D == 0.000:
                                l2.RemoveElementAt(i)
                                rem_l2 += 1
                                self.success.Content += '\nLine 2 - removed ' + str(rem_l2) + ' Segments'

                        l1_elements = l1.GetElements()
                        l2_elements = l2.GetElements()

                    s1 = l1_elements[0].Position
                    e1 = l1_elements[l1_elements.Count - 1].Position
                    s2 = l2_elements[0].Position
                    e2 = l2_elements[l2_elements.Count - 1].Position

                    s1s2 = Vector3D(s1,s2).Length
                    s1e2 = Vector3D(s1,e2).Length
                    e1s2 = Vector3D(e1,s2).Length
                    e1e2 = Vector3D(e1,e2).Length

                    # reverse the lines if necessary
                    if (s1s2 < s1e2) and (s1s2 < e1s2) and (s1s2 < e1e2):
                        l1.ReverseDirection()
                    if (s1e2 < s1s2) and (s1e2 < e1s2) and (s1e2 < e1e2):
                        l1.ReverseDirection()
                        l2.ReverseDirection()
                    if (e1s2 < s1s2) and (e1s2 < s1e2) and (e1s2 < e1e2):
                        pass
                    if (e1e2 < s1s2) and (e1e2 < s1e2) and (e1e2 < e1s2):
                        l2.ReverseDirection()

                    # in case the lines touch each other we'd now have L1-End and L2-Start facing each other -> Zero-Gap
                    # double check and reverse lines in that case
                    l1_elements = l1.GetElements()
                    l2_elements = l2.GetElements()
                    e1 = l1_elements[l1_elements.Count - 1].Position
                    s2 = l2_elements[0].Position
                    if Vector3D(e1, s2).Length == 0.0:
                        l1.ReverseDirection()
                        l2.ReverseDirection()

                    polyseg1 = l1.ComputePolySeg()
                    polyseg1 = polyseg1.ToWorld()
                    polyseg1_v = l1.ComputeVerticalPolySeg()
                    polyseg2 = l2.ComputePolySeg()
                    polyseg2 = polyseg2.ToWorld()


                    if polyseg1.LastSegment.Intersect(polyseg2.FirstSegment, True, intersections):
                        i2 = 0
                        foundvalidintersection = False
                        for i in intersections: # if we get multiple intersections we want the shortest one
                            ### if i.T2 >= 0 and i.T2 <=1:    
                            foundvalidintersection = True
                            if i2 == 0:
                                ip = i.Point
                                i2 += 1
                            else:
                                if polyseg1.LastNode.Point.Distance(i.Point) < polyseg1.LastNode.Point.Distance2D.Distance(ip):
                                    ip = i.Point
                        
                        if foundvalidintersection:
                            pnew = ip
                            if  polyseg1.LastNode.Point.Z == polyseg2.FirstNode.Point.Z and not self.extendl1grade.IsChecked:
                                pnew.Z = polyseg1.LastNode.Point.Z
                            else:
                                if self.extendl1grade.IsChecked:
                                    # Segment.SlopeRate delivers only Zero ???
                                    # have to compute it manually
                                    dh = polyseg1.LastSegment.EndPoint.Z - polyseg1.LastSegment.BeginPoint.Z
                                    fulld2d = polyseg1.LastSegment.Length2D
                                    partd2d = polyseg1.LastNode.Point.Distance2D(pnew)
                                    
                                    pnew.Z = polyseg1.LastNode.Point.Z + (dh / fulld2d) * partd2d
                                else:
                                    # complete code for interpolation
                                    dh = polyseg2.FirstNode.Point.Z - polyseg1.LastNode.Point.Z
                                    fulld2d = polyseg1.LastNode.Point.Distance2D(pnew) + polyseg2.FirstNode.Point.Distance2D(pnew)
                                    partd2d = polyseg1.LastNode.Point.Distance2D(pnew)

                                    pnew.Z = polyseg1.LastNode.Point.Z + (dh / fulld2d) * partd2d

                            e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                            e.Position = pnew
                            l1.AppendElement(e)       
                            e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                            e.Position = polyseg2.FirstNode.Point
                            l1.AppendElement(e)       


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

    


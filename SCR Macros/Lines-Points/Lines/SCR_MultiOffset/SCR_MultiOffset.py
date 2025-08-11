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
    cmdData.Key = "SCR_MultiOffset"
    cmdData.CommandName = "SCR_MultiOffset"
    cmdData.Caption = "_SCR_MultiOffset"
    cmdData.UIForm = "SCR_MultiOffset"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "Lines"
        cmdData.ShortCaption = "Multi Offset"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.03
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "offset multiple lines at once"
        cmdData.ToolTipTextFormatted = "offset multiple lines at once"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_MultiOffset(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_MultiOffset.xaml") as s:
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
        self.objs.IsEntityValidCallback = self.IsValid

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        #self.lfp = self.lunits.Properties.Copy()
        linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation
        #self.lfp.AddSuffix = False
        self.hzoffsetlabel.Content = "horizontal Offset [" + linearsuffix + "]"
        self.voffsetlabel.Content = "vertical Offset [" + linearsuffix + "]"
        
		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def SetDefaultOptions(self):
        
        settingserial = OptionsManager.GetUint("SCR_MultiOffset.layerpicker", 8) # 8 is FixedSerial for Layer Zero
        o = self.currentProject.Concordance.Lookup(settingserial)
        if o != None:
            if isinstance(o.GetSite(), LayerCollection):    
                self.layerpicker.SetSelectedSerialNumber(settingserial, InputMethod(3))
            else:                       
                self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:                       
            self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))

        self.hzoffset.Distance = OptionsManager.GetDouble("SCR_MultiOffset.hzoffset", 0.0000)
        self.voffset.Distance = OptionsManager.GetDouble("SCR_MultiOffset.voffset", 0.0000)

        self.left.IsChecked = OptionsManager.GetBool("SCR_MultiOffset.left", True)
        self.right.IsChecked = OptionsManager.GetBool("SCR_MultiOffset.right", False)
        self.both.IsChecked = OptionsManager.GetBool("SCR_MultiOffset.both", False)



    def SaveOptions(self):

        OptionsManager.SetValue("SCR_MultiOffset.layerpicker", self.layerpicker.SelectedSerialNumber)
        OptionsManager.SetValue("SCR_MultiOffset.hzoffset", self.hzoffset.Distance)
        OptionsManager.SetValue("SCR_MultiOffset.voffset", self.voffset.Distance)

        OptionsManager.SetValue("SCR_MultiOffset.left", self.left.IsChecked)
        OptionsManager.SetValue("SCR_MultiOffset.right", self.right.IsChecked)
        OptionsManager.SetValue("SCR_MultiOffset.both", self.both.IsChecked)

    def IsValid(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        return False

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()


    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content = ''
        self.success.Content = ''

        wv = self.currentProject [Project.FixedSerial.WorldView]
        
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())

        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                voffset = self.voffset.Distance
                if math.isnan(voffset): voffset = 0.0
                    
                for o in self.objs:
                    if isinstance(o, self.lType):

                        
                        polyseg = o.ComputePolySeg().Clone()
                        polyseg = polyseg.ToWorld()
                        polyseg_v = o.ComputeVerticalPolySeg()
                        
                        if polyseg.NumberOfNodes >= 3 and polyseg.IsClosed and polyseg.IsClockWise():
                            polyseg.Reverse()
                            if polyseg_v:
                                polyseg_v.Reverse()
                        
                        polyseg_org_lin = polyseg.Linearize(0.0001, 0.0001, 10000.0, polyseg_v, True)
                        #polyseg_v_org_lin = polyseg_org_lin.Clone()
                        #polyseg_v_org_lin.ConvertPolysegToStationElevation(1.0)

                        #tt0 = polyseg.ComputeStationing()
                        #limits = polyseg_v.BoundingBox.ptMax

                        outPointOnCL = clr.StrongBox[Point3D]()
                        station = clr.StrongBox[float]()

                        #if polyseg_v:
                        #    polyseg_v = polyseg_v.Clone()
                        #    tt1 = polyseg_v.LastNode.Point.X
                        #    #polyseg_v.BeginStation = polyseg.BeginStation
                        #    polyseg_v.ShiftVerticalPolyseg(-polyseg.BeginStation)
                        #    polyseg_v.Translate(polyseg_v.FirstNode, polyseg_v.LastNode, Vector3D(0, voffset, 0))
                        #else:
                        #    if not polyseg.AllPointsAre3D:
                        #        polyseg_v = PolySeg.PolySeg()
                        #        polyseg_v.Add(Point3D(0, voffset, 0))
                        #    else:
                        #        polyseg.Translate(polyseg.FirstNode, polyseg.LastNode, Vector3D(0, 0, voffset))
                        

                        if self.left.IsChecked or self.both.IsChecked:
                            l = wv.Add(clr.GetClrType(Linestring))
                            # l.Name = newname

                            polyseg_ol = polyseg.Offset(Side.Left, abs(self.hzoffset.Distance))[1]
                            polyseg_ol_v = self.verticalatoffset(Side.Left, polyseg_ol, polyseg_org_lin)
                            #polyseg_ol_v = polyseg_ol.ComputeElevationOverrideOnOffsetPolyseg(polyseg_org_lin, polyseg_v_org_lin)

                            l.Layer = self.layerpicker.SelectedSerialNumber
                            l.Append(polyseg_ol, polyseg_ol_v, False, False)
                            #l.Append(polyseg_ol, polyseg_ol_v, False, False)
                        
                        if self.right.IsChecked or self.both.IsChecked:
                            l = wv.Add(clr.GetClrType(Linestring))
                            # l.Name = newname
                            polyseg_or = polyseg.Offset(Side.Right, abs(self.hzoffset.Distance))[1]
                            polyseg_or_v = self.verticalatoffset(Side.Right, polyseg_or, polyseg_org_lin)
                            #polyseg_or_v = polyseg_or.ComputeElevationOverrideOnOffsetPolyseg(polyseg_org_lin, polyseg_v_org_lin)

                            l.Layer = self.layerpicker.SelectedSerialNumber
                            l.Append(polyseg_or, polyseg_or_v, False, False)
                            #l.Append(polyseg_or, polyseg_or_v, False, False)

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
        
            
        self.error.Content += '\ncheck elevations - probably not correct'

        self.SaveOptions()

    def verticalatoffset(self, side, polyseg_os, polyseg_org_lin):

        verticalnodes = List[Point3D]()
        
        # get the first two segments from the original polyseg
        prev_s = polyseg_org_lin.FirstSegment
        # compute the location from the original polyseg onto the offset line
        verticalnodes.Add(Point3D(self.computechainage(polyseg_os, prev_s.BeginPoint), prev_s.BeginPoint.Z, 0))
        curr_s = polyseg_org_lin.Next(prev_s)
        while not curr_s == None:

            # offset the segments
            prev_s_o = prev_s.Clone()
            prev_s_o.Offset(side, abs(self.hzoffset.Distance))
            curr_s_o = curr_s.Clone()
            curr_s_o.Offset(side, abs(self.hzoffset.Distance))
        
            intersections = Intersections()
            slopeprev = Plane3D(prev_s_o.BeginPoint, prev_s_o.EndPoint, prev_s.EndPoint)[0]
            slopecurr = Plane3D(curr_s_o.BeginPoint, curr_s_o.EndPoint, curr_s.BeginPoint)[0]
            
            # intersect the offset segments
            if prev_s_o.Intersect(curr_s_o, True, intersections):

                ips = List[Point3D]()
                ip1 = curr_s.BeginPoint
                ip2 = intersections[0].Point
                ip2.Z = ip1.Z
                ips.Add(ip1)
                
                tt = abs(abs(slopeprev.Slope) - abs(slopecurr.Slope))
                # in case the slope is changing enough we can compute the plane intersection
                if abs(abs(slopeprev.Slope) - abs(slopecurr.Slope)) > 0.00001:
                    
                    ray = Plane3D.IntersectPlanes(slopeprev, slopecurr).Direction.ToVector3D()
                    ip2 = ip1 + ray
                    ips.Add(ip2)
                    # create a polyseg along the plane intersection
                    seg = PolySeg.PolySeg()
                    seg.Add(ips)
            
                    intersections.Clear()
                    maxit = 0
                    # extend the polyseg until we find an intersection with the unlinerarized offset polyseg
                    while not seg.Intersect(polyseg_os, True, intersections):
                        seg.ExtendStart(False, False, abs(self.hzoffset.Distance), True)
                        seg.ExtendEnd(False, False, abs(self.hzoffset.Distance), True)

                        maxit += 1
                        if maxit == 50: break
                
                    wv = self.currentProject [Project.FixedSerial.WorldView]
                    #l = wv.Add(clr.GetClrType(Linestring))
                    #l.Append(seg, None, False, False)
                    #l.Color = Color.Blue

                    if intersections.Count > 0:

                        pt = seg.FirstSegment.ComputePoint(intersections[0].T1)
                        
                        #cadPoint = wv.Add(clr.GetClrType(CadPoint))
                        #cadPoint.Point0 = pt[1]
                
                        verticalnodes.Add(Point3D(self.computechainage(polyseg_os, pt[1]), pt[1].Z, 0))

            #else: # no initial intersection = outside, need to add 3 values
            #
            #    # now intersect with extended offset segments
            #    prev_s_o.Intersect(curr_s_o, True, intersections)
            #
            #    verticalnodes.Add(Point3D(self.computechainage(polyseg_os, prev_s_o.EndPoint), prev_s_o.EndPoint.Z, 0))
            #    #verticalnodes.Add(Point3D(self.computechainage(polyseg_os, intersections[0].Point), prev_s_o.EndPoint.Z, 0))
            #    verticalnodes.Add(Point3D(self.computechainage(polyseg_os, curr_s_o.BeginPoint), curr_s_o.BeginPoint.Z, 0))
            #
            #    tt = 1

            intersections.Clear()
            prev_s = curr_s
            curr_s = polyseg_org_lin.Next(prev_s)

        verticalnodes.Add(Point3D(self.computechainage(polyseg_os, prev_s.EndPoint), prev_s.EndPoint.Z, 0))
        finalpolyseg = PolySeg.PolySeg()
        finalpolyseg.Add(verticalnodes)

        return finalpolyseg

    def computechainage(self, poly, p):

        outPointOnCL = clr.StrongBox[Point3D]()
        station = clr.StrongBox[float]()

        poly.FindPointFromPoint(p, outPointOnCL, station)

        return station.Value

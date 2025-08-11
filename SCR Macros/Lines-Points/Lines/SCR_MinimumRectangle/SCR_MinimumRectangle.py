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
    cmdData.Key = "SCR_MinimumRectangle"
    cmdData.CommandName = "SCR_MinimumRectangle"
    cmdData.Caption = "_SCR_MinimumRectangle"
    cmdData.UIForm = "SCR_MinimumRectangle"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "Lines"
        cmdData.ShortCaption = "Minimum Rectangle"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.03
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "create Minimum Rectangle around Polyline"
        cmdData.ToolTipTextFormatted = "create Minimum Rectangle around Polyline"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png") 
        cmdData.ImageSmall = b
    except:
        pass


class SCR_MinimumRectangle(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_MinimumRectangle.xaml") as s:
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

        self.linepicker.IsEntityValidCallback=self.IsValid
        self.lType = clr.GetClrType(IPolyseg)

        self.linepicker.AutoTab = False
        
		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def SetDefaultOptions(self):
        lserial = OptionsManager.GetUint("SCR_MinimumRectangle.layerpicker", 8)
        o = self.currentProject.Concordance.Lookup(lserial) # get the object withj the saved serial number
        if o != None:   # check if we actually got an object back
            if isinstance(o.GetSite(), LayerCollection):    # check if the object is a layer
                self.layerpicker.SetSelectedSerialNumber(lserial, InputMethod(3))
            else:  # if not then set Layer zero                       
                self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:   # if not then set Layer zero               
            self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))

    def SaveOptions(self):
        OptionsManager.SetValue("SCR_MinimumRectangle.layerpicker", self.layerpicker.SelectedSerialNumber)


    def IsValid(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        self.success.Content = '\nplease select a valid Line'
        return False

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''

        self.success.Content = ""

        wv = self.currentProject [Project.FixedSerial.WorldView]
        
        o = self.linepicker.Entity
        if o:
            self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
            UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
            try:
                # the "with" statement will unroll any changes if something go wrong
                with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                    polyseg1 = o.ComputePolySeg()
                    polyseg1 = polyseg1.ToWorld()

                    # we can use any point as rotation center, we just use the first node of the line
                    rot_p = polyseg1.FirstNode.Point
                    rot_p.Z = 0

                    # we use the built-in boundaries that the polyseg provides
                    # the draw back is that the boundaries are always square to the coordinate system N/S and E/W
                    # so, we rotate the polyseg, while keeping track of the rotation and compute the X/Y area of the bounding box

                    # start the recursive search for the minimum area 
                    min_rot = self.findmin(o, rot_p, 0, math.pi/2, math.pi/2/90) # line object, center of rot, start rot, end rot, rot increment
          
                    # we rotate the polyseg to the position with the minimum bounding box area
                    polyseg1.Rotate(rot_p, min_rot[1])
                    
                    # we get the bounding box once more and use the coordinates to draw a rectangle
                    bb = polyseg1.BoundingBox
                    l = wv.Add(clr.GetClrType(Linestring))
                    for i in range(0,4):
                        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                        p = bb.Vertices[i]
                        p.Z = 0
                        e.Position = p
                        l.AppendElement(e)
        
                    e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                    p = bb.Vertices[0]
                    p.Z = 0
                    e.Position = p
                    l.AppendElement(e)

                    # we build the transformation matrix to rotate the bounding box line, which is exactly N/S and E/W to it's actual final position
                    tm = Matrix4D.BuildTransformMatrix(rot_p, rot_p, -min_rot[1], 1, 1, 1)
                    
                    # now that we have the rectangle square to the coordinate system we rotate it back to the polysegs original position
                    l.Transform(TransformData(tm, None))
                    #l.Color = Color.Lime
                    l.Layer = self.layerpicker.SelectedSerialNumber

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
        
        else:
            self.success.Content = '\nplease select a valid Line'
 
        self.SaveOptions()
            
    def findmin(self, o, rot_cent, startrot, endrot, rot_inc):
        
        # create empty arrays
        single_rot = Array[float]([0.000]) + Array[float]([0.000])
        all_rot = []
        
        # reset the polyseg to original values
        polyseg1 = o.ComputePolySeg()
        polyseg1 = polyseg1.ToWorld()
        
        # if the end rot value is smaller than the start value switch places
        if endrot < startrot:
            endrot, startrot = startrot, endrot
        
        rot_cur = startrot
        
        #rotate the polyseg to the minimum rotation of the previous loop
        polyseg1.Rotate(rot_cent, startrot)

        # do this computation for all the other rotations within the current computation limits
        while rot_cur <= endrot:
            
            bb = polyseg1.BoundingBox # get the bounds of that rotated polyseg
    
            # single_rot (area, rotation)
            # compute the area and save it together with the current rotation into the array
            single_rot[0] = abs(bb.Vertices[0].X - bb.Vertices[2].X) * abs(bb.Vertices[0].Y - bb.Vertices[2].Y)
            single_rot[1] = rot_cur
            all_rot.Add(single_rot.Clone())
            
            # increment values, computation and saving happens in next for loop
            polyseg1.Rotate(rot_cent, rot_inc) # rotate the polyseg a step further
            rot_cur += rot_inc # increment the running rotation
        
        # sort the results by column 0 , which is the area
        all_rot.sort(key=lambda x: x[0])
        if all_rot[1][0] - all_rot[0][0] < 0.000001: # if the difference between the two smallest areas is less than this, consider it done
            return all_rot[0]
        else:
            # otherwise call the same function with new start values
            # -previous increment < rotation of previous area minimum < +previous increment
            # and also reduce the increment
            return self.findmin(o, rot_cent, all_rot[0][1] - 2 * rot_inc, all_rot[0][1] + 2 * rot_inc, rot_inc/10)


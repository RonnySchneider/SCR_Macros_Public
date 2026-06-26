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

from System.Collections.Generic import List, IEnumerable
exec(open(r"C:\ProgramData\Trimble\MacroCommands3\SCR Macros\SCR_Imports.py").read())

_OPTIONS = {
    "layerpicker": 8, "offset_v": 0.0,
}

def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_LinePerpVertOffset"
    cmdData.CommandName = "SCR_LinePerpVertOffset"
    cmdData.Caption = "_SCR_LinePerpVertOffset"
    cmdData.UIForm = "SCR_LinePerpVertOffset"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "Lines"
        cmdData.ShortCaption = "LinePerpVertOffset"
        cmdData.DefaultRibbonToolSize = 0 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.07
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "offset a Lines vertical profile in perpendicular direction"
        cmdData.ToolTipTextFormatted = "offset a Lines vertical profile in perpendicular direction"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_LinePerpVertOffset(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_LinePerpVertOffset.xaml") as s:
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
        #types = Array [Type] ([CadPoint]) + Array [Type] ([Point3D])    # we fill an array with TBC object types, we could combine different types

        self.objs.IsEntityValidCallback = self.IsValid
        self.lType = clr.GetClrType(IPolyseg)

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        #self.lfp = self.lunits.Properties.Copy()
        self.linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation

        self.labeloffset.Content = 'vertical perpendicular Offset [' + self.linearsuffix + ']'

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def SetDefaultOptions(self):
        SCROptions.LoadMacroOptions(self, "SCR_LinePerpVertOffset", _OPTIONS, self.currentProject)

    def SaveOptions(self):
        SCROptions.SaveMacroOptions(self, "SCR_LinePerpVertOffset", _OPTIONS)

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

        wv = self.currentProject [Project.FixedSerial.WorldView]
        # bc = self.currentProject [Project.FixedSerial.BlockCollection]    # getting all blocks as collection
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)

        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                for o in self.objs:
                    if isinstance(o, self.lType):


                        #offset_h = self.offset_h.Value
                        offset_v = self.offset_v.Distance

                        # get the line data as polyseg, in world coordinates
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
                                # add a point with chainage and elevation to the vertical 
                                polyseg1_v.Add(Point3D(station1.Value, node.Z, 0))

                        # offset the line vertically
                        offset_v_polyseg = polyseg1_v.Offset(Side.Left, offset_v)

                        container = o.GetSite()
                        ls = container.Add(Linestring)
                        ls.Append(polyseg1, offset_v_polyseg[1], True, False)
                        ls.Layer = self.layerpicker.SelectedSerialNumber



                        # compute the vertical polyseg for the horizontal offset' one
                        #polyseg1_hoffset_v = ls.ComputeVerticalPolySeg()
                        
                        #polyseg1_voffset = ls.ComputePolySeg()
                        #polyseg1_voffset_v = ls.ComputeVerticalPolySeg()

                        #container.Remove(ls.SerialNumber)

                        ## offset the line horizontally
                        #if offset_h < 0:
                        #    offset_h_polyseg = polyseg1_voffset.Offset(Side.Left, offset_h, PolySeg.PolySeg())
                        #else:
                        #    offset_h_polyseg = polyseg1_voffset.Offset(Side.Right, offset_h, PolySeg.PolySeg())
                        #
                        #
                        #
                        #
                        #ls = container.Add(Linestring)
                        #ls.Append(offset_h_polyseg[1], polyseg1_voffset_v, True, False)

                        
                         #ls.Append(None, tt[1], False, False)
                failGuard.Commit()
                self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
                UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
        
        except Exception as e:
            tt = sys.exc_info()
            exc_type, exc_obj, exc_tb = sys.exc_info()
            # EndMark MUST be set no matter what
            # otherwise TBC won't work anymore and needs to be restarted
            self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
            UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
            self.error.Content += '\nan Error occurred - Result probably incomplete\n' + str(exc_type) + '\n' + str(exc_obj) + '\nLine ' + str(exc_tb.tb_lineno)
        
        #wv.PauseGraphicsCache(False)

        self.SaveOptions()

        
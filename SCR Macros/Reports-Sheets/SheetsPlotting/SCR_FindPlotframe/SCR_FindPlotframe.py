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
    cmdData.Key = "SCR_FindPlotframe"
    cmdData.CommandName = "SCR_FindPlotframe"
    cmdData.Caption = "_SCR_FindPlotframe"
    cmdData.UIForm = "SCR_FindPlotframe"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Reports"
        cmdData.DefaultTabGroupKey = "Sheets and Dynaviews"
        cmdData.ShortCaption = "Find Plotframe"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.04
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Find Plotframe based on Coordinate"
        cmdData.ToolTipTextFormatted = "Find Plotframe based on Coordinate"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png") 
        cmdData.ImageSmall = b
    except:
        pass


class SCR_FindPlotframe(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_FindPlotframe.xaml") as s:
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

        optionMenu = SelectionContextMenuHandler()
        # remove options that don't apply here
        optionMenu.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        self.objs.ButtonContextMenu = optionMenu

        self.objs.IsEntityValidCallback=self.IsValid

        self.coordpick1.ShowElevationIf3D = True
        self.coordpick1.ValueChanged += self.coordpick1Changed

        self.coordpick1.AutoTab = False

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def SetDefaultOptions(self):
        self.exportonly.IsChecked = OptionsManager.GetBool("SCR_FindPlotframe.exportonly", False)
        self.findframes.IsChecked = OptionsManager.GetBool("SCR_FindPlotframe.findframes", True)


    def SaveOptions(self):
        OptionsManager.SetValue("SCR_FindPlotframe.exportonly", self.exportonly.IsChecked)    
        OptionsManager.SetValue("SCR_FindPlotframe.findframes", self.findframes.IsChecked)    


    def IsValid(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        return False

    def coordpick1Changed(self, ctrl, e):

        if e.Cause == InputMethod.Mouse:     
            self.OkClicked(None, None)


    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''

        wv = self.currentProject [Project.FixedSerial.WorldView]
        
        if self.exportonly.IsChecked:
        
            filename = os.path.expanduser('~/Downloads/Frame-Coordinates.csv')
            if File.Exists(filename):
                File.Delete(filename)

            with open(filename, 'w') as f:
                self.success.Content += "\nwriting to /Downloads/Frame-Coordinates.csv\n"
                for o in self.objs:
                    if isinstance(o, IPolyseg):

                        polyseg = o.ComputePolySeg()
                        polyseg_points = polyseg.ToPoint3DArray()

                        if polyseg_points.Count > 0:

                            outputline = ""
                            sourcename = self.getlinename(o)

                            if sourcename == "":
                                sourcename = "object has no name"

                            outputline += sourcename

                            for p in polyseg_points:
                                outputline += "," + str(p.X)
                                outputline += "," + str(p.Y)



                    f.write(outputline + "\n")
                        
            f.close()
            self.success.Content += "\nDone"
        else:

            p1 = self.coordpick1.Coordinate

            if not math.isnan(p1.X):
                self.success.Content = "searching\n"
                for o in wv:
                    if isinstance(o, self.lType):
                        sourcename = ""

                        polyseg = o.ComputePolySeg()      # we convert the line object to a more unified polyseg, for which we have nifty functions available
                        
                        if polyseg.IsClosed:

                            sourcename = self.getlinename(o)

                            if polyseg.PointInPolyseg(p1) == Side.In:
                                if sourcename == "":
                                    sourcename = "found object but it has no name"
                                self.success.Content += "\n" + sourcename

                Keyboard.Focus(self.coordpick1)
                self.success.Content += "\n\nDone"
            else:
                self.success.Content += "\nno Coordinate entered"
    
        self.SaveOptions()


    def getlinename(self, o):
        sourcename = ""
        # for linestrings .Name is set only, we have to use Anchorname
        # # but CadPolyline and Arc don't have Anchorname, here we'd have to use Name
        try:
            sourcename = o.Name
        except:
            try:
                sourcename = o.AnchorName
            except:
                pass

        return sourcename
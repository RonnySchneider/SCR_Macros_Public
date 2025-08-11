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
    cmdData.Key = "SCR_PowerlineSway"
    cmdData.CommandName = "SCR_PowerlineSway"
    cmdData.Caption = "_SCR_PowerlineSway"
    cmdData.UIForm = "SCR_PowerlineSway"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "Lines"
        cmdData.ShortCaption = "Powerline Sway"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.13
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Powerline Sway"
        cmdData.ToolTipTextFormatted = "Powerline Sway"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png") 
        cmdData.ImageSmall = b
    except:
        pass

class SCR_PowerlineSway(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_PowerlineSway.xaml") as s:
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

        self.hortol.DistanceMin = 0.0
        self.hortol.Allow3DDistance = True
        self.vertol.DistanceMin = 0.0
        self.vertol.Allow3DDistance = True
        self.nodespacing.DistanceMin = 0.0
        self.nodespacing.Allow3DDistance = True
        
        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        #self.lfp = self.lunits.Properties.Copy()
        self.linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation

        self.toleranceheader.Header = 'define Chording Tolerance [' + self.linearsuffix + ']'

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def SetDefaultOptions(self):

        self.hortol.SetDistance(OptionsManager.GetDouble("SCR_PowerlineSway.hortol", 0.0001), self.currentProject)
        self.vertol.SetDistance(OptionsManager.GetDouble("SCR_PowerlineSway.vertol", 0.0001), self.currentProject)
        self.nodespacing.SetDistance(OptionsManager.GetDouble("SCR_PowerlineSway.nodespacing", 10000.0), self.currentProject)
        self.rotation.SetAngle(OptionsManager.GetDouble("SCR_PowerlineSway.rotation", math.pi/2), self.currentProject)

    def SaveOptions(self):

        OptionsManager.SetValue("SCR_PowerlineSway.hortol", self.hortol.Distance)
        OptionsManager.SetValue("SCR_PowerlineSway.vertol", self.vertol.Distance)
        OptionsManager.SetValue("SCR_PowerlineSway.nodespacing", self.nodespacing.Distance)
        OptionsManager.SetValue("SCR_PowerlineSway.rotation", self.rotation.Angle)


    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content = ''
        self.success.Content = ''
        
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)

        wv = self.currentProject [Project.FixedSerial.WorldView]
        
        # pause the graphics so every change doesn't draw
        #wv.PauseGraphicsCache(True)

        try:

            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                inputok=True

                l1 = self.linepicker1.Entity
                if l1 == None: 
                    self.success.Content += '\nno Line selected'
                    inputok=False
                else:
                    polyseg1 = l1.ComputePolySeg()
                    polyseg1 = polyseg1.ToWorld()
                    polyseg1_v = l1.ComputeVerticalPolySeg()
                    if polyseg1.IsClosed:
                        self.success.Content += '\nselected Line is closed'
                        inputok=False

                if inputok:

                    # chord the line
                    polyseg1 = polyseg1.Linearize(self.hortol.Distance, self.vertol.Distance, self.nodespacing.Distance, polyseg1_v, False)

                    p1 = polyseg1.FirstNode.Point
                    p2 = polyseg1.LastNode.Point
                    
                    
                    # BiVector3D(Trimble.Vce.Geometry.Vector3D rotationAxis, double rotationAngle)
                    # Spinor3D(Trimble.Vce.Geometry.BiVector3D biVector)
                    targetrot = Spinor3D(BiVector3D(Vector3D(p1, p2), self.rotation.Angle))
                    # BuildTransformMatrix(Vector3D fromPoint, Vector3D translation, Spinor3D rotation, Vector3D scale)
                    targetmatrix = Matrix4D.BuildTransformMatrix(Vector3D(p1), Vector3D(0,0,0), targetrot, Vector3D(1,1,1))

                    # temp cad-point
                    cadPoint = wv.Add(clr.GetClrType(CadPoint))

                    nodes = polyseg1.ToPoint3DArray()

                    l = wv.Add(clr.GetClrType(Linestring))
                    l.Layer = l1.Layer
                    for node in nodes:

                        cadPoint.Point0 = node
                        #Transform(Trimble.Vce.Interfaces.SnapIn.TransformData transform, double scale, double rotation)
                        cadPoint.Transform(TransformData(targetmatrix, Matrix4D(Vector3D.Zero)))

                        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                        e.Position = cadPoint.Point0
                        l.AppendElement(e)

                    # remove temp CAD-Point
                    osite = cadPoint.GetSite()    # we find out in which container the serial number reside
                    osite.Remove(cadPoint.SerialNumber)   # we delete the object from that container

                    self.linepicker1.SerialNumber = l.SerialNumber

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

        #wv.PauseGraphicsCache(False)
        self.SaveOptions()

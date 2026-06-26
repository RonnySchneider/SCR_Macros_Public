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
    "hortol": 0.0001, "vertol": 0.0001, "nodespacing": 2.0,
    "sourcelayer": True, "setlayer": False, "layerpicker": 8,
    "setprefixsuffix": False, "prefix": "", "suffix": "",
    "changeexisting": True, "deletesource": False,
}

def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_ChordLines"
    cmdData.CommandName = "SCR_ChordLines"
    cmdData.Caption = "_SCR_ChordLines"
    cmdData.UIForm = "SCR_ChordLines"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "Lines"
        cmdData.ShortCaption = "Chord Lines"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.13
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "convert curved Lines to chorded Polylines"
        cmdData.ToolTipTextFormatted = "convert curved Lines to chorded Polylines"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_ChordLines(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_ChordLines.xaml") as s:
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

        #self.hortol.NumberOfDecimals = 4
        #self.vertol.NumberOfDecimals = 4
        #self.nodespacing.NumberOfDecimals = 4

        self.objs.IsEntityValidCallback=self.IsValid
        optionMenu = SelectionContextMenuHandler()
        # remove options that don't apply here
        optionMenu.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        self.objs.ButtonContextMenu = optionMenu
        
        self.lType = clr.GetClrType(IPolyseg)

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        #self.lfp = self.lunits.Properties.Copy()
        self.linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation

        self.toleranceheader.Header = 'define Computation Tolerance [' + self.linearsuffix + ']'

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def changeexistingChanged(self, sender, e):
        if self.changeexisting.IsChecked:
            self.deletesource.IsEnabled = False
        else:
            self.deletesource.IsEnabled = True

    def SetDefaultOptions(self):
        SCROptions.LoadMacroOptions(self, "SCR_ChordLines", _OPTIONS, self.currentProject)

    def SaveOptions(self):
        SCROptions.SaveMacroOptions(self, "SCR_ChordLines", _OPTIONS)

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
        lgc = LayerGroupCollection.GetLayerGroupCollection(self.currentProject, False)
                
        wv.PauseGraphicsCache(True)

        self.success.Content=''
        # self.label_benchmark.Content = ''

        # settings = Model3DCompSettings.ProvideSettingsObject(self.currentProject)
        ProgressBar.TBC_ProgressBar.Title = "chording Lines"
        time1 = datetime.now()

        try:

            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                objs_count = self.objs.Count # if delete source is ticked the self.objs.Count would go down to Zero and produce a divison by zero

                selectionserials = []
                deletedserials = []
                createdserials = []

                # save the serials before doing something to them
                # should be faster than updating properties and shoe line direction every time
                for o in self.objs:
                    selectionserials.Add(o.SerialNumber)

                GlobalSelection.Clear()

                j = 0
                for sn in selectionserials:

                    o = self.currentProject.Concordance[sn]

                    if isinstance(o, self.lType):

                        t1 = abs(self.hortol.Distance)
                        t2 = abs(self.vertol.Distance)
                        t3 = abs(self.nodespacing.Distance)

                        polyseg1_v = None

                        if o.Normal == Vector3D(0, 0, 1):

                            polyseg1 = o.ComputePolySeg()
                            polyseg1 = polyseg1.ToWorld()
                            polyseg1_v = o.ComputeVerticalPolySeg()
                            polyseg_new = polyseg1.Linearize(t1, t2, t3, polyseg1_v, False)
                        
                        else:

                            # first need to compute the transformation matrix to flat and back
                            # don't know if there is a more elegant built-in way
                            centerp = o.OriginOfUcs
                            nv = o.Normal.Clone()
                            # will only work if Normal != 0, 0, 1
                            vx = nv.Clone()
                            # rotate 90 degrees around world-Z axis and make it level with world horizon
                            vx.RotateAboutZ(math.pi/2)
                            vx.Horizon = 0
                            # don't really need vy
                            # clone the axis and rotate it 90 degrees around Normal
                            vy = vx.Clone()
                            vy.Rotate(BiVector3D(nv, math.pi/2))
                            # vx, vy and Normal are now square to each other

                            # compute matrix to line up UCS-X,Z with World-X,Z axis
                            rottozero = Spinor3D.ComputeRotation(vx, nv, Vector3D(1,0,0), Vector3D(0,0,1))
                            # transformation to 0, 0, 0
                            matrixtozero = Matrix4D.BuildTransformMatrix(Vector3D(centerp), Vector3D(centerp, Point3D(0, 0, 0)), rottozero, Vector3D(1,1,1))
                            #matrixbackfromzero = Matrix4D.Inverse(matrixtozero)
                            # transformation without shift - just flatten it
                            matrixtoflat = Matrix4D.BuildTransformMatrix(Vector3D(centerp), Vector3D(0, 0, 0), rottozero, Vector3D(1,1,1))
                            matrixbackfromflat = Matrix4D.Inverse(matrixtoflat)

                            #self.debugxyz(centerp, vx, vy, nv)

                            #matrixtest = Matrix4D(centerp, vx, nv)
                            #matrixtestback = Matrix4D.Inverse(matrixtest)

                            # remove the UCS from the object - tilts it to flat
                            orgNormal = o.Normal.Clone()
                            o.Normal = Vector3D(0, 0, 1)
                            
                            # compute and linearize the flat polyseg
                            polyseg1 = o.ComputePolySeg().Clone()
                            polyseg1_v = o.ComputeVerticalPolySeg()
                            polyseg_new = polyseg1.Linearize(t1, t2, t3, polyseg1_v, False)
                            # transform the chorded polyseg to the UCS
                            polyseg_new.Transform(matrixbackfromflat)

                            # restore the objects UCS
                            o.Normal = orgNormal


                        if polyseg_new != None:       # if that worked

                            if self.changeexisting.IsChecked:

                                if isinstance(o, Linestring):
                                    for i in reversed(range(1, o.GetElements().Count)):
                                        o.RemoveElementAt(i)
                                    l = o
                                else:
                                    self.error.Content += '\ncan\'t change non-Linestring object'
                                    l = None
                            else:
                                l = wv.Add(clr.GetClrType(Linestring))      # we start a new string line
                                createdserials.Add(l.SerialNumber)

                                SnapInAttributeExtension.CopyUserAttributes(o, l)

                            if l:
                                oname = IName.Name.__get__(o)
                                if oname == '':
                                    l.Name = l.Name + "chorded"
                                else:
                                    l.Name = oname + " - chorded"

                                if self.setlayer.IsChecked:
                                    l.Layer = self.layerpicker.SelectedSerialNumber
                                    try:
                                        l.Color = o.Color
                                        l.Weight = o.Weight
                                    except: pass
                                
                                if self.sourcelayer.IsChecked:
                                    l.Layer = o.Layer
                                    try:
                                        l.Color = o.Color
                                        l.Weight = o.Weight
                                    except: pass

                                if self.setprefixsuffix.IsChecked:
                                    inputlayer = self.currentProject.Concordance.Lookup(o.Layer)
                                    inputlayergroup = self.currentProject.Concordance.Lookup(inputlayer.LayerGroupSerial)
                                    outputlayer = Layer.FindOrCreateLayer(self.currentProject, self.prefix.Text + inputlayer.Name + self.suffix.Text)
                                    if inputlayergroup: # if the source layer is in a layer group
                                        # we check if the group exists, otherwise it is created
                                        outputlayergroup = lgc.FindOrCreateLayerGroup(self.prefix.Text + inputlayergroup.Name + self.suffix.Text)
                                        # we set the outputlayer group the the one we might just have created
                                        outputlayer.LayerGroupSerial = outputlayergroup.SerialNumber
                                    # setting the values for the layer itself
                                    outputlayer.DefaultColor = inputlayer.DefaultColor
                                    outputlayer.LineStyle = inputlayer.LineStyle
                                    outputlayer.LineWeight = inputlayer.LineWeight
                                    # in case the line settings are not ByLayer we have to set them as well
                                    l.Layer = outputlayer.SerialNumber
                                    try:
                                        l.Color = o.Color
                                        l.LineStyle = o.LineStyle
                                        l.Weight = o.Weight
                                    except: pass

                            
                                #for i in range(0, polyseg_new.NumberOfNodes):       # and add all nodes of the profile as new nodes of that string line
                                #    e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                                #    e.Position = polyseg_new[i].Point  # we draw that string line segment
                                #    l.AppendElement(e)
                       
                                l.Append(polyseg_new, None, False, False)
                                l.Color = o.Color

                        if self.deletesource.IsChecked and not self.changeexisting.IsChecked: # can't delete if we change existing

                            deletedserials.Add(o.SerialNumber)
                            osite = o.GetSite()    # we find out in which container the serial number reside
                            osite.Remove(o.SerialNumber)   # we delete the object from that container

                        j += 1
                        if (datetime.now() - time1).seconds > 0.5:
                            if ProgressBar.TBC_ProgressBar.SetProgress(j * 100 // selectionserials.Count):
                                break   # function returns true if user pressed cancel
                            time1 = datetime.now()

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
        

        # reinstate old selection
        ProgressBar.TBC_ProgressBar.Title = "reinstating selection"
        GlobalSelection.Items(self.currentProject).Set(list(set(selectionserials) - set(deletedserials)) + createdserials)

        self.success.Content += '\nDone'
        ProgressBar.TBC_ProgressBar.Title = ""
        
        wv.PauseGraphicsCache(False)


        self.SaveOptions()

    def debugxyz(self, centerp, vx,vy, nv):  
        
        wv = self.currentProject [Project.FixedSerial.WorldView] 
        
        l = wv.Add(clr.GetClrType(Linestring))
        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
        e.Position = centerp
        l.AppendElement(e)
        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
        e.Position = centerp + nv
        l.AppendElement(e)       
        l.Color = Color.Red
        l.Weight = 100
        #l.Layer = Layer.FindOrCreateLayer(self.currentProject, "Best-Fit Coordinate-System").SerialNumber
        l.Layer = self.layerpicker.SelectedSerialNumber
        
        l = wv.Add(clr.GetClrType(Linestring))
        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
        e.Position = centerp
        l.AppendElement(e)
        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
        e.Position = centerp + vx
        l.AppendElement(e)       
        l.Color = Color.Blue
        l.Weight = 100
        #l.Layer = Layer.FindOrCreateLayer(self.currentProject, "Best-Fit Coordinate-System").SerialNumber
        l.Layer = self.layerpicker.SelectedSerialNumber
        
        l = wv.Add(clr.GetClrType(Linestring))
        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
        e.Position = centerp
        l.AppendElement(e)
        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
        e.Position = centerp + vy
        l.AppendElement(e)       
        l.Color = Color.Green
        l.Weight = 100
        #l.Layer = Layer.FindOrCreateLayer(self.currentProject, "Best-Fit Coordinate-System").SerialNumber
        l.Layer = self.layerpicker.SelectedSerialNumber
        

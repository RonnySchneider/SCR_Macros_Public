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
    cmdData.Key = "SCR_PointBlock2Point"
    cmdData.CommandName = "SCR_PointBlock2Point"
    cmdData.Caption = "_SCR_PointBlock2Point"
    cmdData.UIForm = "SCR_PointBlock2Point"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Lines/Points"
        cmdData.DefaultTabGroupKey = "Points"
        cmdData.ShortCaption = "Point Block 2 Point"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.04
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Point Block 2 Point"
        cmdData.ToolTipTextFormatted = "Point Block 2 Point"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png") 
        cmdData.ImageSmall = b
    except:
        pass


class SCR_PointBlock2Point(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_PointBlock2Point.xaml") as s:
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

        optionMenu = SelectionContextMenuHandler()
        # remove options that don't apply here
        optionMenu.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        self.objs.ButtonContextMenu = optionMenu

        self.samplepoint.IsEntityValidCallback=self.IsValidPoint
        self.samplepointnumber.IsEntityValidCallback=self.IsValidText
        self.samplepointelevation.IsEntityValidCallback=self.IsValidText
        self.samplepointattribute.IsEntityValidCallback=self.IsValidText

        self.objs.IsEntityValidCallback = self.IsValid

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def SetDefaultOptions(self):

        # output-layer
        lserial = OptionsManager.GetUint("SCR_PointBlock2Point.layerpicker", 8)
        o = self.currentProject.Concordance.Lookup(lserial)
        if o != None:
            if isinstance(o.GetSite(), LayerCollection):    
                self.layerpicker.SetSelectedSerialNumber(lserial, InputMethod(3))
            else:                       
                self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:                       
            self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        
        self.prefix.Text = OptionsManager.GetString("SCR_PointBlock2Point.prefix", "")
        self.suffix.Text = OptionsManager.GetString("SCR_PointBlock2Point.suffix", "")
        self.newcode.Text = OptionsManager.GetString("SCR_PointBlock2Point.newcode", "")

        self.datainblocks.IsChecked = OptionsManager.GetBool("SCR_PointBlock2Point.datainblocks", True)
        # name-layer
        lserial = OptionsManager.GetUint("SCR_PointBlock2Point.namelayerpicker", 8)
        o = self.currentProject.Concordance.Lookup(lserial)
        if o != None:
            if isinstance(o.GetSite(), LayerCollection):    
                self.namelayerpicker.SetSelectedSerialNumber(lserial, InputMethod(3))
            else:                       
                self.namelayerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:                       
            self.namelayerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        # elev-layer
        lserial = OptionsManager.GetUint("SCR_PointBlock2Point.elevlayerpicker", 8)
        o = self.currentProject.Concordance.Lookup(lserial)
        if o != None:
            if isinstance(o.GetSite(), LayerCollection):    
                self.elevlayerpicker.SetSelectedSerialNumber(lserial, InputMethod(3))
            else:                       
                self.elevlayerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:                       
            self.elevlayerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        # attr-layer
        lserial = OptionsManager.GetUint("SCR_PointBlock2Point.attrlayerpicker", 8)
        o = self.currentProject.Concordance.Lookup(lserial)
        if o != None:
            if isinstance(o.GetSite(), LayerCollection):    
                self.attrlayerpicker.SetSelectedSerialNumber(lserial, InputMethod(3))
            else:                       
                self.attrlayerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:                       
            self.attrlayerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        
        self.dataintexts.IsChecked = OptionsManager.GetBool("SCR_PointBlock2Point.dataintexts", False)
        self.useelevationfromtext.IsChecked = OptionsManager.GetBool("SCR_PointBlock2Point.useelevationfromtext", True)

    def SaveOptions(self):
        OptionsManager.SetValue("SCR_PointBlock2Point.layerpicker", self.layerpicker.SelectedSerialNumber)
        OptionsManager.SetValue("SCR_PointBlock2Point.prefix", self.prefix.Text)
        OptionsManager.SetValue("SCR_PointBlock2Point.suffix", self.suffix.Text)
        OptionsManager.SetValue("SCR_PointBlock2Point.newcode", self.newcode.Text)
        OptionsManager.SetValue("SCR_PointBlock2Point.datainblocks", self.datainblocks.IsChecked)
        OptionsManager.SetValue("SCR_PointBlock2Point.namelayerpicker", self.namelayerpicker.SelectedSerialNumber)
        OptionsManager.SetValue("SCR_PointBlock2Point.elevlayerpicker", self.elevlayerpicker.SelectedSerialNumber)
        OptionsManager.SetValue("SCR_PointBlock2Point.attrlayerpicker", self.attrlayerpicker.SelectedSerialNumber)
        OptionsManager.SetValue("SCR_PointBlock2Point.dataintexts", self.dataintexts.IsChecked)
        OptionsManager.SetValue("SCR_PointBlock2Point.useelevationfromtext", self.useelevationfromtext.IsChecked)


    def IsValidPoint(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, CadPoint):
            return True
        return False

    def IsValidText(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, CadText):
            return True
        return False

    def IsValid(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, CadPoint):
            return True
        if isinstance(o, BlockReference):
            return True
        return False

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''

        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())

        wv = self.currentProject [Project.FixedSerial.WorldView]
       
        #find PointManager as object
        for o in self.currentProject:
            if isinstance(o, PointManager):
                pm = o

        prefix = self.prefix.Text
        suffix = self.suffix.Text
        newcode = self.newcode.Text

        try:
            # the "with" statement will unroll any changes if something go wrong
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
                # if data is inside blocks
                if self.datainblocks.IsChecked:
                    # go through all selected objects and double check that it is a block
                    for o in self.objs:
                        if isinstance(o, BlockReference):
                            # set the initial variables, in case we can't find a match in the attributes
                            pname = ""
                            pelev = o.InsertionPoint.Z
                            pattr = ""
                            # go through all the block attributes and try to match the layer
                            # if the layer matches get the values
                            for a in o.AttributeList:
                                attr = self.currentProject.Concordance.Lookup(a)
                                if attr.Layer == self.namelayerpicker.SelectedSerialNumber:
                                    pname = attr.TextString
                                if attr.Layer == self.elevlayerpicker.SelectedSerialNumber:
                                    try:
                                        pelev = self.lunits.Convert(self.lunits.DisplayType, float(attr.TextString), self.lunits.InternalType)
                                    except:
                                        pass
                                if attr.Layer == self.attrlayerpicker.SelectedSerialNumber:
                                    pattr = attr.TextString


                            # create a new internal point
                            pnew = Point3D()
                            pnew.X = o.InsertionPoint.X
                            pnew.Y = o.InsertionPoint.Y
                            pnew.Z = pelev

                            # create a new world point in the database
                            pnew_wv = CoordPoint.CreatePoint(self.currentProject, prefix + pname + suffix)
                            # change its FeatureCode by using the PointManager
                            if newcode == "":
                                pm.SetFeatureCodeAtPoint(pnew_wv.SerialNumber, pattr)
                            else:    
                                pm.SetFeatureCodeAtPoint(pnew_wv.SerialNumber, newcode)
                            # set its position and layer
                            pnew_wv.AddPosition(pnew)
                            pnew_wv.Layer = self.layerpicker.SelectedSerialNumber

                # data is in texts
                if self.dataintexts.IsChecked and self.samplepoint.Entity != None:
                    
                    # compute the deltas between the samples and the point
                    samplepoint = self.samplepoint.Entity
                    if self.samplepointnumber.Entity != None: # if a sample was selected
                        samplepointnumber = self.samplepointnumber.Entity
                        ptnumberdx = round(samplepointnumber.AlignmentPoint.X - samplepoint.Position.X, 6)
                        ptnumberdy = round(samplepointnumber.AlignmentPoint.Y - samplepoint.Position.Y, 6)
                    else:
                        ptnumberdx = 99999.
                        ptnumberdy = 99999.
                    
                    if self.useelevationfromtext.IsChecked:
                        if self.samplepointelevation.Entity != None: # if a sample was selected
                            samplepointelevation = self.samplepointelevation.Entity
                            ptelevdx = round(samplepointelevation.AlignmentPoint.X - samplepoint.Position.X, 6)
                            ptelevdy = round(samplepointelevation.AlignmentPoint.Y - samplepoint.Position.Y, 6)
                        else:
                            ptelevdx = 99999.
                            ptelevdy = 99999.
                    
                    if self.samplepointattribute.Entity != None: # if a sample was selected
                        samplepointattribute = self.samplepointattribute.Entity
                        ptattrdx = round(samplepointattribute.AlignmentPoint.X - samplepoint.Position.X, 6)
                        ptattrdy = round(samplepointattribute.AlignmentPoint.Y - samplepoint.Position.Y, 6)
                    else:
                        ptattrdx = 99999.
                        ptattrdy = 99999.
                    
                    for o in self.objs:
                        if isinstance(o, CadPoint):
                            p = o.Position
                            pname = "couldn't find match"
                            pelev = p.Z
                            pattr = ""
                            
                            # go through elements in the world view
                            for t in wv:
                                # if it's a text test if it has the same deltas as the sample
                                if isinstance(t, CadText):
                                    # round the values a bit, I had issue when data was imported as foot
                                    tdx = round(t.AlignmentPoint.X - p.X, 6)
                                    tdy = round(t.AlignmentPoint.Y - p.Y, 6)
                    
                                    if tdx == ptnumberdx and tdy == ptnumberdy:
                                        pname = t.TextString
                                        continue
                    
                                    if self.useelevationfromtext.IsChecked and tdx == ptelevdx and tdy == ptelevdy:
                                        try:
                                            pelev = self.lunits.Convert(self.lunits.DisplayType, float(t.TextString), self.lunits.InternalType)
                                        except:
                                            pass
                                        continue
                                    
                                    if tdx == ptattrdx and tdy == ptattrdy:
                                        pattr = t.TextString
                                        continue
                            
                            # create a new internal point
                            pnew = Point3D()
                            pnew.X = p.X
                            pnew.Y = p.Y
                            pnew.Z = pelev

                            # create a new world point in the database
                            pnew_wv = CoordPoint.CreatePoint(self.currentProject, prefix + pname + suffix)
                            # change its FeatureCode by using the PointManager
                            if newcode == "":
                                pm.SetFeatureCodeAtPoint(pnew_wv.SerialNumber, pattr)
                            else:    
                                pm.SetFeatureCodeAtPoint(pnew_wv.SerialNumber, newcode)
                            # set its position and layer
                            pnew_wv.AddPosition(pnew)
                            pnew_wv.Layer = self.layerpicker.SelectedSerialNumber

                # you MUST commit the change to DB or they will all be un-done (inside the "with" scope)
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

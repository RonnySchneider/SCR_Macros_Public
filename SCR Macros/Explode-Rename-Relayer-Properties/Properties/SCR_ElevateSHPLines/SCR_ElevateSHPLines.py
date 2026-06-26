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
    "layermode": True,
    "layerpickerin": 8,
    "manualmode": False,
    "featureattr": "Elevation",
    "layerpickerout": 8,
}


def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_ElevateSHPLines"
    cmdData.CommandName = "SCR_ElevateSHPLines"
    cmdData.Caption = "_SCR_ElevateSHPLines"
    cmdData.UIForm = "SCR_ElevateSHPLines"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Expld-SNR-Relay-Prop"
        cmdData.DefaultTabGroupKey = "Properties"
        cmdData.ShortCaption = "elevate SHP-Lines"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.08
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "elevate lines based on their feature attribute"
        cmdData.ToolTipTextFormatted = "elevate lines based on their feature attribute"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png") # we have to include a icon revision, otherwise TBC might not show the new one
        cmdData.ImageSmall = b
    except:
        pass


class SCR_ElevateSHPLines(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_ElevateSHPLines.xaml") as s:
            wpf.LoadComponent (self, s)
        self.currentProject = currentProject

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
        self.objs.IsEntityValidCallback = self.IsValid

        self.lType = clr.GetClrType(IPolyseg)
        #self.compType = clr.GetClrType(IPolyseg)

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear

        #self.lfp = self.lunits.Properties.Copy()
        self.linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation

        self.SetDefaultOptions()

    def SetDefaultOptions(self):
        SCROptions.LoadMacroOptions(self, "SCR_ElevateSHPLines", _OPTIONS, self.currentProject)

    def SaveOptions(self):
        SCROptions.SaveMacroOptions(self, "SCR_ElevateSHPLines", _OPTIONS)

    def IsValid(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        return False

    def OkClicked(self, cmd, e):

        Keyboard.Focus(self.okBtn)  # a trick to evaluate all input fields before execution, otherwise you'd have to click in another field first
        
        self.success.Content = ""
        self.error.Content = ""


        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)


        try:
            # the "with" statement will unroll any changes if something go wrong
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:


                #for o in self.currentProject:
                ##find FeatureManager as object
                #    if isinstance(o, FeatureManager):
                #        fm = o
                fm = FeatureManager.Provide(self.currentProject)
                
                serials=[]
                if self.layermode.IsChecked:
                    serials = self.currentProject.Concordance[self.layerpickerin.SelectedSerialNumber].Members

                elif self.manualmode.IsChecked:
                    serials = self.objs.SelectedSerials

                ProgressBar.TBC_ProgressBar.Title = "elevating " + str(serials.Count) + " lines" # set the progressbar description
                time1 = datetime.now()
                j = 0
                for sn in serials:
                    j += 1
                    if (datetime.now() - time1).seconds > 0.5:
                        if ProgressBar.TBC_ProgressBar.SetProgress(j * 100 // serials.Count):
                            break   # function returns true if user pressed cancel
                        time1 = datetime.now()

                    o = self.currentProject.Concordance[sn]
                    if isinstance(o, ICompositeGeometry):
                        foundelev = False

                        # need to jump throughs hoops since the featurecode is protected in composite geometry
                        # need to do it in reverse, go through all featurecodes and look if it's referring to the current object
                        for f in fm:
                            for e in self.currentProject.Concordance.GetObserversOf(f.SerialNumber):

                                if e == o:
                            
                                    fc = f
                                    break
                        if fc:
                            for attr in fc.Attributes:
                                if attr.Name == self.featureattr.Text:
                                    elev = self.lunits.Convert(self.lunits.DisplayType, float(attr.Value), self.lunits.InternalType)
                                    foundelev = True
                                    break

                            if foundelev:

                                for e in o:
                                    if isinstance(e.SnapIn, self.lType):
                                        self.drawline(e.SnapIn, elev)

                            else:
                                o.Color = Color.Red

                        
                    elif isinstance(o, self.lType):

                        # in case the user has selected a combination of linestrings and composites
                        # we'd color linestrings red if we wouldn't check if they are part of a composite
                        # only the composite contains the attribute
                        partofcomposite = False
                        for observedby in self.currentProject.Concordance.GetObserversOf(o.SerialNumber):
                            if isinstance(observedby, ICompositeGeometry):
                                partofcomposite = True
                                break

                        if not partofcomposite:
                            foundelev = False
                            # get the line feature code
                            for observes in self.currentProject.Concordance.GetIsObservedBy(o.SerialNumber):
                                if observes and isinstance(observes, LineFeature):
                                    for attr in observes.Attributes:
                                        if attr.Name == self.featureattr.Text:
                                            elev = self.lunits.Convert(self.lunits.DisplayType, float(attr.Value), self.lunits.InternalType)
                                            foundelev = True
                                            break

                            if foundelev:

                                self.drawline(o, elev)

                            else:
                                o.Color = Color.Red

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


                    #self.success.Content += '\n' + o.GetType().Name + ' - SN#: ' + str(o.SerialNumber)
                    #tt = o.GetSite()
                    #tt2 = o

        self.SaveOptions()

    def drawline(self, l, elev):

        wv = self.currentProject [Project.FixedSerial.WorldView]

        polyseg = l.ComputePolySeg()
        polyseg = polyseg.ToWorld()
        polyseg_v = PolySeg.PolySeg()
        polyseg_v.Add(Point3D(0, elev, 0))
        
        ls = wv.Add(clr.GetClrType(Linestring))
        ls.Append(polyseg, polyseg_v, False, False)
        ls.Layer = self.layerpickerout.SelectedSerialNumber

        return


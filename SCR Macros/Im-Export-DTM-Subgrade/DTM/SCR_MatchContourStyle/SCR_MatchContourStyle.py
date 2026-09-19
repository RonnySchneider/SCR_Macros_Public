#   GNU GPLv3
#   <this is an add-on Script/Macro for the geospatial software "Trimble Business Center" aka TBC>
#   <you'll need at least the "Survey Advanced" licence of TBC in order to run this script>
#	<see the ToolTip section below for a brief explanation what the script does>
#	<see the Help-Files for more details>
#   Copyright (C) 2026 Ronny Schneider
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
    "contourinterval":        True,
    "indexfrequency":         True,
    "labelalllines":          True,
    "normalcolour":           True,
    "normallineweight":       True,
    "indexcolour":            True,
    "indexlineweight":        True,
    "colourbyelevation":      True,
    "rebuildmethod":          True,
    "textstyle":              True,
    "distancebetweenlabels":  True,
    "minimumlength":          True,
    "minimumarea":            True,
    "labelends":              True,
    "smoothcontours":         True,
}

def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_MatchContourStyle"
    cmdData.CommandName = "SCR_MatchContourStyle"
    cmdData.Caption = "_SCR_MatchContourStyle"
    cmdData.UIForm = "SCR_MatchContourStyle"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "DTM"
        cmdData.ShortCaption = "Match Contour Style"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.035
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""

        cmdData.ToolTipTitle = "match the contour representation style"
        cmdData.ToolTipTextFormatted = "match the contour representation style"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_MatchContourStyle(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_MatchContourStyle.xaml") as s:
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

        self.selectedSource = None
        self.sourceTree.SelectedItemChanged += self.SourceTreeSelectionChanged
        # the hidden helper picker needs its WinForms handle created (host attached to the visual
        # tree, layout run) before EntitySerialNumbers is reliably populated - defer past that
        self.Dispatcher.BeginInvoke(DispatcherPriority.Loaded, Action(self.BuildSourceTree))

        self.objs.IsEntityValidCallback = self.IsValidContour
        self.substitutingTargetSelection = False
        self.objs.ValueChanged += self.ObjsValueChanged

        self.SetDefaultOptions()

    def ObjsValueChanged(self, sender, e):
        # a click in the view/Explorer adds the contour's CAD polyline/linestring, not the
        # Model3DContours entity itself - swap the raw pick(s) for their resolved/deduped parent
        # contour(s) live, so the list actually shows what will be matched. With only one
        # MemberSelection in this dialog now (no more focus-arbitration between two of them, which
        # is what made an earlier attempt at this unreliable) this can run synchronously and safely.
        if self.substitutingTargetSelection:
            return
        try:
            contours = self.ResolveContours(self.objs)
            wanted = list(contours.keys())
            current = [o.SerialNumber for o in self.objs]
            if sorted(wanted) != sorted(current):
                self.substitutingTargetSelection = True
                try:
                    if wanted:
                        self.objs.SetSelection(self.currentProject, Array[UInt32](wanted))
                    else:
                        self.objs.ResetSelection()
                finally:
                    self.substitutingTargetSelection = False
        except:
            pass    # cosmetic only - OkClicked always re-resolves from scratch regardless

    def BuildSourceTree(self):
        try:
            self._BuildSourceTree()
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            self.error.Content += '\ncould not build the Source Contours list\n' + str(exc_type) + '\n' + str(exc_obj) + '\nLine ' + str(exc_tb.tb_lineno)

    def _BuildSourceTree(self):
        # ComboBoxEntityPicker can't show more than the entity's own name for each dropdown row -
        # several Contours entities in a project routinely share the same name (e.g. every surface's
        # default "Contours"), so we build our own tree instead: one bold node per Surface, with its
        # Contours entities listed underneath, indented, so they're actually distinguishable
        picker, innerlist = SCREntityPicker.create(self.sourcehelperhost, Project.FixedSerial.WorldView,
                                                    clr.GetClrType(Model3DContours), multi_select=False)

        bysurface = {}
        for kvp in picker.EntitySerialNumbers:
            contour = self.currentProject.Concordance.Lookup(kvp.Value)
            if not isinstance(contour, Model3DContours):
                continue
            surface = contour.Surface
            surfacename = surface.Name if surface is not None else '(no surface)'
            bysurface.setdefault(surfacename, []).append(contour)

        self.sourceTree.Items.Clear()
        for surfacename in sorted(bysurface.keys()):
            surfacenode = TreeViewItem()
            header = TextBlock()
            header.Text = surfacename
            header.FontWeight = FontWeights.Bold
            # block selection on the label itself only - not on the row's expand/collapse toggle
            header.PreviewMouseLeftButtonDown += self.RejectSurfaceNodeClick
            surfacenode.Header = header
            surfacenode.IsExpanded = True
            surfacenode.Focusable = False    # a Surface header is a grouping label only, not a pickable item
            self.sourceTree.Items.Add(surfacenode)

            for contour in sorted(bysurface[surfacename], key=lambda c: c.Name or ''):
                contournode = TreeViewItem()
                contournode.Header = contour.Name if contour.Name else '(unnamed)'
                contournode.FontWeight = FontWeights.Normal
                contournode.Tag = contour
                surfacenode.Items.Add(contournode)

    def RejectSurfaceNodeClick(self, sender, e):
        # stop the click from reaching TreeView's own selection handling at all - toggling
        # IsSelected reactively from SelectedItemChanged instead crashes TBC (a known WPF
        # TreeView bug in HandleSelectionAndCollapsed), so the click must never be allowed through
        e.Handled = True

    def SourceTreeSelectionChanged(self, sender, e):
        selected = self.sourceTree.SelectedItem
        contour = selected.Tag if selected is not None else None
        self.selectedSource = contour if isinstance(contour, Model3DContours) else None

    def ResolveContour(self, o):
        # a click in the view usually hits a contour's CAD polyline/linestring rather than
        # the Model3DContours entity itself - walk up via GetSite() to find its parent contour
        if o is None:
            return None
        if isinstance(o, Model3DContours):
            return o
        try:
            site = o.GetSite()
        except:
            return None
        if isinstance(site, Model3DContours):
            return site
        return None

    def IsValidContour(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        return self.ResolveContour(o) is not None

    def ResolveContours(self, picker):
        # resolve every picked entity to its owning contour and dedupe by serial - clicking in
        # the view will usually pick individual contour lines rather than the Model3DContours itself
        result = {}
        for entity in picker:
            contour = self.ResolveContour(entity)
            if contour is not None:
                result[contour.SerialNumber] = contour
        return result

    def SetDefaultOptions(self):
        SCROptions.LoadMacroOptions(self, "SCR_MatchContourStyle", _OPTIONS, self.currentProject)

    def SaveOptions(self):
        SCROptions.SaveMacroOptions(self, "SCR_MatchContourStyle", _OPTIONS)

    def CancelClicked(self, thisCmd, args):
        thisCmd.CloseUICommand ()

    def OkClicked(self, thisCmd, e):
        self.error.Content = ''
        self.success.Content = ''

        source = self.selectedSource

        if source is None or not isinstance(source, Model3DContours):
            self.error.Content = '\nplease select a Source Contours entity (a Surface node by itself does not count)'
            return

        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)

        count = 0

        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                targets = self.ResolveContours(self.objs)
                targets.pop(source.SerialNumber, None)

                for target in targets.values():

                    if self.contourinterval.IsChecked:
                        target.ElevationInterval = source.ElevationInterval
                    if self.indexfrequency.IsChecked:
                        target.MajorContourFrequency = source.MajorContourFrequency
                    if self.labelalllines.IsChecked:
                        target.LabelAllContours = source.LabelAllContours
                    if self.normalcolour.IsChecked:
                        target.MinorContoursColor = source.MinorContoursColor
                    if self.normallineweight.IsChecked:
                        target.MinorContoursWeight = source.MinorContoursWeight
                    if self.indexcolour.IsChecked:
                        target.MajorContoursColor = source.MajorContoursColor
                    if self.indexlineweight.IsChecked:
                        target.MajorContoursWeight = source.MajorContoursWeight
                    if self.colourbyelevation.IsChecked:
                        target.ColorByElevation = source.ColorByElevation
                    if self.rebuildmethod.IsChecked:
                        target.RebuildMethod = source.RebuildMethod
                    if self.textstyle.IsChecked:
                        target.TextStyleSerial = source.TextStyleSerial
                    if self.distancebetweenlabels.IsChecked:
                        target.LabelContourDistance = source.LabelContourDistance
                    if self.minimumlength.IsChecked:
                        target.MinLength = source.MinLength
                    if self.minimumarea.IsChecked:
                        target.MinArea = source.MinArea
                    if self.labelends.IsChecked:
                        target.LabelContourEnds = source.LabelContourEnds
                    if self.smoothcontours.IsChecked:
                        target.SmoothContours = source.SmoothContours

                    target.BuildContours()    # unconditional rebuild - RebuildSnapIn() skips it for cosmetic-only changes or manual RebuildMethod

                    count += 1

                failGuard.Commit()
                self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
                UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())

                self.success.Content = '\nmatched style on ' + str(count) + ' target contour(s)'

        except Exception as e:
            tt = sys.exc_info()
            exc_type, exc_obj, exc_tb = sys.exc_info()
            # EndMark MUST be set no matter what
            # otherwise TBC won't work anymore and needs to be restarted
            self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
            UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
            self.error.Content += '\nan Error occurred - Result probably incomplete\n' + str(exc_type) + '\n' + str(exc_obj) + '\nLine ' + str(exc_tb.tb_lineno)

        self.SaveOptions()

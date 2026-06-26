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
    "createnewtext": False,
    "findattr": "{lot}\\n{plan}",
    "changeexistingtext": True,
    "namemode_layer": False,
    "namemode_line": True,
    "namemode_original": False,
    "checkBox_relayer": False,
    "optimizetextsize": False,
    "optimizetextposition": False,
    "textBox1": "",
    "textBox2": "",
}

def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_ChangeTextToPolyname"
    cmdData.CommandName = "SCR_ChangeTextToPolyname"
    cmdData.Caption = "_SCR_ChangeTextToPolyname"
    cmdData.UIForm = "SCR_ChangeTextToPolyname"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Expld-SNR-Relay-Prop"
        cmdData.DefaultTabGroupKey = "Renaming"
        cmdData.ShortCaption = "Text from Polyname"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.07
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Texts get Polylinename"
        cmdData.ToolTipTextFormatted = "all Texts inside a Polyline get changed to the name of the Polyline"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_ChangeTextToPolyname(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_ChangeTextToPolyname.xaml") as s:
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
        self.objs.IsEntityValidCallback = self.IsValid
        optionMenu = SelectionContextMenuHandler()
        # remove options that don't apply here
        optionMenu.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        self.objs.ButtonContextMenu = optionMenu
        #self.poly3dType = clr.GetClrType(Poly3D)
        #self.polylineType = clr.GetClrType(PolyLine)
        #self.linestringType = clr.GetClrType(Linestring)
        #self.circleType = clr.GetClrType(CadCircle)
        self.lType = clr.GetClrType(IPolyseg)

        SCRExpanders.wire_pairs([
            (self.expander_createnewtext, self.createnewtext),
            (self.expander_changeexistingtext, self.changeexistingtext),
        ])
        try:
            self.SetDefaultOptions()
        except:
            pass

    def SetDefaultOptions(self):
        SCROptions.LoadProjectOptions(self, "SCR_ChangeTextToPolyname", _OPTIONS, self.currentProject)
        settingserial = OptionsManager.GetUint("SCR_ChangeTextToPolyname.textstylepicker", 24)
        o = self.currentProject.Concordance[settingserial]
        if o is None or not isinstance(o.GetSite(), TextStyleCollection):
            settingserial = 24
        self.textstylepicker.SetSelectedSerialNumber(settingserial, InputMethod(3))

    def SaveOptions(self):
        SCROptions.SaveProjectOptions(self, "SCR_ChangeTextToPolyname", _OPTIONS, self.currentProject)
        OptionsManager.SetValue("SCR_ChangeTextToPolyname.textstylepicker", self.textstylepicker.SelectedSerialNumber)

    def IsValid(self, serial):
        o = self.currentProject.Concordance[serial]
        #if isinstance(o, self.poly3dType):
        #    return True
        #if isinstance(o, self.polylineType):
        #    return True
        #if isinstance(o, self.linestringType):
        #    return True
        #if isinstance(o, self.circleType):
        #    return True
        if isinstance(o, self.lType):
            return True
        return False
        

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()


    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Text = ''
        self.success.Content = ''

        if not self.createnewtext.IsChecked and not self.changeexistingtext.IsChecked:
            self.error.Text = 'select an operation first'
            return

        wv = self.currentProject[Project.FixedSerial.WorldView] # Worldview
        activeview = TrimbleOffice.TheOffice.MainWindow.AppViewManager.ActiveView
        activeviewfilter = self.currentProject.Concordance[activeview.ViewFilter]
        self.fm = FeatureManager.Provide(self.currentProject)
        self.fcm = FeatureCodeManager.Provide(self.currentProject)            
        pm = RawDataContainer.ProvideRawDataContainer(self.currentProject).PointManager

        j = 0
        time1 = datetime.now()
        ProgressBar.TBC_ProgressBar.Title = "Creating Texts and optimizing Size/Position"

        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(Client.CommandGranularity.Command, self.Caption)

        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                for o in self.objs.SelectedMembers(self.currentProject): # we go through all the selected Lines

                    j += 1
                    if (datetime.now() - time1).seconds > 0.5:
                        if ProgressBar.TBC_ProgressBar.SetProgress(j * 100 // self.objs.Count):
                            self.fullbreak = True
                            break   # function returns true if user pressed cancel
                        time1 = datetime.now()

                    if isinstance(o, self.lType) or isinstance(o, ICompositeGeometry): # if they have the right type then continue


                        
                        if self.createnewtext.IsChecked:

                            polyseg = o.ComputePolySeg()

                            if polyseg.IsClosed:

                                attrs = self.getattributes(o)
                                newtext = self.resolve_template(self.findattr.Text, attrs)
                                
                                if newtext != "":
                                    t = wv.Add(clr.GetClrType(MText))
                                    t.AlignmentPoint = self.find_best_placement(polyseg)
                                    t.AttachPoint = AttachmentPoint.MiddleMid
                                    t.TextStyleSerial = self.textstylepicker.SelectedSerialNumber
                                    t.Height = 1
                                    t.Layer = o.Layer
                                    t.TextString = newtext   

                                    self.optimize_text_height(t, polyseg)

                        elif self.changeexistingtext.IsChecked:

                            if self.namemode_layer.IsChecked:
                                linename = self.currentProject.Concordance[o.Layer].Name
                            elif self.namemode_line.IsChecked:
                                linename = IName.Name.__get__(o)
                            else:
                                linename = None  # resolved per text object below

                            polyseg = o.ComputePolySeg()      # we convert the line object to a more unified polyseg, for which we have nifty functions available

                            if polyseg.IsClosed:            # if the line is closed then continue

                                for o2 in wv:               # we go through all drawing objects

                                    if isinstance(o2, CadText) or isinstance(o2, MText):      # in case it is some kind of text we continue

                                        if activeviewfilter.LayerOverrides.Contains(o2.Layer) and not activeviewfilter.LayerOverrides[o2.Layer].Visible:
                                            continue

                                        if polyseg.PointInPolyseg(o2.AlignmentPoint) == Side.In:      # we compute if the text anchor point is inside the polyseg
                                            name = o2.TextString if linename is None else linename
                                            o2.TextString = self.textBox1.Text + name + self.textBox2.Text
                                            if self.checkBox_relayer.IsChecked == True: # if the checkbox is checked we also change the layer to that of the line
                                                o2.Layer = o.Layer
                                            if self.optimizetextsize.IsChecked and isinstance(o2, MText):
                                                self.optimize_text_height(o2, polyseg)
                                            if self.optimizetextposition.IsChecked:
                                                o2.AlignmentPoint = self.find_best_placement(polyseg)
                                                try:
                                                    o2.AttachPoint = AttachmentPoint.MiddleMid
                                                except:
                                                    pass

                                    if isinstance(o2, AlignmentLabel):      # in case of aligment label text we have to unwrap it first

                                        if activeviewfilter.LayerOverrides.Contains(o2.Layer) and not activeviewfilter.LayerOverrides[o2.Layer].Visible:
                                            continue

                                        labelseriallist = o2.ContainedSerials # gets us a list with all the text serial numbers
                                        for i in range(0, o2.Count):         # we count through the serial numbers
                                            labelserial = labelseriallist[i]  # we get us one serial number
                                            if isinstance(o2[labelserial], CadText): #the label container can also contain lines, so we have to make sure we work on a text
                                                if polyseg.PointInPolyseg(o2[labelserial].AlignmentPoint) == Side.In:     # we compute if the text anchor point is inside the polyseg
                                                    name = o2[labelserial].TextString if linename is None else linename
                                                    o2[labelserial].TextString = self.textBox1.Text + name + self.textBox2.Text
                                                    if self.checkBox_relayer.IsChecked == True: # if the checkbox is checked we also change the layer to that of the line
                                                        o2[labelserial].Layer = o.Layer

                if self.objs.SelectedMembers(self.currentProject).Count == 0:
                    self.success.Content += 'nothing selected'
                else:
                    self.success.Content = 'Success'

                failGuard.Commit()

        except Exception as _err:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            self.error.Text += '\nan Error occurred - Result probably incomplete\n' + str(exc_type) + '\n' + str(exc_obj) + '\nLine ' + str(exc_tb.tb_lineno)

        finally:
            self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
            UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())



        ProgressBar.TBC_ProgressBar.Title = ""
        self.SaveOptions()


    def resolve_template(self, template, attrs):
        import re
        text = template.replace('\\n', '\n')
        # split into alternating [literal, {placeholder}, literal, {placeholder}, ..., literal]
        parts = re.split(r'(\{[^}]+\})', text)

        segments = []
        any_found = False
        for i in range(0, len(parts) - 1, 2):
            literal = parts[i]
            key = parts[i + 1][1:-1].strip(': ').lower()  # strip { }, whitespace and colons
            if key in attrs:
                value = attrs[key].strip()
                if not value:
                    return ""   # attribute exists but is empty — skip creating text
                segments.append(literal + value)
                any_found = True
            # if key not in attrs: skip both the separator text and the placeholder

        if not any_found:
            return ""

        if parts[-1]:   # trailing literal after the last placeholder
            segments.append(parts[-1])

        return ''.join(segments).strip()

    def getattributes(self, o):
        attrs = {}

        if isinstance(o, ICompositeGeometry):
            for observes in self.currentProject.Concordance.GetIsObservedBy(o.SerialNumber):
                if observes and isinstance(observes, Feature):
                    tt = observes.Attributes
            # need to jump through hoops since the featurecode is protected in composite geometry
            # need to do it in reverse, go through all featurecodes and look if it's referring to the current object
            fc = None
            for f in self.fm:
                for e in self.currentProject.Concordance.GetObserversOf(f.SerialNumber):
                    if e == o:
                        fc = f
                        break
                if fc:
                    break
            if fc:
                for attr in fc.Attributes:
                    attrs[attr.Name.lower()] = str(attr.Value).strip()

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

                # get the line feature code
                for observes in self.currentProject.Concordance.GetIsObservedBy(o.SerialNumber):
                    if observes and isinstance(observes, Feature):
                        for attr in observes.Attributes:
                            attrs[attr.Name.lower()] = str(attr.Value).strip()

        return attrs

    def text_fits_in_polyseg(self, text_obj, polyseg):
        for corner in text_obj.GetTextBoxBounds():
            if polyseg.PointInPolyseg(corner) != Side.In:
                return False
        return True
    
    def optimize_text_height(self, text_obj, polyseg, tolerance=0.05):
        bb = polyseg.BoundingBox
        max_h = min(bb.ptMax.X - bb.ptMin.X, bb.ptMax.Y - bb.ptMin.Y)
        min_h = 0.01
    
        text_obj.Height = min_h
        if not self.text_fits_in_polyseg(text_obj, polyseg):
            return False  # won't fit at any size
    
        lo, hi = min_h, max_h
        while hi - lo > tolerance:
            mid = (lo + hi) / 2.0
            text_obj.Height = mid
            if self.text_fits_in_polyseg(text_obj, polyseg):
                lo = mid
            else:
                hi = mid
    
        text_obj.Height = lo
        return True

    def find_best_placement(self, polyseg):
        """For quadrilaterals use diagonal intersection; otherwise polylabel."""
        verts = list(polyseg.Point3Ds())
        # remove duplicate closing vertex
        if len(verts) > 1 and abs(verts[0].X - verts[-1].X) < 1e-9 and abs(verts[0].Y - verts[-1].Y) < 1e-9:
            verts = verts[:-1]
        if len(verts) == 4:
            pt = self._quad_diagonal_center(verts)
            if pt is not None and polyseg.PointInPolyseg(pt) == Side.In:
                return pt
        return self._polylabel(polyseg)

    def _quad_diagonal_center(self, verts):
        v0, v1, v2, v3 = verts
        dx1 = v2.X - v0.X; dy1 = v2.Y - v0.Y
        dx2 = v3.X - v1.X; dy2 = v3.Y - v1.Y
        denom = dx1 * dy2 - dy1 * dx2
        if abs(denom) < 1e-12:
            return None
        t = ((v1.X - v0.X) * dy2 - (v1.Y - v0.Y) * dx2) / denom
        return Point3D(v0.X + t * dx1, v0.Y + t * dy1, 0)

    def _polylabel(self, polyseg):
        import heapq
        SQRT2 = math.sqrt(2)

        bb = polyseg.BoundingBox
        w = bb.ptMax.X - bb.ptMin.X
        h = bb.ptMax.Y - bb.ptMin.Y
        cell_size = min(w, h)
        if cell_size == 0:
            return Point3D(bb.ptMin.X, bb.ptMin.Y, 0)

        precision = cell_size / 50.0

        def signed_dist(cx, cy):
            p = Point3D(cx, cy, 0)
            inside = polyseg.PointInPolyseg(p) == Side.In
            (found, closest, sta) = polyseg.FindPointFromPoint(p)
            if not found:
                return 0.0
            d = math.sqrt((cx - closest.X)**2 + (cy - closest.Y)**2)
            return d if inside else -d

        # seed best with centroid
        (area, cgrav, side) = polyseg.Area()
        best_x = cgrav.X if not math.isnan(cgrav.X) else bb.ptMin.X + w / 2.0
        best_y = cgrav.Y if not math.isnan(cgrav.Y) else bb.ptMin.Y + h / 2.0
        best_d = signed_dist(best_x, best_y)

        # initial grid of square cells sized to the narrow dimension
        queue = []
        half = cell_size / 2.0
        x = bb.ptMin.X + half
        while x < bb.ptMax.X:
            y = bb.ptMin.Y + half
            while y < bb.ptMax.Y:
                d = signed_dist(x, y)
                max_d = d + half * SQRT2
                if max_d > best_d:
                    heapq.heappush(queue, (-max_d, x, y, half, d))
                if d > best_d:
                    best_d, best_x, best_y = d, x, y
                y += cell_size
            x += cell_size

        while queue:
            neg_max_d, cx, cy, ch, d = heapq.heappop(queue)

            if d > best_d:
                best_d, best_x, best_y = d, cx, cy

            if -neg_max_d - best_d <= precision:
                continue

            nh = ch / 2.0
            for ox, oy in ((-nh, -nh), (nh, -nh), (-nh, nh), (nh, nh)):
                nd = signed_dist(cx + ox, cy + oy)
                nmax = nd + nh * SQRT2
                if nmax > best_d:
                    heapq.heappush(queue, (-nmax, cx + ox, cy + oy, nh, nd))

        return Point3D(best_x, best_y, 0)
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
    "deletesource": False,
    "changetextlayer": False,
    "newtextlayerpicker": 8,
    "changetextcolor": False,
    "textcolorpicker": Color.Red,
    "changetextfont": False,
    "fontpicker": "tmodelf.fnt",
    "changetextheight": False,
    "textheightdist": 1.0,
    "changetextweight": False,
    "textweightpicker": 0,
}

def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_ExplodeTextToLines"
    cmdData.CommandName = "SCR_ExplodeTextToLines"
    cmdData.Caption = "_SCR_ExplodeTextToLines"
    cmdData.UIForm = "SCR_ExplodeTextToLines"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Expld-SNR-Relay-Prop"
        cmdData.DefaultTabGroupKey = "Explode"
        cmdData.ShortCaption = "Explode Text to Linework"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.20
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Explode Text"
        cmdData.ToolTipTextFormatted = "Explode Text to Linework"

    except:
        pass
    
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_ExplodeTextToLines(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_ExplodeTextToLines.xaml") as s:
            wpf.LoadComponent (self, s)
        self.currentProject = currentProject
        self.macroFileFolder = macroFileFolder

    def HelpClicked(self, cmd, e):
        webbrowser.open("C:\ProgramData\Trimble\MacroCommands3\SCR Macros\MacroHelp\MacroHelp.htm#" + type(self).__name__)

    def IsValid(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.mtextType):
            return True
        if isinstance(o, self.cadtextType):
            return True
        return False

    def OnLoad(self, cmd, buttons, event):
        self.okBtn = buttons[0]
        buttons[2].Content = "Help"
        buttons[2].Visibility = Visibility.Visible
        buttons[2].Click += self.HelpClicked
        self.Caption = cmd.Command.Caption

        self.objs.IsEntityValidCallback=self.IsValid
        optionMenu = SelectionContextMenuHandler()
        # remove options that don't apply here
        optionMenu.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        self.objs.ButtonContextMenu = optionMenu
        self.mtextType = clr.GetClrType(MText)
        self.cadtextType = clr.GetClrType(CadText)
        
        fontlist = [x for x in os.listdir("C:\\ProgramData\\Trimble\\Fonts") if x.endswith(".fnt")]
        if fontlist.Count > 0:
            for f in fontlist:
                # testload font
                cachedfont = StrokeFontManager.LoadFontFile(f)
                if math.isnan(cachedfont) == False:
                    if cachedfont.Characters.Count > 0:
                        item = ComboBoxItem()
                        item.Content = f
                        item.FontSize = 12
                        self.fontpicker.Items.Add(item)

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        #self.lfp = self.lunits.Properties.Copy()
        self.linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation

        self.changetextheight.Content = 'Change Text Height [' + self.linearsuffix + ']'

            
		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()
    
    # code for enabling disabling the pickers can found below the main loop


    # main loop
    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''


        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        
        wv = self.currentProject [Project.FixedSerial.WorldView]
        wv.PauseGraphicsCache(True)

        ProgressBar.TBC_ProgressBar.Title = self.Caption
        
        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                # get/set font - we don't want to load it multiple times, so we get it outside the loop
                if self.changetextfont.IsChecked and self.fontpicker.Items.Count > 0:
                    cachedfont = StrokeFontManager.LoadFontFile(self.fontpicker.SelectedItem.Content)
                else:
                    cachedfont = StrokeFontManager.LoadFontFile("tmodelf.fnt")
                
                # objs_count = self.objs.Count # since we might delete the texts in the loop the count would change and mess up the percentage    
                
                objlist = []
                for o in self.objs.SelectedMembers(self.currentProject):
                    objlist.Add(o.SerialNumber)
                GlobalSelection.Clear()
                
                j = 0
                time1 = datetime.now()
                for sn in objlist:
                    texto = self.currentProject.Concordance.Lookup(sn)
                    container = texto.GetSite()

                    j += 1
                    if (datetime.now() - time1).seconds > 0.2:
                        if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(j * 100 / objlist.Count)):
                            break   # function returns true if user pressed cancel
                        time1 = datetime.now()

                    if isinstance(texto, CadLabel):
                        self.success.Content ='found Labels - explode them with CAD-Explode first'
                        #texto.ExplodeAbsolutely(wv)


                    # check if it's a text
                    if (isinstance(texto, CadText) or isinstance(texto, MText)):
                        
                        textlines = texto.GetTextSegments   # get an array of the texlines, makes it easier for multiline
                        
                        ### experimenting with text to number parsing
                        ### tt = texto.TextString
                        ### #tt2 = TextUtilities.ParseElevationText(texto.TextString, self.currentProject)
                        ### if not TextUtilities.XTextContainsFormatting(texto.TextString):
                        ###     el1 = TextUtilities.ParseElevationText(texto.TextString, self.currentProject)
                        ### else:
                        ###     # standard multi-line text is also considered a XText
                        ###     # first strip the codes
                        ###     tt = TextUtilities.GetPlainStringFromXText(texto.TextString, None)
                        ###     # parse the remaining string into a number
                        ###     el2 = TextUtilities.ParseElevationText(tt, self.currentProject)
                        
                        # get /set Layer
                        if self.changetextlayer.IsChecked:
                            newtextlayer = self.newtextlayerpicker.SelectedSerialNumber
                        else:
                            newtextlayer = texto.Layer
                        # get/set color
                        if self.changetextcolor.IsChecked:
                            newtextcolor = self.textcolorpicker.SelectedColor
                        else:
                            newtextcolor = texto.Color
                        # get/set height
                        if self.changetextheight.IsChecked:
                            try: newtextheight = self.textheightdist.Distance
                            except: newtextheight = texto.Height
                        else:
                            newtextheight = texto.Height
                        # get/set weight
                        if self.changetextweight.IsChecked:
                            newtextweight = self.textweightpicker.Lineweight
                        else:
                            newtextweight = texto.Weight
                        
                        polysegs = List[PolySeg.PolySeg]()
                        polyseg = None
                        polysegnodes = List[Point3D]()

                        # When the font changes, InsertPoint is pre-computed for the old font's
                        # metrics. Non-left-baseline texts (center, middle-center, etc.) shift
                        # because the InsertPoint encodes an offset proportional to old glyph width.
                        # Compute a correction so all segments stay anchored to AlignmentPoint.
                        insert_corrections = None
                        insert_corrections = self._compute_insert_correction(texto, textlines, cachedfont, newtextheight)

                        # go through all the single lines
                        seg_idx = 0
                        for singletext in textlines:

                            # singletext -> textnodelist
                            # textnodelist -> polysegnodes
                            # polysegnodes -> polysegs
                            # draw -> polysegs
                            # explode each textline into a textnodelist (which can be scrambled with multi element gaps and unnecessary 1 element entries)
                            # parse the textnodelist and add elements to polysegnodes
                            # if we have at least 2 consecutive nodes create a new polyseg and add it to polysegs
                            # draw the polysegs

                            corr = insert_corrections[seg_idx] if (insert_corrections is not None and seg_idx < len(insert_corrections)) else None
                            if corr is not None:
                                use_insert = Point3D(singletext.InsertPoint.X + corr[0],
                                                     singletext.InsertPoint.Y + corr[1],
                                                     singletext.InsertPoint.Z)
                            else:
                                use_insert = singletext.InsertPoint
                            seg_idx += 1

                            textnodelist = cachedfont.DrawText(use_insert, newtextheight, singletext.WidthFactor, singletext.ObliqueAngle, singletext.RotateAngle, singletext.TextString)
                            tt = singletext.TextString

                            for i in range (0, textnodelist.Count):

                                # clean up if we reached the end of the list                                
                                if i == textnodelist.Count - 1 and polysegnodes.Count > 1:
                                        polyseg = PolySeg.PolySeg()
                                        polyseg.Add(polysegnodes.ToArray())
                                        polysegs.Add(polyseg.Clone())
                                        continue
                                
                                # in case of a gap
                                if textnodelist[i].IsUndefined:
                                    # if we have at least 2 nodes we create a new line
                                    if polysegnodes.Count >= 2:
                                        polyseg = PolySeg.PolySeg()
                                        polyseg.Add(polysegnodes.ToArray())
                                        polysegs.Add(polyseg.Clone())

                                        polysegnodes.Clear()
                                        continue
                                    else: # clean up, we don't want those single coordinates from textnodelist
                                        polysegnodes.Clear()
                                
                                # no gap, but a valid point
                                else:
                                    # can't recall if it always was like this
                                    # lately the textnodelist comes back at elevation 0
                                    ip = textnodelist[i]
                                    ip.Z = singletext.InsertPoint.Z
                                    polysegnodes.Add(ip)
                                    continue
                            
                            # use build in function to combine the segments as much as possible, spares us to do a manual Project-Cleanup
                            ttcount = PolySeg.PolySeg.JoinTouchingPolysegs(polysegs)

                            # draw the lines
                            for p in polysegs:
                                if p and p.NumberOfNodes > 1: # final double check that we don't create a single node line
                                    self.CreateLinestring(p, container, newtextlayer, newtextcolor, newtextweight)
                            
                            # cleanup the arrays, otherwise it could happen that we drag unwanted stuff into the next line, if it is a multiline text
                            textnodelist.Clear()
                            polysegnodes.Clear()
                            polysegs.Clear()

                        # delete the source-text if ticked
                        if self.deletesource.IsChecked:
                            tt = container.Remove(texto.SerialNumber)
                    
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

        ProgressBar.TBC_ProgressBar.Title = ""
        self.success.Content += '\nDone'
        self.SaveOptions()           

        wv.PauseGraphicsCache(False)

    def CreateLinestring(self, p, container, layer, color, weight):
        ls = container.Add(Linestring)
        ls.Layer = layer
        ls.Color = color
        ls.Weight = weight
        try:
            ls.Append(p, None, False, False)
            #return ls
        except:
            # some objects with funny UCS throw an error when trying to append to
            # linestring
            container.Remove(ls.SerialNumber)
            #return None
    
        
    def _compute_insert_correction(self, texto, textlines, cachedfont, newtextheight):
        """
        Returns a list of (dx, dy) or None items, one per segment in textlines order.
        Each segment's InsertPoint is pre-computed by GetTextSegments for the old font,
        so for non-left-baseline attachments it needs adjusting when the draw font differs.
        Per-segment corrections are needed because each line has a different width.

        Attachment grids (from SCR_RotateText comments):
          MText  AttachmentPoint: 1=TL 2=TC 3=TR / 4=ML 5=MC 6=MR / 7=BL 8=BC 9=BR
          CadText TextJustification:  6=TL 7=TC 8=TR / 3=ML 4=MC 5=MR / 0=BL 1=BC 2=BR
        """
        try:
            if isinstance(texto, MText):
                attach_str = str(texto.AttachPoint).lower()
            else:
                attach_str = str(texto.Alignment).lower()

            # Parse h/v from the enum NAME string — robust against unknown underlying int values.
            if attach_str.endswith("mid") or attach_str.endswith("center"):
                h = 0.5
            elif attach_str.endswith("right"):
                h = 1.0
            else:
                h = 0.0

            if attach_str.startswith("top"):
                v = 1.0
            elif attach_str.startswith("middle"):
                v = 0.5
            else:
                v = 0.0

            all_segs = list(textlines)
            seg_count = len(all_segs)

            # For multiline blocks the block height >> one-line cap height, so the
            # single-line formula for v gives a wildly wrong Y correction.
            # Each line's InsertPoint already encodes line spacing — only correct X (h).
            if seg_count > 1:
                v = 0.0

            if seg_count == 0 or (h == 0.0 and v == 0.0):
                return None

            align_pt = texto.AlignmentPoint
            seg0 = all_segs[0]
            angle = seg0.RotateAngle
            cos_a = math.cos(angle)
            sin_a = math.sin(angle)

            # For single-line v-correction we also need the glyph bbox height.
            y_max_seg0 = None
            if v != 0.0:
                test_nodes0 = cachedfont.DrawText(
                    Point3D(0, 0, 0), newtextheight,
                    seg0.WidthFactor, seg0.ObliqueAngle, 0, seg0.TextString)
                ys0 = [n.Y for n in test_nodes0 if not n.IsUndefined]
                y_max_seg0 = max(ys0) if ys0 else None

            # Per-segment corrections: each line may have a different width in the new
            # font, so h-centering/right-aligning requires individual dx per segment.
            corrections = []
            any_correction = False
            for _seg in all_segs:
                ip = _seg.InsertPoint
                test_nodes = cachedfont.DrawText(
                    Point3D(0, 0, 0), newtextheight,
                    _seg.WidthFactor, _seg.ObliqueAngle, 0, _seg.TextString)
                valid_xs = [n.X for n in test_nodes if not n.IsUndefined]
                valid_ys = [n.Y for n in test_nodes if not n.IsUndefined]

                if not valid_xs:
                    corrections.append(None)
                    continue

                x_max = max(valid_xs)
                y_max = max(valid_ys) if valid_ys else (y_max_seg0 or 0.0)

                local_dx = -h * x_max
                local_dy = -v * y_max  # 0 when v=0 (multiline)

                old_disp_x = ip.X - align_pt.X
                old_disp_y = ip.Y - align_pt.Y
                old_local_text = cos_a * old_disp_x + sin_a * old_disp_y
                old_local_perp = -sin_a * old_disp_x + cos_a * old_disp_y

                new_local_text = local_dx
                new_local_perp = local_dy if v != 0.0 else old_local_perp

                delta_text = new_local_text - old_local_text
                delta_perp = new_local_perp - old_local_perp  # 0 for multiline

                dx = cos_a * delta_text - sin_a * delta_perp
                dy = sin_a * delta_text + cos_a * delta_perp

                corr = (dx, dy) if (abs(dx) > 1e-9 or abs(dy) > 1e-9) else None
                corrections.append(corr)
                if corr is not None:
                    any_correction = True

            return corrections if any_correction else None
        except:
            return None

    def SetDefaultOptions(self):
        SCROptions.LoadMacroOptions(self, "SCR_ExplodeTextToLines", _OPTIONS, self.currentProject)

    def SaveOptions(self):
        SCROptions.SaveMacroOptions(self, "SCR_ExplodeTextToLines", _OPTIONS)

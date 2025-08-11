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

        cmdData.Version = 1.15
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


        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        
        wv = self.currentProject [Project.FixedSerial.WorldView]

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
                for sn in objlist:
                    texto = self.currentProject.Concordance.Lookup(sn)
                    container = texto.GetSite()

                    j += 1
                    if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(j * 100 / objlist.Count)):
                        break   # function returns true if user pressed cancel

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

                        # go through all the single lines
                        for singletext in textlines:

                            # singletext -> textnodelist
                            # textnodelist -> polysegnodes
                            # polysegnodes -> polysegs
                            # draw -> polysegs
                            # explode each textline into a textnodelist (which can be scrambled with multi element gaps and unnecessary 1 element entries)
                            # parse the textnodelist and add elements to polysegnodes
                            # if we have at least 2 consecutive nodes create a new polyseg and add it to polysegs
                            # draw the polysegs

                            textnodelist = cachedfont.DrawText(singletext.InsertPoint, newtextheight, singletext.WidthFactor, singletext.ObliqueAngle, singletext.RotateAngle, singletext.TextString)
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

        ProgressBar.TBC_ProgressBar.Title = ""
        self.SaveOptions()           

    def CreateLinestring(self, p, container, layer, color, weight):
        ls = container.Add(Linestring)
        ls.Layer = layer
        ls.Color = color
        ls.Weight = weight
        try:
            ls.Append(p, None, False, False)
            return ls
        except:
            # some objects with funny UCS throw an error when trying to append to
            # linestring
            container.Remove(ls.SerialNumber)
            return None
    
        
    def SetDefaultOptions(self):
        #   text settings
        # delete source
        self.deletesource.IsChecked = OptionsManager.GetBool("SCR_ExplodeTextToLines.deletesource", False)
        #   text layer
        self.changetextlayer.IsChecked = OptionsManager.GetBool("SCR_ExplodeTextToLines.changetextlayer", False)
        settingserial = OptionsManager.GetUint("SCR_ExplodeTextToLines.newtextlayerpicker", 8) # 8 is FixedSerial for Layer Zero
        o = self.currentProject.Concordance.Lookup(settingserial)
        if o != None:
            if isinstance(o.GetSite(), LayerCollection):    
                self.newtextlayerpicker.SetSelectedSerialNumber(settingserial, InputMethod(3))
            else:                       
                self.newtextlayerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:                       
            self.newtextlayerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        #   text color
        self.changetextcolor.IsChecked = OptionsManager.GetBool("SCR_ExplodeTextToLines.changetextcolor", False)
        self.textcolorpicker.SelectedColor = Color.FromArgb(OptionsManager.GetInt("SCR_ExplodeTextToLines.textcolorpicker", Color.Red.ToArgb()))
        #   text font
        self.changetextfont.IsChecked = OptionsManager.GetBool("SCR_ExplodeTextToLines.changetextfont", False)
        try: self.fontpicker.SelectedValue = OptionsManager.GetString("SCR_ExplodeTextToLines.fontpicker", "tmodelf.fnt")
        except: pass
        #   text height
        self.changetextheight.IsChecked = OptionsManager.GetBool("SCR_ExplodeTextToLines.changetextheight", False)
        self.textheightdist.Distance = OptionsManager.GetDouble("SCR_ExplodeTextToLines.textheightdist", 1.000)
        #   text weight
        self.changetextweight.IsChecked = OptionsManager.GetBool("SCR_ExplodeTextToLines.changetextweight", False)
        self.textweightpicker.Lineweight = OptionsManager.GetInt("SCR_ExplodeTextToLines.textweightpicker", 0)

    def SaveOptions(self):
        # text settings
        OptionsManager.SetValue("SCR_ExplodeTextToLines.deletesource", self.deletesource.IsChecked)
        
        OptionsManager.SetValue("SCR_ExplodeTextToLines.changetextlayer", self.changetextlayer.IsChecked)
        OptionsManager.SetValue("SCR_ExplodeTextToLines.newtextlayerpicker", self.newtextlayerpicker.SelectedSerialNumber)

        OptionsManager.SetValue("SCR_ExplodeTextToLines.changetextcolor", self.changetextcolor.IsChecked)
        OptionsManager.SetValue("SCR_ExplodeTextToLines.textcolorpicker", self.textcolorpicker.SelectedColor.ToArgb())

        OptionsManager.SetValue("SCR_ExplodeTextToLines.changetextfont", self.changetextfont.IsChecked)
        OptionsManager.SetValue("SCR_ExplodeTextToLines.fontpicker", self.fontpicker.SelectedValue)

        OptionsManager.SetValue("SCR_ExplodeTextToLines.changetextheight", self.changetextheight.IsChecked)
        OptionsManager.SetValue("SCR_ExplodeTextToLines.textheightdist", self.textheightdist.Distance)

        OptionsManager.SetValue("SCR_ExplodeTextToLines.changetextweight", self.changetextweight.IsChecked)
        OptionsManager.SetValue("SCR_ExplodeTextToLines.textweightpicker", self.textweightpicker.Lineweight)


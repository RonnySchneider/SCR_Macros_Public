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
    cmdData.Key = "SCR_DTMSlopeColored"
    cmdData.CommandName = "SCR_DTMSlopeColored"
    cmdData.Caption = "_SCR_DTMSlopeColored"
    cmdData.UIForm = "SCR_DTMSlopeColored"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "DTM"
        cmdData.ShortCaption = "Slope-Colored 3D-Faces"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.11
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Slope-Colored 3D-Faces"
        cmdData.ToolTipTextFormatted = "create Linestrings, IFC-Faces and set their color based on a table\nor apply colored materials to a surface"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png") 
        cmdData.ImageSmall = b
    except:
        pass


class SCR_DTMSlopeColored(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_DTMSlopeColored.xaml") as s:
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

        # a hidden field to catch changes to the materials list - is then trigering an update of the normal combobox
        self.hiddenmateriallist.FilterByEntityTypes = Array[Type]([clr.GetClrType(MiscMaterial)])
        self.hiddenmateriallist.ValueChanged += self.materialsChanged
        
        types = Array[Type](SurfaceTypeLists.AllWithCutFillMap)+Array[Type]([clr.GetClrType(ProjectedSurface)])
        #types.extend (Array[Type]([clr.GetClrType(ProjectedSurface)]))
        #self.surfacepicker.FilterByEntityTypes = types
        #self.surfacepicker.AllowNone = False

        self.surfaceticklist.SearchContainer = Project.FixedSerial.WorldView
        self.surfaceticklist.UseSelectionEngine = False
        self.surfaceticklist.SetEntityType(types, self.currentProject)
        #tt = self.surfaceticklist
        #self.surfaceticklist.Height = 200
        #self.surfaceticklist.Width = 200

        # Description is the actual property name inside the item in ticklist.Content.Items
        sd = System.ComponentModel.SortDescription("Description", System.ComponentModel.ListSortDirection.Ascending)
        self.surfaceticklist.Content.Items.SortDescriptions.Add(sd)

        self.ticklistfilter.TextChanged += self.FilterChanged


        # find the the SiteImprovementMaterialCollection as object
        for o in self.currentProject:
            if isinstance(o, ConstructionMaterialCollection):
                self.cmc = o

        # populate material category list the first time
        self.materialsChanged(None, None)

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def FilterChanged(self, ctrl, e):
        
        exclude = []
        self.surfaceticklist.SetExcludedEntities(exclude)

        tt = self.ticklistfilter.Text.lower()
        ticklistfilter = tt.split()

        for i in self.surfaceticklist.EntitySerialNumbers:
            for f in ticklistfilter:
               if not f in i.Key.Description.lower():
                    exclude.Add(i.Value)

        self.surfaceticklist.SetExcludedEntities(exclude)

        self.SaveOptions()


    def newsurfaceChanged(self, sender, e):
        if self.newsurface.IsChecked:
            self.nonewsurface.IsEnabled = False
        else:
            self.nonewsurface.IsEnabled = True

    def materialsChanged(self, sender, e):

        # to have the drop down list sorted we fill a templist first, sort it and then fill the dropdownbox
        templist = []
        for m in self.cmc:
            # make sure we only show the pit types that fit the network type gravity/pressure etc.
            if isinstance(m, MiscMaterial):
                if not m.UserCategory in templist:
                    templist.Add(m.UserCategory)
        templist.sort()

        self.usermaterialpicker.Items.Clear() # need to clear, otherwise the list would get longer and longer
        for p in templist:
            item = ComboBoxItem()
            item.Content = p
            item.FontSize = 12
            self.usermaterialpicker.Items.Add(item)


    def SetDefaultOptions(self):
        ## Select surface
        #try:    self.surfacepicker.SelectIndex(OptionsManager.GetInt("SCR_DTMSlopeColored.surfacepicker", 0))
        #except: self.surfacepicker.SelectIndex(0)

        self.newsurface.IsChecked = OptionsManager.GetBool("SCR_DTMSlopeColored.newsurface", False)
        self.usecsv.IsChecked = OptionsManager.GetBool("SCR_DTMSlopeColored.usecsv", True)
        self.csvfilename.Text = OptionsManager.GetString("SCR_DTMSlopeColored.csvfilename", os.path.abspath(os.path.dirname(__file__)))
        self.usecmc.IsChecked = OptionsManager.GetBool("SCR_DTMSlopeColored.usecmc", False)

        try: self.usermaterialpicker.Text = OptionsManager.GetString("SCR_DTMSlopeColored.usermaterialpicker")
        except: pass

        self.interpolatecolor.IsChecked = OptionsManager.GetBool("SCR_DTMSlopeColored.interpolatecolor", False)

        #self.remapcolor.IsChecked = OptionsManager.GetBool("SCR_DTMSlopeColored.remapcolor", False)
        self.recolordtm.IsChecked = OptionsManager.GetBool("SCR_DTMSlopeColored.recolordtm", True)
        self.drawifc.IsChecked = OptionsManager.GetBool("SCR_DTMSlopeColored.drawifc", False)
        self.drawlines.IsChecked = OptionsManager.GetBool("SCR_DTMSlopeColored.drawlines", False)

        lserial = OptionsManager.GetUint("SCR_DTMSlopeColored.layerpicker", 8)
        o = self.currentProject.Concordance.Lookup(lserial) # get the object with the saved serial number
        if o != None:   # check if we actually got an object back
            if isinstance(o.GetSite(), LayerCollection):    # check if the object is a layer
                self.layerpicker.SetSelectedSerialNumber(lserial, InputMethod(3))
            else:  # if not then set Layer zero                       
                self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:   # if not then set Layer zero               
            self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))

        # need to restore that one last, since it also saves, it would clear all the other fields
        self.ticklistfilter.Text = OptionsManager.GetString("SCR_DTMSlopeColored.ticklistfilter", "")


    def SaveOptions(self):

        #try:    # if nothing is selected it would throw an error
        #   OptionsManager.SetValue("SCR_DTMSlopeColored.surfacepicker", self.surfacepicker.SelectedIndex)
        #except:
        #    pass
        OptionsManager.SetValue("SCR_DTMSlopeColored.ticklistfilter", self.ticklistfilter.Text)
        OptionsManager.SetValue("SCR_DTMSlopeColored.newsurface", self.newsurface.IsChecked)  
        OptionsManager.SetValue("SCR_DTMSlopeColored.usecsv", self.usecsv.IsChecked)  
        OptionsManager.SetValue("SCR_DTMSlopeColored.csvfilename", self.csvfilename.Text)
        OptionsManager.SetValue("SCR_DTMSlopeColored.usecmc", self.usecmc.IsChecked)  
        OptionsManager.SetValue("SCR_DTMSlopeColored.usermaterialpicker", self.usermaterialpicker.SelectionBoxItem)

        OptionsManager.SetValue("SCR_DTMSlopeColored.interpolatecolor", self.interpolatecolor.IsChecked)  


        #OptionsManager.SetValue("SCR_DTMSlopeColored.remapcolor", self.remapcolor.IsChecked)  
        OptionsManager.SetValue("SCR_DTMSlopeColored.recolordtm", self.recolordtm.IsChecked)  
        OptionsManager.SetValue("SCR_DTMSlopeColored.drawifc", self.drawifc.IsChecked)  
        OptionsManager.SetValue("SCR_DTMSlopeColored.drawlines", self.drawlines.IsChecked)  
        OptionsManager.SetValue("SCR_DTMSlopeColored.layerpicker", self.layerpicker.SelectedSerialNumber)

    def csvbutton_Click(self, sender, e):
        dialog = OpenFileDialog()
        dialog.InitialDirectory = os.path.dirname(self.csvfilename.Text)
        dialog.Filter=("CSV|*.csv")
        
        tt=dialog.ShowDialog()
        if tt==DialogResult.OK:
            self.csvfilename.Text = dialog.FileName

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content = ''
        self.success.Content = ''

        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())

        wv = self.currentProject [Project.FixedSerial.WorldView]

        try:
            # the "with" statement will unroll any changes if something go wrong
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                for i in self.surfaceticklist.Content.SelectedItems:

                    # get the surface geometry
                    surface = wv.Lookup(i.EntitySerialNumber)
                    gem = surface.Gem
                    if isinstance(surface, ProjectedSurface):
                        projected=True
                    else:
                        projected=False


                    if self.newsurface.IsChecked:
                        if not projected:
                            self.createnewsurface(surface)
                        else:
                            self.error.Content += '\nnew projected surface not supported yet'
                    else:
                        # prepare the list with the triangle information and get max-min slope
                        trianglelist = []
                        minslope = 1000000000.0
                        maxslope = 0.0
                        ProgressBar.TBC_ProgressBar.Title = "compute Slopes"
                        # go through the triangles and compute the slope
                        if gem.NumberOfTriangles > 0:
                            for i in range(gem.NumberOfTriangles):

                                if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(i * 100 / gem.NumberOfTriangles)):
                                    break   # function returns true if user pressed cancel
                                
                                if projected == True:

                                    c1 = surface.TransformPointToWorldDelegate(gem.GetTrianglePoint(i,0))
                                    c2 = surface.TransformPointToWorldDelegate(gem.GetTrianglePoint(i,1))
                                    c3 = surface.TransformPointToWorldDelegate(gem.GetTrianglePoint(i,2))

                                else:

                                    c1 = gem.GetTrianglePoint(i,0)
                                    c2 = gem.GetTrianglePoint(i,1)
                                    c3 = gem.GetTrianglePoint(i,2)

                                p = Plane3D(c1, c2, c3)[0]

                                triangle = []
                                triangle = [p.Slope, c1, c2, c3, 0] # the 0 is a place holder for the material index

                                trianglelist.Add(triangle)

                                # Plane3D.Slope - Gets the steepest slope on the plane in percent.
                                if p.Slope > maxslope: maxslope = p.Slope
                                if p.Slope < minslope: minslope = p.Slope

                            self.success.Content += '\nMin-Slope: ' + str(minslope)
                            self.success.Content += '\nMax-Slope: ' + str(maxslope)

                        #restrictmaxslope = True

                        #if self.remapcolor.IsChecked and not self.usecmc.IsChecked:
                        #    
                        #    self.success.Content += '\n\n\Remapping Slope-Color-Table'
                        #
                        #    if restrictmaxslope and maxslope > 150:
                        #        maxslope = 150
                        #        self.success.Content += '\nrestricting Max-Slope'
                        #    
                        #    slopeinc = (maxslope - minslope) / (slopelist.Count - 2)
                        #
                        #    slopelist[1][0] = minslope
                        #    for i in range(2, slopelist.Count):
                        #        slopelist[i][0] = slopelist[i - 1][0] + slopeinc
                        #        self.success.Content += '\n' + str(slopelist[i][0]) + '% - ' + \
                        #                                str(slopelist[i][1]) + '-' + str(slopelist[i][2]) + '-' + str(slopelist[i][3])

                        # get the color/slope list from MSI or CSV
                        slopelist = self.getcolors()

                        if slopelist.Count > 0:
                            # set materials to surface
                            if trianglelist.Count > 0 and self.recolordtm.IsChecked:
                                # prepare the cmcslopelist - maybe create new materials if interpolation is ticked
                                trianglelist = self.preparematerials(gem, trianglelist, slopelist)
                                self.setdtmcolor(surface, trianglelist)
                            
                            # draw linestrings or IFC faces
                            if trianglelist.Count > 0 and not self.recolordtm.IsChecked:
                                self.drawnewentities(gem, trianglelist, slopelist, surface.Name)
                      

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

    def createnewsurface(self, surface):

        wv = self.currentProject [Project.FixedSerial.WorldView]
        surName = Model3D.GetUniqueName(surface.Name, None, wv) #make sure name is unique
        newSurface = wv.Add(clr.GetClrType(Model3D))
        newSurface.Name = surName
        builder = newSurface.GetGemBatchBuilder()
        b = DTMSharpness.eSharpAndTextureBndy
        
        ProgressBar.TBC_ProgressBar.Title = "create new Surface"
        # go through the triangles and compute the slope
        if surface.NumberOfTriangles > 0:
            for i in range(surface.NumberOfTriangles):

                if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(i * 100 / surface.NumberOfTriangles)):
                    break   # function returns true if user pressed cancel

                if surface.IsTriangleMaterialPresent(i):

                    tvertices = List[Point3D]()
                    c1 = surface.GetVertexPoint(surface.GetTriangleIVertex(i,0))
                    c2 = surface.GetVertexPoint(surface.GetTriangleIVertex(i,1))
                    c3 = surface.GetVertexPoint(surface.GetTriangleIVertex(i,2))

                    #tvertices.Add(c1)
                    #tvertices.Add(c2)
                    #tvertices.Add(c3)
                    
                    polyseg = PolySeg.PolySeg()
                    polyseg.Add(c1)
                    polyseg.Add(c2)
                    polyseg.Add(c3)
                    polyseg.Close(True)
                    poly_offset = polyseg.Offset(Side.Right, 0.001)
                    shrinkpoints = poly_offset[1].ToPoint3DArray()

                    if shrinkpoints.Count == 4:
                        p = Plane3D(c1, c2, c3)[0]

                        for i in range(shrinkpoints.Count):
                            pt = shrinkpoints[i]
                            pt.Z = p.Slope
                            shrinkpoints[i] = pt

                        for j in range(3):
                            v1 = builder.AddVertex(shrinkpoints[j])
                            v2 = builder.AddVertex(shrinkpoints[j + 1])
                            builder.AddBreakline(Byte(b), v1[0], v2[0])
                    else:
                        tt = 1

        builder.Construction()
        builder.Commit()
        #Mode = SmoothShading | ElevationContours | ShowBackfaces | FilledTriangles | ElevationContours_2D | FilledTriangles_2D
        newSurface.Mode = DisplayMode(1 +64 + 128 + 512 + 1024 + 65536)

    def preparematerials(self, gem, trianglelist, slopelist):
        
        if self.interpolatecolor.IsChecked:
            # create new category name - need to do it outside of the triangle loop
            now = datetime.now()
            self.usematcat += ' - interpolated - ' + now.strftime("%Y%m%d%H%M%S")

        ProgressBar.TBC_ProgressBar.Title = "prepare Materials"
        i = 0
        # trianglelist [p.Slope, c1, c2, c3, 0] the 0 is a place holder for the material index
        for triangle in trianglelist:
            if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(i * 100 / trianglelist.Count)):
                break   # function returns true if user pressed cancel
            i += 1
            
            if not gem.IsTriangleMaterialPresent(i - 1): # skip "holes"
                continue
            
            if not self.interpolatecolor.IsChecked:
                if triangle[0] > float(slopelist[slopelist.Count - 1][0]):
                    triangle[4] = gem.MaterialIndex(slopelist[slopelist.Count - 1][4])
                else:
                    # go through the slopelist and get the material SN for the triangle
                    for j in range(1, slopelist.Count):
        
                        if triangle[0] >= float(slopelist[j - 1][0]) and triangle[0] <= float(slopelist[j][0]):
                            
                            # slopelist(%, R, G, B, 0) # 0 is placeholder for global material SN

                            # gem.MaterialIndex checks if the material exists already in the surface
                            # if yes it returns the material index inside the surface
                            # if not it will add the material first and returns the internal index
                            triangle[4] = gem.MaterialIndex(slopelist[j - 1][4])

                            break # no need to continue through the whole colorlist

            else: # need to interpolate

                ## slopelist(%, R, G, B, 0)
                # go through the slopelist and get the interpolated RGB value
                # trianglelist [p.Slope, c1, c2, c3, 0] the 0 is a place holder for the material index
                r, g, b = 0, 0, 0
                if triangle[0] > float(slopelist[slopelist.Count - 1][0]):
                    r = int(slopelist[slopelist.Count - 1][1])
                    g = int(slopelist[slopelist.Count - 1][2])
                    b = int(slopelist[slopelist.Count - 1][3])
                else:
                    for j in range(1, slopelist.Count):
        
                        if triangle[0] >= float(slopelist[j - 1][0]) and triangle[0] <= float(slopelist[j][0]):
        
                            r = self.mapvalue(float(slopelist[j - 1][0]), float(slopelist[j][0]), float(slopelist[j - 1][1]), float(slopelist[j][1]), triangle[0])
                            g = self.mapvalue(float(slopelist[j - 1][0]), float(slopelist[j][0]), float(slopelist[j - 1][2]), float(slopelist[j][2]), triangle[0])
                            b = self.mapvalue(float(slopelist[j - 1][0]), float(slopelist[j][0]), float(slopelist[j - 1][3]), float(slopelist[j][3]), triangle[0])
        
                            break # no need to continue through the whole colorlist

                # now check if a material with that RGB values exists already
                # check if the material for that slope exists in that material category
                foundmatsn = None
                for mat in self.cmc: # cmc - ConstructionMaterialCollection
                    if mat.UserCategory == self.usematcat:
                        #tt = float(mat.Name)
                        if mat.Color == Color.FromArgb(r, g, b):
                            foundmatsn = mat.SerialNumber
                            break # no need to check th rest of the list
        
                if foundmatsn:
                     # check/add global material SN to surface - get internal mat index
                    gemmati = gem.MaterialIndex(foundmatsn)
                    # we could now set the index directlyto the triangle, but that's done in a separate step anyway
                    triangle[4] = gemmati
                else:
                    # need to create a new material in that category
                    newmat = self.cmc.Add(MiscMaterial)
                    newmat.Name = str(triangle[0])
                    newmat.Color = Color.FromArgb(r, g, b)
                    newmat.UserCategory = self.usematcat
                    # check/add global material SN to surface - get internal mat index
                    gemmati = gem.MaterialIndex(newmat.SerialNumber)
                    # we could now set the index directly to the triangle, but that's done in a separate step anyway
                    triangle[4] = gemmati
        


        return trianglelist

    def setdtmcolor(self, surface, trianglelist):

        gem = surface.Gem
        # if we recolor the gem via materials from the MSI Manager
        # we set the material directly to the triangle
        # standard material index after DTM creation is 0

        # trianglelist [p.Slope, c1, c2, c3, materialindex]

        ProgressBar.TBC_ProgressBar.Title = "set DTM Materials"
        for i in range(trianglelist.Count):
            if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(i * 100 / trianglelist.Count)):
                break   # function returns true if user pressed cancel

            if not surface.IsTriangleMaterialPresent(i): # skip "holes"
                continue

            gem.SetTriangleMaterialIndex(i, trianglelist[i][4])

        # SmoothShading 1 + Materials 8 + ShowBackfaces 128 + FilledTriangles 512 + Materials_2D 8192 + FilledTriangles_2D 65536
        surface.Mode = DisplayMode(1 + 8 + 128 + 512 + 8192 + 65536)
             

        # trigger classification forth and back, otherwise the material might not show
        tt = surface.Classification
        tt2 = SurfaceClassification(2)
        for i in range(0,11):
            if not surface.Classification == SurfaceClassification(i):
                surface.Classification = SurfaceClassification(i)
                break
        surface.Classification = tt


                
    def getcolors(self):
        if self.usecmc.IsChecked:
            # slopelist(%, R, G, B, 0) # 0 is placeholder for global material SN
            slopelist = []
            
            for mat in self.cmc: # cmc - ConstructionMaterialCollection
                
                if mat.UserCategory == self.usermaterialpicker.SelectionBoxItem:

                    # gem.MaterialIndex checks if the material exists already in the surface
                    # if yes it returns the material index inside the surface
                    # if not it will add the material first and returns the internal index
                    
                    slopelist.Add([float(mat.Name), mat.Color.R, mat.Color.G, mat.Color.B, mat.SerialNumber])
            slopelist.sort() # depending on how the values had been entered in the MSI Manager (strings) the numerical order might now be wrong

            self.usematcat = self.usermaterialpicker.SelectionBoxItem

        else: # from CSV
            # slopelist(%, R, G, B, 0) # 0 is placeholder for global material SN
            filename = self.csvfilename.Text

            slopelist=[]
            
            if File.Exists(filename): # if file still exists
                with open(filename, 'r') as csvfile:
                    reader = csv.reader(csvfile, delimiter=',') 
                    # This skips the first row of the CSV file.
                    next(reader)
                    for row in reader:
                        row[0] = float(row[0])
                        row[1] = float(row[1])
                        row[2] = float(row[2])
                        row[3] = float(row[3])
                        if row.Count < 5:
                            row.Add(0)
                        slopelist.Add(row)
            else:
                self.error.Content += "\ncouldn't find file"
            #slopelist.pop(0) # remove the header line
            slopelist.sort() # just in case

            if self.recolordtm.IsChecked and slopelist.Count > 0: # create materials for direct DTM colorizing
                # create new category name
                now = datetime.now()
                newcat = os.path.basename(filename) + ' - ' + now.strftime("%Y%m%d%H%M%S")
                # slopelist(%, R, G, B, 0)
                for s in slopelist:
                    # create a new material in that category
                    newmat = self.cmc.Add(MiscMaterial)
                    newmat.Name = str(s[0])
                    newmat.Color = Color.FromArgb(int(s[1]), int(s[2]), int(s[3]))
                    newmat.UserCategory = newcat

                    s[4] = newmat.SerialNumber

                self.usematcat = newcat

        return slopelist


    def drawnewentities(self, gem, trianglelist, slopelist, surfacename):

        wv = self.currentProject [Project.FixedSerial.WorldView]

        i = 1
        ProgressBar.TBC_ProgressBar.Title = "create colored Triangles"
        for triangle in trianglelist:
            if ProgressBar.TBC_ProgressBar.SetProgress(math.floor(i * 100 / trianglelist.Count)):
                break   # function returns true if user pressed cancel
            i += 1

            if not gem.IsTriangleMaterialPresent(i - 2): # skip "holes"
                continue

            r, g, b = 0, 0, 0
            if triangle[0] > float(slopelist[slopelist.Count - 1][0]):
                r = int(slopelist[slopelist.Count - 1][1])
                g = int(slopelist[slopelist.Count - 1][2])
                b = int(slopelist[slopelist.Count - 1][3])
            else:
                # go through the slopelist and get the color for the triangle
                for j in range(1, slopelist.Count):
        
                    if triangle[0] >= float(slopelist[j - 1][0]) and triangle[0] <= float(slopelist[j][0]):
                        
                        # interpolate color if checked
                        if self.interpolatecolor.IsChecked:
                            r = self.mapvalue(float(slopelist[j - 1][0]), float(slopelist[j][0]), float(slopelist[j - 1][1]), float(slopelist[j][1]), triangle[0])
                            g = self.mapvalue(float(slopelist[j - 1][0]), float(slopelist[j][0]), float(slopelist[j - 1][2]), float(slopelist[j][2]), triangle[0])
                            b = self.mapvalue(float(slopelist[j - 1][0]), float(slopelist[j][0]), float(slopelist[j - 1][3]), float(slopelist[j][3]), triangle[0])
                        else:
                            r = int(slopelist[j-1][1])
                            g = int(slopelist[j-1][2])
                            b = int(slopelist[j-1][3])

                        break # no need to continue through the whole colorlist

            # draw filled stringlines
            if self.drawlines.IsChecked:
                l = wv.Add(clr.GetClrType(Linestring))
                
                e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                e.Position = triangle[1]
                l.AppendElement(e)  
                e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                e.Position = triangle[2]
                l.AppendElement(e)  
                e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                e.Position = triangle[3]
                l.AppendElement(e)  
                
                l.Closed = True
                l.DrawSolidFill = True
                l.Layer = self.layerpicker.SelectedSerialNumber
                l.Name = 'Slope: ' + str(triangle[0]) + '%'
                #l.FillColor = Color.Magenta
                
                l.FillTransparency = 0
                l.FillColor = Color.FromArgb(r, g, b)
                l.Color = Color.FromArgb(r, g, b)
            
            # draw IFC shapes
            if self.drawifc.IsChecked:
                # prepare the empty arrays
                ifcvertices = Array[Point3D]([Point3D()] * 3)
                ifcfacelist = Array[Int32]([Int32()] * 4) # 1 triangles * 4 entries
                ifcnormals = Array[Point3D]([Point3D()]*0)

                ifcvertices[0] = triangle[1]
                ifcvertices[1] = triangle[2]
                ifcvertices[2] = triangle[3]

                ifcfacelist[0] = 3 # always 3 vertices per triangle
                ifcfacelist[1] = 0
                ifcfacelist[2] = 1
                ifcfacelist[3] = 2

                ### mesh = wv.Add(clr.GetClrType(IFCMesh))
                ### 
                ### mesh.CreateShell(0, Guid.NewGuid(), ifcvertices, ifcfacelist, ifcnormals, Color.FromArgb(r, g, b).ToArgb(), Matrix4D())
                ### #mesh._TriangulatedFaceList = ifcfacelist # otherwise GetTriangulatedFaceList() in the perpDist macro will not get the necessary values
                ### mesh.Layer = self.layerpicker.SelectedSerialNumber
                ### mesh.Name = 'Slope-IFC-Faces'
                ### mesh.Description = 'Slope: ' + str(triangle[0]) + '%'
                ### #Mode = SmoothShading | BreakLines | ElevationContours | ShowBackfaces | FilledTriangles | FilledTriangles_2D | BreakLines_2D | DrapeLines_2D
                ### mesh.Mode = DisplayMode(1 + 128 + 512 + 65536 + 4 + 4096)

                self.createifcobject("SlopeColoredDTM - " + surfacename, "", 'Slope: ' + str(triangle[0]) + '%', ifcvertices, ifcfacelist, ifcnormals, Color.FromArgb(r, g, b).ToArgb(), Matrix4D(), self.layerpicker.SelectedSerialNumber, DisplayMode(1 + 128 + 512 + 65536 + 4 + 4096))
        
        
    def mapvalue(self, input_start, input_end, output_start, output_end, x):

        tt = (x - input_start) / (input_end - input_start) * (output_end - output_start) + output_start
        tt2 = int(round(tt, 0))

        return tt2

    def createifcobject(self, ifcprojectname, ifcshellgroupname, ifcshellname, ifcvertices, ifcfacelist, ifcnormals, color, transform, layer, meshmode):
  
        # get the BIM entity collection from the world view - if it doesn't exist yet it is created
        bimEntityColl = BIMEntityCollection.ProvideEntityCollection(self.currentProject, True)
        shellMeshDataColl = ShellMeshDataCollection.ProvideShellMeshDataCollection(self.currentProject, True)

        # theoretical IFC hierarchy in TBC
        # IFCPROJECT
        #   IFCSITE
        #       IFCBUILDING
        #           IFCBUILDINGSTOREY
        #               IFCBUILDINGELEMENTPROXY
        
        # test if provided project name exists, if yes then add to it
        found = False
    
        for e in bimEntityColl:
            if e.Name == ifcprojectname:
                bimprojectEntity = e
                found = True
                continue
        if not found:
            bimprojectEntity = bimEntityColl.Add(clr.GetClrType(BIMEntity))
            bimprojectEntity.EntityType = "IFCPROJECT"
            bimprojectEntity.Description = ifcprojectname # must be set, otherwise export fails
            bimprojectEntity.BIMGuid = Guid.NewGuid()

        #bimsiteEntity = bimprojectEntity.Add(clr.GetClrType(BIMEntity))
        #bimsiteEntity.EntityType = "IFCSITE"
        #bimsiteEntity.Description = "TestSite" # must be set, otherwise export fails
        #bimsiteEntity.BIMGuid = Guid.NewGuid()

        #bimbuildingEntity = bimprojectEntity.Add(clr.GetClrType(BIMEntity))
        #bimbuildingEntity.EntityType = "IFCBUILDING"
        #bimbuildingEntity.Description = "TestBuilding" # must be set, otherwise export fails
        #bimbuildingEntity.BIMGuid = Guid.NewGuid()

        #bimbuildingstoreyEntity = bimbuildingEntity.Add(clr.GetClrType(BIMEntity))
        #bimbuildingstoreyEntity.EntityType = "IFCBUILDINGSTOREY"
        #bimbuildingstoreyEntity.Description = "TestStorey" # must be set, otherwise export fails
        #bimbuildingstoreyEntity.BIMGuid = Guid.NewGuid()
        
        ## find unique group name
        #while found == True:
        #    found = False
        #    for e in bimprojectEntity:
        #        if e.Name == ifcshellgroupname:
        #            found = True
        #    if found:
        #        ifcshellgroupname = PointHelper.Helpers.Increment(ifcshellgroupname, None)
        
        # if we create at least IFCPROJECT and IFCBUILDINGELEMENTPROXY (which includes the IFCMesh)
        # we can export it properly into an IFC file but keep the meshes separate
        # upon import TBC will structure it as follows
        # IFCPROJECT --> IFCPROJECT
        # everything will end up under "Default Site - IFCSITE"
        # the IFCBUILDINGELEMENTPROXY with the subsequent IFCMesh get combined into separate BIM Objects

        bimobjectEntity = bimprojectEntity.Add(clr.GetClrType(BIMEntity))
        bimobjectEntity.EntityType = "IFCBUILDINGELEMENTPROXY"
        bimobjectEntity.Description = ifcshellname
        bimobjectEntity.BIMGuid = Guid.NewGuid()
        bimobjectEntity.Mode = DisplayMode(1 + 64 + 128 + 512 + 4096) 
        bimobjectEntity.MeshColor = TrimbleColorCodeValues.AsImported.ToArgb()
        # MUST set layer, otherwise the properties manager will falsly show layer "0"
        # but the object is invisible until manually changing the layer
        bimobjectEntity.Layer = layer

        shellmeshdata = shellMeshDataColl.AddShellMeshData(self.currentProject)
        shellmeshdata.CreateShellMeshData(ifcvertices, ifcfacelist, ifcnormals)
        shellmeshdata.SetVolumeCalculationShell(ifcvertices, ifcfacelist)

        meshInstance = bimobjectEntity.Add(clr.GetClrType(ShellMeshInstance))
        meshInstance.CreateShell(0, shellmeshdata.SerialNumber, color, Matrix4D())
        #mesh._TriangulatedFaceList = ifcfacelist # otherwise GetTriangulatedFaceList() in the perpDist macro will not get the necessary values

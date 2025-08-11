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
    cmdData.Key = "SCR_ProjectedSurfaceContour"
    cmdData.CommandName = "SCR_ProjectedSurfaceContour"
    cmdData.Caption = "_SCR_ProjectedSurfaceContour"
    cmdData.UIForm = "SCR_ProjectedSurfaceContour"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "DTM"
        cmdData.ShortCaption = "Projected Surface Contour"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.15
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Create Contour on Projected Surface"
        cmdData.ToolTipTextFormatted = "Creates a Contour on a projected Surface"

    except:
        pass
    try:
        b = Bitmap(macroFileFolder + "\\" + cmdData.Key + ".png") 
        cmdData.ImageSmall = b
    except:
        pass

class SCR_ProjectedSurfaceContour(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader(macroFileFolder + r"\SCR_ProjectedSurfaceContour.xaml") as s:
            wpf.LoadComponent(self, s)
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
        types = Array[Type]([clr.GetClrType(ProjectedSurface)])    # we fill an array with TBC object types, we could combine different types in that list
                                                                   # that would add normal Surfaces to the list ->> +Array[Type](SurfaceTypeLists.AllWithCutFillMap)

        self.surfacedroplist.FilterByEntityTypes = types    # we fill the dropdownlist by applying that types array as filter
        self.surfacedroplist.AllowNone = False              # our list shall not show an empty field
                                                            # I haven't found a way yet to auto-select the top most value in the list, it says it's read only

        self.ElCtl1.ValueChanged += self.El1Changed         # not 100% sure how that works, but with those two lines we add to our first
        self.surfacedroplist.ValueChanged += self.surfacedroplistChanged    # elevation field and the dropdown list the ability to react to changes

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def SetDefaultOptions(self):
        try: self.surfacedroplist.SelectIndex(OptionsManager.GetInt("SCR_ProjectedSurfaceContour.surfacedroplist", 0))
        except: self.surfacedroplist.SelectIndex(0)
        
        lserial = OptionsManager.GetUint("SCR_ProjectedSurfaceContour.layerpicker", 8)
        o = self.currentProject.Concordance.Lookup(lserial)
        if o != None:
            if isinstance(o.GetSite(), LayerCollection):    
                self.layerpicker.SetSelectedSerialNumber(lserial, InputMethod(3))
            else:                       
                self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:                       
            self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))


    def SaveOptions(self):
        OptionsManager.SetValue("SCR_ProjectedSurfaceContour.surfacedroplist", self.surfacedroplist.SelectedIndex)
        OptionsManager.SetValue("SCR_ProjectedSurfaceContour.layerpicker", self.layerpicker.SelectedSerialNumber)
        

    def surfacedroplistChanged(self, sender, e):        # in case we select a new surface from the list we update the min/max textfields
        wv = self.currentProject[Project.FixedSerial.WorldView]
        if self.surfacedroplist.SelectedSerial!=0:
            surface = wv.Lookup(self.surfacedroplist.SelectedSerial) # we get our selected surface as object
            
            self.LabelMin.Content='Minimum Elevation: ' + str("%4.4f" % surface.MinElevation) # we get some surface values into some labels
            self.LabelMax.Content='Maximum Elevation: ' + str("%4.4f" % surface.MaxElevation)
        
        

    def CheckBoxChanged(self, sender, e):       # if the interval checkbox changes we enable/disable some input fields
        if self.checkBox.IsChecked==True:
            self.ElCtl1.IsEnabled=False
            self.ElStart.IsEnabled=True
            self.ElEnd.IsEnabled=True
            self.ElInterval.IsEnabled=True
        else:
            self.ElCtl1.IsEnabled=True
            self.ElStart.IsEnabled=False
            self.ElEnd.IsEnabled=False
            self.ElInterval.IsEnabled=False

    def El1Changed(self, ctrl, e):              # in case of no interval we draw a new line everytime we click with the mouse and get a new elevation
        if e.Cause == InputMethod.Mouse:     
            self.OkClicked(None, None)
        
    def CancelClicked(self, thisCmd, args):
        thisCmd.CloseUICommand()

    def OkClicked(self, thisCmd, e):
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())

        wv = self.currentProject[Project.FixedSerial.WorldView]

        ctrl = self.ElCtl1

        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
            
                surface = wv.Lookup(self.surfacedroplist.SelectedSerial)    # we get our selected surface as object
                nTri = surface.NumberOfTriangles    # we need the number of triangles in the surface to count through them
                
                tri_intersect=Array[Point3D]([Point3D()]*2) # we have to define an array that will hold our intersection points

                interval = self.ElInterval.Elevation    # we read the interval in order to check it, to avoid a divide by zero

                if self.checkBox.IsChecked==True:       # in case the interval box is checked we compute the number of lines we need to draw

                   start_el = self.ElStart.Elevation       # we read the interval start/end values
                   end_el = self.ElEnd.Elevation

                   if interval!=0 and math.isnan(interval)==False:   #we check if the interval is zero or not set at all

                      if end_el>start_el:             # we check in which direction we have to 
                          interval = abs(interval) # just in case the interval was negative although start/end are normal
                          self.ElInterval.Elevation = interval # we update the interval field, just in case it was negative
                          contour_count = int(math.floor((end_el-start_el)/interval)+1)
                      else:
                          interval = -abs(interval) # just in case the interval was positive although start/end upside down
                          self.ElInterval.Elevation = interval # we update the interval field, just in case it was positive
                          contour_count = int(math.floor((start_el-end_el)/abs(interval))+1)  

                   else:        # in case the interval is zero but the checkbox is checked we draw one line at the start elevation
                      interval=0
                      contour_count=1
    
                else:           # no interval, just a single line
                    start_el = self.ElCtl1.Elevation
                    interval=0
                    contour_count=1

                #layer_sn = self.layerpicker.SelectedSerialNumber

                for i in range(1,contour_count+1): # we draw the number of lines we've computed above, at least one

                    linetuples = DynArray()

                    for i2 in range(nTri): # we go through all triangles of our surface
                        
                        if isinstance(surface,ProjectedSurface): # that check isn't really necessary since our drop down list is filtered to projected surfaces only
                                                                 # that was used for a test when the list contained all types of surfaces
                            # we get the three corners of the triangle
                            # since we are working with a projected surface we have to convert the UCS coordinate to World coordinates
                            # in case of a standard surface we would just use
                            # v1 = surface.GetVertexPoint(surface.GetTriangleIVertex(i,0))
                            # that's why I had the if up there

                            v1 = surface.TransformPointToWorldDelegate(surface.GetVertexPoint(surface.GetTriangleIVertex(i2,0)))
                            v2 = surface.TransformPointToWorldDelegate(surface.GetVertexPoint(surface.GetTriangleIVertex(i2,1)))
                            v3 = surface.TransformPointToWorldDelegate(surface.GetVertexPoint(surface.GetTriangleIVertex(i2,2)))
                            
                            # we use the built in function to compute intersections
                            test_intersection=Triangle3D.IntersectElevationNoTolerance(v1,v2,v3,start_el,tri_intersect)
                    
                            if test_intersection[4] == 2: # we are only interested if we have two intersections
                                
                                linetuples.Add(tri_intersect.Clone()) # we have to add a clone of the result, that's an Object oriented language issue
                                                                      # it only stores a link back the tri_intersect object, which does change in each iteration,
                                                                      # but the link is always the same, the array would show us the same/last value over and over again


                    self.connect_linetuples(linetuples) # we connect the lines as much as possible

                    start_el=start_el+interval  # we increment the elevation to draw the next line
            
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

    
    def connect_linetuples(self, linetuples):                  
    
        wv = self.currentProject[Project.FixedSerial.WorldView]
        #layer_sn = self.layerpicker.SelectedSerialNumber

        polysegs = List[PolySeg.PolySeg]()
        polyseg = None

        for i in range(0, linetuples.Count):
            polyseg = PolySeg.PolySeg()
            polyseg.Add(linetuples[i][0])
            polyseg.Add(linetuples[i][1])
            polysegs.Add(polyseg.Clone())

        # use build in function to combine the segments as much as possible, spares us to do a manual Project-Cleanup
        ttcount = PolySeg.PolySeg.JoinTouchingPolysegs(polysegs)

        # draw the lines
        for p in polysegs:
            if p and p.NumberOfNodes > 1: # final double check that we don't create a single node line
                l = wv.Add(Linestring)
                l.Layer = self.layerpicker.SelectedSerialNumber
                l.Append(p, None, False, False)

        # cleanup the arrays, otherwise it could happen that we drag unwanted stuff into the next line, if it is a multiline text
        polysegs.Clear()

        # old code manually connecting line tuples
        #linestart=Point3D()
        #lineend=Point3D()
        #
        #while linetuples.Count>0:  # we've got at least one tuple/line
        #    
        #    if linetuples.Count==1: # if it's just one we make it simple
        #        l = wv.Add(clr.GetClrType(Linestring))
        #        l.Layer = layer_sn
        #        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
        #        e.Position = linetuples[0][0]
        #        l.AppendElement(e)
        #        e.Position = linetuples[0][1]
        #        l.AppendElement(e)
        #        linetuples.RemoveAt(0)
        #    
        #    else:
        #        # if we have more we start with adding the first line and store start/end
        #        l = wv.Add(clr.GetClrType(Linestring))
        #        l.Layer = layer_sn
        #        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
        #        e.Position = linetuples[0][0]
        #        linestart=linetuples[0][0]
        #        l.AppendElement(e)
        #        e.Position = linetuples[0][1]
        #        lineend=linetuples[0][1]
        #        l.AppendElement(e)
        #        linetuples.RemoveAt(0) # and delete it from the list
        #
        #        # and now we try to find tuples that are connected to this one
        #        foundprevious=True # we use this to jump over voids, if our list still contains tuples, 
        #                           # but we couldn't find one that is attached we better start a new line
        #                           # we have to set it True here since we just started a new line and want to find attached tuples
        #
        #        while linetuples.Count>0 and foundprevious==True: # if our tuple list still contains values, and we did find an attached one in the previous loop,
        #                                                          # we try to find another one
        #                                                          # if we can't find another one, but the list contains values,
        #                                                          # we revert to the loop above and start a new line
        #            
        #            foundprevious=False
        #            
        #            for i in range(0,linetuples.Count): 
        #
        #                if linestart==linetuples[i][0]: # a tuple that connects to the start of the line
        #                   e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
        #                   e.Position = linetuples[i][1]                        
        #                   l.InsertElementAt(e,0)
        #                   linestart=linetuples[i][1]
        #                   linetuples.RemoveAt(i)
        #                   foundprevious=True # we set this to true in order to keep the loop running and try to find another attached tuple
        #                   break # we stop searching since there can't be another one in this application
        #
        #                if linestart==linetuples[i][1]: # a tuple that connects to the start of the line
        #                   e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
        #                   e.Position = linetuples[i][0]                        
        #                   l.InsertElementAt(e,0)
        #                   linestart=linetuples[i][0]
        #                   linetuples.RemoveAt(i)
        #                   foundprevious=True # we set this to true in order to keep the loop running and try to find another attached tuple
        #                   break # we stop searching since there can't be another one in this application
        #
        #                if lineend==linetuples[i][0]: # a tuple that connects to the end of the line
        #                   e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
        #                   e.Position = linetuples[i][1]                        
        #                   l.AppendElement(e)
        #                   lineend=linetuples[i][1]
        #                   linetuples.RemoveAt(i)
        #                   foundprevious=True # we set this to true in order to keep the loop running and try to find another attached tuple
        #                   break # we stop searching since there can't be another one in this application
        #
        #                if lineend==linetuples[i][1]: # a tuple that connects to the end of the line
        #                   e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
        #                   e.Position = linetuples[i][0]                        
        #                   l.AppendElement(e)
        #                   lineend=linetuples[i][0]
        #                   linetuples.RemoveAt(i)
        #                   foundprevious=True # we set this to true in order to keep the loop running and try to find another attached tuple
        #                   break # we stop searching since there can't be another one in this application
        #
        #
                       
                
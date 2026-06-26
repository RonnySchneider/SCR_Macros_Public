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
    "firstviewpicker":  "",
    "secondviewpicker": "",
    "thirdviewenabled": False,
    "thirdviewpicker":  "",
    "lockenabled":      False,
    "lockscale":        False,
}

def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_LockViews"
    cmdData.CommandName = "SCR_LockViews"
    cmdData.Caption = "_SCR_LockViews"
    #cmdData.UIForm = "SCR_LockViews"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
                                                      # if you enable or disable this line, you MUST restart TBC
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.Version = 1.07
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.DefaultTabKey = "SCR Expld-SNR-Relay-Prop"
        cmdData.DefaultTabGroupKey = "View Tools"
        cmdData.ShortCaption = "Lock Views"
        cmdData.ToolTipTitle = "Lock Plan Views"
        cmdData.ToolTipTextFormatted = "move/zoom 2 Planviews with different Viewfilters simultanously"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3


    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass

def Execute(cmd, currentProject, macroFileFolder, parameters):
    form = SCR_LockViewsDialog(currentProject, macroFileFolder).Show()
    return
    # .Show() - is non modal - you can interact with the drawing window
    # .ShowDialog() - is modal - you CAN NOT interact with the drawing window

class SCR_LockViewsDialog(Window): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):

        with StreamReader(macroFileFolder + r"\SCR_LockViews.xaml") as s:
            wpf.LoadComponent(self, s) 

        ElementHost.EnableModelessKeyboardInterop(self)
        
        self.currentProject = currentProject
        self.macroFileFolder = macroFileFolder

        self.ViewOverlay = Guid.NewGuid()
        self.overlayBag = OverlayBag(self.ViewOverlay)

        self.wv = self.currentProject[Project.FixedSerial.WorldView]
        ModelEvents.ProjectUnloading += self.projectischanging # close macro on project change
        self.avm = TrimbleOffice.TheOffice.MainWindow.AppViewManager
        #self.avm.ViewAdded += self.viewchangeevent
        UIEvents.UIViewFilterChanged += self.viewchangeevent
        #UIEvents.UIViewFilterChanged += self.testevent
        
        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear

        self.v1 = None
        self.v2 = None
        self.v3 = None


        #self.firstviewpicker.FilterByEntityTypes = Array[Type]([clr.GetClrType(Hoops2dView)])
        #self.secondviewpicker.FilterByEntityTypes = Array[Type]([clr.GetClrType(Hoops2dView)])
        self.thirdviewenabled.Checked += self.viewpickerchangeevent
        self.thirdviewenabled.Unchecked += self.viewpickerchangeevent
        self.lockenabled.Checked += self.lockchangeevent
        self.lockenabled.Unchecked += self.lockchangeevent
        self.lockscale.Checked += self.lockchangeevent
        self.lockscale.Unchecked += self.lockchangeevent
        
        self.firstviewpicker.SelectionChanged += self.viewpickerchangeevent
        self.secondviewpicker.SelectionChanged += self.viewpickerchangeevent
        self.thirdviewpicker.SelectionChanged += self.viewpickerchangeevent
        
        self.firstviewpicker.DropDownOpened += self.updateviewpickers
        self.secondviewpicker.DropDownOpened += self.updateviewpickers
        self.thirdviewpicker.DropDownOpened += self.updateviewpickers
        
        self.colorpicker.ValueChanged += self.lockchangeevent

        #UIEvents.AfterDataProcessing += self.changeevent

        self.time1 = datetime.now()

        self.Loaded += self.SetDefaultOptions
        self.Closing += self.SaveOptions

    def testevent(self, sender,e):

        tt1 = 1

    def viewpickerchangeevent(self, sender, e):

        # need to get rid of any attached event triggers before reassigning the view variables
        # they'll be reset during lockchangeevent
        #for i in range(100):
        try:
            self.v1.WindowViewChanged -= self.hoopsviewevent
            self.v2.WindowViewChanged -= self.hoopsviewevent
            self.v3.WindowViewChanged -= self.hoopsviewevent
        except:
            pass
        
        viewselectionsok = True
        try:
            if self.firstviewpicker.SelectedIndex != -1 and \
            self.firstviewpicker.SelectedItem.Content != self.secondviewpicker.SelectedItem.Content:
                pass
            else:
                viewselectionsok = False
            if self.secondviewpicker.SelectedIndex != -1 and \
            self.firstviewpicker.SelectedItem.Content != self.secondviewpicker.SelectedItem.Content:
                pass
            else:
                viewselectionsok = False
            if self.thirdviewenabled.IsChecked:
                if self.thirdviewpicker.SelectedIndex != -1 and self.thirdviewpicker.SelectedItem.Content != '':
                    pass
                else:
                    viewselectionsok = False
        except:
            viewselectionsok = False

        if viewselectionsok:
            for v in self.avm.Views:
                if v.Text == self.firstviewpicker.SelectedItem.Content:
                    self.v1 = v
                if v.Text == self.secondviewpicker.SelectedItem.Content:
                    self.v2 = v
                if self.thirdviewenabled.IsChecked:
                    if v.Text == self.thirdviewpicker.SelectedItem.Content:
                        self.v3 = v

        self.lockchangeevent(None, None)

    def viewchangeevent(self, sender, e):
        
        #self.avm.ViewAdded -= self.viewchangeevent
        UIEvents.UIViewFilterChanged -= self.viewchangeevent
        
        self.updateviewpickers("viewfilter", None)
    
        #self.avm.ViewAdded += self.viewchangeevent
        UIEvents.UIViewFilterChanged += self.viewchangeevent

    def lockchangeevent(self, sender, e):

        #for i in range(100):
        try:
            self.v1.WindowViewChanged -= self.hoopsviewevent
            self.v2.WindowViewChanged -= self.hoopsviewevent
            self.v3.WindowViewChanged -= self.hoopsviewevent
        except:
            pass
        
        if self.lockenabled.IsChecked and self.v1 and self.v2:

            self.v1.WindowViewChanged += self.hoopsviewevent
            self.v2.WindowViewChanged += self.hoopsviewevent

            if self.thirdviewenabled.IsChecked and self.v3:
                self.v3.WindowViewChanged += self.hoopsviewevent

        self.drawoverlay()
 
    def hoopsviewevent(self, sender, e):

        if self.lockenabled.IsChecked:
            
            if self.v1 and self.v2:

                # up is the plan view rotation
                # get the senders camera settings
                p1, t1, up1, w1, h1, pr1 = sender.GetCamera()
                
                #tt = p1 + sender.PageToWorld.Translate
                #tt2 = 1

                #tt = sender.ComputeCenterPoint() # is suppossed to be world coord, but isnt
                #tt2 = sender.TransformScreenToWorld(sender.ComputeCenterPoint()) # is suppossed to be world coord, but isnt
                #tt3 = sender.GetWorldViewContainer() # is suppossed to be world coord, but isnt
                #tt4 = sender.GetBounds()
                #tt5 = sender.DebugDrawGraphicsInfo()
                #tt5 = sender.GetCameraField()
                #
                #tt6 = self.v1.View.PageToWorld
                #tt7 = self.v2.View.PageToWorld


                # for some weird reason the computecenterpoint and transforming the camera point
                # don't work properly in a project with proper coordinate system
                # everything was shifted by several metres and drifting all over the place
                bl = sender.TransformScreenToWorld(Point3D(sender.Left,  sender.Bottom))
                tr = sender.TransformScreenToWorld(Point3D(sender.Right,  sender.Top))
                c = Point3D.MidPoint(bl, tr)

                #tt11 = sender.GetCamera()


                self.v1.WindowViewChanged -= self.hoopsviewevent
                self.v2.WindowViewChanged -= self.hoopsviewevent
                try:
                    self.v3.WindowViewChanged -= self.hoopsviewevent
                except:
                    pass

                # go through the TBC views
                # and if it's not the sender itself then set the camera values to those of the sender
                for v in self.avm.Views:

                    if isinstance(v, Hoops2dView) and v.Text != sender.Parent.Text:

                        # don't not to worry about not existing views here since compare names with existing ones
                        # in case a self.v1-3 doesn't exist, is "", we won't have an issue
                        # we only need to check in drawoverlay
                        if v.Text == self.v1.Text or v.Text == self.v2.Text or \
                           (self.thirdviewenabled.IsChecked and self.v3 and v.Text == self.v3.Text):
                
                            p2, t2, up2, w2, h2, pr2 = v.Viewing.GetCamera()
                            if self.lockscale.IsChecked:
                                # keep the plan view rotation as it is, but adopt the scale from the sender
                                #self.wv.PauseGraphicsCache(True)
                                v.Viewing.SetCamera(p2, t2, up2, w1, h1, pr2)
                                #v.Viewing.GridSpacing = sender.GridSpacing
                                v.Viewing.CenterView(c)
                                #self.wv.PauseGraphicsCache(False)
                            else:
                                #v.Viewing.SetCamera(p1, t1, up2, w2, h2, pr2)
                                v.Viewing.CenterView(c)
                
                #tt5 = self.v2.Viewing.DebugDrawGraphicsInfo()
                
                self.v1.WindowViewChanged += self.hoopsviewevent
                self.v2.WindowViewChanged += self.hoopsviewevent
                if self.thirdviewenabled.IsChecked:
                    if self.v3:
                        self.v3.WindowViewChanged += self.hoopsviewevent

            self.drawoverlay()

    def updateviewpickers(self, sender, e):

        self.firstviewpicker.SelectionChanged -= self.viewpickerchangeevent
        self.secondviewpicker.SelectionChanged -= self.viewpickerchangeevent
        self.thirdviewpicker.SelectionChanged -= self.viewpickerchangeevent

        if isinstance(sender, ComboBox): # if it's triggered by opening the drop down list
            self.fillviewpicker(sender, False) # we don't want to set the index in this case, could lead to an endless -1 loop
        else:
            self.fillviewpicker(self.firstviewpicker, True) # only set the index after a global update, i.e. viewfilter change
            self.fillviewpicker(self.secondviewpicker, True)
            self.fillviewpicker(self.thirdviewpicker, True)

        self.firstviewpicker.SelectionChanged += self.viewpickerchangeevent
        self.secondviewpicker.SelectionChanged += self.viewpickerchangeevent
        self.thirdviewpicker.SelectionChanged += self.viewpickerchangeevent

        self.viewpickerchangeevent(None, None)

    def fillviewpicker(self, picker, setindex):
        
        oldindex = -1
        oldindex = picker.SelectedIndex
        
        picker.SelectionChanged -= self.viewpickerchangeevent
        
        picker.Items.Clear()
        for v in self.avm.Views:
            if isinstance(v, Hoops2dView):
                item = ComboBoxItem()
                item.Content = v.Text
                item.FontSize = 12
                picker.Items.Add(item)

        if setindex:
            try: picker.SelectedIndex = oldindex
            except: pass

        picker.SelectionChanged += self.viewpickerchangeevent
        return
    
    def computeviewcorners(self, v):

        # [bl1, br1, tr1, tl1, bl1]
        frame1 = List[Point3D]()
        frame1.Add(v.View.TransformScreenToWorld(Point3D(v.View.Left,  v.View.Bottom))  )
        frame1.Add(v.View.TransformScreenToWorld(Point3D(v.View.Right, v.View.Bottom))  )
        frame1.Add(v.View.TransformScreenToWorld(Point3D(v.View.Right, v.View.Top))     )
        frame1.Add(v.View.TransformScreenToWorld(Point3D(v.View.Left,  v.View.Top))     )
        frame1.Add(v.View.TransformScreenToWorld(Point3D(v.View.Left,  v.View.Bottom))  )

        return frame1

    def drawoverlay(self):

        # don't redraw overlay too often
        #tt = datetime.now() - time1 # delivers .days .microseconds . seconds
        if (datetime.now() - self.time1).microseconds > 1/30: # FPS
        
            #wv = self.currentProject [Project.FixedSerial.WorldView]
            TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)
            self.overlayBag = OverlayBag(self.ViewOverlay) # create a new/empty overlaybag
            
            if self.lockenabled.IsChecked and not self.lockscale.IsChecked:
                if self.v1 and self.v2:
                    if self.v1.Text != "" and self.v2.Text != "": # need to make sure not to write to a view that might not exist anymore

                        frames = []
                        areas = []

                        frames.Add(self.computeviewcorners(self.v1))
                        seg1 = PolySeg.PolySeg()
                        seg1.Add(frames[0])
                        a1, c1, d1 = seg1.Area()
                        areas.Add(a1)

                        frames.Add(self.computeviewcorners(self.v2))
                        seg2 = PolySeg.PolySeg()
                        seg2.Add(frames[1])
                        a2, c2, d2 = seg2.Area()
                        areas.Add(a2)

                        if self.thirdviewenabled.IsChecked and self.v3:
                            if self.v3.Text != "":
                                frames.Add(self.computeviewcorners(self.v3))
                                seg3 = PolySeg.PolySeg()
                                seg3.Add(frames[2])
                                a3, c3, d3 = seg3.Area()       
                                areas.Add(a3)

                        mina = areas.index(min(areas)) # index if minimum area, used to address frame coordinate list
                        color = self.colorpicker.SelectedColor.ToArgb()

                        self.overlayBag.AddPolyline(Array[Point3D](frames[mina]), color, 3)
                        self.overlayBag.AddPolyline(Array[Point3D]([Point3D.MidPoint(frames[mina][0], frames[mina][3]), Point3D.MidPoint(frames[mina][1], frames[mina][2])]), color, 1)
                        self.overlayBag.AddPolyline(Array[Point3D]([Point3D.MidPoint(frames[mina][0], frames[mina][1]), Point3D.MidPoint(frames[mina][2], frames[mina][3])]), color, 1)
                        
                        self.v1.AddOverlayGeometry(self.overlayBag)
                        self.v2.AddOverlayGeometry(self.overlayBag)
                        if self.thirdviewenabled.IsChecked and self.v3:
                            if self.v3.Text != "":
                                self.v3.AddOverlayGeometry(self.overlayBag)

                        # if it only needs to be visible in all Planview then remove the Hoops3DViewGUID
                        #array = Array[Guid]([DisplayWindow.HoopsPlanViewGUID])
                        #TrimbleOffice.TheOffice.MainWindow.AppViewManager.AddOverlayGeometry(array, self.overlayBag)
        
            self.time1 = datetime.now()
        return

    def projectischanging(self, sender, e):
        self.Close()

    def Dispose(self, thisCmd, disposing):
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)

    def cleanupeventtriggers(self):

        self.thirdviewenabled.Checked -= self.viewpickerchangeevent
        self.thirdviewenabled.Unchecked -= self.viewpickerchangeevent
        self.lockenabled.Checked -= self.lockchangeevent
        self.lockenabled.Unchecked -= self.lockchangeevent
        self.lockscale.Checked -= self.lockchangeevent
        self.lockscale.Unchecked -= self.lockchangeevent
        
        self.firstviewpicker.SelectionChanged -= self.viewpickerchangeevent
        self.secondviewpicker.SelectionChanged -= self.viewpickerchangeevent
        self.thirdviewpicker.SelectionChanged -= self.viewpickerchangeevent
        self.colorpicker.ValueChanged -= self.lockchangeevent

        for i in range(100):
            try:
                self.v1.WindowViewChanged -= self.hoopsviewevent
                self.v2.WindowViewChanged -= self.hoopsviewevent
                self.v3.WindowViewChanged -= self.hoopsviewevent
            except:
                pass

        self.firstviewpicker.DropDownOpened -= self.updateviewpickers
        self.secondviewpicker.DropDownOpened -= self.updateviewpickers
        self.thirdviewpicker.DropDownOpened -= self.updateviewpickers

        ModelEvents.ProjectUnloading -= self.projectischanging
       # UIEvents.AfterDataProcessing -= self.changeevent
        UIEvents.UIViewFilterChanged -= self.viewchangeevent
        #self.avm.ViewAdded -= self.viewchangeevent
        #UIEvents.UIViewFilterChanged -= self.testevent

    def SetDefaultOptions(self, sender, e):
        SCROptions.LoadWindowState(self, "SCR_LockViews", default_width=400)
        self.updateviewpickers(None, None)
        SCROptions.LoadMacroOptions(self, "SCR_LockViews", _OPTIONS)
        try:    self.colorpicker.SelectedColor = Color.FromArgb(OptionsManager.GetInt("SCR_LockViews.colorpicker"))
        except: self.colorpicker.SelectedColor = Color.Purple
        self.viewpickerchangeevent(None, None)

    def SaveOptions(self, sender, e):
        SCROptions.SaveWindowState(self, "SCR_LockViews")
        SCROptions.SaveMacroOptions(self, "SCR_LockViews", _OPTIONS)
        self.cleanupeventtriggers()
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)
        OptionsManager.SetValue("SCR_LockViews.colorpicker", self.colorpicker.SelectedColor.ToArgb())


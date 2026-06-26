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

def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_DuplicateSurface"
    cmdData.CommandName = "SCR_DuplicateSurface"
    cmdData.Caption = "_SCR_DuplicateSurface"
    cmdData.UIForm = "SCR_DuplicateSurface"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "DTM"
        cmdData.ShortCaption = "Duplicate DTM"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.02
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Duplicate a Surface"
        cmdData.ToolTipTextFormatted = "Duplicate a Surface, including Composites"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_DuplicateSurface(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_DuplicateSurface.xaml") as s:
            wpf.LoadComponent (self, s)
        self.currentProject = currentProject
        self.macroFileFolder = macroFileFolder

        self.ViewOverlay = Guid.NewGuid()
        self.overlayBag = OverlayBag(self.ViewOverlay)

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


		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass

    def IsValid(self, serial):
        
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, Model3D) and not isinstance(o, clr.GetClrType(ProjectedSurface)):
            return True
        return False


    def SetDefaultOptions(self):
        pass

    def SaveOptions(self):
        pass

    def CancelClicked(self, thisCmd, args):
        thisCmd.CloseUICommand()

    def Dispose(self, cmd, disposing):
        TrimbleOffice.TheOffice.MainWindow.AppViewManager.RemoveOverlayGeometry(self.ViewOverlay)


    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''

        wv = self.currentProject [Project.FixedSerial.WorldView]
        dp = self.currentProject.CreateDuplicator()

        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
        try:
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                for o in self.objs:
                    
                    newsurface = None
                    
                    if isinstance(o, clr.GetClrType(CompositeSurface)):

                        newsurface = o.GetSite().Add(CompositeSurface)
                        newsurface.Name = CompositeSurface.GetUniqueName(o.Name + ' - copy', None, o.GetSite())

                        for sn in reversed(o.SourceSurfacesSerials):
                            newsurface.AddSourceSurface(0, sn)


                    elif isinstance(o, clr.GetClrType(Model3D)) and not isinstance(o, clr.GetClrType(ProjectedSurface)):

                        newsurface = o.GetSite().Add(clr.GetClrType(Model3D))
                        newsurface.Name = Model3D.GetUniqueName(o.Name, None, o.GetSite())
                        newsurface.Gem.CopyFrom(o.Gem)
                        newsurface.Gem.CopySettingsFrom(o.Gem)
                        #newsurface.CopyBody(self.currentProject.Concordance, self.currentProject.TransactionManager, o, dp)           

                        #newsurfacebuilder = newsurface.GetGemBatchBuilder()
                        #
                        ## use a copy of the original surface data to create a new surface
                        #tmpgem = o.GemCopy
                        #tmpgem.External = False # make all data internal
                        #tmpgem.IsLimited = False # recompute the min/max limits 
                        #newsurface.Gem = tmpgem
                        #
                        #newsurfacebuilder = self.addModel3D(newsurfacebuilder, newsurface)
                        #
                        #newsurfacebuilder.Construction()
                        #newsurfacebuilder.Commit()

                    if newsurface:

                        newsurface.Mode = o.Mode
                        newsurface.Color = o.Color
                        newsurface.RebuildMethod = o.RebuildMethod
                        newsurface.TransparencyPercentage = o.TransparencyPercentage


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

        self.SaveOptions()

    def addModel3D(self, builder, surface: Model3D):
        
        mapVertices = {}
        nVertices = surface.NumberOfVertices
        nTri = surface.NumberOfTriangles
        for i in range(nVertices):
            if not surface.IsVertexPresent(i):
                continue
            if not surface.Gem.IsVertexTriangulated(i):
                continue
            p = surface.GetVertexPoint(i)
            v = builder.AddVertex(p)
            mapVertices.Add(i, v[0])
        for t in range(nTri):
            if not surface.IsTriangleMaterialPresent(t):
                continue
            for side in range(3):
                isOuter = surface.GetTriangleOuterSide(t, side)
                if not isOuter:
                    (ok, tAdj, sideAdj) = surface.GetTriangleAdjacent(t, side)
                    if not surface.IsTriangleMaterialPresent(tAdj):
                        isOuter = True # treat edges next to null as valid
                iVertexA = surface.GetTriangleIVertex(t, side)
                nextSide = side + 1
                if nextSide == 3:
                    nextSide = 0
                iVertexB = surface.GetTriangleIVertex(t, nextSide)
                if isOuter or iVertexA < iVertexB:
                    b = DTMSharpness.eSoft
                    if isOuter:
                        b = DTMSharpness.eSharpAndTextureBndy
                    builder.AddBreakline(Byte(b), mapVertices[iVertexA], mapVertices[iVertexB])
        
        return builder


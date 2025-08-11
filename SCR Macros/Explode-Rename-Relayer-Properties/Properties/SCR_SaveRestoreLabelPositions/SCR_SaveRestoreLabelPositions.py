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
    cmdData.Key = "SCR_SaveRestoreLabelPositions"
    cmdData.CommandName = "SCR_SaveRestoreLabelPositions"
    cmdData.Caption = "_SCR_SaveRestoreLabelPositions"
    cmdData.UIForm = "SCR_SaveRestoreLabelPositions"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Expld-SNR-Relay-Prop"
        cmdData.DefaultTabGroupKey = "Properties"
        cmdData.ShortCaption = "Save or Restore Point Label Positions"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.02
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Save or Restore Point Label Positions"
        cmdData.ToolTipTextFormatted = "Save or Restore Point Label Positions"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png") 
        cmdData.ImageSmall = b
    except:
        pass


class SCR_SaveRestoreLabelPositions(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_SaveRestoreLabelPositions.xaml") as s:
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


    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''


        cmd.CloseUICommand ()

    
    def openbutton_Click(self, sender, e):
        wv = self.currentProject [Project.FixedSerial.WorldView]
        wv.PauseGraphicsCache(True)


        # filename = os.path.expanduser('~/Downloads/LabelPositions.csv')
        filename = os.path.dirname(self.currentProject.ProjectName) + '/LabelPositions.csv'

       
        if File.Exists(filename): # if file still exists
            
            labellist=[]

            with open(filename, 'r') as csvfile: 
                reader = csv.reader(csvfile, delimiter=',') 
                for row in reader:
                    labellist.Add(row)
        
        for o in wv: # and try to find a matching label
            if isinstance(o, CadLabel):
                if o.IsPointLabel(): # could be an alignment label
                    
                    foundincsv = False
                    
                    for label in labellist: # we go through all lines in the CSV
                        if self.currentProject.Concordance.Lookup(o.PointSerial).AnchorName == label[0]: # we compare the name
                            
                            foundincsv = True

                            anchor = Point3D()
                            anchor.X = float(label[1])
                            anchor.Y = float(label[2])
                            anchor.Z = float(label[3])
                            delta = Point3D()
                            delta.X = float(label[4])
                            delta.Y = float(label[5])
                            delta.Z = float(label[6])
                            point0 = Point3D()
                            point0.X = float(label[7])
                            point0.Y = float(label[8])
                            point0.Z = float(label[9])
                            hasleader = label[10]

                            #o.ObjectAnchorPosition = anchor
                            o.Delta = delta
                            o.Point0 = point0

                            if hasleader == "True":

                                leaderpointcount = int((label.Count - 1 - 10) / 3)
                                leaderpoints = List[Point3D]()
                                for i in range(0, leaderpointcount):
                                    tmp = Point3D()
                                    tmp.X = float(label[11 + i*3 + 0])
                                    tmp.Y = float(label[11 + i*3 + 1])
                                    tmp.Z = float(label[11 + i*3 + 2])
                                    leaderpoints.Add(tmp)

                                newleader = wv.Add(clr.GetClrType(Leader))
                                newleader.AnnotationSerial = o.SerialNumber
                                newleader.Points = leaderpoints
                                newleader.ScaleFactor = 1
                                newleader.TextGap = 1.6
                                newleader.Layer = o.Layer
                            tt = labellist
                            tt2 = tt
                    
                    if not foundincsv:
                        o.Color = Color.Magenta
                        self.success.Content = 'Magenta - Labels without Information in CSV'

        wv.PauseGraphicsCache(False)
        self.currentProject.Calculate(False)
        
    def savebutton_Click(self, sender, e):
        wv = self.currentProject [Project.FixedSerial.WorldView]

        # filename = os.path.expanduser('~/Downloads/LabelPositions.csv')
        filename = os.path.dirname(self.currentProject.ProjectName) + '/LabelPositions.csv'

        if File.Exists(filename):
                File.Delete(filename)
        
        with open(filename, 'w') as f:            
        
            for o in wv:
                
                # Label.ObjectAnchorPosition - coordinate of the point itself
                # Label.Point0
        
                if isinstance(o, CadLabel):
                    if o.IsPointLabel(): # could be an alignment label
                        pointname = self.currentProject.Concordance.Lookup(o.PointSerial).AnchorName
                        anchor = o.ObjectAnchorPosition
                        delta = o.Delta
                        point0 = o.Point0
        
                        leader = Leader.GetRelatedLeader(self.currentProject, o.SerialNumber)
                        if leader:
                            hasleader = True
                        else:
                            hasleader = False
                        
                        # we compile a string with all the information
                        outputstring =  pointname + ',' + \
                                        str(anchor.X) + ',' + str(anchor.Y) + ',' + str(anchor.Z) + ',' + \
                                        str(delta.X) + ',' + str(delta.Y) + ',' + str(delta.Z) + ',' + \
                                        str(point0.X) + ',' + str(point0.Y) + ',' + str(point0.Z) + ',' + \
                                        str(hasleader)

                        # if it has a leader line we also store the points of that one
                        if hasleader:
                            for p in leader.Points:
                                outputstring += ',' + str(p.X) + ',' + str(p.Y) + ',' + str(p.Z)

                        f.write(outputstring + "\n") # we write the string to the file
                        
            self.success.Content +='data written to file'
        
        f.close()

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
    "colonmodeoff": True,
    "colonmodeon":  False,
}

def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_ExportPointFC2CSV"
    cmdData.CommandName = "SCR_ExportPointFC2CSV"
    cmdData.Caption = "_SCR_ExportPointFC2CSV"
    cmdData.UIForm = "SCR_ExportPointFC2CSV"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "Export"
        cmdData.ShortCaption = "Point FC to CSV"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.06
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Export Point Features into CSV Columns"
        cmdData.ToolTipTextFormatted = "Export Point Features into CSV Columns"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_ExportPointFC2CSV(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_ExportPointFC2CSV.xaml") as s:
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


        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        self.lfp = self.lunits.Properties.Copy()
        self.lfp.AddSuffix = False
        linearsuffix = self.lunits.Units[self.lunits.DisplayType].Abbreviation
        #self.hztollabel.Content = "horizontal tolerance [" + linearsuffix + "] <="
        #self.vztollabel.Content = "vertical tolerance [" + linearsuffix + "] <="


        self.SetDefaultOptions()

    def IsValid(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, CoordPoint):
            return True
        return False


    def SetDefaultOptions(self):
        SCROptions.LoadMacroOptions(self, "SCR_ExportPointFC2CSV", _OPTIONS)

    def SaveOptions(self):
        SCROptions.SaveMacroOptions(self, "SCR_ExportPointFC2CSV", _OPTIONS)

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def OkClicked(self, cmd, e):

        Keyboard.Focus(self.okBtn)
        self.error.Content=''
        self.success.Content = ''


        wv = self.currentProject [Project.FixedSerial.WorldView]
        wv.PauseGraphicsCache(True)

        for o in self.currentProject:
        #find PointManager as object
            if isinstance(o, PointManager):
                pm = o
        
        codesdict = {}
        pointdatadict = {}
        
        # go through the selected point and save their FC and attributes in a dictionary
        # while doing that create a dictionary with all FC's and attributes
        # need this to group FC-columns in the header
        for o in self.objs:

            if isinstance(o, CoordPoint):
                features = pm.AssociatedRDFeatures(o.SerialNumber)

                pointdatadict.update({o.Name : {}}) # add point name to pointdata

                pointdatadict[o.Name].update({"FC" : {}}) # add FC key to pointdata
                pointdatadict[o.Name].update({"Coords" : {}}) # add Coords key to pointdata
                pointdatadict[o.Name]["Coords"].update({"X" : self.lunits.Format(o.Position.X, self.lfp)}) # add coords to Coords key
                pointdatadict[o.Name]["Coords"].update({"Y" : self.lunits.Format(o.Position.Y, self.lfp)}) # add coords to Coords key
                pointdatadict[o.Name]["Coords"].update({"Z" : self.lunits.Format(o.Position.Z, self.lfp)}) # add coords to Coords key
                
                # prepare dictionary for header and later lookup
                for fc in features:

                    # add the FC names to the dictionaries
                    codesdict.update({fc.Code : {}})
                    pointdatadict[o.Name]["FC"].update({fc.Code : {}})

                    # add the attribute names to the dictionaries; for the pointdata add the values as well
                    for att in fc.Attributes:

                        codesdict[fc.Code].update({att.Name : ""})
                        pointdatadict[o.Name]["FC"][fc.Code].update({att.Name : att.Value})

        # now prepare the outputline per point
        headerstring = 'Pointnumber'
        headercomplete = False
        pointstring = ''
        
        # get the point number and add basic data to the pointstring
        # then go through the combined FC dictionary and try to read the attribute value from the point
        # if the point doesn't have the attribute add just a comma
        # this way we get a properly structured and FC/attribute separated output
        # dictionaries don't stay properly sorted - need to use sorted lookup list while looping through them
        for p in sorted(pointdatadict.items()):

            pointstring += p[0]

            pointstring += "," + pointdatadict.get(p[0], {}).get("Coords", {}).get("X", "NA")
            pointstring += "," + pointdatadict.get(p[0], {}).get("Coords", {}).get("Y", "NA")
            pointstring += "," + pointdatadict.get(p[0], {}).get("Coords", {}).get("Z", "NA")

            if not headercomplete:
                headerstring += ",Easting"
                headerstring += ",Northing"
                headerstring += ",Elevation"

            for code in sorted(codesdict.items()):

                # applies only if colon mode is off
                # need to add an extra column with the header description
                # and if the point has that code enter the value into the column
                if self.colonmodeoff.IsChecked:
                    if not headercomplete:
                        headerstring += ",Featurecode"
                    
                    pointstring += ","
                    if str(pointdatadict.get(p[0], {}).get("FC", {}).get(code[0], "NA")) != "NA":
                        pointstring += code[0]
                
                else:
                    # in case we have a code without attributes the for loop below wouldn't run
                    # and we'd lose the information
                    if codesdict[code[0]].Count == 0:
                        if not headercomplete:
                            headerstring += "," + code[0] + ":"
                        pointstring += ","
                                 
                for att in sorted(codesdict[code[0]].items()):
                    
                    if not headercomplete:
                        if self.colonmodeon.IsChecked:
                            headerstring += "," + code[0] + ":" + att[0]
                        else:
                           headerstring += "," + att[0]

                    pointstring += ","

                    # try to read the attribute from the point dictionary
                    # will add an empty string if not existent
                    # a standalone comma will just add an empty column - exactly what we want
                    pointstring += str(pointdatadict.get(p[0], {}).get("FC", {}).get(code[0], {}).get(att[0], ""))

            pointstring += "\n"
            headercomplete = True

        headerstring += "\n"


        tt2 = 321
        #featurelist = list(set(featurelist))
        #featurelist.sort()
                
        #codesdict = {"P1": {
        #                
        #                "fc1" : {
        #                    "att1" : "fjksdfh", "att2" : "984732", "att3" : "h342k"
        #                    },
        #                "fc2" : {
        #                    "att1" : "3232", "att2" : "984732", "att3" : "h342k"
        #                        }
        #                }                       
        #            }
        filename = os.path.expanduser('~/Downloads/PointFC-Output.csv')
        try:
            if File.Exists(filename):
                File.Delete(filename)
            with open(filename,'w', newline='') as f:
                f.write(headerstring)
                f.write(pointstring)
        except Exception as e:
            tt = sys.exc_info()
            exc_type, exc_obj, exc_tb = sys.exc_info()
            self.error.Content += '\nan Error occurred - Result probably incomplete\n' + str(exc_type) + '\n' + str(exc_obj) + '\nLine ' + str(exc_tb.tb_lineno)



        self.SaveOptions()           
        
        ProgressBar.TBC_ProgressBar.Title = ""
        
        wv.PauseGraphicsCache(False)


    

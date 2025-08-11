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
    cmdData.Key = "SCR_CSVtoTable"
    cmdData.CommandName = "SCR_CSVtoTable"
    cmdData.Caption = "_SCR_CSVtoTable"
    cmdData.UIForm = "SCR_CSVtoTable"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "Import/Export"
        cmdData.ShortCaption = "CSV to Table"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.02
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Import CSV-Content into Table"
        cmdData.ToolTipTextFormatted = "Import CSV-Content into Table"

    except:
        pass
    
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_CSVtoTable(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_CSVtoTable.xaml") as s:
            wpf.LoadComponent (self, s)
        self.currentProject = currentProject
        self.macroFileFolder = macroFileFolder
        if self.openfilename.Text=='': self.openfilename.Text = macroFileFolder

    def HelpClicked(self, cmd, e):
        webbrowser.open("C:\ProgramData\Trimble\MacroCommands3\SCR Macros\MacroHelp\MacroHelp.htm#" + type(self).__name__)

    def OnLoad(self, cmd, buttons, event):
        self.okBtn = buttons[0]
        buttons[2].Content = "Help"
        buttons[2].Visibility = Visibility.Visible
        buttons[2].Click += self.HelpClicked
        self.Caption = cmd.Command.Caption

		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block
        try:
            self.SetDefaultOptions()
        except:
            pass
        self.unitssetup(None, None)

    def unitssetup(self, sender, e):
        # setup everything for the unit conversions

        # get the units for linear distance
        self.lunits = self.currentProject.Units.Linear
        self.lfp = self.lunits.Properties.Copy() # create a copy in order to set the decimals and enable/disable the suffix
        self.lfp.AddSuffix = False # disable suffix, we need to set it manually, it would always add the current projects units

        self.unitschanged(None, None)
    
    def unitschanged(self, sender, e):

        # loop through all objects of self and set the properties for all DistanceEdits
        # the code is slower than doing it manually for each single one
        # but more convenient since we don't have to worry about how many DistanceEdit Controls we have in the UI
        tt = self.__dict__.items()
        for i in self.__dict__.items():
            if i[1].GetType() == TBCWpf.DistanceEdit().GetType():
                i[1].DisplayUnit = LinearType(self.lunits.DisplayType)
                i[1].ShowControlIcon(False)
                i[1].FormatProperty.AddSuffix = ControlBoolean(1)
                #i[1].FormatProperty.NumberOfDecimals = self.lfp.NumberOfDecimals
    
    def SetDefaultOptions(self):
        self.openfilename.Text = OptionsManager.GetString("SCR_CSVtoTable.openfilename", os.path.expanduser('~\\Downloads'))
        settingserial = OptionsManager.GetUint("SCR_CSVtoTable.textstylepicker", 24) # 24 is FixedSerial for Standard text style
        o = self.currentProject.Concordance.Lookup(settingserial)
        if o != None:
            if isinstance(o.GetSite(), TextStyleCollection):    
                self.textstylepicker.SetSelectedSerialNumber(settingserial, InputMethod(3))
            else:                       
                self.textstylepicker.SetSelectedSerialNumber(24, InputMethod(3))
        else:                       
            self.textstylepicker.SetSelectedSerialNumber(24, InputMethod(3))
        self.textheight.Distance = OptionsManager.GetDouble("SCR_CSVtoTable.textheight", 0.1)
        self.textweightpicker.Lineweight = OptionsManager.GetInt("SCR_CSVtoTable.textweightpicker", 0)
        self.marginlr.Distance = OptionsManager.GetDouble("SCR_CSVtoTable.marginlr", 0.5)
        self.margintb.Distance = OptionsManager.GetDouble("SCR_CSVtoTable.margintb", 0.2)
        self.drawrectangle.IsChecked = OptionsManager.GetBool("SCR_CSVtoTable.drawrectangle", True)

        settingserial = OptionsManager.GetUint("SCR_CSVtoTable.layerpicker", 8) # 8 is FixedSerial for Layer Zero
        o = self.currentProject.Concordance.Lookup(settingserial)
        if o != None:
            if isinstance(o.GetSite(), LayerCollection):    
                self.layerpicker.SetSelectedSerialNumber(settingserial, InputMethod(3))
            else:                       
                self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
        else:                       
            self.layerpicker.SetSelectedSerialNumber(8, InputMethod(3))
            
    def SaveOptions(self):
        OptionsManager.SetValue("SCR_CSVtoTable.openfilename", self.openfilename.Text)
        OptionsManager.SetValue("SCR_CSVtoTable.textstylepicker", self.textstylepicker.SelectedSerialNumber)
        OptionsManager.SetValue("SCR_CSVtoTable.textheight", self.textheight.Distance)
        OptionsManager.SetValue("SCR_CSVtoTable.textweightpicker", self.textweightpicker.Lineweight)
        OptionsManager.SetValue("SCR_CSVtoTable.marginlr", self.marginlr.Distance)
        OptionsManager.SetValue("SCR_CSVtoTable.margintb", self.margintb.Distance)
        OptionsManager.SetValue("SCR_CSVtoTable.drawrectangle", self.drawrectangle.IsChecked)
        OptionsManager.SetValue("SCR_CSVtoTable.layerpicker", self.layerpicker.SelectedSerialNumber)

    def IsValid(self, serial):
        o=self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.lType):
            return True
        return False

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

  
    def openbutton_Click(self, sender, e):
        dialog=OpenFileDialog()
        dialog.InitialDirectory = self.openfilename.Text
        dialog.Filter=("CSV|*.csv")
        
        tt=dialog.ShowDialog()
        if tt==DialogResult.OK:
            self.openfilename.Text = dialog.FileName

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Content=''
        self.success.Content=''

        layer_sn = self.layerpicker.SelectedSerialNumber

        
        wv = self.currentProject [Project.FixedSerial.WorldView]
        inputok = True
        # we check if we can work with the input values
        textheight = abs(self.textheight.Distance)
        if textheight == 0.0:
            self.error.Content = '\nTextheight was set to Zero'

        marginlr = abs(self.marginlr.Distance)
        margintb = abs(self.margintb.Distance)
        
        topleft = self.coordpick1.Coordinate
        if math.isnan(topleft):
            inputok = False
            self.success.Content = '\nno Table Location picked'

        stationlist=[]   
        columnwidth=[]
        
        if inputok:
            self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
            UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
            try:
                # we read the CSV file into an array
                with open(self.openfilename.Text,'r') as csvfile: 
                    reader = csv.reader(csvfile, delimiter=',', quotechar='|') 
                    for row in reader:
                        stationlist.Add(row)
                # some rows might not have the same amount of values, so we have to find the max numbers in x/y
                maxrowcount = stationlist.Count
                maxcolumncount = stationlist[0].Count
                for i in range(0, maxrowcount):
                    if stationlist[i].Count > maxcolumncount: maxcolumncount = stationlist[i].Count
                
                # we create one textelement and give it our values
                t = wv.Add(clr.GetClrType(MText))
                t.TextStyleSerial=self.textstylepicker.SelectedSerialNumber
                t.Height = textheight
                t.AttachPoint = AttachmentPoint.BottomLeft
                t.Weight=self.textweightpicker.Lineweight
                
                # now we have to go through all the CSV values, give that text element that string and get us the maximum columnwidth
                for x in range(0, maxcolumncount):
                    l=0
                    for y in range(0, maxrowcount):
                        if x < stationlist[y].Count:
                            t.TextString = stationlist[y][x]
                            tt = t.MTextLength(stationlist[y][x], textheight)[0]
                            if t.MTextLength(stationlist[y][x], textheight)[0] > l: l = t.MTextLength(stationlist[y][x], textheight)[0]
                            
                    columnwidth.Add(l)  # now we have an array with the width of each column that accomodates the longest string in each column
                
                with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:

                    # draw the vertical lines of the table
                    p1 = Point3D(topleft)
                    p2 = Point3D(p1)
                    p2.Y = p2.Y - (textheight*1 + 2*margintb) * maxrowcount

                    for i in range(0, maxcolumncount):
                        if columnwidth[i] > 0:
                            p1.X = p1.X + columnwidth[i] + 2*marginlr
                        else:
                            p1.X = p1.X + 1
                        p2.X = p1.X
                        if i < maxcolumncount-1:
                            l = wv.Add(clr.GetClrType(Linestring))
                            e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                            e.Position = p1 
                            l.AppendElement(e)
                            e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                            e.Position = p2
                            l.AppendElement(e)       
                            l.Layer = self.layerpicker.SelectedSerialNumber
                            l.Weight = self.textweightpicker.Lineweight
                        bottomright = Point3D(p2)

                    # draw the horizontal lines of the table
                    p1 = Point3D(topleft)
                    p2 = Point3D(bottomright)
                    p2.Y = p1.Y

                    for i in range(0, maxrowcount):
                        p1.Y = p1.Y - (textheight*1 + 2*margintb)
                        p2.Y = p1.Y
                        if i < maxrowcount-1:
                            l = wv.Add(clr.GetClrType(Linestring))
                            e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                            e.Position = p1 
                            l.AppendElement(e)
                            e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                            e.Position = p2
                            l.AppendElement(e)
                            l.Layer = self.layerpicker.SelectedSerialNumber
                            l.Weight = self.textweightpicker.Lineweight

                    # draw the rectangle
                    if self.drawrectangle.IsChecked:
                        p1 = Point3D(topleft)
                        
                        l = wv.Add(clr.GetClrType(Linestring))
                        l.Layer = self.layerpicker.SelectedSerialNumber 
                        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                        e.Position = p1 
                        l.AppendElement(e)
                        
                        p1.X = bottomright.X
                        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                        e.Position = p1 
                        l.AppendElement(e)
                        
                        p1.Y = bottomright.Y
                        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                        e.Position = p1 
                        l.AppendElement(e)
                        
                        p1.X = topleft.X
                        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                        e.Position = p1 
                        l.AppendElement(e)
                        
                        p1.Y = topleft.Y
                        e = ElementFactory.Create(clr.GetClrType(IStraightSegment), clr.GetClrType(IXYZLocation))
                        e.Position = p1 
                        l.AppendElement(e)

                    # draw the text
                    tpos = Point3D(topleft)
                    tpos.Y = tpos.Y - margintb - textheight
                    for y in range(0, maxrowcount):
                        tpos.X = topleft.X + marginlr
                        for x in range(0, maxcolumncount):
                            if x < stationlist[y].Count:
                                if x == 0 and y == 0:
                                    t.TextString = stationlist[0][0]
                                    t.AlignmentPoint = tpos
                                    t.Layer = self.layerpicker.SelectedSerialNumber
                                    tpos.X += columnwidth[x] + marginlr + marginlr
                                else:
                                    t = wv.Add(clr.GetClrType(MText))
                                    t.TextStyleSerial=self.textstylepicker.SelectedSerialNumber
                                    t.Height = textheight
                                    t.AttachPoint = AttachmentPoint.BottomLeft
                                    t.Weight = self.textweightpicker.Lineweight
                                    t.TextString = stationlist[y][x]
                                    t.Layer = self.layerpicker.SelectedSerialNumber
                                    t.AlignmentPoint = tpos
                                    tpos.X += columnwidth[x] + marginlr + marginlr

                        tpos.Y = tpos.Y + textheight - textheight*1 - 2*margintb - textheight
                    
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

        self.success.Content='Done'
        Keyboard.Focus(self.coordpick1)
        self.SaveOptions()


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
    cmdData.Key = "SCR_Inspect12dAttributes"
    cmdData.CommandName = "SCR_Inspect12dAttributes"
    cmdData.Caption = "_SCR_Inspect12dAttributes"
    #cmdData.UIForm = "SCR_Inspect12dAttributes"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
                                                      # if you enable or disable this line, you MUST restart TBC
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.Version = 1.11  # included metre to project conversion
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.DefaultTabKey = "SCR Expld-SNR-Relay-Prop"
        cmdData.DefaultTabGroupKey = "Properties"
        cmdData.ShortCaption = "Inspect/Check 12d Attributes"
        cmdData.ToolTipTitle = "Inspect/Check 12d Attributes"
        cmdData.ToolTipTextFormatted = "Inspect/Check 12d Attributes"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3


    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass

def Execute(cmd, currentProject, macroFileFolder, parameters):
    form = SCR_Inspect12dAttributesDialog(currentProject, macroFileFolder).Show()
    return
    # .Show() - is non modal - you can interact with the drawing window
    # .ShowDialog() - is modal - you CAN NOT interact with the drawing window

class SCR_Inspect12dAttributesDialog(Window): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):

        with StreamReader(macroFileFolder + r"\SCR_Inspect12dAttributes.xaml") as s:
            wpf.LoadComponent(self, s) 

        ElementHost.EnableModelessKeyboardInterop(self)
        
        self.currentProject = currentProject
        self.macroFileFolder = macroFileFolder

    #def OnLoad(self, cmd, buttons, event):
    #    self.okBtn = buttons[0]
    #    buttons[2].Content = "Help"
    #    buttons[2].Visibility = Visibility.Visible
    #    buttons[2].Click += self.HelpClicked
    #    self.Caption = cmd.Command.Caption

        #self.CloseBtn.Click += self.Button_Click_Cancel
        self.buttonsavesel.Click += self.saveselection_Click
        self.buttonrestoresel.Click += self.restoreselection_Click
        self.buttonupdatelist.Click += self.updatelist_Click
        
        
        self.savedserials = []
        self.datagrid.SelectedCellsChanged += self.dataselectionchanged
        self.datagrid.CellEditEnding += self.cellcontentchanged
        #self.datagrid.CellEditEnding += self.findexception
    
        self.polyType = clr.GetClrType(IPolyseg)
        
        #CalculateProjectEvents.CalculateProjectEvent += self.backgroundupdateend
        ModelEvents.EntityDeleted += self.objectdeleted
        ModelEvents.ProjectUnloading += self.projectischanging # close macro on project change
        #self.deletecount = 0
        UIEvents.AfterDataProcessing += self.backgroundupdateend

        #self.objs.IsEntityValidCallback = self.IsValid
    
        optionMenu = SelectionContextMenuHandler()
        # remove options that don't apply here
        optionMenu.ExcludedCommands = "SelectObservations | SelectPoints | SelectDuplicatePoints"
        self.objs.ButtonContextMenu = optionMenu
    
		# after changing the input fields in a lot of macros from the old textboxes to floating point number or distance edits
		# it could happen that old settings, saved as strings, would throw a type cast error
		# hence it's better to have it in a try block

        self.Loaded += self.SetDefaultOptions
        self.Closing += self.SaveOptions # includes cleanupeventtriggers

    def Button_Click_Cancel(self, sender, e):
        self.Close()
        pass

    def SetDefaultOptions(self, sender, e):
        SCROptions.LoadWindowState(self, "SCR_Inspect12dAttributesDialog", default_width=1000, default_height=500)

    def SaveOptions(self, sender, e):
        SCROptions.SaveWindowState(self, "SCR_Inspect12dAttributesDialog")
        self.cleanupeventtriggers()

    def cleanupeventtriggers(self):
        ModelEvents.EntityDeleted -= self.objectdeleted
        UIEvents.AfterDataProcessing -= self.backgroundupdateend
        ModelEvents.ProjectUnloading -= self.projectischanging

    def HelpClicked(self, cmd, e):
        webbrowser.open("C:\ProgramData\Trimble\MacroCommands3\SCR Macros\MacroHelp\MacroHelp.htm#" + type(self).__name__)

    def Dispose(self, thisCmd, disposing):
        pass
        #self.SaveOptions()
    def projectischanging(self, sender, e):
        self.Close()

    def backgroundupdateend(self, sender, e):
        #tt = sender.Caption
        #if self.deletecount > 0:
        #    #VceMessageBox.Show(TrimbleOffice.TheOffice.MainWindow.Form, "you've deleted objects")
        #    self.deletecount = 0
        #    return
        
        # must avoid list update when we did something, interferes with datatable update and produces exception
        try:
            if not sender.Caption == "_SCR_Inspect12dAttributes":
                self.updatelist()
        except:
            pass
        
    def objectdeleted(self, sender, e):

        self.error.Content += "\n!!! YOU'VE DELETED AN OBJECT !!!"

    def updatelist_Click(self, sender, e):

        self.error.Content = ''
        self.error2.Content = ''
        self.updatelist()


    def restoreselection_Click(self, sender, e):

        GlobalSelection.Clear()
        GlobalSelection.Items(self.currentProject).Set(self.savedserials)

    def saveselection_Click(self, sender, e):

        self.savedserials.Clear()
        
        for o in GlobalSelection.SelectedMembers(self.currentProject):
            self.savedserials.Add(o.SerialNumber)
        
        self.savedserials.sort()
        self.savecount.Content = str(self.savedserials.Count)

        #GlobalSelection.Clear()
        self.SaveOptions(None, None)
        self.updatelist()

    def dataselectionchanged(self, sender, e):

        if self.datagrid.CurrentItem: # during the update of the list it can happen that the list is empty, hence no item can possibly be selcted
            s = self.datagrid.CurrentItem.Row.ItemArray[0]
            s = []
            for cell in self.datagrid.SelectedCells:
                s.Add(cell.Item.Row.ItemArray[0])

            GlobalSelection.Clear()
            GlobalSelection.Items(self.currentProject).Set(s)   # Set([s])

    def cellcontentchanged(self, sender, e):
        
        #self.error2.Content = ''
        uda = self.currentProject.Lookup(27) # User Defined Attributes are always Serial# 27
        changedrows = []

        UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
        self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, "Change12dAttributes")#self.Caption)
    
        try:
            # the "with" statement will unroll any changes if something go wrong
            with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
    
                for c in self.datagrid.SelectedCells:
                    crow = c.Item.Row.ItemArray[0] # selected cell row serial
                    erow = e.Row.DataContext.Row.ItemArray[0] # event sending cell row serial
                    ccol = c.Column.DisplayIndex # selected cell column
                    ecol = e.Column.DisplayIndex # event sending cell column
                    #if not (crow == erow and ccol ==ecol): # we don't want to programmatically change the cell the user is changing
                    # need to find the row index
                    for i in range(self.dt.Rows.Count):
                        if crow == self.dt.Rows[i][0]:
                            cindex = i

                    self.dt.Rows[cindex][ccol] = e.EditingElement.Text
                    changedrows.Add(cindex)

                changedrows = list(set(changedrows)) # remove duplicates


                # apply the changes to the objects
                for i in changedrows:
                    
                    dic = []
                    for j in range(1, self.dt.Columns.Count, 3):
                        if self.dt.Rows[i][j]:
                            if not self.dt.Rows[i][j + 1]:
                                self.dt.Rows[i][j + 1] = ""
                            if not self.dt.Rows[i][j + 2]:
                                self.dt.Rows[i][j + 2] = ""

                            tmp = '\"'
                            dic.Add({'ValueForOutput' : tmp + self.dt.Rows[i][j + 2] + tmp, 'Name' : self.dt.Rows[i][j], 'Type' : self.dt.Rows[i][j + 1], 'Value' : self.dt.Rows[i][j + 2]})

                    o = uda.GetAssociatedAttributes(self.dt.Rows[i][0])
                    o['Sitech12dAttributes'] = json.dumps(dic)
    
    
                failGuard.Commit()
                self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
                UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
        
        except Exception as err:
            tt = sys.exc_info()
            exc_type, exc_obj, exc_tb = sys.exc_info()
            # EndMark MUST be set no matter what
            # otherwise TBC won't work anymore and needs to be restarted
            self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
            UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
            self.error2.Content += '\nan Error occurred - Result probably incomplete\n' + str(exc_type) + '\n' + str(exc_obj) + '\nLine ' + str(exc_tb.tb_lineno)

        #self.updatelist()
        self.cellsanitycheck()

    def cellsanitycheck(self):
        
        self.error.Content = ''
        errorlist = []
        # go through all rows
        for r in self.dt.Rows:
        
            # check for double attribute names
            for i in range(1, self.dt.Columns.Count, 3):
                for j in range(i + 3, self.dt.Columns.Count, 3):
                    if r[i] and r[i] == r[j]:
                        errorlist.Add("Serial# " + str(r[0]) + " - duplicate Attribute " + str(r[i]))
            # if attribute is set
            # check for wrong Type
            for i in range(1, self.dt.Columns.Count, 3):
                if r[i] and len(r[i]) > 0: # if attribute is set
                    if not r[i + 1] or len(r[i + 1]) == 0: # check for missing Type
                        tt = str(r[i])
                        errorlist.Add("Serial# " + str(r[0]) + " - Attribute '" + str(r[i]) + "' - Type Entry missing")
                    elif not r[i + 1] == "Text" and not r[i + 1] == "Real" and not r[i + 1] == "Integer":
                        errorlist.Add("Serial# " + str(r[0]) + " - Attribute '" + str(r[i]) + "' - unknown Type entered - '" + str(r[i + 1]) + "'")

                    if not r[i + 2] or len(r[i + 2]) == 0: # check for missing Value
                        errorlist.Add("Serial# " + str(r[0]) + " - Attribute '" + str(r[i]) + "' - Value Entry missing")
                    else:   # we do have a value, but is it the correct type
                        valuetype = 0 # 0 - Text, 1 - Real, 2 - Integer
                        #if re.search('[a-zA-Z]', r[i + 2]):
                        if "." in r[i + 2] and r[i + 2].replace('.','',1).isdigit():
                            valuetype = 1
                        elif r[i + 2].isdigit():
                            valuetype = 2

                        if valuetype == 0 and not r[i + 1] == "Text":
                            errorlist.Add("Serial# " + str(r[0]) + " - Attribute '" + str(r[i]) + "' - possible Type and Value mismatch")
                        if valuetype == 1 and not r[i + 1] == "Real":
                            errorlist.Add("Serial# " + str(r[0]) + " - Attribute '" + str(r[i]) + "' - possible Type and Value mismatch")
                        if valuetype == 2 and not r[i + 1] == "Integer":
                            errorlist.Add("Serial# " + str(r[0]) + " - Attribute '" + str(r[i]) + "' - possible Type and Value mismatch")

                            

        errorlist.sort()
        for err in errorlist:
            self.error.Content += "\n" + err

    def IsValid(self, serial):
        o = self.currentProject.Concordance.Lookup(serial)
        if isinstance(o, self.polyType):
            return True
        return False
    
    def updatelist(self):

        uda = self.currentProject.Lookup(27) # User Defined Attributes are always Serial# 27

        # find out how many different attributes we're dealing with
        foundattrnames = []
        for s in self.savedserials:
            o = self.currentProject.Concordance.Lookup(s)
            if o:
                try:
                    dic = json.loads(uda.GetAssociatedAttributes(o.SerialNumber)['Sitech12dAttributes'])
                #liststring = ''
                
                    for e in sorted(dic, key=lambda k: k['Name'].lower()):
                        #if not len(liststring) == 0:
                        #    liststring += ' ||| '
                        #liststring += "Attr-Name: " + e['Name']
                        #liststring += ' - ' + "Attr-Type: " + e['Type']
                        #liststring += ' - ' + "Attr-Value: " + e['Value']
                        #tt = {'Name': "A", 'Type':  "b", 'Value': '20'}
                        
                        if not e['Name'] in foundattrnames:
                            foundattrnames.Add(e['Name'])
                except:
                    pass
        
        foundattrnames = sorted(foundattrnames, key=lambda k: k.lower())

        self.dt = self.MakeColumnsTable(foundattrnames.Count)
                
        self.datagrid.DataContext = self.dt
        # stupid bug - need to set the header caption manually
        for i in range(self.dt.Columns.Count):
            tt = self.datagrid.Columns
            self.datagrid.Columns[i].Header = self.dt.Columns[i].Caption
            tt = i % 6
            
            if 0 < i % 6 < 4:
                self.datagrid.Columns[i].HeaderStyle = self.coloredHeader

        #self.datagrid.Columns[0].Visibility = System.Windows.Visibility.Hidden # set serial columns to not show
        # Once a table has been created, use the
        # NewRow to create a DataRow.
        for s in self.savedserials:
            o = self.currentProject.Concordance.Lookup(s)
            
            if o:
                row = self.dt.NewRow()
                row["serial"] = o.SerialNumber

                try:
                    dic = json.loads(uda.GetAssociatedAttributes(o.SerialNumber)['Sitech12dAttributes'])
                    
                    for e in sorted(dic, key=lambda k: k['Name'].lower()):
                        ci = foundattrnames.index(e['Name']) * 3 + 1
                    
                        # Then add the new row to the collection.
                        row["c" + str(ci)] = e['Name']
                        row["c" + str(ci + 1)] = e['Type']
                        row["c" + str(ci + 2)] = e['Value']
                except:
                    pass
                
                self.dt.Rows.Add(row)

        #self.datagrid.Columns[0].SortDirection = System.ComponentModel.ListSortDirection.Ascending
        self.cellsanitycheck()
        # fill the grid
        # the headers
        #for i in range(1, self.savedserials.Count, 3):
        #    t1 = TextBlock()

    #def CancelClicked(self, cmd, args):
    #
    #    cmd.CloseUICommand ()

    #def OkClicked(self, cmd, e):
    #    Keyboard.Focus(self.okBtn)
    #    self.error.Content=''
    #
    #
    #    self.currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, self.Caption)
    #    UIEvents.RaiseBeforeDataProcessing(self, UIEventArgs())
    #
    #    wv = self.currentProject [Project.FixedSerial.WorldView]
    #
    #    try:
    #        # the "with" statement will unroll any changes if something go wrong
    #        with TransactMethodCall(self.currentProject.TransactionCollector) as failGuard:
    #
    #
    #
    #            failGuard.Commit()
    #            UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
    #            self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
    #    
    #    except Exception as e:
    #        tt = sys.exc_info()
    #        exc_type, exc_obj, exc_tb = sys.exc_info()
    #        # EndMark MUST be set no matter what
    #        # otherwise TBC won't work anymore and needs to be restarted
    #        self.currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
    #        UIEvents.RaiseAfterDataProcessing(self, UIEventArgs())
    #        self.error.Content += '\nan Error occurred - Result probably incomplete\n' + str(exc_type) + '\n' + str(exc_obj) + '\nLine ' + str(exc_tb.tb_lineno)
    #
    #    self.SaveOptions()

    def MakeColumnsTable(self, attrcount):

        # Create a new DataTable titled 'Names.'
        cTable = DataTable("Names")
        ## Create an array for DataColumn objects.
        keys = Array[DataColumn]([DataColumn()] * (attrcount * 3 + 1))
        
        c = 0
        serialColumn = cTable.Columns.Add("serial", System.Type.GetType("System.UInt32"))
        serialColumn.Caption = "Serial#"
        serialColumn.ReadOnly = True
        keys[c] = serialColumn

        for i in range(attrcount):    
            # Add three column objects to the table.
            c += 1
            nameColumn = cTable.Columns.Add("c" + str(c), System.Type.GetType("System.String"))
            nameColumn.Caption = "Attribute-Name"
            keys[c] = nameColumn
            c += 1
            typeColumn = cTable.Columns.Add("c" + str(c), System.Type.GetType("System.String"))
            typeColumn.Caption = "Attribute-Type"
            keys[c] = typeColumn
            c += 1
            valueColumn = cTable.Columns.Add("c" + str(c), System.Type.GetType("System.String"))
            valueColumn.Caption = "Attribute-Value"
            keys[c] = valueColumn

    
   
        #cTable.PrimaryKey = keys
    
        # Return the new DataTable.
        return cTable



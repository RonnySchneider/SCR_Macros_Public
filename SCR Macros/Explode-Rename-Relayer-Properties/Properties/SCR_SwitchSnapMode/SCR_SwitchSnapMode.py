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
    cmdData.Key = "SCR_SwitchSnapMode"
    cmdData.CommandName = "SCR_SwitchSnapMode"
    cmdData.Caption = "_SCR_SwitchSnapMode"
    cmdData.UIForm = "SCR_SwitchSnapMode"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
                                                      # if you enable or disable this line, you MUST restart TBC
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Expld-SNR-Relay-Prop"
        cmdData.DefaultTabGroupKey = "Properties"
        cmdData.ShortCaption = "Switch Snap Mode"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.04
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Switch Snap Mode"
        cmdData.ToolTipTextFormatted = "Switch Snap Mode"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png") # we have to include a icon revision, otherwise TBC might not show the new one
        cmdData.ImageSmall = b
    except:
        pass

#def Execute(cmd, currentProject, macroFileFolder, parameters):
#    form = SCR_SwitchSnapMode(currentProject, macroFileFolder).Show()
#    return
#    # .Show() - is non modal - you can interact with the drawing window
#    # .ShowDialog() - is modal - you CAN NOT interact with the drawing window


class SCR_SwitchSnapMode(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_SwitchSnapMode.xaml") as s:
            wpf.LoadComponent (self, s)
        self.currentProject = currentProject



    def OnLoad(self, cmd, buttons, event):
        self.okBtn = buttons[0]
        self.Caption = cmd.Command.Caption

        InputSettings.RunningSnapsChanged += self.snapsettingupdate

        self.Loaded += self.SetDefaultOptions
        #self.Closing += self.SaveOptions


    def snapsettingupdate(self, sender, e):

        self.buttonnosnap.IsChecked = False
        self.snaplist = InputSettings.GetSnapModes(self.currentProject.Guid) # order/priority and enabled/disabled of snap modes
        self.setbuttons()
        
        # save to settings string - this is used to restore it later
        tt = [str(element) for element in self.snaplist]
        self.snapstring = ",".join(tt)
        
        # trigger save
        self.SaveOptions(None, None)


    def Dispose(self, thisCmd, disposing):

        InputSettings.RunningSnapsChanged -= self.snapsettingupdate

    def CancelClicked(self, cmd, args):
        self.SaveOptions(None, None)

        cmd.CloseUICommand ()

    def SetDefaultOptions(self, sender, e):
        settings = ConstructionCommandsSettings.ProvideObject(self.currentProject)

        ### for fixing a mess up when I somehow multiplied point snaps
        ##self.snapstring = "0,1,4098,3,4,5,7,8,9,11,4108,19,20"
        ##self.castsnapstringtoarray()
        ##tt1 = InputSettings.GetSnapModes(self.currentProject.Guid)
        ##InputSettings.SaveSnapModes(self.currentProject.Guid, self.snaplist)
        ##tt2 = InputSettings.GetSnapModes(self.currentProject.Guid)
        ##tt3 = InputSettings.RunningSnaps
        ##testarray = Array[SnapMode]([SnapMode(2), SnapMode(12)])
        ##InputSettings.RunningSnaps = testarray
        ##tt3 = InputSettings.RunningSnaps

        self.snapstring = settings.GetString("SCR_SwitchSnapMode.snaps", "") #"0,1,2,3,4,5,7,8,9,11,4108,19,20")
        if self.snapstring != None and self.snapstring != "":
            self.castsnapstringtoarray()
        else:
            self.snaplist = InputSettings.GetSnapModes(self.currentProject.Guid) # order/priority and enabled/disabled of snap modes in menu
            self.snapstickedarray = InputSettings.RunningSnaps # actually active snaps - no idea why they keep two separate lists
            
            # in case the snap window was never open and OK'ed in the project - the self.snaplist will be empty
            # cast array to string and save now, otherwise we'll run into an issue once we try to reactivate them
            if self.snaplist.Count == 0:
                self.castsnaparraytostring()
            self.SaveOptions(None, None)

        self.setbuttons()

        # set snap settings
        InputSettings.SaveSnapModes(self.currentProject.Guid, self.snaplist)

        InputSettings.RunningSnapsChanged -= self.snapsettingupdate
        InputSettings.RunningSnaps = self.snapstickedarray
        InputSettings.RunningSnapsChanged += self.snapsettingupdate

        # restore no snap button at last
        #tt = settings.GetBoolean("SCR_SwitchSnapMode.buttonnosnap")
        self.buttonnosnap.IsChecked = settings.GetBoolean("SCR_SwitchSnapMode.buttonnosnap", False)
        # and trigger the opposite - for quick change via keyboard shortcut
        self.buttonnosnap.IsChecked = not self.buttonnosnap.IsChecked

    def SaveOptions(self, sender, e):

        settings = ConstructionCommandsSettings.ProvideObject(self.currentProject)

        settings.SetBoolean("SCR_SwitchSnapMode.buttonnosnap", self.buttonnosnap.IsChecked)
        if not self.buttonnosnap.IsChecked: # otherwise it will overwrite it with 'Free' only
            settings.SetString("SCR_SwitchSnapMode.snaps", self.snapstring)

    def castsnaparraytostring(self):

        # this method is only triggered when the snapstring and snaplist are empty
        self.snapstring = "0,1,2,3,4,5,7,8,9,11,12,19,20" # set standard priority
        tt = self.snapstring.split(",")

        # in theorey only the 5 default snaps should be set to enabled at this stage
        # but just in case we go through the list and match the string to the array
        for s in self.snapstickedarray:
            sint = int(s)
            for i in range(tt.Count):
                if tt[i] == str(sint):
                    tt[i] = str(sint + 4096)
        
        self.snapstring = ",".join(tt)    


    def castsnapstringtoarray(self):
        
        # InputSettings.SaveSnapModes is very picky when it comes to the lists it accepts
        # it won't accept tt below, although it looks the same with GetType as self.snaplist
        self.snaplist = Array[UInt16]([UInt16()] * 13)
        tt = self.snapstring.split(",")
        tt = [System.UInt16.Parse(element) for element in tt]
        snapsticked = []
        
        for i in range(tt.Count):
            self.snaplist[i] = tt[i]

            if tt[i] >= 4096:
                snapsticked.Add(tt[i] - 4096)

        # prepare a list with what is actually used for snapping
        # for some stupid reason the menu and what is used are two different things
        # you can in fact set InputSettings.RunningSnaps without updating the menu
        # but I like to keep the menu consistant, hence the need to keep track of two lists
                
        if snapsticked.Count > 0:
            self.snapstickedarray = Array[SnapMode]([SnapMode(13)] * snapsticked.Count)
            tt = self.snapstickedarray
            for i in range(snapsticked.Count):
                self.snapstickedarray[i] = SnapMode(snapsticked[i])
        else:
            self.snapstickedarray = Array[SnapMode]([SnapMode(13)]) # set at least no snap


    def setbuttons(self):
        for snap in self.snaplist:
            snap = int(snap)
            if   snap == 0:         self.buttonpoint.IsChecked = False
            elif snap == 0 + 4096:  self.buttonpoint.IsChecked = True
            elif snap == 1:         self.buttonend.IsChecked = False
            elif snap == 1 + 4096:  self.buttonend.IsChecked = True
            elif snap == 2:         self.buttonmid.IsChecked = False
            elif snap == 2 + 4096:  self.buttonmid.IsChecked = True
            elif snap == 3:         self.buttonnear.IsChecked = False
            elif snap == 3 + 4096:  self.buttonnear.IsChecked = True
            elif snap == 4:         self.buttoncenter.IsChecked = False
            elif snap == 4 + 4096:  self.buttoncenter.IsChecked = True
            elif snap == 5:         self.buttoninsertion.IsChecked = False
            elif snap == 5 + 4096:  self.buttoninsertion.IsChecked = True
            elif snap == 7:         self.buttonperp.IsChecked = False
            elif snap == 7 + 4096:  self.buttonperp.IsChecked = True
            elif snap == 8:         self.buttontangent.IsChecked = False
            elif snap == 8 + 4096:  self.buttontangent.IsChecked = True
            elif snap == 9:         self.buttonnode.IsChecked = False
            elif snap == 9 + 4096:  self.buttonnode.IsChecked = True
            elif snap == 11:        self.buttonintersection.IsChecked = False
            elif snap == 11 + 4096: self.buttonintersection.IsChecked = True
            elif snap == 12:        self.buttonfree.IsChecked = False
            elif snap == 12 + 4096: self.buttonfree.IsChecked = True
            elif snap == 19:        self.buttonquad.IsChecked = False
            elif snap == 19 + 4096: self.buttonquad.IsChecked = True
            elif snap == 20:        self.buttonimage.IsChecked = False
            elif snap == 20 + 4096: self.buttonimage.IsChecked = True

    def nosnapChanged(self, sender, e):
        # method is toggled by the XAML directly
        if self.buttonnosnap.IsChecked:
            self.buttonlist.IsEnabled = False
        
            # get current project settings
            self.snaplist = InputSettings.GetSnapModes(self.currentProject.Guid) # order/priority and enabled/disabled of snap modes

            # save to settings string - this is used to restore it later
            tt = [str(element) for element in self.snaplist]
            self.snapstring = ",".join(tt)
            
            # the snaplist is what is used to set the snaps in the menu
            for i in range(self.snaplist.Count):

                # keep the order, but deactivate everything except 4108 which is 'Free'
                if not self.snaplist[i] == 4108:
                    if self.snaplist[i] == 12:
                        self.snaplist[i] = 4108
                    elif self.snaplist[i] >= 4096:
                        self.snaplist[i] = self.snaplist[i] - 4096 # deactivate
                    else:
                        pass # it's already deactivated

            # set snap settings in menu
            InputSettings.SaveSnapModes(self.currentProject.Guid, self.snaplist)

            # now actually enable the snaps
            # disable the event trigger first in order to avoid false looping/trigger/retrigger
            InputSettings.RunningSnapsChanged -= self.snapsettingupdate
            InputSettings.RunningSnaps = Array[SnapMode]([SnapMode(12)])
            InputSettings.RunningSnapsChanged += self.snapsettingupdate
            # trigger save
            self.SaveOptions(None, None)

        else: # restore previous snap settings
            self.buttonlist.IsEnabled = True

            self.castsnapstringtoarray()
            self.setbuttons()

            InputSettings.SaveSnapModes(self.currentProject.Guid, self.snaplist)

            InputSettings.RunningSnapsChanged -= self.snapsettingupdate
            InputSettings.RunningSnaps = self.snapstickedarray
            InputSettings.RunningSnapsChanged += self.snapsettingupdate

            self.SaveOptions(None, None)

    def snapmodeClicked(self, sender, e):
            
            # with click it's only triggered if the users clicks the button and not of the program changes the state, i.e. during startup
            button = int(sender.Uid)
            buttonpressed = sender.IsChecked
            
            for i in range(self.snaplist.Count):

                if self.snaplist[i] == button or self.snaplist[i] == button + 4096:

                    if buttonpressed:

                        self.snaplist[i] = button + 4096

                    else:

                        self.snaplist[i] = button


            InputSettings.SaveSnapModes(self.currentProject.Guid, self.snaplist)

            # save to settings string - this is used to restore it later
            tt = [str(element) for element in self.snaplist]
            self.snapstring = ",".join(tt)

            self.castsnapstringtoarray() # need to run in order to get the updated ticked array


            InputSettings.RunningSnapsChanged -= self.snapsettingupdate
            InputSettings.RunningSnaps = self.snapstickedarray
            InputSettings.RunningSnapsChanged += self.snapsettingupdate

            self.SaveOptions(None, None)



    def OkClicked(self, cmd, e):
        
        self.CancelClicked(cmd, None)
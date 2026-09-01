#   GNU GPLv3
#   <this is an add-on Script/Macro for the geospatial software "Trimble Business Center" aka TBC>
#   <you'll need at least the "Survey Advanced" licence of TBC in order to run this script>
#	<see the ToolTip section below for a brief explanation what the script does>
#	<see the Help-Files for more details>
#   Copyright (C) 2026 Ronny Schneider
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


class TrimbleConnectClient(object):
    """Thin wrapper around Trimble.Vce.Services.Construction.TrimbleConnect.TrimbleConnectService,
    reusing TBC's own already-authenticated Trimble ID session - no separate OAuth flow needed. GetFileList(parent, progressBar) is the one
    primitive the service exposes for browsing: called with None it returns the region list; called with
    a region it returns the projects in that region; called with a project or folder it returns that
    folder's immediate children (one level at a time - there is no single call that returns a whole tree,
    so deep folders must be browsed into rather than eagerly flattened)."""

    def __init__(self):
        self._svc = None

    def _get_service(self):
        if self._svc is None:
            clr.AddReference("Trimble.Vce.Services.Construction")
            from Trimble.Vce.Services.TrimbleConnect import TrimbleConnectService
            svc = TrimbleConnectService()
            svc.Connect(svc.Settings, None)
            self._svc = svc
        return self._svc

    def get_regions(self):
        return list(self._get_service().GetFileList(None, None) or [])

    def get_children(self, parent):
        return list(self._get_service().GetFileList(parent, None) or [])

    def save_file_remotely(self, localFile, parentFolderOrProject):
        # ProgressBar.TBC_ProgressBar lives in the same Trimble.Vce.Interfaces.ProgressBar namespace as
        # the IProgressBarControl this method expects, so it's very likely the intended control to pass -
        # fall back to no progress bar if it turns out not to satisfy the interface
        try:
            self._get_service().SaveFileRemotely(localFile, parentFolderOrProject, ProgressBar.TBC_ProgressBar)
        except Exception:
            self._get_service().SaveFileRemotely(localFile, parentFolderOrProject, None)

    def download_file(self, remoteFile, localFilePath, progressBar=None):
        # remoteFile is the IFileInformation node exactly as returned by GetFileList - no id lookup
        # needed. The ID-based DownloadFile/DownloadFileAsync(path, fileId, token) API looked like the
        # obvious choice but needs a RemoteFileId that browsed files don't carry (that interface appears
        # to be for TBC's internal "referenced file" tracking, not general Trimble Connect browsing) -
        # it silently no-oped (no exception, no file) and even deadlocked the UI thread once, presumably
        # while stuck retrying with a bad/empty id. BeginDownloadFile/EndDownloadFile is the plain
        # IAsyncResult (APM) pattern that takes the file node directly, and EndDownloadFile blocks on a
        # wait handle rather than an awaited continuation, so it's safe to call synchronously.
        svc = self._get_service()
        asyncResult = svc.BeginDownloadFile(localFilePath, remoteFile, None, None, progressBar)
        return svc.EndDownloadFile(asyncResult)

    def create_folder(self, folderName, parentFolder):
        # returns True if the folder was created OR already existed - CreateFolder is its own
        # existence check, so callers don't need to look before creating
        return self._get_service().CreateFolder(folderName, parentFolder)

    def remove_file(self, fileName, parentFolder):
        return self._get_service().RemoveFile(fileName, parentFolder, None)

    def remove_folder(self, folderName, parentFolder):
        return self._get_service().RemoveFolder(folderName, parentFolder, None)

    def get_or_create_child_folder(self, parentFolder, folderName):
        # there is no API that hands back the folder node CreateFolder just created/confirmed (it only
        # returns a bool), so after ensuring it exists we re-list parentFolder's children to find it
        self.create_folder(folderName, parentFolder)
        for child in self.get_children(parentFolder):
            if child.IsFolder and str(child.FileName) == folderName:
                return child
        raise Exception("Trimble Connect folder '" + folderName + "' could not be created or found.")

    def get_access_token(self):
        # TrimbleConnectService.AuthenticationInfo is documented only as "authentication information...
        # usually acquired by signing in", with no public type/shape - but Connect() takes a
        # Trimble.Vce.Services.Auth.TrimbleIdAuthInfo, which does publicly expose an AccessToken (a TID
        # bearer JWT), so AuthenticationInfo is almost certainly that same object. Returns None (rather
        # than raising) on any mismatch, so callers can treat "no token" as "Web API not available here".
        try:
            authInfo = self._get_service().AuthenticationInfo
            return authInfo.AccessToken if authInfo is not None else None
        except Exception:
            return None


def trimble_connect_region_base_url(remoteItem):
    # confirmed via live diagnostic: a browsed item's RootOriginBaseURL is a region-specific, protocol-
    # relative host (e.g. "//app32.connect.trimble.com" for this account's region) - NOT the generic
    # "app.connect.trimble.com" gateway the .NET SDK docs quote as the default production endpoint, so
    # calls need to go to the item's own shard rather than the generic host
    rootOrigin = getattr(remoteItem, "RootOriginBaseURL", None)
    if not rootOrigin:
        return None
    return "https:" + rootOrigin + "/tc/api/2.0"


class TrimbleConnectWebApi(object):
    """Thin wrapper around the real Trimble Connect REST API, used only for the one operation
    TrimbleConnectService has no primitive for: actually moving a file or folder to a different parent
    (PATCH .../{id} with a "parentId" body, confirmed against Trimble's public tcps/2.0 API spec and, for
    the id shape, against a live item dump - RemoteFileId matches that item's own "ID" field exactly).
    Everything else (browsing, download, upload, create/remove) goes through TrimbleConnectClient/
    TrimbleConnectService instead, since that's the one already proven to work with TBC's own signed-in
    session."""

    def __init__(self, accessToken, baseUrl):
        self.accessToken = accessToken
        self.baseUrl = baseUrl

    def _patch(self, path, bodyDict):
        request = HttpWebRequest.Create(self.baseUrl + path)
        request.Method = "PATCH"
        request.ContentType = "application/json"
        request.Accept = "application/json"
        request.Headers.Add("Authorization", "Bearer " + self.accessToken)

        payload = Encoding.UTF8.GetBytes(json.dumps(bodyDict))
        request.ContentLength = payload.Length
        stream = request.GetRequestStream()
        stream.Write(payload, 0, payload.Length)
        stream.Close()

        try:
            request.GetResponse().Close()
        except WebException as ex:
            detail = ""
            if ex.Response is not None:
                with StreamReader(ex.Response.GetResponseStream(), Encoding.UTF8) as sr:
                    detail = sr.ReadToEnd()
            raise Exception("Trimble Connect Web API PATCH " + path + " failed: " + str(ex.Message) + " " + detail)

    def move_file(self, fileId, newParentId):
        self._patch("/files/" + str(fileId), {"parentId": str(newParentId)})

    def move_folder(self, folderId, newParentId):
        self._patch("/folders/" + str(folderId), {"parentId": str(newParentId)})


class ExcelJobRegister(object):
    """Writes a value into the next empty row of one column in an existing Excel workbook, via late-bound
    COM automation (Excel.Application) - just needs Excel installed, no PIA reference or extra DLL. The
    workbook is left open and visible afterwards (not saved-and-closed) so the user can fill in the rest
    of that row by hand - reuses an already-running Excel/already-open copy of the workbook if there is
    one (the user's own, or one left open by a previous sync), rather than opening a second copy."""

    XL_UP = -4162

    def append_value(self, workbookPath, columnLetter, value, minimumRow=1):
        excelApp, weOwnTheApp = self._get_or_create_application()
        workbook = None
        try:
            workbook = self._find_open_workbook(excelApp, workbookPath)
            if workbook is None:
                workbook = excelApp.Workbooks.Open(workbookPath)

            worksheet = workbook.Worksheets[1]
            nextRow = max(self._find_next_empty_row(worksheet, columnLetter), minimumRow)
            cell = worksheet.Range(columnLetter + str(nextRow))
            cell.Value2 = value

            workbook.Save()

            excelApp.Visible = True
            workbook.Activate()
            worksheet.Activate()
            cell.Select()
            return nextRow
        finally:
            # deliberately no workbook.Close()/excelApp.Quit() - the whole point is to leave it open for
            # the user. Releasing our own COM references is still correct hygiene: Excel is a genuinely
            # separate process (out-of-process COM), so dropping our proxy to it doesn't close anything
            # the user can see - it just means this macro is no longer holding a handle to it.
            if workbook is not None:
                try:
                    Marshal.ReleaseComObject(workbook)
                except Exception:
                    pass
            try:
                Marshal.ReleaseComObject(excelApp)
            except Exception:
                pass
            GC.Collect()
            GC.WaitForPendingFinalizers()

    def _get_or_create_application(self):
        # (app, weOwnIt) - weOwnIt is False when we attached to an Excel the user already had running,
        # which matters only in spirit here (we never Quit() either way) but documents the intent
        try:
            return Marshal.GetActiveObject("Excel.Application"), False
        except Exception:
            excelType = Type.GetTypeFromProgID("Excel.Application")
            return Activator.CreateInstance(excelType), True

    def _find_open_workbook(self, excelApp, workbookPath):
        targetPath = os.path.normcase(os.path.abspath(workbookPath))
        for workbook in excelApp.Workbooks:
            try:
                if os.path.normcase(os.path.abspath(workbook.FullName)) == targetPath:
                    return workbook
            except Exception:
                continue
        return None

    def _find_next_empty_row(self, worksheet, columnLetter):
        # standard Ctrl+Up idiom: from the very bottom row, jump up to the last non-empty cell - if the
        # whole column is empty, that jump lands back on row 1, so that row-1 case needs its own check
        # (a non-empty row 1 means row 1 IS the last used row; an empty row 1 means nothing is used yet).
        # This already copes with gaps/leading blanks - it always finds the true last used row from the
        # bottom regardless of what's empty above it - so minimumRow (above) exists only as an explicit
        # floor for the edge case where row 1 itself is blank but shouldn't be written into (e.g. a
        # header whose cell is technically empty, like a merged cell).
        bottomRow = worksheet.Rows.Count
        lastUsedRow = worksheet.Range(columnLetter + str(bottomRow)).End(self.XL_UP).Row
        firstCellValue = worksheet.Range(columnLetter + "1").Value2
        if lastUsedRow == 1 and (firstCellValue is None or str(firstCellValue).strip() == ""):
            return 1
        return lastUsedRow + 1


def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_DayJobSync"
    cmdData.CommandName = "SCR_DayJobSync"
    cmdData.Caption = "_SCR_DayJobSync"
    #cmdData.UIForm = "SCR_DayJobSync"      # left disabled - this is a fully independent floating window, not a TBC-managed dialog
                                                        # if you enable or disable this line, you MUST restart TBC
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "0"

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "Update/Transfer"
        cmdData.ShortCaption = "DayJob Sync"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3
        cmdData.EnableNoProject       = True

        cmdData.Version = 1.01
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""

        cmdData.ToolTipTitle = "DayJobSync"
        cmdData.ToolTipTextFormatted = "log in to Trimble Connect and browse its folder tree"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png") # we have to include a icon revision, otherwise TBC might not show the new one
        cmdData.ImageSmall = b
    except:
        pass

def Execute(cmd, currentProject, macroFileFolder, parameters):
    form = SCR_DayJobSyncDialog(currentProject, macroFileFolder).Show()
    return
    # .Show() - is non modal - you can interact with the drawing window
    # .ShowDialog() - is modal - you CAN NOT interact with the drawing window


class SCR_DayJobSyncDialog(Window): # this inherits from the WPF Window control - a fully independent floating window
    def __init__(self, currentProject, macroFileFolder):

        with StreamReader(macroFileFolder + r"\SCR_DayJobSync.xaml") as s:
            wpf.LoadComponent(self, s)

        ElementHost.EnableModelessKeyboardInterop(self)

        self.currentProject = currentProject
        self.macroFileFolder = macroFileFolder

        self.trimbleConnectClient = TrimbleConnectClient()
        self.projectRoot = None
        self.folderStack = []

        self.folderStructureBox.Text = OptionsManager.GetString("SCR_DayJobSync.folderstructure", "{YYYY}/{YYMM}/{Job}")
        self.controllerDataBox.Text = OptionsManager.GetString("SCR_DayJobSync.controllerdataname", "Controller Data")
        self.tbcBox.Text = OptionsManager.GetString("SCR_DayJobSync.tbcname", "TBC")
        self.controllerDataEnabledCheckbox.IsChecked = OptionsManager.GetString("SCR_DayJobSync.controllerdataenabled", "True") != "False"
        self.tbcEnabledCheckbox.IsChecked = OptionsManager.GetString("SCR_DayJobSync.tbcenabled", "True") != "False"
        self.oldJobFolderBox.Text = OptionsManager.GetString("SCR_DayJobSync.oldjobfoldername", "Old Jobs")
        self.oldJobFolderEnabledCheckbox.IsChecked = OptionsManager.GetString("SCR_DayJobSync.oldjobfolderenabled", "True") != "False"
        self.localSyncFolder = OptionsManager.GetString("SCR_DayJobSync.localsyncfolder", "")
        self.update_local_folder_label()

        self.jobRegisterPath = OptionsManager.GetString("SCR_DayJobSync.jobregisterpath", "")
        self.jobRegisterColumnBox.Text = OptionsManager.GetString("SCR_DayJobSync.jobregistercolumn", "A")
        self.jobRegisterStartRowBox.Text = OptionsManager.GetString("SCR_DayJobSync.jobregisterstartrow", "1")
        self.update_job_register_label()

        self.folderStructureBox.TextChanged += self.folder_structure_changed
        self.controllerDataBox.TextChanged += self.controller_data_name_changed
        self.tbcBox.TextChanged += self.tbc_name_changed
        self.controllerDataEnabledCheckbox.Checked += self.controller_data_enabled_changed
        self.controllerDataEnabledCheckbox.Unchecked += self.controller_data_enabled_changed
        self.tbcEnabledCheckbox.Checked += self.tbc_enabled_changed
        self.tbcEnabledCheckbox.Unchecked += self.tbc_enabled_changed
        self.oldJobFolderBox.TextChanged += self.old_job_folder_name_changed
        self.oldJobFolderEnabledCheckbox.Checked += self.old_job_folder_enabled_changed
        self.oldJobFolderEnabledCheckbox.Unchecked += self.old_job_folder_enabled_changed
        self.browseLocalFolderBtn.Click += self.browse_local_folder_clicked
        self.browseJobRegisterBtn.Click += self.browse_job_register_clicked
        self.jobRegisterColumnBox.TextChanged += self.job_register_column_changed
        self.jobRegisterStartRowBox.TextChanged += self.job_register_start_row_changed

        self.reloadBtn.Click += self.reload_regions_clicked
        self.regionCombo.SelectionChanged += self.region_selection_changed
        self.projectCombo.SelectionChanged += self.project_selection_changed
        self.upBtn.Click += self.up_clicked
        self.fileList.MouseDoubleClick += self.file_list_double_click

        self.syncBtn.Click += self.sync_clicked
        self.helpBtn.Click += self.help_clicked

        self.Loaded += self.SetDefaultOptions
        self.Closing += self.SaveOptions

        self.reload_regions_clicked(None, None)


    # ---------- local sync folder / subfolder-name settings ----------

    def update_local_folder_label(self):
        self.localFolderLabel.Text = self.localSyncFolder if self.localSyncFolder else "(not set)"

    def browse_local_folder_clicked(self, sender, e):
        dlg = FolderBrowserDialog()
        dlg.Description = "Choose the local sync folder"
        if self.localSyncFolder and os.path.isdir(self.localSyncFolder):
            dlg.SelectedPath = self.localSyncFolder

        if dlg.ShowDialog() == DialogResult.OK:
            self.localSyncFolder = dlg.SelectedPath
            OptionsManager.SetValue("SCR_DayJobSync.localsyncfolder", self.localSyncFolder)
            self.update_local_folder_label()

    def folder_structure_changed(self, sender, e):
        OptionsManager.SetValue("SCR_DayJobSync.folderstructure", self.folderStructureBox.Text)

    def controller_data_name_changed(self, sender, e):
        OptionsManager.SetValue("SCR_DayJobSync.controllerdataname", self.controllerDataBox.Text)

    def tbc_name_changed(self, sender, e):
        OptionsManager.SetValue("SCR_DayJobSync.tbcname", self.tbcBox.Text)

    def controller_data_enabled_changed(self, sender, e):
        OptionsManager.SetValue("SCR_DayJobSync.controllerdataenabled", str(bool(self.controllerDataEnabledCheckbox.IsChecked)))

    def tbc_enabled_changed(self, sender, e):
        OptionsManager.SetValue("SCR_DayJobSync.tbcenabled", str(bool(self.tbcEnabledCheckbox.IsChecked)))

    def old_job_folder_name_changed(self, sender, e):
        OptionsManager.SetValue("SCR_DayJobSync.oldjobfoldername", self.oldJobFolderBox.Text)

    def old_job_folder_enabled_changed(self, sender, e):
        OptionsManager.SetValue("SCR_DayJobSync.oldjobfolderenabled", str(bool(self.oldJobFolderEnabledCheckbox.IsChecked)))


    # ---------- job register (Excel log) settings ----------

    def update_job_register_label(self):
        self.jobRegisterLabel.Text = self.jobRegisterPath if self.jobRegisterPath else "(not set)"

    def browse_job_register_clicked(self, sender, e):
        dlg = OpenFileDialog()
        dlg.Title = "Choose the job register Excel workbook"
        dlg.Filter = "Excel Workbooks (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls|All files (*.*)|*.*"
        if self.jobRegisterPath and os.path.isfile(self.jobRegisterPath):
            dlg.InitialDirectory = os.path.dirname(self.jobRegisterPath)
            dlg.FileName = os.path.basename(self.jobRegisterPath)

        if dlg.ShowDialog() == DialogResult.OK:
            self.jobRegisterPath = dlg.FileName
            OptionsManager.SetValue("SCR_DayJobSync.jobregisterpath", self.jobRegisterPath)
            self.update_job_register_label()

    def job_register_column_changed(self, sender, e):
        OptionsManager.SetValue("SCR_DayJobSync.jobregistercolumn", self.jobRegisterColumnBox.Text)

    def job_register_start_row_changed(self, sender, e):
        OptionsManager.SetValue("SCR_DayJobSync.jobregisterstartrow", self.jobRegisterStartRowBox.Text)


    # ---------- window position/size ----------

    def SetDefaultOptions(self, sender, e):
        SCROptions.LoadWindowState(self, "SCR_DayJobSync", default_width=340, default_height=460)

    def SaveOptions(self, sender, e):
        SCROptions.SaveWindowState(self, "SCR_DayJobSync")


    # ---------- remembered combo selections ----------

    def get_saved(self, key):
        return OptionsManager.GetString("SCR_DayJobSync." + key, "")

    def save_selected(self, key, value):
        OptionsManager.SetValue("SCR_DayJobSync." + key, str(value) if value is not None else "")

    def select_preferred_by_content(self, combo, optionKey):
        # Trimble Connect regions/projects have no stable id we can persist (region.ID is always None),
        # so their display name is the only usable key
        preferred = self.get_saved(optionKey) or None

        target = None
        if preferred is not None:
            for item in combo.Items:
                if str(item.Content) == str(preferred):
                    target = item
                    break

        if target is not None:
            combo.SelectedItem = target
        elif combo.Items.Count > 0:
            combo.SelectedIndex = 0


    # ---------- Trimble Connect login + folder tree ----------
    # login is implicit: TrimbleConnectService.Connect() reuses TBC's own already-authenticated Trimble ID
    # session, so the first call that touches the service (get_regions, below) is effectively "logging in".
    # Browsing is one level at a time via double-click/Up, since GetFileList only returns one folder's
    # immediate children per call (no single "whole tree" call).

    def reload_regions_clicked(self, sender, e):
        if Keyboard.IsKeyDown(Key.LeftShift) or Keyboard.IsKeyDown(Key.RightShift):
            self.run_diagnostic(sender, e)
            return

        self.error.Content = ""
        self.statusLabel.Text = "Connecting to Trimble Connect..."
        self.Dispatcher.Invoke(DispatcherPriority.Render, Action(lambda: None))

        self.regionCombo.Items.Clear()
        self.projectCombo.Items.Clear()
        self.fileList.Items.Clear()
        self.projectRoot = None
        self.folderStack = []

        try:
            regions = self.trimbleConnectClient.get_regions()
        except Exception as ex:
            self.statusLabel.Text = ""
            self.error.Content = str(ex)
            return

        for r in sorted(regions, key=lambda x: str(x.FileName)):
            item = ComboBoxItem()
            item.Content = str(r.FileName)
            item.Tag = r
            self.regionCombo.Items.Add(item)

        self.statusLabel.Text = "Connected to Trimble Connect."
        self.select_preferred_by_content(self.regionCombo, "selectedregionname")


    # ---------- diagnostics (Shift-click the reload/login button) ----------
    # dumps svc's own fields/properties (public AND private - TBC's actual token/authenticator is very
    # likely held in a private field, not exposed on TrimbleConnectService's public surface) plus the
    # currently selected Connect item, to find the real access-token source and item-id shape rather than
    # continuing to guess offline. Never calls RetrieveToken()/similar itself - only reads already-set
    # values - so it can't trigger a fresh login prompt or block on anything.

    def run_diagnostic(self, sender, e):
        bf = BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public
        lines = []

        lines.append("=== TrimbleConnectService (svc) ===")
        try:
            svc = self.trimbleConnectClient._get_service()
            self._dump_object(svc, lines, bf)
        except Exception as ex:
            lines.append("ERROR getting service: " + str(ex))

        lines.append("")
        lines.append("=== get_access_token() result ===")
        try:
            token = self.trimbleConnectClient.get_access_token()
            lines.append("token=" + (str(token)[:40] + "...(truncated)" if token else str(token)))
        except Exception as ex:
            lines.append("ERROR: " + str(ex))

        selectedItem = self.fileList.SelectedItem
        if selectedItem is not None:
            lines.append("")
            lines.append("=== Currently selected Connect item ===")
            self._dump_object(selectedItem.Tag, lines, bf)

        path = r"C:\temp\SCR_DayJobSync_diag.txt"
        folder = os.path.dirname(path)
        if not os.path.isdir(folder):
            os.makedirs(folder)
        with open(path, "w") as f:
            f.write("\n".join(lines))
        MessageBox.Show("Diagnostics written to " + path, "SCR_DayJobSync [DIAG]")

    def _dump_object(self, obj, lines, bf, prefix="  "):
        if obj is None:
            lines.append(prefix + "None")
            return
        t = obj.GetType()
        lines.append(prefix + "type=" + t.FullName)
        for p in t.GetProperties(bf):
            try:
                if len(p.GetIndexParameters()) > 0:
                    continue  # skip indexers (e.g. this[int]) - GetValue needs index args we don't have
                val = p.GetValue(obj, None)
                lines.append(prefix + "  [P] " + p.Name + " (" + p.PropertyType.Name + ") = " + str(val))
            except Exception as ex:
                lines.append(prefix + "  [P] " + p.Name + " ERR: " + str(ex))
        for f in t.GetFields(bf):
            try:
                val = f.GetValue(obj)
                lines.append(prefix + "  [F] " + f.Name + " (" + f.FieldType.Name + ") = " + str(val))
            except Exception as ex:
                lines.append(prefix + "  [F] " + f.Name + " ERR: " + str(ex))

    def region_selection_changed(self, sender, e):
        self.projectCombo.Items.Clear()
        self.fileList.Items.Clear()
        self.projectRoot = None
        self.folderStack = []
        self.error.Content = ""

        regionItem = self.regionCombo.SelectedItem
        if regionItem is None:
            return
        self.save_selected("selectedregionname", regionItem.Content)

        try:
            projects = self.trimbleConnectClient.get_children(regionItem.Tag)
        except Exception as ex:
            self.error.Content = str(ex)
            return

        for p in sorted(projects, key=lambda x: str(x.FileName)):
            item = ComboBoxItem()
            item.Content = str(p.FileName)
            item.Tag = p
            self.projectCombo.Items.Add(item)

        self.select_preferred_by_content(self.projectCombo, "selectedprojectname")

    def project_selection_changed(self, sender, e):
        self.fileList.Items.Clear()
        self.folderStack = []
        self.error.Content = ""

        projectItem = self.projectCombo.SelectedItem
        if projectItem is None:
            self.projectRoot = None
            self.tcPathLabel.Text = "/"
            return
        self.save_selected("selectedprojectname", projectItem.Content)

        self.projectRoot = projectItem.Tag
        self.restore_folder_stack()
        self.refresh_file_list()

    def current_folder(self):
        return self.folderStack[-1] if self.folderStack else self.projectRoot

    def save_folder_path(self):
        names = [str(f.FileName) for f in self.folderStack]
        OptionsManager.SetValue("SCR_DayJobSync.lastfolderpath", json.dumps(names))

    def restore_folder_stack(self):
        # walks the saved folder path (a list of folder names from the project root down to the last
        # folder the user browsed into) one level at a time via get_children - best effort, stopping at
        # whichever level no longer matches (e.g. a different project, or the folder was renamed/removed)
        raw = self.get_saved("lastfolderpath")
        if not raw:
            return
        try:
            names = json.loads(raw)
        except Exception:
            return

        current = self.projectRoot
        stack = []
        for name in names:
            try:
                children = self.trimbleConnectClient.get_children(current)
            except Exception:
                break
            match = None
            for c in children:
                if c.IsFolder and str(c.FileName) == name:
                    match = c
                    break
            if match is None:
                break
            stack.append(match)
            current = match

        self.folderStack = stack

    def refresh_file_list(self):
        self.fileList.Items.Clear()
        self.tcPathLabel.Text = "/" + "/".join(str(f.FileName) for f in self.folderStack)

        folder = self.current_folder()
        if folder is None:
            return

        try:
            children = self.trimbleConnectClient.get_children(folder)
        except Exception as ex:
            self.error.Content = str(ex)
            return

        for child in sorted(children, key=lambda c: (not c.IsFolder, str(c.FileName))):
            item = ListBoxItem()
            item.Content = ("[folder] " if child.IsFolder else "") + str(child.FileName)
            item.Tag = child
            self.fileList.Items.Add(item)

    def file_list_double_click(self, sender, e):
        item = self.fileList.SelectedItem
        if item is None or not item.Tag.IsFolder:
            return
        self.folderStack.append(item.Tag)
        self.save_folder_path()
        self.refresh_file_list()

    def up_clicked(self, sender, e):
        if self.folderStack:
            self.folderStack.pop()
            self.save_folder_path()
            self.refresh_file_list()


    # ---------- sync ----------

    def render_folder_structure(self, yy, mm, jobName):
        # user-defined template like "{YYYY}/{YYMM}/{Job}" - {YYYY} is assumed to belong to the 2000s
        # (20yy), fine for any realistic day-job date. Split on "/" (and "\" in case someone uses that)
        # into path segments, dropping empty ones, so os.path.join(*segments) builds the actual folders.
        # {YYMM} is replaced before the standalone {YY}/{MM} so it isn't first broken up into "{YY}{MM}"
        # by those replacements running first.
        template = self.folderStructureBox.Text.strip() or "{YYYY}/{YYMM}/{Job}"
        rendered = template.replace("{YYYY}", "20" + yy).replace("{YYMM}", yy + mm) \
                            .replace("{YY}", yy).replace("{MM}", mm).replace("{Job}", jobName)
        segments = [p for p in rendered.replace("\\", "/").split("/") if p]
        if not segments:
            raise Exception("Folder Structure resolved to nothing - check the template: " + template)
        return segments

    def sync_clicked(self, sender, e):
        self.error.Content = ""
        self.statusLabel.Text = ""

        item = self.fileList.SelectedItem
        if item is None:
            self.error.Content = "Select a file in the Trimble Connect folder tree first."
            return

        connectFile = item.Tag
        if connectFile.IsFolder:
            self.error.Content = "Select a file, not a folder."
            return

        if not self.localSyncFolder or not os.path.isdir(self.localSyncFolder):
            self.error.Content = "Set a valid local sync folder first."
            return

        fileName = str(connectFile.FileName)
        match = re.match(r"^(\d{2})(\d{2})\d{2}", fileName)
        if match is None:
            self.error.Content = "Filename does not start with a 6-digit date (YYMMDD): " + fileName
            return
        yy, mm = match.group(1), match.group(2)

        controllerDataEnabled = bool(self.controllerDataEnabledCheckbox.IsChecked)
        tbcEnabled = bool(self.tbcEnabledCheckbox.IsChecked)
        controllerDataName = self.controllerDataBox.Text.strip()
        tbcName = self.tbcBox.Text.strip()
        if controllerDataEnabled and not controllerDataName:
            self.error.Content = "Set the 'Controller Data' subfolder name, or untick its checkbox."
            return
        if tbcEnabled and not tbcName:
            self.error.Content = "Set the 'TBC' subfolder name, or untick its checkbox."
            return

        folderName = os.path.splitext(fileName)[0]
        try:
            targetFolder = os.path.join(self.localSyncFolder, *self.render_folder_structure(yy, mm, folderName))
        except Exception as ex:
            self.error.Content = str(ex)
            return
        controllerDataFolder = os.path.join(targetFolder, controllerDataName) if controllerDataEnabled else targetFolder
        # if Controller Data isn't ticked, the job (and its Files subfolder) is placed directly in
        # targetFolder instead of a Controller Data subfolder - controllerDataFolder just aliases
        # targetFolder in that case, so every downstream use of it already does the right thing

        try:
            foldersToCreate = [targetFolder]
            if controllerDataEnabled:
                foldersToCreate.append(controllerDataFolder)
            if tbcEnabled:
                foldersToCreate.append(os.path.join(targetFolder, tbcName))
            for folder in foldersToCreate:
                if not os.path.isdir(folder):
                    os.makedirs(folder)
        except Exception as ex:
            self.error.Content = str(ex)
            return

        downloadPath = os.path.join(controllerDataFolder, fileName)

        self.syncBtn.IsEnabled = False
        self.statusLabel.Text = "Downloading " + fileName + "..."
        self.Dispatcher.Invoke(DispatcherPriority.Render, Action(lambda: None))

        try:
            self.trimbleConnectClient.download_file(connectFile, downloadPath)
        except Exception as ex:
            self.syncBtn.IsEnabled = True
            self.error.Content = "Folders created, but download failed: " + str(ex)
            return

        if not self.wait_for_file(downloadPath):
            self.syncBtn.IsEnabled = True
            self.error.Content = "Download reported success, but the file is not at: " + downloadPath
            return

        # a job file can have an accompanying "<name without suffix> Files" subfolder alongside it on
        # Connect (e.g. "260831SCR1.job" -> "260831SCR1 Files") holding its supporting files - if one
        # exists next to the job file, copy the whole thing into the same destination folder
        filesFolderName = folderName + " Files"
        parentFolder = self.current_folder()
        try:
            siblings = self.trimbleConnectClient.get_children(parentFolder)
        except Exception as ex:
            self.syncBtn.IsEnabled = True
            self.error.Content = "Job file downloaded, but couldn't check for '" + filesFolderName + "': " + str(ex)
            return

        filesFolder = next((c for c in siblings if c.IsFolder and str(c.FileName) == filesFolderName), None)

        if filesFolder is not None:
            self.statusLabel.Text = "Downloading " + filesFolderName + "..."
            self.Dispatcher.Invoke(DispatcherPriority.Render, Action(lambda: None))
            try:
                self.download_connect_folder(filesFolder, os.path.join(controllerDataFolder, filesFolderName))
            except Exception as ex:
                self.syncBtn.IsEnabled = True
                self.error.Content = "Job file downloaded, but '" + filesFolderName + "' folder copy failed: " + str(ex)
                return

        # archive the job (and its Files subfolder, if any) on Trimble Connect itself, into an "old job"
        # folder alongside where they were found. get_or_create_child_folder does its own "does it already
        # exist" check via CreateFolder. Each item is archived independently (see archive_item_on_connect):
        # a real move via the Connect Web API is tried first, and only that specific item falls back to
        # upload+delete if the move doesn't work out - so a large "Files" folder that successfully moves
        # is never re-uploaded just because something else went wrong.
        oldJobFolderEnabled = bool(self.oldJobFolderEnabledCheckbox.IsChecked)
        oldJobFolderName = self.oldJobFolderBox.Text.strip()
        if oldJobFolderEnabled and oldJobFolderName:
            self.statusLabel.Text = "Archiving on Trimble Connect..."
            self.Dispatcher.Invoke(DispatcherPriority.Render, Action(lambda: None))
            try:
                oldJobFolder = self.trimbleConnectClient.get_or_create_child_folder(parentFolder, oldJobFolderName)
                self.archive_item_on_connect(connectFile, False, fileName, parentFolder, oldJobFolder, downloadPath)
                if filesFolder is not None:
                    self.archive_item_on_connect(filesFolder, True, filesFolderName, parentFolder, oldJobFolder,
                                                  os.path.join(controllerDataFolder, filesFolderName))
            except Exception as ex:
                self.syncBtn.IsEnabled = True
                self.error.Content = "Downloaded locally, but archiving on Trimble Connect failed: " + str(ex)
                self.refresh_file_list()
                return

            self.refresh_file_list()

        registerError = None
        if self.jobRegisterPath:
            columnLetter = self.jobRegisterColumnBox.Text.strip().upper()
            if not columnLetter:
                registerError = "job register column is not set"
            else:
                try:
                    minimumRow = int(self.jobRegisterStartRowBox.Text.strip() or "1")
                except ValueError:
                    minimumRow = 1

                self.statusLabel.Text = "Logging to job register..."
                self.Dispatcher.Invoke(DispatcherPriority.Render, Action(lambda: None))
                try:
                    ExcelJobRegister().append_value(self.jobRegisterPath, columnLetter, folderName, minimumRow)
                except Exception as ex:
                    registerError = str(ex)

        self.syncBtn.IsEnabled = True
        self.statusLabel.Text = "Downloaded to: " + downloadPath
        if filesFolder is not None:
            self.statusLabel.Text += " (with " + filesFolderName + ")"
        if oldJobFolderEnabled and oldJobFolderName:
            self.statusLabel.Text += " - archived on Connect in '" + oldJobFolderName + "'"
        if registerError is not None:
            self.error.Content = "Synced, but logging to job register failed: " + registerError
        elif self.jobRegisterPath:
            self.statusLabel.Text += " - logged to job register"

    def archive_item_on_connect(self, remoteItem, isFolder, name, parentFolder, oldJobFolder, localPathForFallback):
        # tries a real move first (PATCH .../{id} with parentId, via the actual Trimble Connect Web API -
        # near-instant regardless of file size, unlike re-uploading). Confirmed via a live diagnostic dump:
        # RemoteFileId matches the item's own "ID" field exactly, and RootOriginBaseURL is the item's real
        # region-specific API host - so any failure here now would mean something else entirely (network,
        # permissions, an expired token), and just falls through to the proven-safe fallback: upload the
        # local copy, then delete the original.
        moved = False
        try:
            accessToken = self.trimbleConnectClient.get_access_token()
            itemId = getattr(remoteItem, "RemoteFileId", None)
            targetFolderId = getattr(oldJobFolder, "RemoteFileId", None)
            baseUrl = trimble_connect_region_base_url(remoteItem) or trimble_connect_region_base_url(oldJobFolder)
            if accessToken and itemId and targetFolderId and baseUrl:
                webApi = TrimbleConnectWebApi(accessToken, baseUrl)
                if isFolder:
                    webApi.move_folder(itemId, targetFolderId)
                else:
                    webApi.move_file(itemId, targetFolderId)
                moved = True
        except Exception:
            moved = False

        if moved:
            return

        if isFolder:
            oldSubFolder = self.trimbleConnectClient.get_or_create_child_folder(oldJobFolder, name)
            self.upload_connect_folder(localPathForFallback, oldSubFolder)
            self.trimbleConnectClient.remove_folder(name, parentFolder)
        else:
            self.trimbleConnectClient.save_file_remotely(localPathForFallback, oldJobFolder)
            self.trimbleConnectClient.remove_file(name, parentFolder)

    def download_connect_folder(self, remoteFolder, localFolder):
        # recursively mirrors a whole Connect folder (e.g. a job's "<name> Files" subfolder) into
        # localFolder, preserving its own subfolder structure
        if not os.path.isdir(localFolder):
            os.makedirs(localFolder)

        for child in self.trimbleConnectClient.get_children(remoteFolder):
            childLocalPath = os.path.join(localFolder, str(child.FileName))
            if child.IsFolder:
                self.download_connect_folder(child, childLocalPath)
            else:
                self.trimbleConnectClient.download_file(child, childLocalPath)

    def upload_connect_folder(self, localFolder, remoteFolder):
        # mirror image of download_connect_folder - re-uploads the local copy of a folder tree (e.g. the
        # "<name> Files" folder we just downloaded) into remoteFolder, recreating subfolders as needed
        for name in os.listdir(localFolder):
            fullLocalPath = os.path.join(localFolder, name)
            if os.path.isdir(fullLocalPath):
                childRemoteFolder = self.trimbleConnectClient.get_or_create_child_folder(remoteFolder, name)
                self.upload_connect_folder(fullLocalPath, childRemoteFolder)
            else:
                self.trimbleConnectClient.save_file_remotely(fullLocalPath, remoteFolder)

    def wait_for_file(self, path, timeoutSeconds=5.0, pollIntervalSeconds=0.2):
        # EndDownloadFile() returning doesn't guarantee the destination file is visible yet - there's a
        # brief race where the download itself has finished but a final rename/flush is still in flight,
        # so a single immediate os.path.isfile() check can report a false failure - poll for a few
        # seconds instead of trusting one snapshot
        deadline = time.time() + timeoutSeconds
        while time.time() < deadline:
            if os.path.isfile(path):
                return True
            time.sleep(pollIntervalSeconds)
        return os.path.isfile(path)

    def help_clicked(self, sender, e):
        webbrowser.open(r"C:\ProgramData\Trimble\MacroCommands3\SCR Macros\MacroHelp\MacroHelp.htm#SCR_DayJobSync")

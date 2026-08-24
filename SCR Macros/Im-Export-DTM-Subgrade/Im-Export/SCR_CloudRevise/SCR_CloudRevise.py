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
import webbrowser
exec(open("C:\ProgramData\Trimble\MacroCommands3\SCR Macros\SCR_Imports.py").read())

CIVILLO_BASE_URL = "https://app.civillo.com/api/v1"
CIVILLO_PLACEHOLDER_KEY = "PASTE_YOUR_API_KEY_HERE"
CIVILLO_PLACEHOLDER_SECRET = "PASTE_YOUR_API_SECRET_HERE"


class CivilloClient(object):
    """Thin wrapper around the Civillo REST API (https://docs.civillo.com/api/) shared by the main
    window and the Add Sync dialog, so both use the same auth/error handling."""

    # class-level (not per-instance) so the 1-request-per-second throttle holds across every
    # CivilloClient created during this session - e.g. reload_orgs_clicked makes a fresh instance
    _last_request_time = None

    def __init__(self, config):
        self.config = config

    def _throttle(self):
        # Civillo's API rejects requests faster than 1/second, so pace every real HTTP call here
        now = time.time()
        if CivilloClient._last_request_time is not None:
            elapsed = now - CivilloClient._last_request_time
            if elapsed < 1.0:
                time.sleep(1.1 - elapsed)
        CivilloClient._last_request_time = time.time()

    def _auth_header(self):
        raw = self.config["api_key"] + ":" + self.config["api_secret"]
        return "Bearer " + Convert.ToBase64String(Encoding.UTF8.GetBytes(raw))

    def get(self, relativePath):
        self._throttle()
        request = HttpWebRequest.Create(CIVILLO_BASE_URL + relativePath)
        request.Method = "GET"
        request.Headers.Add("Authorization", self._auth_header())
        return self._read_json_response(request, relativePath)

    def post_json(self, relativePath, bodyDict):
        self._throttle()
        request = HttpWebRequest.Create(CIVILLO_BASE_URL + relativePath)
        request.Method = "POST"
        request.ContentType = "application/json"
        request.Headers.Add("Authorization", self._auth_header())

        payload = Encoding.UTF8.GetBytes(json.dumps(bodyDict))
        request.ContentLength = payload.Length
        stream = request.GetRequestStream()
        stream.Write(payload, 0, payload.Length)
        stream.Close()

        return self._read_json_response(request, relativePath)

    def upload_file(self, processPath, token, application, jobId, localPath, uploadFileName=None):
        # step 2 of the "Create or revise a layer" flow - multipart upload to the returned processPath
        # uploadFileName overrides the name reported in the upload (must match fileNames from step 1)
        fileName = uploadFileName if uploadFileName is not None else os.path.basename(localPath)
        return self.upload_files(processPath, token, application, jobId, [(localPath, fileName)])

    def upload_files(self, processPath, token, application, jobId, files):
        # files: list of (localPath, uploadFileName) tuples - each becomes its own "file" part in one
        # multipart POST. The job's fileNames (from step 1) must list exactly these uploadFileNames,
        # in the same count - Civillo's processor rejects the job if the uploaded files don't match.
        self._throttle()
        primaryExt = os.path.splitext(files[0][1])[1].lstrip(".")
        url = processPath + "?token=" + token + "&application=" + application + "&jobID=" + str(jobId) + "&filetype=" + primaryExt

        boundary = "----SCRCloudReviseBoundary" + str(Guid.NewGuid()).replace("-", "")
        footer = "\r\n--" + boundary + "--\r\n"
        footerBytes = Encoding.UTF8.GetBytes(footer)

        parts = []
        totalLength = 0
        for localPath, uploadFileName in files:
            header = "--" + boundary + "\r\n" + 'Content-Disposition: form-data; name="file"; filename="' + uploadFileName + '"\r\n' + "Content-Type: application/octet-stream\r\n\r\n"
            headerBytes = Encoding.UTF8.GetBytes(header)
            fileBytes = File.ReadAllBytes(localPath)
            parts.append((headerBytes, fileBytes))
            totalLength += headerBytes.Length + fileBytes.Length
        totalLength += footerBytes.Length

        request = HttpWebRequest.Create(url)
        request.Method = "POST"
        request.ContentType = "multipart/form-data; boundary=" + boundary
        request.ContentLength = totalLength

        stream = request.GetRequestStream()
        for headerBytes, fileBytes in parts:
            stream.Write(headerBytes, 0, headerBytes.Length)
            stream.Write(fileBytes, 0, fileBytes.Length)
        stream.Write(footerBytes, 0, footerBytes.Length)
        stream.Close()

        return self._read_json_response(request, processPath, allowNonJson=True)

    def _read_json_response(self, request, relativePath, allowNonJson=False):
        try:
            response = request.GetResponse()
        except WebException as ex:
            detail = ""
            if ex.Response is not None:
                with StreamReader(ex.Response.GetResponseStream(), Encoding.UTF8) as sr:
                    detail = sr.ReadToEnd()
            raise Exception("Civillo API call to " + relativePath + " failed: " + str(ex.Message) + " " + detail)

        with StreamReader(response.GetResponseStream(), Encoding.UTF8) as sr:
            body = sr.ReadToEnd()

        if body == "":
            return {}
        try:
            return json.loads(body)
        except Exception:
            if allowNonJson:
                return {"raw": body}
            raise


def join_folder_and_file(folderPath, fileName):
    if folderPath.endswith("/"):
        return folderPath + fileName
    return folderPath + "/" + fileName


def collect_civillo_files(node, results):
    """Recursively walks a /layer-directory response tree, collecting {"display", "layerId", "folderPath",
    "layerName"} dicts for every file. folderPath/layerName are kept separate (alongside the combined
    "display" string) so callers can render/store them differently, e.g. coloring the folder path and
    the file name differently in a list item."""
    folderPath = node.get("path", "/")
    for layer in node.get("layers", []):
        results.append({
            "display": join_folder_and_file(folderPath, layer["layerName"]),
            "layerId": layer["layerId"],
            "folderPath": folderPath,
            "layerName": layer["layerName"],
        })
    for d in node.get("directories", []):
        collect_civillo_files(d, results)


def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_CloudRevise"
    cmdData.CommandName = "SCR_CloudRevise"
    cmdData.Caption = "_SCR_CloudRevise"
    #cmdData.UIForm = "SCR_CloudRevise"      # left disabled - this is a fully independent floating window, not a TBC-managed dialog
                                                        # if you enable or disable this line, you MUST restart TBC
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "0"

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "Update/Transfer"
        cmdData.ShortCaption = "Cloud Revise"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.05
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""

        cmdData.ToolTipTitle = "CloudRevise"
        cmdData.ToolTipTextFormatted = "transfer new files and revise Civillo layers"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png") # we have to include a icon revision, otherwise TBC might not show the new one
        cmdData.ImageSmall = b
    except:
        pass

def Execute(cmd, currentProject, macroFileFolder, parameters):
    form = SCR_CloudReviseDialog(currentProject, macroFileFolder).Show()
    return
    # .Show() - is non modal - you can interact with the drawing window
    # .ShowDialog() - is modal - you CAN NOT interact with the drawing window


class SCR_CloudReviseDialog(Window): # this inherits from the WPF Window control - a fully independent floating window
    def __init__(self, currentProject, macroFileFolder):

        with StreamReader(macroFileFolder + r"\SCR_CloudRevise.xaml") as s:
            wpf.LoadComponent(self, s)

        ElementHost.EnableModelessKeyboardInterop(self)

        self.currentProject = currentProject
        self.macroFileFolder = macroFileFolder

        self.orgCombo.SelectionChanged += self.org_selection_changed
        self.projectCombo.SelectionChanged += self.project_selection_changed
        self.reloadOrgsBtn.Click += self.reload_orgs_clicked
        self.addScheduleBtn.Click += self.add_schedule_clicked
        self.runSchedulesBtn.Click += self.run_schedules_clicked
        self.helpBtn.Click += self.help_clicked
        self.configPathLabel.MouseLeftButtonDown += self.config_path_clicked

        self.Loaded += self.SetDefaultOptions
        self.Closing += self.SaveOptions

        configPath = self.get_config_path()
        self.configPathLabel.Text = "config: " + configPath
        self.civilloConfig = self.load_or_create_config(configPath)
        self.civilloClient = CivilloClient(self.civilloConfig)

        self.schedules = self.load_schedules()
        self.backfill_schedule_names()
        self.refresh_schedules_ui()

        if self.is_placeholder_config(self.civilloConfig):
            self.error.Content = "Edit the config file above with your Civillo API key/secret, then click 'Reload Organizations'."
        else:
            self.load_organizations()


    # ---------- window position/size ----------

    def SetDefaultOptions(self, sender, e):
        SCROptions.LoadWindowState(self, "SCR_CloudRevise", default_width=300, default_height=460)

    def SaveOptions(self, sender, e):
        SCROptions.SaveWindowState(self, "SCR_CloudRevise")


    # ---------- org/project selection persistence ----------

    def get_saved_selection(self):
        nickname = OptionsManager.GetString("SCR_CloudRevise.selectedorgnickname", "")
        projectId = OptionsManager.GetString("SCR_CloudRevise.selectedprojectid", "")
        return nickname, projectId

    def save_selection(self, nickname, projectId):
        OptionsManager.SetValue("SCR_CloudRevise.selectedorgnickname", str(nickname) if nickname is not None else "")
        OptionsManager.SetValue("SCR_CloudRevise.selectedprojectid", str(projectId) if projectId is not None else "")


    # ---------- config handling ----------

    def get_config_path(self):
        appdata = os.environ.get("APPDATA")
        folder = os.path.join(appdata, "SCR Macros", "SCR_CloudRevise")
        return os.path.join(folder, "civillo_config.json")

    def load_or_create_config(self, path):
        folder = os.path.dirname(path)
        if not os.path.exists(folder):
            os.makedirs(folder)

        if not os.path.exists(path):
            template = {
                "api_key": CIVILLO_PLACEHOLDER_KEY,
                "api_secret": CIVILLO_PLACEHOLDER_SECRET,
            }
            with open(path, "w") as f:
                f.write(json.dumps(template, indent=2))
            return template

        with open(path, "r") as f:
            return json.load(f)

    def is_placeholder_config(self, cfg):
        return cfg.get("api_key", "") == CIVILLO_PLACEHOLDER_KEY or cfg.get("api_secret", "") == CIVILLO_PLACEHOLDER_SECRET

    def config_path_clicked(self, sender, e):
        subprocess.Popen(["notepad.exe", self.get_config_path()])


    # ---------- schedules persistence ----------

    def get_schedules_path(self):
        return os.path.join(os.path.dirname(self.get_config_path()), "schedules.json")

    def load_schedules(self):
        path = self.get_schedules_path()
        if not os.path.exists(path):
            return []
        with open(path, "r") as f:
            return json.load(f)

    def save_schedules(self):
        path = self.get_schedules_path()
        folder = os.path.dirname(path)
        if not os.path.exists(folder):
            os.makedirs(folder)
        with open(path, "w") as f:
            f.write(json.dumps(self.schedules, indent=2))

    def backfill_schedule_names(self):
        # (Re)resolve orgName/projectName for every schedule entry from the live API, so the Schedules
        # tab group header always matches the plain names shown in the Settings tab dropdowns - this
        # both fills in older entries that only have the raw orgNickname/projectId, and corrects any
        # entries that briefly picked up an "extended" name format from a since-reverted change.
        nicknames = set(entry["orgNickname"] for entry in self.schedules if entry.get("orgNickname"))
        if not nicknames:
            return

        orgNameByNickname = {}
        try:
            orgs = self.civilloClient.get("/applications")
            for org in orgs:
                orgNameByNickname[org["nickname"]] = org["name"]
        except Exception:
            return  # API not reachable right now (e.g. placeholder config) - leave as-is for next time

        projectNameByKey = {}
        for nickname in nicknames:
            try:
                projects = self.civilloClient.get("/" + nickname + "/projects")
                for proj in projects:
                    projectNameByKey[(nickname, proj["id"])] = proj["name"]
            except Exception:
                pass

        changed = False
        for entry in self.schedules:
            nickname = entry.get("orgNickname")
            if nickname in orgNameByNickname and entry.get("orgName") != orgNameByNickname[nickname]:
                entry["orgName"] = orgNameByNickname[nickname]
                changed = True
            key = (nickname, entry.get("projectId"))
            if key in projectNameByKey and entry.get("projectName") != projectNameByKey[key]:
                entry["projectName"] = projectNameByKey[key]
                changed = True

        if changed:
            self.save_schedules()


    # ---------- UI actions ----------

    def help_clicked(self, sender, e):
        webbrowser.open(r"C:\ProgramData\Trimble\MacroCommands3\SCR Macros\MacroHelp\MacroHelp.htm#SCR_CloudRevise")

    def reload_orgs_clicked(self, sender, e):
        configPath = self.get_config_path()
        self.civilloConfig = self.load_or_create_config(configPath)
        self.civilloClient = CivilloClient(self.civilloConfig)

        if self.is_placeholder_config(self.civilloConfig):
            self.error.Content = "Edit the config file above with your Civillo API key/secret, then click 'Reload Organizations'."
            return

        self.load_organizations()

    def load_organizations(self):
        self.error.Content = ""
        self.orgCombo.Items.Clear()
        self.projectCombo.Items.Clear()

        try:
            orgs = self.civilloClient.get("/applications")
        except Exception as ex:
            self.error.Content = str(ex)
            return

        for org in orgs:
            item = ComboBoxItem()
            item.Content = org["name"]
            item.Tag = org["nickname"]
            self.orgCombo.Items.Add(item)

        savedNickname, savedProjectId = self.get_saved_selection()
        target = None
        for item in self.orgCombo.Items:
            if item.Tag == savedNickname:
                target = item
                break

        if target is not None:
            self.orgCombo.SelectedItem = target
        elif self.orgCombo.Items.Count > 0:
            self.orgCombo.SelectedIndex = 0

    def org_selection_changed(self, sender, e):
        self.projectCombo.Items.Clear()
        self.error.Content = ""

        selected = self.orgCombo.SelectedItem
        if selected is None:
            return

        nickname = selected.Tag

        try:
            projects = self.civilloClient.get("/" + nickname + "/projects")
        except Exception as ex:
            self.error.Content = str(ex)
            return

        for proj in projects:
            item = ComboBoxItem()
            item.Content = proj["name"]
            item.Tag = proj["id"]
            self.projectCombo.Items.Add(item)

        savedNickname, savedProjectId = self.get_saved_selection()
        target = None
        if nickname == savedNickname:
            for item in self.projectCombo.Items:
                if str(item.Tag) == savedProjectId:
                    target = item
                    break

        if target is not None:
            self.projectCombo.SelectedItem = target
        elif self.projectCombo.Items.Count > 0:
            self.projectCombo.SelectedIndex = 0

    def project_selection_changed(self, sender, e):
        orgItem = self.orgCombo.SelectedItem
        projectItem = self.projectCombo.SelectedItem

        if orgItem is not None and projectItem is not None:
            self.save_selection(orgItem.Tag, projectItem.Tag)


    # ---------- schedules UI ----------

    def refresh_schedules_ui(self):
        self.schedulesPanel.Children.Clear()

        # group rows by org/project (not by insertion order) so an "Org - Site" header always
        # sits above every contiguous block of that org/project's entries, however they were added
        indices = list(range(len(self.schedules)))
        indices.sort(key=lambda i: self.schedule_group_key(self.schedules[i]))

        lastGroupKey = None
        for i in indices:
            entry = self.schedules[i]
            groupKey = self.schedule_group_key(entry)
            if groupKey != lastGroupKey:
                self.schedulesPanel.Children.Add(self.build_schedule_group_header(entry))
                lastGroupKey = groupKey
            self.schedulesPanel.Children.Add(self.build_schedule_row(i))

    def schedule_group_key(self, entry):
        return (entry.get("orgName") or entry.get("orgNickname") or "", entry.get("projectName") or str(entry.get("projectId") or ""))

    def build_schedule_group_header(self, entry):
        orgLabel = entry.get("orgName") or entry.get("orgNickname") or "Unknown Org"
        siteLabel = entry.get("projectName") or str(entry.get("projectId") or "Unknown Site")

        header = TextBlock()
        header.Text = orgLabel + " - " + siteLabel
        header.FontWeight = FontWeights.Bold
        header.Margin = Thickness(0, 8, 0, 4)
        return header

    def build_schedule_row(self, index):
        entry = self.schedules[index]
    
        row = Grid()
        row.Margin = Thickness(0, 0, 0, 8)
    
        colCheck = ColumnDefinition()
        colCheck.Width = GridLength(1, GridUnitType.Auto)
    
        colText = ColumnDefinition()
        colText.Width = GridLength(1, GridUnitType.Star)
    
        colEdit = ColumnDefinition()
        colEdit.Width = GridLength(1, GridUnitType.Auto)
    
        colDelete = ColumnDefinition()
        colDelete.Width = GridLength(1, GridUnitType.Auto)
    
        row.ColumnDefinitions.Add(colCheck)
        row.ColumnDefinitions.Add(colText)
        row.ColumnDefinitions.Add(colEdit)
        row.ColumnDefinitions.Add(colDelete)
    
        chk = CheckBox()
        chk.IsChecked = bool(entry.get("enabled", True))
        Grid.SetColumn(chk, 0)
        chk.Checked += lambda s, e, i=index: self.schedule_enabled_toggled(i, True)
        chk.Unchecked += lambda s, e, i=index: self.schedule_enabled_toggled(i, False)
        row.Children.Add(chk)
    
        # Local paths
        local1Run = Run(entry.get("localPath1", ""))
        local1Run.Foreground = SolidColorBrush(Colors.DarkGreen)
    
        local2Run = Run(entry.get("localPath2", ""))
        local2Run.Foreground = SolidColorBrush(Colors.Brown)
    
        # Split Civillo path into folder and name
        civilloPath = entry.get("civilloPath", "")
        slashIndex = civilloPath.rfind("/")
    
        if slashIndex >= 0:
            civilloFolderText = civilloPath[:slashIndex + 1]
            civilloNameText = civilloPath[slashIndex + 1:]
        else:
            civilloFolderText = ""
            civilloNameText = civilloPath
    
        civilloFolderRun = Run(civilloFolderText)
        civilloFolderRun.Foreground = SolidColorBrush(Colors.Black)
    
        civilloNameRun = Run(civilloNameText)
        civilloNameRun.Foreground = SolidColorBrush(Colors.SteelBlue)
    
        # Text block
        text = TextBlock()
        text.TextWrapping = TextWrapping.Wrap
        text.Margin = Thickness(6, 0, 6, 0)
    
        # Local heading
        localHeader = Run("Local")
        localHeader.FontWeight = FontWeights.Bold
        localHeader.FontStyle = FontStyles.Italic
        text.Inlines.Add(localHeader)
        text.Inlines.Add(LineBreak())
    
        # Local Path 1
        text.Inlines.Add(local1Run)
    
        # Local Path 2 (optional)
        if entry.get("localPath2", ""):
            text.Inlines.Add(LineBreak())
            text.Inlines.Add(local2Run)
    
        # Spacer
        text.Inlines.Add(LineBreak())
    
        # Civillo heading
        civilloHeader = Run("Civillo")
        civilloHeader.FontWeight = FontWeights.Bold
        civilloHeader.FontStyle = FontStyles.Italic
        text.Inlines.Add(civilloHeader)
        text.Inlines.Add(LineBreak())
    
        # Civillo path
        text.Inlines.Add(civilloFolderRun)
        text.Inlines.Add(civilloNameRun)
    
        Grid.SetColumn(text, 1)
        row.Children.Add(text)
    
        editBtn = Button()
        editBtn.Content = "Edit"
        editBtn.Width = 40
        editBtn.Margin = Thickness(0, 0, 4, 0)
        Grid.SetColumn(editBtn, 2)
        editBtn.Click += lambda s, e, i=index: self.edit_schedule_clicked(i)
        row.Children.Add(editBtn)
    
        delBtn = Button()
        delBtn.Content = "X"
        delBtn.Width = 24
        Grid.SetColumn(delBtn, 3)
        delBtn.Click += lambda s, e, i=index: self.delete_schedule_clicked(i)
        row.Children.Add(delBtn)
    
        return row

    def schedule_enabled_toggled(self, index, value):
        if 0 <= index < len(self.schedules):
            self.schedules[index]["enabled"] = value
            self.save_schedules()

    def delete_schedule_clicked(self, index):
        if 0 <= index < len(self.schedules):
            del self.schedules[index]
            self.save_schedules()
            self.refresh_schedules_ui()

    def add_schedule_clicked(self, sender, e):
        orgItem = self.orgCombo.SelectedItem
        projectItem = self.projectCombo.SelectedItem
        nickname = orgItem.Tag if orgItem is not None else None
        projectId = projectItem.Tag if projectItem is not None else None

        dlg = SCR_CloudReviseAddSyncDialog(self.macroFileFolder, self.civilloClient, nickname, projectId)
        dlg.Owner = self
        dlg.ShowDialog()

        if dlg.result is not None:
            entry = {
                "enabled": True,
                "localPath1": dlg.result["localPath1"],
                "localPath2": dlg.result["localPath2"],
                "civilloPath": dlg.result["civilloPath"],
                "civilloLayerId": dlg.result["civilloLayerId"],
                "orgNickname": dlg.result["orgNickname"],
                "projectId": dlg.result["projectId"],
                "orgName": str(orgItem.Content) if orgItem is not None else "",
                "projectName": str(projectItem.Content) if projectItem is not None else "",
            }
            self.schedules.append(entry)
            self.save_schedules()
            self.refresh_schedules_ui()

    def edit_schedule_clicked(self, index):
        if not (0 <= index < len(self.schedules)):
            return

        existingEntry = self.schedules[index]

        dlg = SCR_CloudReviseAddSyncDialog(
            self.macroFileFolder, self.civilloClient,
            existingEntry.get("orgNickname"), existingEntry.get("projectId"),
            existingEntry=existingEntry
        )
        dlg.Owner = self
        dlg.ShowDialog()

        if dlg.result is not None:
            existingEntry["localPath1"] = dlg.result["localPath1"]
            existingEntry["localPath2"] = dlg.result["localPath2"]
            existingEntry["civilloPath"] = dlg.result["civilloPath"]
            existingEntry["civilloLayerId"] = dlg.result["civilloLayerId"]
            existingEntry["orgNickname"] = dlg.result["orgNickname"]
            existingEntry["projectId"] = dlg.result["projectId"]
            existingEntry.pop("lastSyncedHash", None)  # force the next run, even if file 1 is unchanged but file 2 was added/removed
            self.save_schedules()
            self.refresh_schedules_ui()

    def run_schedules_clicked(self, sender, e):
        self.error.Content = ""
        results = []

        for entry in self.schedules:
            if not entry.get("enabled", True):
                continue
            label = os.path.basename(entry.get("localPath1", ""))
            try:
                status = self.run_single_sync(entry)
                if status == "skipped":
                    results.append(label + " -> skipped (unchanged)")
                else:
                    results.append(label + " -> OK")
            except Exception as ex:
                results.append(label + " -> FAILED: " + str(ex))

        if len(results) == 0:
            MessageBox.Show("No enabled sync entries to run.", "SCR_CloudRevise")
        else:
            MessageBox.Show("\n".join(results), "SCR_CloudRevise - Run Results")

    def compute_file_hash(self, path):
        hasher = hashlib.sha256()
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(65536), b""):
                hasher.update(chunk)
        return hasher.hexdigest()

    def run_single_sync(self, entry):
        localPath1 = entry.get("localPath1", "")
        localPath2 = entry.get("localPath2", "")

        if not os.path.isfile(localPath1):
            raise Exception("Local file not found: " + localPath1)
        if localPath2 and not os.path.isfile(localPath2):
            raise Exception("Local file not found: " + localPath2)

        currentHash = self.compute_file_hash(localPath1)
        lastHash = entry.get("lastSyncedHash")

        if lastHash is not None and lastHash == currentHash:
            return "skipped"

        # Civillo's auto-detection of the source SRS is currently unavailable server-side, so we
        # must pass the project's own default projection explicitly instead of relying on srid=-1 (auto).
        projectInfo = self.civilloClient.get("/" + entry["orgNickname"] + "/projects/" + str(entry["projectId"]))
        srid = projectInfo.get("defaultProjection", -1)

        # local file 2 (e.g. a linestyle file) is optional - only include it if the user supplied one
        files = [(localPath1, os.path.basename(localPath1))]
        if localPath2:
            files.append((localPath2, os.path.basename(localPath2)))

        body = {
            "mode": 1,  # 1 = revise an existing layer
            "fileNames": [name for _, name in files],
            "replaceLayerId": entry["civilloLayerId"],
            "fileLastModifieds": [int(os.path.getmtime(path)) for path, _ in files],
            "srid": srid,
        }

        initResp = self.civilloClient.post_json(
            "/" + entry["orgNickname"] + "/projects/" + str(entry["projectId"]) + "/layers",
            body
        )

        self.civilloClient.upload_files(
            initResp["processPath"], initResp["token"], initResp["application"], initResp["job"],
            files
        )

        entry["lastSyncedHash"] = currentHash
        self.save_schedules()

        return "synced"


class SCR_CloudReviseAddSyncDialog(Window):
    """Modal dialog for picking a local source file and an existing Civillo target file/layer to pair up as a sync entry."""

    def __init__(self, macroFileFolder, civilloClient, orgNickname, projectId, existingEntry=None):

        with StreamReader(macroFileFolder + r"\SCR_CloudRevise_AddSync.xaml") as s:
            wpf.LoadComponent(self, s)

        ElementHost.EnableModelessKeyboardInterop(self)

        self.civilloClient = civilloClient
        self.orgNickname = orgNickname
        self.projectId = projectId
        self.existingEntry = existingEntry
        self.result = None

        existingLocalPath1 = existingEntry.get("localPath1", "") if existingEntry is not None else ""
        existingLocalPath2 = existingEntry.get("localPath2", "") if existingEntry is not None else ""

        self.localFolder1 = os.path.dirname(existingLocalPath1) if existingLocalPath1 else None
        self.localFolder2 = os.path.dirname(existingLocalPath2) if existingLocalPath2 else None

        # separate from localFolder1/2 (which drive the file list contents) - only used to seed the
        # folder browse dialogs, so a cleared file still leaves an empty list but Browse opens where you left off
        self.lastBrowseFolder1 = OptionsManager.GetString("SCR_CloudRevise_AddSync.lastlocalfolder1", "") or None
        self.lastBrowseFolder2 = OptionsManager.GetString("SCR_CloudRevise_AddSync.lastlocalfolder2", "") or None

        if existingEntry is not None:
            self.Title = "Edit Sync"
            self.okBtn.Content = "Apply"
        else:
            self.Title = "Add Sync"

        self.browseLocalFolderBtn1.Click += self.browse_local_folder1_clicked
        self.browseLocalFolderBtn2.Click += self.browse_local_folder2_clicked
        self.refreshCivilloFilesBtn.Click += self.refresh_civillo_files_clicked
        self.clearLocalFile1Btn.Click += self.clear_local_file1_clicked
        self.clearLocalFile2Btn.Click += self.clear_local_file2_clicked
        self.clearCivilloFileBtn.Click += self.clear_civillo_file_clicked
        self.okBtn.Click += self.ok_clicked
        self.cancelBtn.Click += self.cancel_clicked

        self.Loaded += self.restore_window_state
        self.Closing += self.save_window_state

        if self.orgNickname and self.projectId is not None:
            self.refresh_civillo_files_clicked(None, None)
        else:
            self.error.Content = "Select an organization and project in Settings first."

        if self.localFolder1:
            self.populate_local_files(self.localFileList1, self.localFolder1)
            self.select_item_by_text(self.localFileList1, os.path.basename(existingLocalPath1))

        if self.localFolder2:
            self.populate_local_files(self.localFileList2, self.localFolder2)
            self.select_item_by_text(self.localFileList2, os.path.basename(existingLocalPath2))


    # ---------- window position/size persistence ----------
    # SizeToContent="Width" (set in the xaml) auto-sizes the window on first run; once the user
    # has resized it, the saved size/position takes over and SizeToContent is switched off so it sticks.

    def restore_window_state(self, sender, e):
        savedWidth = OptionsManager.GetDouble("SCR_CloudRevise_AddSync.windowwidth", 0)
        savedHeight = OptionsManager.GetDouble("SCR_CloudRevise_AddSync.windowheight", 0)
        savedLeft = OptionsManager.GetDouble("SCR_CloudRevise_AddSync.windowleft", -1)
        savedTop = OptionsManager.GetDouble("SCR_CloudRevise_AddSync.windowtop", -1)

        if savedWidth > 0:
            self.SizeToContent = SizeToContent.Manual
            self.Width = savedWidth
        if savedHeight > 0:
            self.Height = savedHeight

        if savedLeft >= 0 and savedTop >= 0:
            self.Left = savedLeft
            self.Top = savedTop
            if not SCROptions._IsWindowOnAnyScreen(self):
                self.Left = 100
                self.Top = 100

        savedLeftStar = OptionsManager.GetDouble("SCR_CloudRevise_AddSync.leftcolumnstar", 0)
        savedRightStar = OptionsManager.GetDouble("SCR_CloudRevise_AddSync.rightcolumnstar", 0)
        if savedLeftStar > 0 and savedRightStar > 0:
            self.rootGrid.ColumnDefinitions[0].Width = GridLength(savedLeftStar, GridUnitType.Star)
            self.rootGrid.ColumnDefinitions[2].Width = GridLength(savedRightStar, GridUnitType.Star)

    def save_window_state(self, sender, e):
        OptionsManager.SetValue("SCR_CloudRevise_AddSync.windowwidth", self.Width)
        OptionsManager.SetValue("SCR_CloudRevise_AddSync.windowheight", self.Height)
        OptionsManager.SetValue("SCR_CloudRevise_AddSync.windowleft", self.Left)
        OptionsManager.SetValue("SCR_CloudRevise_AddSync.windowtop", self.Top)

        OptionsManager.SetValue("SCR_CloudRevise_AddSync.leftcolumnstar", self.rootGrid.ColumnDefinitions[0].Width.Value)
        OptionsManager.SetValue("SCR_CloudRevise_AddSync.rightcolumnstar", self.rootGrid.ColumnDefinitions[2].Width.Value)

    def browse_local_folder1_clicked(self, sender, e):
        dlg = FolderBrowserDialog()
        startFolder1 = self.localFolder1 or self.lastBrowseFolder1
        if startFolder1 and os.path.isdir(startFolder1):
            dlg.SelectedPath = startFolder1
        if dlg.ShowDialog() == DialogResult.OK:
            self.localFolder1 = dlg.SelectedPath
            self.error.Content = ""
            self.populate_local_files(self.localFileList1, self.localFolder1)
            OptionsManager.SetValue("SCR_CloudRevise_AddSync.lastlocalfolder1", self.localFolder1)

    def browse_local_folder2_clicked(self, sender, e):
        dlg = FolderBrowserDialog()
        startFolder2 = self.localFolder2 or self.lastBrowseFolder2
        if startFolder2 and os.path.isdir(startFolder2):
            dlg.SelectedPath = startFolder2
        if dlg.ShowDialog() == DialogResult.OK:
            self.localFolder2 = dlg.SelectedPath
            self.error.Content = ""
            self.populate_local_files(self.localFileList2, self.localFolder2)
            OptionsManager.SetValue("SCR_CloudRevise_AddSync.lastlocalfolder2", self.localFolder2)

    def populate_local_files(self, listBox, folder):
        listBox.Items.Clear()
        if not folder:
            return
        try:
            for name in os.listdir(folder):
                full = os.path.join(folder, name)
                if os.path.isfile(full):
                    listBox.Items.Add(name)
        except Exception as ex:
            self.error.Content = str(ex)

    def select_item_by_text(self, listBox, targetName):
        if not targetName:
            return
        for item in listBox.Items:
            if str(item) == targetName:
                listBox.SelectedItem = item
                listBox.UpdateLayout()  # force layout so ScrollIntoView works before the window is shown
                listBox.ScrollIntoView(item)
                break

    def clear_local_file1_clicked(self, sender, e):
        self.localFileList1.SelectedItem = None

    def clear_local_file2_clicked(self, sender, e):
        self.localFileList2.SelectedItem = None

    def clear_civillo_file_clicked(self, sender, e):
        self.civilloFileList.SelectedItem = None

    def refresh_civillo_files_clicked(self, sender, e):
        self.civilloFileList.Items.Clear()
        self.error.Content = ""

        if not self.orgNickname or self.projectId is None:
            self.error.Content = "Select an organization and project in Settings first."
            return

        try:
            tree = self.civilloClient.get("/" + self.orgNickname + "/projects/" + str(self.projectId) + "/layer-directory")
        except Exception as ex:
            self.error.Content = str(ex)
            return

        files = []
        collect_civillo_files(tree, files)

        for f in files:
            folderText = f["folderPath"] if f["folderPath"].endswith("/") else f["folderPath"] + "/"

            folderRun = Run(folderText)
            folderRun.Foreground = SolidColorBrush(Colors.Black)  # matches the "->" separator color on the Schedules tab

            nameRun = Run(f["layerName"])
            nameRun.Foreground = SolidColorBrush(Colors.SteelBlue)  # matches the Civillo path color on the Schedules tab

            textBlock = TextBlock()
            textBlock.Inlines.Add(folderRun)
            textBlock.Inlines.Add(nameRun)

            item = ListBoxItem()
            item.Content = textBlock
            item.Tag = {"layerId": f["layerId"], "display": f["display"]}
            self.civilloFileList.Items.Add(item)

        if self.existingEntry is not None:
            targetLayerId = self.existingEntry.get("civilloLayerId")
            for item in self.civilloFileList.Items:
                if item.Tag["layerId"] == targetLayerId:
                    self.civilloFileList.SelectedItem = item
                    self.civilloFileList.UpdateLayout()  # force layout so ScrollIntoView works before the window is shown
                    self.civilloFileList.ScrollIntoView(item)
                    break

    def ok_clicked(self, sender, e):
        local1Item = self.localFileList1.SelectedItem
        local2Item = self.localFileList2.SelectedItem
        civilloItem = self.civilloFileList.SelectedItem

        if local1Item is None or civilloItem is None:
            self.error.Content = "Select at least local file 1 and a Civillo target file."
            return

        localPath1 = os.path.join(self.localFolder1, str(local1Item))
        localPath2 = os.path.join(self.localFolder2, str(local2Item)) if local2Item is not None else ""

        self.result = {
            "localPath1": localPath1,
            "localPath2": localPath2,
            "civilloPath": civilloItem.Tag["display"],
            "civilloLayerId": civilloItem.Tag["layerId"],
            "orgNickname": self.orgNickname,
            "projectId": self.projectId,
        }
        self.Close()

    def cancel_clicked(self, sender, e):
        self.result = None
        self.Close()

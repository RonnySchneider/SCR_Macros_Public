#   GNU GPLv3
#   <this is an add-on Script/Macro for the geospatial software "Trimble Business Center" aka TBC>
#   <you'll need at least the "Survey Advanced" licence of TBC in order to run this script>
#   <see the Help-Files for more details>
#   Copyright (C) 2023 Ronny Schneider
#
#   This program is free software: you can redistribute it and/or modify
#   it under the terms of the GNU General Public License as published by
#   the Free Software Foundation, either version 3 of the License, or
#   (at your option) any later version.

import clr
import json
import os
import re
import shutil
import threading
import webbrowser

clr.AddReference("IronPython.Wpf")
import wpf

clr.AddReference("System.Data")
clr.AddReference("System.Drawing")
clr.AddReference("WindowsFormsIntegration")

import System
from System import Action, Array, Byte
from System.Data import DataTable
from System.Drawing import Bitmap
from System.IO import File, MemoryStream, StreamReader
from System.Net import HttpWebRequest
from System.Security.Cryptography import SHA1
from System.Text import Encoding
from System.Windows import Visibility, Window
from System.Windows.Threading import DispatcherPriority
from System.Windows.Forms.Integration import ElementHost


# ── embedded XAML ─────────────────────────────────────────────────────────────

_XAML = """<Window Background="#FFF3F3F7"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="SCR Macro Update – GitHub" Height="600" Width="1000"
        MinHeight="320" MinWidth="520">

    <Window.Resources>
        <Style TargetType="DataGridColumnHeader" x:Key="hdrStyle">
            <Setter Property="HorizontalContentAlignment" Value="Center"/>
            <Setter Property="BorderBrush"    Value="Gray"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
            <Setter Property="Background"     Value="#FFCBD8EA"/>
            <Setter Property="FontWeight"     Value="Bold"/>
            <Setter Property="Padding"        Value="4,3"/>
        </Style>
        <Style TargetType="DataGridRow">
            <Style.Triggers>
                <DataTrigger Binding="{Binding [status]}" Value="Up to date">
                    <Setter Property="Background" Value="#FFD4F0D4"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding [status]}" Value="Update available">
                    <Setter Property="Background" Value="#FFFFF0C0"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding [status]}" Value="New">
                    <Setter Property="Background" Value="#FFD4E8FF"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding [status]}" Value="Incomplete">
                    <Setter Property="Background" Value="#FFFFD8A0"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding [status]}" Value="Local is newer">
                    <Setter Property="Background" Value="#FFECE0FF"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding [status]}" Value="Modified">
                    <Setter Property="Background" Value="#FFE8E8E8"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding [status]}" Value="Local only">
                    <Setter Property="Background" Value="#FFFFD8B0"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding [status]}" Value="Deleted">
                    <Setter Property="Background" Value="#FFE0E0E0"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding [status]}" Value="Error">
                    <Setter Property="Background" Value="#FFFFCCCC"/>
                </DataTrigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>

    <Grid Margin="8">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <UniformGrid Grid.Row="0" Rows="1" Columns="4" Margin="0,0,0,6">
            <Button x:Name="fetchbtn"       Content="&#x2193;  Fetch File List from GitHub" Height="30" Margin="0,0,4,0" FontWeight="Bold" Background="#FFBDE0FF"/>
            <Button x:Name="selectallbtn"   Content="Select All"   Height="30" Margin="2,0"/>
            <Button x:Name="deselectallbtn" Content="Deselect All" Height="30" Margin="2,0"/>
            <Button x:Name="helpbtn"        Content="Help"         Height="30" Margin="4,0,0,0"/>
        </UniformGrid>

        <DataGrid Grid.Row="1" x:Name="datagrid"
                  AutoGenerateColumns="False" ItemsSource="{Binding}"
                  CanUserAddRows="False" CanUserDeleteRows="False" CanUserReorderColumns="False"
                  GridLinesVisibility="All"
                  BorderBrush="Gray" BorderThickness="1"
                  ColumnHeaderStyle="{StaticResource hdrStyle}"
                  HorizontalScrollBarVisibility="Auto"
                  VerticalScrollBarVisibility="Auto"
                  EnableRowVirtualization="False">
            <DataGrid.Columns>
                <DataGridTemplateColumn Header="&#x2193;" Width="Auto" CanUserSort="False">
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <CheckBox IsChecked="{Binding [download], Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
                                      HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                <DataGridTextColumn Header="File Path"     Binding="{Binding [path]}"         SortMemberPath="path"          Width="*"           MinWidth="200" IsReadOnly="True"/>
                <DataGridTextColumn Header="Local Version" Binding="{Binding [local_version]}" SortMemberPath="local_version" Width="SizeToCells" MinWidth="90"  IsReadOnly="True"/>
                <DataGridTextColumn Header="Status"        Binding="{Binding [status]}"        SortMemberPath="status"        Width="SizeToCells" MinWidth="80"  IsReadOnly="True"/>
            </DataGrid.Columns>
        </DataGrid>

        <Button Grid.Row="2" x:Name="downloadbtn"
                Content="&#x2193;  Download selected   –   ⚠ ticked &#x27;Local only&#x27; files will be deleted"
                Height="34" Margin="0,6,0,0"
                FontWeight="Bold" Background="#FFC0E8A0"/>

        <ProgressBar Grid.Row="3" x:Name="progressbar"
                     Height="18" Margin="0,6,0,0"
                     Minimum="0" Value="0"
                     Visibility="Collapsed"/>

        <StackPanel Grid.Row="4" Margin="0,4,0,0">
            <Label x:Name="error"   Content="" Foreground="#FFCC0000" FontWeight="Bold" Padding="2"/>
            <Label x:Name="success" Content="" Foreground="#FF006600" Padding="2"/>
        </StackPanel>
    </Grid>
</Window>"""

# filenames always downloaded silently alongside any macro batch, never shown in the grid
_HELPER_FILENAMES      = {"scr_imports.py", "scr_globalhelpers.py"}
# resource files always silently copied with any macro download, never shown in the grid
_ALWAYS_COPY_FILENAMES = {"ab arrow block.vcl"}


def Setup(cmdData, macroFileFolder):
    cmdData.Key         = "SCR_MacroUpdate"
    cmdData.CommandName = "SCR_MacroUpdate"
    cmdData.Caption     = "_SCR_MacroUpdate"
    cmdData.HelpFile    = "Macros.chm"
    cmdData.HelpTopic   = "22602"

    try:
        cmdData.DefaultTabKey         = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey    = "Update/Transfer"
        cmdData.ShortCaption          = "Update from GitHub"
        cmdData.DefaultRibbonToolSize = 3
        cmdData.EnableNoProject       = True

        cmdData.Version     = 1.14
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo   = r""

        cmdData.ToolTipTitle         = "Update SCR Macros from GitHub"
        cmdData.ToolTipTextFormatted = "Download and update SCR Macros from a GitHub repository"
    except:
        pass
    try:
        cmdData.ImageSmall = Bitmap(macroFileFolder + "\\" + cmdData.Key + ".png")
    except:
        pass


def Execute(cmd, currentProject, macroFileFolder, parameters):
    SCR_MacroUpdateDialog(currentProject, macroFileFolder).Show()


class SCR_MacroUpdateDialog(Window):

    def __init__(self, currentProject, macroFileFolder):
        stream = MemoryStream(Encoding.UTF8.GetBytes(_XAML))
        wpf.LoadComponent(self, StreamReader(stream))

        ElementHost.EnableModelessKeyboardInterop(self)

        self.currentProject  = currentProject
        self.macroFileFolder = macroFileFolder
        self._dt             = None
        self._repo_tree      = []
        self._repo_owner     = ""
        self._repo_name      = ""

        self._repo_url       = "https://github.com/RonnySchneider/SCR_Macros_Public"
        self._repo_subfolder = "SCR Macros"   # e.g. "subfolder" to sync only that subtree; "" = whole repo
        self._local_root     = r"C:\ProgramData\Trimble\MacroCommands3\SCR Macros"
        self._propagating    = False
        self._downloading    = False
        self._cancel         = False

        self.fetchbtn.Click       += self._fetch_click
        self.selectallbtn.Click   += self._select_all
        self.deselectallbtn.Click += self._deselect_all
        self.helpbtn.Click        += self._help_click
        self.downloadbtn.Click    += self._download_click

    # ── helpers ────────────────────────────────────────────────────────────

    def _parse_repo(self, url):
        url = url.strip().rstrip("/")
        m = re.search(r'github\.com/([^/]+)/([^/]+?)(?:\.git)?$', url)
        return (m.group(1), m.group(2)) if m else (None, None)

    def _api_get(self, url):
        req = HttpWebRequest.Create(url)
        req.Method    = "GET"
        req.UserAgent = "SCR-MacroUpdate/2.0"
        req.Accept    = "application/vnd.github.v3+json"
        req.Timeout   = 20000
        resp = req.GetResponse()
        with StreamReader(resp.GetResponseStream()) as rdr:
            return json.loads(rdr.ReadToEnd())

    _TEXT_EXTENSIONS       = {'.py', '.xaml', '.htm', '.html', '.txt', '.json', '.xml', '.csv', '.md'}
    _REQUIRED_EXTENSIONS   = {'.py', '.xaml', '.png', '.htm', '.html'}  # only these trigger Incomplete
    _SKIP_EXTENSIONS       = {'.p', '.recursiveversion', '.bak', '.tmp'}  # dev artifacts, never download

    def _git_blob_sha(self, file_path):
        try:
            raw = File.ReadAllBytes(file_path)
            ext = os.path.splitext(file_path)[1].lower()
            if ext in self._TEXT_EXTENSIONS:
                content = Encoding.UTF8.GetBytes(
                    Encoding.UTF8.GetString(raw).replace("\r\n", "\n").replace("\r", "\n"))
            else:
                content = raw
            prefix = Encoding.ASCII.GetBytes("blob " + str(len(content)) + "\0")
            ms = MemoryStream()
            ms.Write(prefix,  0, len(prefix))
            ms.Write(content, 0, len(content))
            return "".join("{0:02x}".format(int(b)) for b in SHA1.Create().ComputeHash(ms.ToArray()))
        except:
            return None

    def _extract_version(self, file_path):
        try:
            if not file_path.lower().endswith(".py"):
                return "-"
            txt = File.ReadAllText(file_path, Encoding.UTF8)
            m = re.search(r'cmdData\.Version\s*=\s*([\d.]+)', txt)
            return m.group(1) if m else "-"
        except:
            return "-"

    def _fetch_remote_version(self, repo_path):
        try:
            prefix = self._repo_subfolder.strip("/")
            repo_path_prefix = prefix + "/" if prefix else ""
            raw_url = "https://raw.githubusercontent.com/{}/{}/HEAD/{}".format(
                self._repo_owner, self._repo_name, repo_path_prefix + repo_path)
            req = HttpWebRequest.Create(raw_url)
            req.Method    = "GET"
            req.UserAgent = "SCR-MacroUpdate/2.0"
            req.Timeout   = 10000
            resp = req.GetResponse()
            with StreamReader(resp.GetResponseStream()) as rdr:
                txt = rdr.ReadToEnd()
            m = re.search(r'cmdData\.Version\s*=\s*([\d.]+)', txt)
            return m.group(1) if m else "-"
        except:
            return "-"

    def _make_table(self):
        dt = DataTable("files")
        dt.Columns.Add("download",      System.Type.GetType("System.Boolean"))
        dt.Columns.Add("path",         System.Type.GetType("System.String"))
        dt.Columns.Add("local_version",System.Type.GetType("System.String"))
        dt.Columns.Add("status",       System.Type.GetType("System.String"))
        dt.Columns.Add("remote_sha",   System.Type.GetType("System.String"))
        dt.ColumnChanged += self._on_download_changed
        return dt

    def _on_download_changed(self, sender, e):
        if self._propagating or e.Column.ColumnName != "download":
            return
        if not self.Dispatcher.CheckAccess():
            return  # called from background download thread – skip multi-row propagation
        if self._dt is None:
            return  # table still being built during populate
        self._propagating = True
        try:
            new_val     = e.Row["download"]
            changed_row = e.Row
            for item in self.datagrid.SelectedItems:
                try:
                    row = item.Row
                    if row != changed_row:
                        row["download"] = new_val
                except Exception:
                    pass
            # helpers are mandatory whenever any non-helper download (not delete) row is ticked
            any_dl = any(
                bool(r["download"])
                for r in self._dt.Rows
                if r["path"].split("/")[-1].lower() not in _HELPER_FILENAMES
                and r["status"] != "Local only"
                and not r["path"].lower().endswith(".htm")
            )
            for r in self._dt.Rows:
                if r["path"].split("/")[-1].lower() in _HELPER_FILENAMES:
                    r["download"] = any_dl
        finally:
            self._propagating = False

    def _populate_grid(self, tree_items, local_root):
        dt         = self._make_table()
        local_root = local_root.replace("/", "\\")

        # repo files
        repo_py_paths = set()
        for item in tree_items:
            if item.get("type") != "blob" or not item["path"].endswith(".py"):
                continue
            path       = item["path"]
            remote_sha = item.get("sha", "")
            local_path = os.path.join(local_root, path.replace("/", "\\"))
            exists     = os.path.isfile(local_path)
            local_sha  = self._git_blob_sha(local_path) if exists else None
            repo_py_paths.add(path.replace("/", "\\").lower())

            if not exists:
                local_ver = "-"
                status    = "New"
                do_dl     = False
            elif local_sha == remote_sha:
                local_ver  = self._extract_version(local_path)
                py_folder  = "/".join(path.split("/")[:-1])
                incomplete = None
                for si in tree_items:
                    if si.get("type") != "blob" or si["path"] == path:
                        continue
                    if os.path.splitext(si["path"])[1].lower() not in self._REQUIRED_EXTENSIONS:
                        continue
                    if "/".join(si["path"].split("/")[:-1]) != py_folder:
                        continue
                    si_local = os.path.join(local_root, si["path"].replace("/", "\\"))
                    if not os.path.isfile(si_local) or self._git_blob_sha(si_local) != si.get("sha", ""):
                        incomplete = si["path"].split("/")[-1]
                        break
                status = "Incomplete ({})".format(incomplete) if incomplete else "Up to date"
                do_dl  = False
            else:
                local_ver  = self._extract_version(local_path)
                remote_ver = self._fetch_remote_version(path)
                try:
                    lv = float(local_ver)  if local_ver  != "-" else 0.0
                    rv = float(remote_ver) if remote_ver != "-" else 0.0
                except ValueError:
                    lv = rv = 0.0
                if rv > lv:
                    status = "Update available"
                elif lv > rv:
                    status = "Local is newer"
                else:
                    status = "Modified"
                do_dl = False

            if path.split("/")[-1].lower() in _HELPER_FILENAMES:
                local_ver = "mandatory on macro update"

            row = dt.NewRow()
            row["download"]      = do_dl
            row["path"]         = path
            row["local_version"]= local_ver
            row["status"]       = status
            row["remote_sha"]   = remote_sha
            dt.Rows.Add(row)

        # .htm files in folders that contain no .py (e.g. MacroHelp)
        repo_py_folders = set(
            "/".join(i["path"].split("/")[:-1])
            for i in tree_items
            if i.get("type") == "blob" and i["path"].endswith(".py")
        )
        for item in tree_items:
            if item.get("type") != "blob" or not item["path"].lower().endswith(".htm"):
                continue
            path   = item["path"]
            folder = "/".join(path.split("/")[:-1])
            if folder in repo_py_folders:
                continue
            remote_sha = item.get("sha", "")
            local_path = os.path.join(local_root, path.replace("/", "\\"))
            exists     = os.path.isfile(local_path)
            local_sha  = self._git_blob_sha(local_path) if exists else None
            folder_prefix = folder + "/" if folder else ""
            if not exists:
                status = "New"
                do_dl  = False
            elif local_sha == remote_sha:
                incomplete = None
                for si in tree_items:
                    if si.get("type") != "blob" or si["path"] == path:
                        continue
                    if os.path.splitext(si["path"])[1].lower() not in self._REQUIRED_EXTENSIONS:
                        continue
                    if not si["path"].startswith(folder_prefix):
                        continue
                    si_local = os.path.join(local_root, si["path"].replace("/", "\\"))
                    if not os.path.isfile(si_local) or self._git_blob_sha(si_local) != si.get("sha", ""):
                        incomplete = si["path"].split("/")[-1]
                        break
                status = "Incomplete ({})".format(incomplete) if incomplete else "Up to date"
                do_dl  = False
            else:
                status = "Update available"
                do_dl  = False
            row = dt.NewRow()
            file_count = sum(1 for i in tree_items
                             if i.get("type") == "blob"
                             and (i["path"] == path or i["path"].startswith(folder_prefix))
                             and not i["path"].lower().endswith(".dict"))
            row["download"]      = do_dl
            row["path"]         = path
            row["local_version"]= "{} files".format(file_count)
            row["status"]       = status
            row["remote_sha"]   = remote_sha
            dt.Rows.Add(row)

        # local-only .py files not present in the repo
        if os.path.isdir(local_root):
            for dirpath, _, filenames in os.walk(local_root):
                for fname in filenames:
                    if not fname.endswith(".py"):
                        continue
                    abs_path = os.path.join(dirpath, fname)
                    rel_path = abs_path[len(local_root):].lstrip("\\")
                    if rel_path.lower() in repo_py_paths:
                        continue
                    row = dt.NewRow()
                    row["download"]      = False
                    row["path"]         = rel_path.replace("\\", "/")
                    row["local_version"]= self._extract_version(abs_path)
                    row["status"]       = "Local only"
                    row["remote_sha"]   = ""
                    dt.Rows.Add(row)

        self._dt = dt
        self.datagrid.DataContext = dt

    def _rebind_grid(self):
        self.datagrid.DataContext = None
        self.datagrid.DataContext = self._dt

    # ── event handlers ─────────────────────────────────────────────────────

    def _fetch_click(self, sender, e):
        self.error.Content   = ""
        self.success.Content = "Fetching file list from GitHub…"
        owner, repo = self._parse_repo(self._repo_url)
        if not owner:
            self.error.Content   = "Invalid GitHub URL – expected https://github.com/owner/repo"
            self.success.Content = ""
            return

        self.fetchbtn.IsEnabled = False
        try:
            data      = self._api_get(
                "https://api.github.com/repos/{}/{}/git/trees/HEAD?recursive=1".format(owner, repo)
            )
            tree      = data.get("tree", [])
            truncated = data.get("truncated", False)

            # optionally restrict to a subfolder, stripping the prefix from all paths
            prefix = self._repo_subfolder.strip("/")
            if prefix:
                pfx = prefix + "/"
                tree = [dict(item, path=item["path"][len(pfx):])
                        for item in tree
                        if item["path"].startswith(pfx) and item["path"] != pfx]

            self._repo_owner = owner
            self._repo_name  = repo
            self._repo_tree  = tree

            self._populate_grid(tree, self._local_root)

            pyfiles = [i for i in tree if i.get("type") == "blob" and i["path"].endswith(".py")]
            msg = "{} macro(s) found in repo".format(len(pyfiles))
            if truncated:
                msg += "  ⚠ result was truncated by GitHub – tree may be incomplete"
            self.success.Content = msg
        except Exception as ex:
            self.error.Content   = "GitHub error: " + str(ex)
            self.success.Content = ""
        finally:
            self.fetchbtn.IsEnabled = True

    def _help_click(self, sender, e):
        webbrowser.open(r"C:\ProgramData\Trimble\MacroCommands3\SCR Macros\MacroHelp\MacroHelp.htm#SCR_MacroUpdate")

    def _select_all(self, sender, e):
        if self._dt is None:
            return
        for row in self._dt.Rows:
            row["download"] = True
        self._rebind_grid()

    def _deselect_all(self, sender, e):
        if self._dt is None:
            return
        for row in self._dt.Rows:
            row["download"] = False
        self._rebind_grid()

    def _ui(self, func):
        self.Dispatcher.BeginInvoke(DispatcherPriority.Normal, Action(func))

    def _download_click(self, sender, e):
        if self._downloading:
            self._cancel = True
            return

        self.error.Content   = ""
        self.success.Content = ""

        if self._dt is None:
            self.error.Content = "Fetch the file list first."
            return

        local_root = self._local_root
        if not os.path.isdir(local_root):
            try:
                os.makedirs(local_root)
            except Exception as ex:
                self.error.Content = "Cannot create folder: " + str(ex)
                return

        # pre-extract everything from the DataTable on the UI thread
        jobs = []
        for row in self._dt.Rows:
            if not bool(row["download"]):
                continue
            py_path = row["path"]
            if row["status"] == "Local only":
                jobs.append({"row": row, "py_path": py_path, "siblings": [], "delete": True, "is_new": False})
            else:
                folder = "/".join(py_path.split("/")[:-1])
                if py_path.lower().endswith(".htm"):
                    # include everything under the folder recursively
                    folder_prefix = folder + "/" if folder else ""
                    siblings = [item["path"] for item in self._repo_tree
                                if item.get("type") == "blob"
                                and (item["path"] == py_path or item["path"].startswith(folder_prefix))
                                and os.path.splitext(item["path"])[1].lower()
                                and os.path.splitext(item["path"])[1].lower() not in self._SKIP_EXTENSIONS]
                else:
                    siblings = [item["path"] for item in self._repo_tree
                                if item.get("type") == "blob"
                                and "/".join(item["path"].split("/")[:-1]) == folder
                                and os.path.splitext(item["path"])[1].lower()
                                and os.path.splitext(item["path"])[1].lower() not in self._SKIP_EXTENSIONS]
                jobs.append({"row": row, "py_path": py_path, "siblings": siblings, "delete": False, "is_new": row["status"] == "New"})

        if not jobs:
            self.error.Content = "No files selected for download."
            return

        # silently prepend resource files that should always accompany any macro download
        if any(not j["delete"] for j in jobs):
            for item in self._repo_tree:
                if item.get("type") == "blob" and item["path"].split("/")[-1].lower() in _ALWAYS_COPY_FILENAMES:
                    jobs.insert(0, {"row": None, "py_path": item["path"], "siblings": [item["path"]], "delete": False, "is_new": False})

        total_files = sum(1 if j["delete"] else len(j["siblings"]) for j in jobs)
        self._downloading             = True
        self._cancel                  = False
        self.fetchbtn.IsEnabled       = False
        self.selectallbtn.IsEnabled   = False
        self.deselectallbtn.IsEnabled = False
        self.downloadbtn.Content      = "✕  Cancel"
        self.progressbar.Maximum      = max(total_files, 1)
        self.progressbar.Value        = 0
        self.progressbar.Visibility   = Visibility.Visible

        t = threading.Thread(
            target=self._do_download,
            args=(jobs, self._repo_owner, self._repo_name, local_root)
        )
        t.daemon = True
        t.start()

    def _do_download(self, jobs, owner, repo, local_root):
        done         = 0
        files_done   = 0
        new_count    = 0
        update_count = 0
        delete_count = 0
        errors       = []
        prefix = self._repo_subfolder.strip("/")
        repo_path_prefix = prefix + "/" if prefix else ""

        try:
            for job in jobs:
                if self._cancel:
                    break
                row        = job["row"]
                py_path    = job["py_path"]
                macro_name = py_path.split("/")[-1]

                n = done + 1
                self._ui(lambda n=n, name=macro_name:
                    setattr(self.success, "Content",
                            "{}/{}: {}".format(n, len(jobs), name)))

                if job["delete"]:
                    macro_dir = os.path.dirname(os.path.join(local_root, py_path.replace("/", "\\")))
                    try:
                        shutil.rmtree(macro_dir)
                        if row is not None:
                            row["status"]   = "Deleted"
                            row["download"] = False
                        delete_count += 1
                    except Exception as ex:
                        errors.append(py_path + ": " + str(ex))
                        if row is not None:
                            row["status"] = "Error"
                    files_done += 1
                    fd = files_done
                    self._ui(lambda fd=fd: setattr(self.progressbar, "Value", fd))
                else:
                    row_ok = True
                    for path in job["siblings"]:
                        if self._cancel:
                            break
                        raw_url  = "https://raw.githubusercontent.com/{}/{}/HEAD/{}".format(owner, repo, repo_path_prefix + path)
                        loc_path = os.path.join(local_root, path.replace("/", "\\"))
                        loc_dir  = os.path.dirname(loc_path)
                        try:
                            if not os.path.isdir(loc_dir):
                                os.makedirs(loc_dir)

                            req = HttpWebRequest.Create(raw_url)
                            req.Method    = "GET"
                            req.UserAgent = "SCR-MacroUpdate/2.0"
                            req.Timeout   = 30000
                            resp   = req.GetResponse()
                            stream = resp.GetResponseStream()
                            ms     = MemoryStream()
                            buf    = Array.CreateInstance(Byte, 8192)
                            while True:
                                nb = stream.Read(buf, 0, 8192)
                                if nb == 0:
                                    break
                                ms.Write(buf, 0, nb)
                            stream.Close()
                            File.WriteAllBytes(loc_path, ms.ToArray())

                        except Exception as ex:
                            errors.append(path + ": " + str(ex))
                            row_ok = False

                        files_done += 1
                        fd = files_done
                        self._ui(lambda fd=fd: setattr(self.progressbar, "Value", fd))

                    if row is not None:
                        py_local = os.path.join(local_root, py_path.replace("/", "\\"))
                        row["local_version"] = self._extract_version(py_local)
                        row["status"]        = "Up to date" if row_ok else "Error"
                        row["download"]      = False
                    if row_ok and row is not None:
                        if job["is_new"]:
                            new_count    += 1
                        else:
                            update_count += 1

                done += 1

        except Exception as ex:
            errors.append("Unexpected error: " + str(ex))

        # silently remove local folders that have no .py and are not in the repo
        repo_dir_paths = set()
        for item in self._repo_tree:
            if item.get("type") == "blob":
                segs = item["path"].split("/")
                for i in range(1, len(segs)):
                    repo_dir_paths.add("/".join(segs[:i]).lower())
        for dirpath, subdirs, filenames in os.walk(local_root, topdown=False):
            if dirpath == local_root:
                continue
            if any(f.endswith(".py") for f in filenames):
                continue
            rel_dir = dirpath[len(local_root):].lstrip("\\").replace("\\", "/")
            if rel_dir.lower() in repo_dir_paths:
                continue
            if filenames or not subdirs:
                try:
                    shutil.rmtree(dirpath)
                except Exception:
                    pass

        parts = []
        if new_count:    parts.append("{} new".format(new_count))
        if update_count: parts.append("{} updated".format(update_count))
        if delete_count: parts.append("{} deleted".format(delete_count))
        ok_text      = "Done – " + ", ".join(parts) + "." if parts else "Done."
        err_text     = "\n".join(errors)
        restart_text = ("⚠  New macros were added – please restart TBC to make them available!"
                        if new_count > 0 else "")

        was_cancelled = self._cancel
        if was_cancelled:
            ok_text      = "Download cancelled – refreshing file list…"
            err_text     = ""
            restart_text = ""

        def finish(e=err_text, o=ok_text, r=restart_text, cancelled=was_cancelled):
            self._downloading             = False
            self._cancel                  = False
            self.error.Content            = "\n".join(x for x in [r, e] if x)
            self.success.Content          = o
            self.progressbar.Visibility   = Visibility.Collapsed
            self.fetchbtn.IsEnabled       = True
            self.downloadbtn.Content      = "↓  Download selected   –   ⚠ ticked 'Local only' files will be deleted"
            self.selectallbtn.IsEnabled   = True
            self.deselectallbtn.IsEnabled = True
            self._rebind_grid()
            if cancelled:
                self._fetch_click(None, None)

        self._ui(finish)

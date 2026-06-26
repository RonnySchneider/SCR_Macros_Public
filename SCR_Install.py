#   GNU GPLv3
#   <this is an add-on Script/Macro for the geospatial software "Trimble Business Center" aka TBC>
#   <you'll need at least the "Survey Advanced" licence of TBC in order to run this script>
#   Copyright (C) 2023 Ronny Schneider
#
#   Single-file installer: downloads SCR Macros from GitHub, then deletes itself.
#   Place this file in a subfolder of the target macro folder, e.g.:
#     <macro root>\SCR Synchtest\SCR_Install\SCR_Install.py
#   After download the entire SCR_Install folder is removed automatically.

import clr
import json
import os
import re
import subprocess
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
from System.Windows.Forms.Integration import ElementHost
from System.Windows.Threading import DispatcherPriority

# ── embedded XAML ─────────────────────────────────────────────────────────────

_XAML = """<Window Background="#FFF3F3F7"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="SCR Macro Installer" Height="600" Width="880"
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

        <UniformGrid Grid.Row="0" Rows="1" Columns="3" Margin="0,0,0,6">
            <Button x:Name="fetchbtn"       Content="&#x2193;  Fetch File List from GitHub" Height="30" Margin="0,0,4,0" FontWeight="Bold" Background="#FFBDE0FF"/>
            <Button x:Name="selectallbtn"   Content="Select All"   Height="30" Margin="2,0"/>
            <Button x:Name="deselectallbtn" Content="Deselect All" Height="30" Margin="4,0,0,0"/>
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
                Content="&#x2193;  Install Selected Macros"
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
_HELPER_FILENAMES = {"scr_imports.py", "scr_globalhelpers.py"}

# sync macro — mandatory install so users can update macros later
_SYNC_FILENAME = "scr_macroupdate.py"

# ── TBC macro entry points ────────────────────────────────────────────────────

def Setup(cmdData, macroFileFolder):
    cmdData.Key         = "SCR_Install"
    cmdData.CommandName = "SCR_Install"
    cmdData.Caption     = "_SCR_Install"
    cmdData.HelpFile    = "Macros.chm"
    cmdData.HelpTopic   = "22602"

    try:
        cmdData.DefaultTabKey         = "SCR Macro Installer"
        cmdData.DefaultTabGroupKey    = "Install"
        cmdData.ShortCaption          = "Install SCR Macros"
        cmdData.DefaultRibbonToolSize = 3
        cmdData.EnableNoProject       = True

        cmdData.Version     = 1.01
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo   = r""

        cmdData.ToolTipTitle         = "Install SCR Macros from GitHub"
        cmdData.ToolTipTextFormatted = "First-time installer: downloads SCR Macros from GitHub and removes itself afterwards"
    except:
        pass
    try:
        b = Bitmap(macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


def Execute(cmd, currentProject, macroFileFolder, parameters):
    SCR_InstallDialog(currentProject, macroFileFolder).Show()


# ── dialog ────────────────────────────────────────────────────────────────────

class SCR_InstallDialog(Window):

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
        self._did_install       = False
        self._did_install_macro = False
        self._propagating       = False

        self._repo_url       = "https://github.com/RonnySchneider/SCR_Macros_Public"
        self._repo_subfolder = "SCR Macros"   # e.g. "subfolder" to sync only that subtree; "" = whole repo
        self._local_root     = r"C:\ProgramData\Trimble\MacroCommands3\SCR Macros"

        self.fetchbtn.Click       += self._fetch_click
        self.selectallbtn.Click   += self._select_all
        self.deselectallbtn.Click += self._deselect_all
        self.downloadbtn.Click    += self._download_click

        self.Closing += self._on_closing

    def _on_closing(self, sender, e):
        if not self._did_install_macro:
            return
        install_py = os.path.join(self.macroFileFolder, "SCR_Install.py")
        subprocess.Popen(
            'cmd /c timeout /t 3 /nobreak > nul & del /f /q "{}"'.format(install_py),
            shell=True
        )

    # ── helpers ────────────────────────────────────────────────────────────

    def _ui(self, func):
        self.Dispatcher.BeginInvoke(DispatcherPriority.Normal, Action(func))

    def _parse_repo(self, url):
        url = url.strip().rstrip("/")
        m = re.search(r'github\.com/([^/]+)/([^/]+?)(?:\.git)?$', url)
        return (m.group(1), m.group(2)) if m else (None, None)

    def _api_get(self, url):
        req = HttpWebRequest.Create(url)
        req.Method    = "GET"
        req.UserAgent = "SCR-Install/1.0"
        req.Accept    = "application/vnd.github.v3+json"
        req.Timeout   = 20000
        resp = req.GetResponse()
        with StreamReader(resp.GetResponseStream()) as rdr:
            return json.loads(rdr.ReadToEnd())

    def _git_blob_sha(self, file_path):
        try:
            raw = File.ReadAllBytes(file_path)
            normalized = Encoding.UTF8.GetString(raw).replace("\r\n", "\n").replace("\r", "\n")
            content    = Encoding.UTF8.GetBytes(normalized)
            prefix     = Encoding.ASCII.GetBytes("blob " + str(len(content)) + "\0")
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
            # helpers and SCR_MacroSync are mandatory whenever any regular macro is ticked
            any_dl = any(
                bool(r["download"])
                for r in self._dt.Rows
                if r["path"].split("/")[-1].lower() not in _HELPER_FILENAMES
                and r["path"].split("/")[-1].lower() != _SYNC_FILENAME
                and not r["path"].lower().endswith(".htm")
            )
            for r in self._dt.Rows:
                fname = r["path"].split("/")[-1].lower()
                if fname in _HELPER_FILENAMES or fname == _SYNC_FILENAME:
                    r["download"] = any_dl
        finally:
            self._propagating = False

    def _populate_grid(self, tree_items, local_root):
        dt         = self._make_table()
        local_root = local_root.replace("/", "\\")

        for item in tree_items:
            if item.get("type") != "blob" or not item["path"].endswith(".py"):
                continue
            path       = item["path"]
            remote_sha = item.get("sha", "")
            local_path = os.path.join(local_root, path.replace("/", "\\"))
            exists     = os.path.isfile(local_path)
            local_sha  = self._git_blob_sha(local_path) if exists else None

            if not exists:
                local_ver = "-"
                status    = "New"
                do_dl     = True
            elif local_sha == remote_sha:
                local_ver = self._extract_version(local_path)
                status    = "Up to date"
                do_dl     = False
            else:
                local_ver = self._extract_version(local_path)
                status    = "Update available"
                do_dl     = True

            if path.split("/")[-1].lower() in _HELPER_FILENAMES or path.split("/")[-1].lower() == _SYNC_FILENAME:
                local_ver = "mandatory on macro install"

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
            if not exists:
                status = "New"
                do_dl  = True
            elif local_sha == remote_sha:
                status = "Up to date"
                do_dl  = False
            else:
                status = "Update available"
                do_dl  = True
            row = dt.NewRow()
            folder_prefix = folder + "/" if folder else ""
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

        self._dt = dt
        self.datagrid.DataContext = dt
        self._resize_window()

    def _resize_window(self):
        self.datagrid.UpdateLayout()
        total_w = sum(col.ActualWidth for col in self.datagrid.Columns) + 42
        self.Width = max(total_w, self.MinWidth)

    def _rebind_grid(self):
        self.datagrid.DataContext = None
        self.datagrid.DataContext = self._dt

    # ── event handlers ─────────────────────────────────────────────────────

    def _fetch_click(self, sender, e):
        self.error.Content   = ""
        self.success.Content = "Fetching file list from GitHub…"
        owner, repo = self._parse_repo(self._repo_url)
        if not owner:
            self.error.Content   = "Invalid GitHub URL."
            self.success.Content = ""
            return

        self.fetchbtn.IsEnabled = False
        try:
            data      = self._api_get(
                "https://api.github.com/repos/{}/{}/git/trees/HEAD?recursive=1".format(owner, repo)
            )
            tree      = data.get("tree", [])
            truncated = data.get("truncated", False)

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
                msg += "  ⚠ result was truncated by GitHub"
            self.success.Content = msg
        except Exception as ex:
            self.error.Content   = "GitHub error: " + str(ex)
            self.success.Content = ""
        finally:
            self.fetchbtn.IsEnabled = True

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

    def _download_click(self, sender, e):
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

        jobs = []
        for row in self._dt.Rows:
            if not bool(row["download"]):
                continue
            py_path = row["path"]
            folder  = "/".join(py_path.split("/")[:-1])
            if py_path.lower().endswith(".htm"):
                folder_prefix = folder + "/" if folder else ""
                siblings = [item["path"] for item in self._repo_tree
                            if item.get("type") == "blob"
                            and (item["path"] == py_path or item["path"].startswith(folder_prefix))
                            and not item["path"].lower().endswith(".dict")]
            else:
                siblings = [item["path"] for item in self._repo_tree
                            if item.get("type") == "blob"
                            and "/".join(item["path"].split("/")[:-1]) == folder
                            and not item["path"].lower().endswith(".dict")]
            jobs.append({"row": row, "py_path": py_path, "siblings": siblings})

        if not jobs:
            self.error.Content = "No files selected."
            return

        total_files = sum(len(j["siblings"]) for j in jobs)
        self.fetchbtn.IsEnabled       = False
        self.downloadbtn.IsEnabled    = False
        self.selectallbtn.IsEnabled   = False
        self.deselectallbtn.IsEnabled = False
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
        errors       = []
        has_macro    = any(j["py_path"].lower().endswith(".py") for j in jobs)
        prefix = self._repo_subfolder.strip("/")
        repo_path_prefix = prefix + "/" if prefix else ""

        try:
            for job in jobs:
                row        = job["row"]
                py_path    = job["py_path"]
                siblings   = job["siblings"]
                macro_name = py_path.split("/")[-1]

                n = done + 1
                self._ui(lambda n=n, name=macro_name:
                    setattr(self.success, "Content",
                            "Installing {}/{}: {}".format(n, len(jobs), name)))

                row_ok = True
                for path in siblings:
                    raw_url  = "https://raw.githubusercontent.com/{}/{}/HEAD/{}".format(owner, repo, repo_path_prefix + path)
                    loc_path = os.path.join(local_root, path.replace("/", "\\"))
                    loc_dir  = os.path.dirname(loc_path)
                    try:
                        if not os.path.isdir(loc_dir):
                            os.makedirs(loc_dir)

                        req = HttpWebRequest.Create(raw_url)
                        req.Method    = "GET"
                        req.UserAgent = "SCR-Install/1.0"
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
                done += 1

        except Exception as ex:
            errors.append("Unexpected error: " + str(ex))

        err_text  = "\n".join(errors)
        if has_macro:
            ok_text      = "Installation complete – {} macro(s) installed.".format(done)
            restart_text = (
                "⚠  Please close this window and restart TBC to load the new macros.\n"
                "⚠  This installer will be deleted automatically when you close this window."
            )
        else:
            ok_text      = "Download complete."
            restart_text = ""

        def finish(e=err_text, o=ok_text, r=restart_text, m=has_macro):
            self._did_install             = True
            self._did_install_macro       = m
            self.error.Content            = "\n".join(x for x in [r, e] if x)
            self.success.Content          = o
            self.progressbar.Visibility   = Visibility.Collapsed
            self.fetchbtn.IsEnabled       = True
            self.downloadbtn.IsEnabled    = True
            self.selectallbtn.IsEnabled   = True
            self.deselectallbtn.IsEnabled = True
            self._rebind_grid()
            if not m:
                htm = os.path.join(self._local_root, "MacroHelp", "MacroHelp.htm")
                if os.path.isfile(htm):
                    webbrowser.open(htm)

        self._ui(finish)

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

from System.Collections.Generic import List, IEnumerable
exec(open(r"C:\ProgramData\Trimble\MacroCommands3\SCR Macros\SCR_Imports.py").read())

_OPTIONS = {
    "sourcefilename": "",
    "targetfilename": "",
}

# ---------------------------------------------------------------------------
# TBC ribbon-customization XML (.NET SoapFormatter multi-ref XML) merge logic.
# Every object is a flat top-level child of SOAP-ENV:Body with a unique
# id="ref-N"; relationships are via href="#ref-N". To graft a custom tab from
# one export into another, we take the BFS closure of the tab's object graph,
# strip legacy markers that break TBC's import validator, renumber ids to
# avoid collisions with the target file, and splice the result into the
# target's Body and RibbonTabCollection.
# ---------------------------------------------------------------------------

TAG_RE = re.compile(r'<(/?)([A-Za-z0-9_:\-]+)((?:\s+[A-Za-z0-9_:\-]+="[^"]*")*)\s*(/?)>')
ID_ATTR_RE = re.compile(r'\bid="(ref-\d+)"')

def build_id_spans(text):
    stack = []
    id_spans = {}
    id_ancestor = {}
    for m in TAG_RE.finditer(text):
        closing = m.group(1) == '/'
        name = m.group(2)
        attrs = m.group(3)
        selfclose = m.group(4) == '/'
        idm = ID_ATTR_RE.search(attrs)
        this_id = idm.group(1) if idm else None
        if selfclose:
            if this_id:
                nearest = None
                for frame in reversed(stack):
                    if frame['id']:
                        nearest = frame['id']; break
                id_spans[this_id] = (m.start(), m.end())
                id_ancestor[this_id] = nearest
        elif not closing:
            stack.append({'name': name, 'tag_start': m.start(), 'id': this_id})
        else:
            frame = stack.pop()
            if frame['name'] != name:
                raise ValueError("Mismatched tag: expected </" + frame['name'] + "> got </" + name + "> at " + str(m.start()))
            if frame['id']:
                nearest = None
                for f2 in reversed(stack):
                    if f2['id']:
                        nearest = f2['id']; break
                id_spans[frame['id']] = (frame['tag_start'], m.end())
                id_ancestor[frame['id']] = nearest
    if stack:
        raise ValueError("Unclosed tags remain: " + str([f['name'] for f in stack]))
    return id_spans, id_ancestor


# legacy markers present on older/stale exports (absent from what TBC's current
# UI produces for genuinely custom items) which cause TBC's Import feature to
# fail with a generic "Unable to import the Ribbon data" dialog with no further
# detail, even though the resulting XML is perfectly well-formed.
STOCKITEM_RE = re.compile(r'[ \t]*<IsStockItem>true</IsStockItem>\r?\n')
MERGEKEYS_RE = re.compile(r'[ \t]*<NonAutogenerateMergeGroupKeys\b.*?</NonAutogenerateMergeGroupKeys>\r?\n', re.S)
INSTANCEPROPS_RE = re.compile(r'[ \t]*<InstanceProps href="#ref-\d+"/>\r?\n')
CUSTOMIZEDCAPTION_RE = re.compile(r'[ \t]*<CustomizedCaption href="#ref-\d+"/>\r?\n')
UNDERLYINGTOOLTYPEID_RE = re.compile(r'[ \t]*<UnderlyingToolTypeID href="#ref-\d+"/>\r?\n')
UNDERLYINGOWNEROCC_RE = re.compile(r'[ \t]*<UnderlyingToolOwnerOccurrence>\d+</UnderlyingToolOwnerOccurrence>\r?\n')
UNDERLYINGOWNERTYPE_RE = re.compile(r'[ \t]*<UnderlyingToolOwnerType>\d+</UnderlyingToolOwnerType>\r?\n')

def clean_block(text):
    text = STOCKITEM_RE.sub('', text)
    text = MERGEKEYS_RE.sub('', text)
    text = INSTANCEPROPS_RE.sub('', text)
    text = CUSTOMIZEDCAPTION_RE.sub('', text)
    text = UNDERLYINGTOOLTYPEID_RE.sub('', text)
    text = UNDERLYINGOWNEROCC_RE.sub('', text)
    text = UNDERLYINGOWNERTYPE_RE.sub('', text)
    return text


def encode_index(i):
    # .NET XML element name digit-escaping used by RibbonTabCollection's
    # indexed entries, e.g. index 35 -> "_x0033_5"
    s = str(i)
    return "_x00" + format(ord(s[0]), '02x') + "_" + s[1:]


def set_caption_text(fragment, new_text):
    m = re.search(r'<Caption[^>]*>[^<]*</Caption>', fragment)
    if not m:
        return fragment
    open_tag = re.match(r'(<Caption[^>]*>)', m.group(0)).group(1)
    return fragment[:m.start()] + open_tag + new_text + '</Caption>' + fragment[m.end():]


def find_tab_captions(text):
    captions = []
    for m in re.finditer(r'<a1:RibbonTab id="ref-\d+"[^>]*>(.*?)</a1:RibbonTab>', text, re.S):
        cm = re.search(r'<Caption[^>]*>([^<]*)</Caption>', m.group(1))
        if cm:
            captions.append(cm.group(1))
    return captions


def merge_ribbon_tabs(srcpath, dstpath, tab_captions):
    with io.open(srcpath, 'r', encoding='utf-8') as f:
        src = f.read()
    with io.open(dstpath, 'r', encoding='utf-8') as f:
        dst = f.read()

    id_spans, id_ancestor = build_id_spans(src)

    tab_ids = {}
    for m in re.finditer(r'<a1:RibbonTab id="(ref-\d+)"[^>]*>', src):
        rid = m.group(1)
        start, end = id_spans[rid]
        block = src[start:end]
        cm = re.search(r'<Caption[^>]*>([^<]*)</Caption>', block)
        if cm and cm.group(1) in tab_captions:
            tab_ids[cm.group(1)] = rid
    for cap in tab_captions:
        if cap not in tab_ids:
            raise Exception("Could not find RibbonTab with caption '" + cap + "' in the source file.")

    # collision handling: if the target already has a tab with this caption,
    # rename the merged copy with an increment and merge it anyway
    existing_captions = set()
    for m in re.finditer(r'<a1:RibbonTab id="ref-\d+"[^>]*>(.*?)</a1:RibbonTab>', dst, re.S):
        cm = re.search(r'<Caption[^>]*>([^<]*)</Caption>', m.group(1))
        if cm:
            existing_captions.add(cm.group(1))

    final_captions = {}
    for cap in tab_captions:
        candidate = cap
        n = 2
        while candidate in existing_captions:
            candidate = cap + " " + str(n)
            n += 1
        final_captions[cap] = candidate
        existing_captions.add(candidate)

    # shared "empty Key" singleton: many groups' DialogBoxLauncherKey (and
    # QuickAccessToolbar.Key) point to one canonical empty <Key id="ref-N"/>
    # object. Cloning a duplicate of this when grafting breaks import, so we
    # detect it in both files and rewire references to the target's instance.
    EMPTYKEY_RE = re.compile(r'<Key id="(ref-\d+)"(?:/>|></Key>)')
    src_emptykey_m = EMPTYKEY_RE.search(src)
    dst_emptykey_m = EMPTYKEY_RE.search(dst)
    external_resolve = {}
    if src_emptykey_m and dst_emptykey_m:
        external_resolve[src_emptykey_m.group(1)] = dst_emptykey_m.group(1)

    # BFS closure over all requested tabs together (shared visited set/queue)
    visited = set(tab_ids.values())
    queue = list(tab_ids.values())
    href_re = re.compile(r'href="#(ref-\d+)"')
    while queue:
        cur = queue.pop()
        start, end = id_spans[cur]
        block = clean_block(src[start:end])
        for hm in href_re.finditer(block):
            target = hm.group(1)
            if target in external_resolve:
                continue
            if target not in visited:
                visited.add(target)
                queue.append(target)

    emit_ids = [rid for rid in visited if (id_ancestor.get(rid) is None or id_ancestor.get(rid) not in visited)]
    emit_ids.sort(key=lambda rid: id_spans[rid][0])

    fragments = [clean_block(src[id_spans[rid][0]:id_spans[rid][1]]) for rid in emit_ids]

    # apply the increment-renamed caption, only to each tab's own top-level fragment
    for cap in tab_captions:
        if final_captions[cap] != cap:
            idx = emit_ids.index(tab_ids[cap])
            fragments[idx] = set_caption_text(fragments[idx], final_captions[cap])

    blob = "\n\t\t".join(fragments)

    external_nums = set(int(rid[4:]) for rid in external_resolve)
    all_nums = set()
    for m in re.finditer(r'id="ref-(\d+)"', blob):
        all_nums.add(int(m.group(1)))
    for m in re.finditer(r'href="#ref-(\d+)"', blob):
        all_nums.add(int(m.group(1)))
    all_nums -= external_nums

    dst_max = max(int(n) for n in re.findall(r'ref-(\d+)', dst))

    remap = {}
    next_new = dst_max + 1
    for old_n in sorted(all_nums):
        remap[old_n] = next_new
        next_new += 1

    def sub_id(m):
        n = int(m.group(1))
        return 'id="ref-' + str(remap[n]) + '"' if n in remap else m.group(0)

    def sub_href(m):
        n = int(m.group(1))
        return 'href="#ref-' + str(remap[n]) + '"' if n in remap else m.group(0)

    blob = re.sub(r'id="ref-(\d+)"', sub_id, blob)
    blob = re.sub(r'href="#ref-(\d+)"', sub_href, blob)
    for src_id, dst_id in external_resolve.items():
        blob = blob.replace('href="#' + src_id + '"', 'href="#' + dst_id + '"')

    new_tab_ids = {}
    for cap in tab_captions:
        new_tab_ids[cap] = "ref-" + str(remap[int(tab_ids[cap][4:])])

    blob_ids = re.findall(r'id="(ref-\d+)"', blob)
    if len(blob_ids) != len(set(blob_ids)):
        raise Exception("Internal error: duplicate ids produced while merging.")
    dst_id_set = set(re.findall(r'id="(ref-\d+)"', dst))
    collisions = set(blob_ids) & dst_id_set
    if collisions:
        raise Exception("Internal error: id collision with target file: " + str(collisions))

    # locate target's Ribbon Tabs collection, add ALL new tabs
    m = re.search(r'<a1:Ribbon id="ref-3".*?<Tabs href="#(ref-\d+)"', dst, re.S)
    if not m:
        raise Exception("Could not locate the Ribbon Tabs collection in the target file.")
    tabs_ref = m.group(1)
    m2 = re.search(r'<a1:(\w+) id="' + tabs_ref + r'"[^>]*>(.*?)</a1:\1>', dst, re.S)
    tag_name = m2.group(1)
    coll_start, coll_end = m2.start(), m2.end()
    coll_text = m2.group(0)
    old_count = int(re.search(r'<Count>(\d+)</Count>', coll_text).group(1))

    new_entries = []
    count = old_count
    for cap in tab_captions:
        idx_name = encode_index(count)
        new_entries.append("\t\t\t<" + idx_name + ' href="#' + new_tab_ids[cap] + '"/>\n\t\t')
        count += 1

    new_coll_text = coll_text.replace("<Count>" + str(old_count) + "</Count>", "<Count>" + str(count) + "</Count>", 1)
    close_tag = "</a1:" + tag_name + ">"
    insert_at = new_coll_text.rindex(close_tag)
    new_coll_text = new_coll_text[:insert_at] + "".join(new_entries) + new_coll_text[insert_at:]

    dst = dst[:coll_start] + new_coll_text + dst[coll_end:]

    body_close = "\t</SOAP-ENV:Body>"
    idx = dst.rindex(body_close)
    insertion = "\t\t" + blob + "\n"
    dst = dst[:idx] + insertion + dst[idx:]

    tag = "+".join(final_captions[c] for c in tab_captions)
    base, ext = os.path.splitext(dstpath)
    outpath = base + " + " + tag + ext

    with io.open(outpath, 'w', encoding='utf-8') as f:
        f.write(dst)

    return outpath


def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_ImportRibbonTab"
    cmdData.CommandName = "SCR_ImportRibbonTab"
    cmdData.Caption = "_SCR_ImportRibbonTab"
    cmdData.UIForm = "SCR_ImportRibbonTab"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "0"

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "Import"
        cmdData.ShortCaption = "Import Ribbon Tab"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3
        cmdData.EnableNoProject = True

        cmdData.Version = 1.01
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""

        cmdData.ToolTipTitle = "Import Ribbon Tab"
        cmdData.ToolTipTextFormatted = "Merge custom ribbon tab(s) from an old ribbon customization export into a newer ribbon customization file"

    except:
        pass

    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass


class SCR_ImportRibbonTab(StackPanel): # this inherits from the WPF StackPanel control
    def __init__(self, currentProject, macroFileFolder):
        with StreamReader (macroFileFolder + r"\SCR_ImportRibbonTab.xaml") as s:
            wpf.LoadComponent (self, s)
        self.currentProject = currentProject
        self.macroFileFolder = macroFileFolder

    def HelpClicked(self, cmd, e):
        webbrowser.open(r"C:\ProgramData\Trimble\MacroCommands3\SCR Macros\MacroHelp\MacroHelp.htm#" + type(self).__name__)

    def OnLoad(self, cmd, buttons, event):
        self.okBtn = buttons[0]
        buttons[2].Content = "Help"
        buttons[2].Visibility = Visibility.Visible
        buttons[2].Click += self.HelpClicked
        self.Caption = cmd.Command.Caption

        try:
            self.SetDefaultOptions()
        except:
            pass

    def SetDefaultOptions(self):
        SCROptions.LoadMacroOptions(self, "SCR_ImportRibbonTab", _OPTIONS, self.currentProject)

    def SaveOptions(self):
        SCROptions.SaveMacroOptions(self, "SCR_ImportRibbonTab", _OPTIONS)

    def CancelClicked(self, cmd, args):
        cmd.CloseUICommand ()

    def browsesource_Click(self, sender, e):
        dialog = OpenFileDialog()
        initdir = self.sourcefilename.Text
        if initdir and os.path.isdir(os.path.dirname(initdir)):
            dialog.InitialDirectory = os.path.dirname(initdir)
        dialog.Filter = ("Ribbon XML|*.xml")

        tt = dialog.ShowDialog()
        if tt == DialogResult.OK:
            self.error.Text = ''
            self.success.Text = ''
            try:
                with io.open(dialog.FileName, 'r', encoding='utf-8') as f:
                    text = f.read()
                captions = find_tab_captions(text)
            except Exception as ex:
                self.error.Text = 'Could not read tabs from that file:\n' + str(ex)
                return

            self.sourcefilename.Text = dialog.FileName
            self.tabslist.Items.Clear()
            for cap in captions:
                self.tabslist.Items.Add(cap)
            if not captions:
                self.error.Text = 'No custom ribbon tabs were found in that file.'

    def browsetarget_Click(self, sender, e):
        dialog = OpenFileDialog()
        initdir = self.targetfilename.Text
        if initdir and os.path.isdir(os.path.dirname(initdir)):
            dialog.InitialDirectory = os.path.dirname(initdir)
        dialog.Filter = ("Ribbon XML|*.xml")

        tt = dialog.ShowDialog()
        if tt == DialogResult.OK:
            self.error.Text = ''
            self.success.Text = ''
            self.targetfilename.Text = dialog.FileName

    def OkClicked(self, cmd, e):
        Keyboard.Focus(self.okBtn)
        self.error.Text = ''
        self.success.Text = ''

        srcpath = self.sourcefilename.Text
        dstpath = self.targetfilename.Text

        if not srcpath or not File.Exists(srcpath):
            self.error.Text = 'Please browse for a valid source ribbon XML file.'
            return
        if not dstpath or not File.Exists(dstpath):
            self.error.Text = 'Please browse for a valid target ribbon XML file.'
            return

        selected = [str(c) for c in self.tabslist.SelectedItems]
        if not selected:
            self.error.Text = 'Please select at least one tab to merge.'
            return

        try:
            ProgressBar.TBC_ProgressBar.Title = "merging ribbon tab(s), this can take a moment for large files"
            outpath = merge_ribbon_tabs(srcpath, dstpath, selected)
            ProgressBar.TBC_ProgressBar.Title = ""
            self.success.Text = "Merged. Wrote:\n" + outpath
            self.SaveOptions()
            try:
                subprocess.Popen(["explorer", os.path.dirname(outpath)])
            except Exception:
                pass
        except Exception as ex:
            ProgressBar.TBC_ProgressBar.Title = ""
            exc_type, exc_obj, exc_tb = sys.exc_info()
            self.error.Text = 'An error occurred:\n' + str(exc_type) + '\n' + str(exc_obj) + '\nLine ' + str(exc_tb.tb_lineno)

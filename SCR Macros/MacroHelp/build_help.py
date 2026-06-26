#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
MacroHelp assembler — per-macro chapter files.

Usage:
    python build_help.py             # assemble MacroHelp.htm from chapters/ + template
    python build_help.py --migrate   # one-time: split existing MacroHelp.htm into per-macro files

─── Chapter file layout ──────────────────────────────────────────────────────
MacroHelp/chapters/
    00_News.htm                 global – always downloaded
    00_Installation.htm         global – always downloaded
    <MacroName>.htm             one file per macro (e.g. SCR_DTMCleanup.htm)

Global files contain the full h1 section HTML.
Per-macro files contain just the h3 tag + content (no h1/h2 wrappers).
The build script generates h1/h2 group headers based on HIERARCHY.

─── Selective install ────────────────────────────────────────────────────────
The installer downloads 00_News.htm + 00_Installation.htm always, plus the
macro chapter files for whatever macros are installed.  build_help.py
assembles MacroHelp.htm from whatever is present in chapters/.

─── Adding a new macro ───────────────────────────────────────────────────────
1. Add it to the correct position in HIERARCHY below.
2. Create chapters/<MacroName>.htm with its h3 + content.
3. Run: python build_help.py
"""

import sys
import os
import glob
import re

SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
CHAPTERS_DIR = os.path.join(SCRIPT_DIR, 'chapters')
TEMPLATE_FILE = os.path.join(SCRIPT_DIR, 'MacroHelp_template.htm')
OUTPUT_FILE  = os.path.join(SCRIPT_DIR, 'MacroHelp.htm')

GENERATED_COMMENT = '<!-- GENERATED FILE – edit chapter files in MacroHelp/chapters/ and run build_help.py -->\n'


# ── Hierarchy ─────────────────────────────────────────────────────────────────
# (h1_id, h1_label, [ (h2_id, h2_label, [macro_ids]) ])
# h1_id / h2_id are used as the HTML id attribute AND as the sidebar href target.

HIERARCHY = [
    ('ExplodeRenameRelayerProperties', 'ExplodeRenameRelayerProperties', [
        ('Explode',           'Explode',              ['SCR_ExplodeAlignlabelToText', 'SCR_ExplodeCADLeader', 'SCR_ExplodeIFC', 'SCR_ExplodeTextToLines']),
        ('Relayer',           'Relayer',              ['SCR_ChangeTextlayerToPolylayer', 'SCR_DeletePopulatedLayers', 'SCR_Relayer', 'SCR_UnhideObjects']),
        ('Rename',            'Rename',               ['SCR_BatchRenameAlignments', 'SCR_ChangeTextToPolyname', 'SCR_RenameLines', 'SCR_SNRLayername', 'SCR_SNRPoints', 'SCR_SNRText']),
        ('Properties',        'Properties',           ['SCR_ChangePropNoSelect', 'SCR_DepthToFC', 'SCR_ElevateSHPLines', 'SCR_ExifGPS', 'SCR_FixMediaLinks', 'SCR_Inspect12dAttributes', 'SCR_SaveRestoreLabelPositions']),
        ('ViewTools',         'View Tools',           ['SCR_LockViews', 'SCR_RotatePlanView', 'SCR_SwitchSnapMode', 'SCR_ViewFilter']),
    ]),
    ('ImExportDTMSubgrade', 'ImExportDTMSubgrade', [
        ('DTM',               'DTM',                  ['SCR_DTMCleanup', 'ExplodeSurfaceToLines', 'SCR_DTMMerge', 'SCR_DTMShot', 'SCR_DTMSlopeColored', 'SCR_DuplicateSurface', 'SCR_ElevateToPlane', 'SCR_IntersectSurfaceAnySurface', 'SCR_ManholeSurfaceRL', 'SCR_MatchSurfaceStyle', 'SCR_MoveAlongVector', 'SCR_PerpDistToDTM', 'SCR_PerpDistToFace', 'SCR_ProjectedSurfaceContour', 'SCR_Rainfall', 'SCR_Remove3DFaces', 'SCR_VolumeContours']),
        ('Export',            'Export',               ['SCR_BatchExportRxl', 'SCR_CreateGeoRefWorldFiles', 'SCR_ExportDEM', 'SCR_ExportPointFC2CSV', 'SCR_ExportStaOffEl']),
        ('Import',            'Import',               ['SCR_CSVtoTable', 'SCR_ImportDEM', 'SCR_ImportOBJ', 'SCR_ImportStaOffEl']),
        ('Subgrade',          'Subgrade',             ['SCR_PavementStep', 'SCR_SubgradeFromFSL', 'SCR_SweepBetweenLines']),
        ('Pointcloud',        'Pointcloud',           ['SCR_AnalyzeCloudAlongLine', 'SCR_BestFitPlane', 'SCR_PointcloudExtractor']),
    ]),
    ('ReportsSheets', 'ReportsSheets', [
        ('Reports',           'Reports',              ['SCR_ABReportElevationDifference', 'SCR_ABReportWithBearing', 'SCR_ABReportWithLines', 'SCR_AreaReport', 'SCR_BearingFlatnessCheck']),
        ('SheetsAndDynaviews','Sheets and Dynaviews', ['SCR_AlignmentPlotFrames', 'SCR_ChangeDynaviewFilters', 'SCR_CopyXSectionsToPlanview', 'SCR_FindPlotframe', 'SCR_LocateDynaviewSource', 'SCR_QuickDynaToSheet', 'SCR_RelocateSheetset']),
        ('Text',              'Text',                 ['SCR_AlignTextToLine', 'SCR_CreateIncrementText', 'SCR_RotateText', 'SCR_ThreshColorText']),
    ]),
    ('LinesPoints', 'LinesPoints', [
        ('IFC',               'IFC',                  ['SCR_PlaceObjects', 'SCR_RotateIFC', 'SCR_SplitCombineIFCMeshes']),
        ('Lines',             'Lines',                ['SCR_3PtsToCircle', 'SCR_AddHogToLine', 'SCR_AutoCorner', 'SCR_BearingPads', 'SCR_Check4DoubleSegments', 'SCR_ChordLines', 'SCR_ConvertLinestringToPointBased', 'SCR_CreateRectangleCLs', 'SCR_EditSideslope', 'SCR_GirderOutline', 'SCR_LineInfoArcs', 'SCR_LineInfoVPI', 'SCR_LineBundleEnvelope', 'SCR_LinePerpVertOffset', 'SCR_MeasureOffsetDistance', 'SCR_MergeLineProfile', 'SCR_MinimumRectangle', 'SCR_MultiOffset', 'SCR_PowerlineSway', 'SCR_SweepShape']),
        ('Points',            'Points',               ['SCR_AutoConnectPoints', 'SCR_CreatePointsOnXlines', 'SCR_P1_XY_P2_Z', 'SCR_PasteCSVPoints', 'SCR_PointBlock2Point', 'SCR_PointFromVectorOffsetStation', 'SCR_PointSurfaceSlopeToCSV', 'SCR_SelectPtsOffLine', 'SCR_SimpleCADPoint']),
    ]),
]

# H1/H2 HTML templates — id doubles as the anchor; no mozToc noise needed.
H1_TEMPLATE = '        <h1 id="{h1_id}" class="header1">{label}</h1>\n'
H2_TEMPLATE = '        <h2 id="{h2_id}" class="header2">{label}</h2>\n'


def _all_macros_set():
    """Return the set of all macro names known to HIERARCHY."""
    return {m for _, _, subs in HIERARCHY for _, _, macros in subs for m in macros}


def _present_macros():
    """Return the set of macro names that have a chapter file in chapters/."""
    present = set()
    for m in _all_macros_set():
        if os.path.isfile(os.path.join(CHAPTERS_DIR, m + '.htm')):
            present.add(m)
    return present


# ── BUILD MODE ─────────────────────────────────────────────────────────────────

def _build_content(present):
    parts = []

    for fname in ('00_News.htm', '00_Installation.htm'):
        p = os.path.join(CHAPTERS_DIR, fname)
        if os.path.isfile(p):
            with open(p, 'r', encoding='utf-8') as f:
                parts.append(f.read())

    for h1_id, h1_label, subcategories in HIERARCHY:
        h1_present = [m for _, _, macros in subcategories for m in macros if m in present]
        if not h1_present:
            continue

        parts.append(H1_TEMPLATE.format(h1_id=h1_id, label=h1_label))

        for h2_id, h2_label, macros in subcategories:
            h2_present = [m for m in macros if m in present]
            if not h2_present:
                continue

            parts.append(H2_TEMPLATE.format(h2_id=h2_id, label=h2_label))

            for macro in macros:
                if macro not in present:
                    continue
                p = os.path.join(CHAPTERS_DIR, macro + '.htm')
                with open(p, 'r', encoding='utf-8') as f:
                    parts.append(f.read())

    return ''.join(parts)


def _build_sidebar(present):
    L = []
    I = '          '   # base indent inside <ul id="mozToc">

    if os.path.isfile(os.path.join(CHAPTERS_DIR, '00_News.htm')):
        L.append(I + '<li><a href="#News">News</a></li>')
        L.append(I + '<br>')
    if os.path.isfile(os.path.join(CHAPTERS_DIR, '00_Installation.htm')):
        L.append(I + '<li><a href="#Installation">Installation and Updating</a></li>')
        L.append(I + '<br>')

    for h1_id, h1_label, subcategories in HIERARCHY:
        h1_macros = [m for _, _, macros in subcategories for m in macros]
        if not any(m in present for m in h1_macros):
            continue

        L.append(I + '<li><a href="#{}">{}</a>'.format(h1_id, h1_label))
        L.append(I + '  <ul class="customListStyle">')

        for h2_id, h2_label, macros in subcategories:
            h2_present = [m for m in macros if m in present]
            if not h2_present:
                continue

            L.append(I + '    <li><a href="#{}">{}</a>'.format(h2_id, h2_label))
            L.append(I + '      <ul class="customListStyle">')

            for macro in macros:
                if macro in present:
                    L.append(I + '        <li><a href="#{}">{}</a></li>'.format(macro, macro))

            L.append(I + '      </ul>')
            L.append(I + '    </li>')

        L.append(I + '  </ul>')
        L.append(I + '</li>')
        L.append(I + '<br>')

    return '\n'.join(L) + '\n'


def build():
    """Assemble MacroHelp.htm from template + available chapter files."""
    if not os.path.isfile(TEMPLATE_FILE):
        sys.exit('ERROR: MacroHelp_template.htm not found. Run --migrate first.')

    present = _present_macros()
    print('Present macro chapters: {}'.format(len(present)))

    content_block = _build_content(present)
    nav_block     = _build_sidebar(present)

    with open(TEMPLATE_FILE, 'r', encoding='utf-8') as f:
        template = f.read()

    result = re.sub(r'[^\S\r\n]*<!-- INSERT-CONTENT -->\r?\n', lambda m: content_block, template)
    result = re.sub(r'[^\S\r\n]*<!-- INSERT-SIDEBAR -->\r?\n', lambda m: nav_block,     result)

    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        f.write(GENERATED_COMMENT)
        f.write(result)

    n_global = sum(1 for f in ('00_News.htm','00_Installation.htm') if os.path.isfile(os.path.join(CHAPTERS_DIR,f)))
    print('Built {} ({} global + {} macro chapters).'.format(OUTPUT_FILE, n_global, len(present)))


# ── MIGRATE MODE ───────────────────────────────────────────────────────────────

def _read_source_lines():
    """Read MacroHelp.htm, stripping any leading GENERATED comment."""
    with open(OUTPUT_FILE, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    if lines and lines[0].startswith('<!-- GENERATED'):
        lines = lines[1:]
    return lines


def _find_h3_start(lines, macro_id, search_lo, search_hi):
    """Return 0-indexed line where macro_id's h3 tag starts, or None."""
    target = 'id="{}"'.format(macro_id)
    for i in range(search_lo, search_hi):
        if target in lines[i]:
            if '<h3' in lines[i]:
                return i
            # anchor is inside h3 on a nearby line – walk backward
            for j in range(i - 1, max(i - 6, search_lo - 1), -1):
                if '<h3' in lines[j]:
                    return j
            return i   # fallback: anchor line itself
    return None


def migrate():
    """
    One-time: split MacroHelp.htm into per-macro chapter files + global files.
    Deletes old group-level chapter files (0N_*.htm where N >= 3).
    """
    os.makedirs(CHAPTERS_DIR, exist_ok=True)
    lines = _read_source_lines()
    print('Read {} lines from {}'.format(len(lines), OUTPUT_FILE))

    # ── Global sections (News + Installation) ────────────────────────────────
    # Content section starts at line 128 (1-indexed) = index 127
    # Installation ends just before ExplodeRenameRelayerProperties at index 478
    GLOBAL_DEFS = [
        ('00_News.htm',         127, 466),   # 0-indexed [start:stop) exclusive
        ('00_Installation.htm', 466, 478),
    ]
    for fname, start, stop in GLOBAL_DEFS:
        content = ''.join(lines[start:stop])
        path = os.path.join(CHAPTERS_DIR, fname)
        with open(path, 'w', encoding='utf-8') as f:
            f.write(content)
        print('Written {}'.format(fname))

    # ── Per-macro sections ────────────────────────────────────────────────────
    # Content area for macro groups: lines[478:2040] (0-indexed)
    CONTENT_LO = 478
    CONTENT_HI = 2040

    all_macros = [m for _, _, subs in HIERARCHY for _, _, macros in subs for m in macros]

    # Build sorted list of (line_index, macro_name)
    macro_positions = []
    not_found = []
    for macro in all_macros:
        idx = _find_h3_start(lines, macro, CONTENT_LO, CONTENT_HI)
        if idx is not None:
            macro_positions.append((idx, macro))
        else:
            not_found.append(macro)

    if not_found:
        print('WARNING: could not locate h3 for: {}'.format(', '.join(not_found)))

    macro_positions.sort(key=lambda x: x[0])

    # Each macro's content = from its h3 start to (but not including) the line
    # that starts the next h3/h2/h1.  We determine end by looking for the next
    # macro's position or for an h2/h1 tag in between.
    def find_section_end(start_idx, next_macro_idx):
        """Return the end index (exclusive) of this macro's content block."""
        limit = next_macro_idx if next_macro_idx is not None else CONTENT_HI
        # Walk forward from start+1; stop if we hit h2 or h1 before next macro
        for i in range(start_idx + 1, limit):
            if re.match(r'\s*<h[12][ >]', lines[i]):
                return i
        return limit

    written = 0
    for pos_idx, (start, macro) in enumerate(macro_positions):
        next_start = macro_positions[pos_idx + 1][0] if pos_idx + 1 < len(macro_positions) else None
        end = find_section_end(start, next_start)

        content = ''.join(lines[start:end])
        path = os.path.join(CHAPTERS_DIR, macro + '.htm')
        with open(path, 'w', encoding='utf-8') as f:
            f.write(content)
        written += 1

    print('Written {} macro chapter files.'.format(written))

    # ── Remove old numbered group-level chapter files ─────────────────────────
    for old in glob.glob(os.path.join(CHAPTERS_DIR, '[0-9][0-9]_*.htm')):
        os.remove(old)
        print('Removed old file: {}'.format(os.path.basename(old)))

    print('\nMigration complete. Run "python build_help.py" to rebuild MacroHelp.htm.')


# ── MAIN ───────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    if '--migrate' in sys.argv:
        migrate()
    else:
        build()

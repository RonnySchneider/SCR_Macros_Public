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
    cmdData.Key = "SCR_UnhideObjects"
    cmdData.CommandName = "SCR_UnhideObjects"
    cmdData.Caption = "_SCR_UnhideObjects"
    #cmdData.UIForm = "SCR_UnhideObjects"      # MUST MATCH NAME FROM CLASS DEFINED BELOW !!!
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "22602"

    try:
        cmdData.DefaultTabKey = "SCR Expld-SNR-Relay-Prop"
        cmdData.DefaultTabGroupKey = "Relayer"
        cmdData.ShortCaption = "Unhide Layers"
        cmdData.DefaultRibbonToolSize = 0 # Default=0, ImageOnly=1, Normal=2, Large=3

        cmdData.Version = 1.04
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""
        
        cmdData.ToolTipTitle = "Unhide Layers of Selected Objects"
        cmdData.ToolTipTextFormatted = "Unhide Layers of Selected Objects"

    except:
        pass
    
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png")
        cmdData.ImageSmall = b
    except:
        pass

# objects whose imitation type cannot be read from the instance (no public/protected ImitationTypeHandle/Guid)
# — mapped by .NET class name to the ViewFilter Manager display name in imicoll
_NET_TO_IMITYPE_NAME = {
    # raw data — container/setup types without ImitationTypeHandle
    "MediaFolder":         "Media Folder",
    "RDTotalStation":      "Total Station",
    "RDTotalStationSetup": "Total Station",
    "RDTraverse":          "Traverse",
    "Baseline":            "Baseline",
    "PPSGVector":          "PP Stop and Go Vector",
    "RDCoordinate":        "Point",
    "RDLevelBag":          "Levelling",
    "RDLevelRun":          "Levelling",
    # photogrammetry (Trimble.Vce.Data.RawData / Trimble.Vce.Flights)
    "RDPhotoStation":               "Photo Station",
    "RDPhotoCamera":                "Photo Station",
    "RDPhotoStationSetup":          "Photo Station",
    "RDImageStation":               "Image Frame",
    "FlightMission":                "Flight Mission",
    "FlightBlock":                  "Flight Block",
    "FlightPlan":                   "Flight Block Plan",
    # referenced images — ImitationTypeGuid is an explicit interface implementation,
    # not reachable via reflection; enable all view categories to be safe
    "RDReferencedImage":            ["Referenced Image (Station View)",
                                     "Referenced Image (Plan View)",
                                     "Referenced Image (3D View)",
                                     "Image Frame"],
    "RDReferencedImageUncompressed":["Referenced Image (Station View)",
                                     "Referenced Image (Plan View)",
                                     "Referenced Image (3D View)",
                                     "Image Frame"],
    "RDPublicMediaFile":            ["Referenced Image (Station View)",
                                     "Referenced Image (Plan View)",
                                     "Referenced Image (3D View)",
                                     "Image Frame"],
    "RDPanorama":                   ["Referenced Image (Station View)",
                                     "Referenced Image (Plan View)",
                                     "Referenced Image (3D View)",
                                     "Image Frame"],
    # CAD — ForeignCad (imported DXF/DWG) and feature-coded entities
    "BlockReference":      "CAD Block",
    "BlockFeature":        "CAD Block",
    "LineFeature":         "CAD Line",
    "PolygonFeature":      "CAD Line",
    "Linestring":          "CAD Line",
    "Line":                "CAD Line",
    "LwPolyLine":          "CAD Line",
    "PolyLine":            "CAD Line",
    "Arc":                 "CAD Line",
    "Circle":              "CAD Line",
    "Spline":              "CAD Line",
    "Ellipse":             "CAD Line",
    "MLine":               "CAD Line",
    "Hatch":               "CAD Line",
    "Text":                "CAD Text",
    "MText":               "CAD Text",
    "TextEntity":          "CAD Text",
    # IFC / BIM
    "BIMEntity":           "Mesh",
}

_prop_cache = {}        # (type_fullname, prop_name) -> property or None
_has_imitype_prop = {} # type_fullname -> bool (False = skip reflection entirely)
_class_imitypes = {}   # short type name -> frozenset of imitype serials from full chain walk

def _find_property(o, prop_name, _bf):
    # GetProperty with NonPublic only searches the declaring type, not base classes.
    # Walk the hierarchy manually to find inherited protected properties.
    t = o.GetType()
    key = (t.FullName, prop_name)
    if key in _prop_cache:
        return _prop_cache[key]
    cur = t
    result = None
    while cur is not None:
        prop = cur.GetProperty(prop_name, _bf | BindingFlags.DeclaredOnly)
        if prop:
            result = prop
            break
        cur = cur.BaseType
    _prop_cache[key] = result
    return result

def _collect_imitype(o, imicoll, imitypeserials, _bf):
    tname = o.GetType().FullName
    if _has_imitype_prop.get(tname) is False:
        return
    found_prop = False
    for prop_name in ("ImitationTypeGuid", "ImitationTypeHandle"):
        try:
            prop = _find_property(o, prop_name, _bf)
            if prop:
                found_prop = True
                val = prop.GetValue(o, None)
                if val:
                    imitypeserials.add(imicoll[val].SerialNumber)
        except:
            pass
    if tname not in _has_imitype_prop:
        _has_imitype_prop[tname] = found_prop

def _run_diagnostic(currentProject):
    selectionSet = GlobalSelection.Items(currentProject)
    imicoll = currentProject.Concordance.Lookup(20)
    _bf = BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public
    activeForm = TrimbleOffice.TheOffice.MainWindow.AppViewManager.ActiveView
    activeViewFilter = currentProject.Concordance.Lookup(activeForm.ViewFilter)

    lines = []
    lines.append("=== All imitation types in imicoll ===")
    try:
        for ite in imicoll:
            try:
                name = ""
                for attr in ("Name", "Description", "TypeName"):
                    try:
                        name = str(getattr(ite, attr))
                        break
                    except:
                        pass
                lines.append("  serial={0}  clrtype={1}  name={2}".format(
                    ite.SerialNumber, ite.GetType().Name, name))
            except Exception as ex:
                lines.append("  ERR: " + str(ex))
    except Exception as ex:
        lines.append("  ERROR iterating imicoll: " + str(ex))

    lines.append("")
    lines.append("=== PointCloudScanImitationType entries (full properties) ===")
    try:
        for ite in imicoll:
            if ite.GetType().Name != "PointCloudScanImitationType":
                continue
            lines.append("  serial={0}  name={1}".format(ite.SerialNumber, ite.Name))
            for p in ite.GetType().GetProperties(_bf):
                try:
                    val = p.GetValue(ite, None)
                    lines.append("    ." + p.Name + " = " + str(val))
                except Exception as ex:
                    lines.append("    ." + p.Name + " ERR: " + str(ex))
    except Exception as ex:
        lines.append("  ERROR: " + str(ex))

    lines.append("")
    lines.append("=== ImitationTypeOverrides in active ViewFilter ===")
    try:
        for ito in activeViewFilter.ImitationTypeOverrides:
            lines.append("  serial={0}  visible={1}".format(ito.ImitationTypeSerial, ito.Visible))
    except Exception as ex:
        lines.append("  ERROR: " + str(ex))

    def _dump_obj(obj, lines, prefix=""):
        t = obj.GetType()
        lines.append(prefix + "type=" + t.FullName)
        for p in t.GetProperties(_bf):
            try:
                val = p.GetValue(obj, None)
                lines.append(prefix + "  ." + p.Name + " = " + str(val))
                if p.Name in ("ImitationTypeHandle", "ImitationTypeGuid", "FilterCollectionGUID"):
                    try:
                        e = imicoll[val]
                        lines.append(prefix + "    -> imicoll serial=" + str(e.SerialNumber))
                    except Exception as ex:
                        lines.append(prefix + "    -> imicoll ERR: " + str(ex))
            except Exception as ex:
                lines.append(prefix + "  ." + p.Name + " ERR: " + str(ex))

    lines.append("")
    lines.append("=== Selected objects + MySite chain + RelatedObjects ===")
    for o in selectionSet:
        _dump_obj(o, lines)
        # dump MySite chain
        try:
            visited = set()
            parent = o.MySite
            depth = 1
            while parent is not None:
                try:
                    sn = parent.SerialNumber
                    if sn in visited:
                        break
                    visited.add(sn)
                    lines.append("  [MySite x{0}]".format(depth))
                    _dump_obj(parent, lines, "  ")
                    parent = parent.MySite
                    depth += 1
                except:
                    break
        except:
            pass
        # dump RelatedObjects, each with their own MySite chain
        try:
            for idx, related in enumerate(o.RelatedObjects):
                try:
                    lines.append("  [RelatedObject #{0}]".format(idx))
                    _dump_obj(related, lines, "  ")
                    visited2 = set()
                    rparent = related.MySite
                    rdepth = 1
                    while rparent is not None:
                        try:
                            rsn = rparent.SerialNumber
                            if rsn in visited2:
                                break
                            visited2.add(rsn)
                            lines.append("    [RelatedObject #{0} MySite x{1}]".format(idx, rdepth))
                            _dump_obj(rparent, lines, "    ")
                            rparent = rparent.MySite
                            rdepth += 1
                        except:
                            break
                except:
                    pass
        except:
            pass

    path = r"C:\temp\SCR_UnhideObjects_diag.txt"
    with open(path, "w") as f:
        f.write("\n".join(lines))
    MessageBox.Show("Diagnostics written to " + path, "SCR_UnhideObjects [DIAG]")

def _collect_selection(currentProject, imicoll, _bf):
    """Return (layerserials list, imitypeserials set) for the current selection."""
    selectionSet = GlobalSelection.Items(currentProject)
    needed = set()
    for v in _NET_TO_IMITYPE_NAME.values():
        if isinstance(v, list):
            needed.update(v)
        else:
            needed.add(v)
    name_to_serial = {}
    ProgressBar.TBC_ProgressBar.Title = "SCR_UnhideObjects"
    ProgressBar.TBC_ProgressBar.SetProgress(0)
    for ite in imicoll:
        try:
            n = ite.Name
            if n in needed and n not in name_to_serial:
                name_to_serial[n] = ite.SerialNumber
                if len(name_to_serial) == len(needed):
                    break
        except:
            pass

    layerserials = set()
    imitypeserials = set()

    ProgressBar.TBC_ProgressBar.Title = "SCR_UnhideObjects"
    ProgressBar.TBC_ProgressBar.SetProgress(0)
    j = 0
    time1 = datetime.now()
    for o in selectionSet:
        j += 1
        if (datetime.now() - time1).seconds > 1:
            if ProgressBar.TBC_ProgressBar.SetProgress(j * 100 // selectionSet.Count):
                break
            time1 = datetime.now()

        try:
            if o.Layer:
                layerserials.add(o.Layer)
        except:
            pass

        type_name = o.GetType().Name
        imitype_entry = _NET_TO_IMITYPE_NAME.get(type_name)
        if imitype_entry:
            # known simple CAD/container type — no chain walking needed
            names = imitype_entry if isinstance(imitype_entry, list) else [imitype_entry]
            for n in names:
                if n in name_to_serial:
                    imitypeserials.add(name_to_serial[n])
            continue

        # unknown type: check class-level cache first
        if type_name in _class_imitypes:
            imitypeserials.update(_class_imitypes[type_name])
            continue

        # first time seeing this class — full chain walk, then cache result
        obj_imitypes = set()
        _collect_imitype(o, imicoll, obj_imitypes, _bf)

        try:
            anchor = o.AnchoringObject
            if anchor:
                _collect_imitype(anchor, imicoll, obj_imitypes, _bf)
                try:
                    if anchor.Layer:
                        layerserials.add(anchor.Layer)
                except:
                    pass
        except:
            pass

        try:
            visited = set()
            parent = o.MySite
            while parent is not None:
                try:
                    oid = id(parent)
                    if oid in visited:
                        break
                    visited.add(oid)
                    _collect_imitype(parent, imicoll, obj_imitypes, _bf)
                    parent = parent.MySite
                except:
                    break
        except:
            pass

        try:
            for related in o.RelatedObjects:
                try:
                    _collect_imitype(related, imicoll, obj_imitypes, _bf)
                    visited = set()
                    parent = related.MySite
                    while parent is not None:
                        try:
                            oid = id(parent)
                            if oid in visited:
                                break
                            visited.add(oid)
                            _collect_imitype(parent, imicoll, obj_imitypes, _bf)
                            parent = parent.MySite
                        except:
                            break
                except:
                    pass
        except:
            pass

        try:
            for scan in o.Scans:
                try:
                    if scan.GetType().Name == "PointCloudScan":
                        _collect_imitype(scan, imicoll, obj_imitypes, _bf)
                except:
                    pass
        except:
            pass

        _class_imitypes[type_name] = frozenset(obj_imitypes)
        imitypeserials.update(obj_imitypes)

    ProgressBar.TBC_ProgressBar.Title = ""
    return layerserials, imitypeserials


def Execute(cmd, currentProject, macroFileFolder, parameters):

    if Control.ModifierKeys == (Keys.Control | Keys.Shift):
        _run_diagnostic(currentProject)
        return

    isolate_mode = (Control.ModifierKeys == Keys.Shift)

    wv = currentProject[Project.FixedSerial.WorldView]
    wv.PauseGraphicsCache(True)

    UIEvents.RaiseBeforeDataProcessing(cmd, UIEventArgs())
    currentProject.TransactionManager.AddBeginMark(CommandGranularity.Command, "SCR_UnhideObjects")

    try:
        with TransactMethodCall(currentProject.TransactionCollector) as failGuard:

            activeForm = TrimbleOffice.TheOffice.MainWindow.AppViewManager.ActiveView
            activeViewFilter = currentProject.Concordance.Lookup(activeForm.ViewFilter)
            imicoll = currentProject.Concordance.Lookup(20)
            _bf = BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public

            layerserials, imitypeserials = _collect_selection(currentProject, imicoll, _bf)

            if isolate_mode:
                lc = currentProject.Concordance.Lookup(Project.FixedSerial.LayerContainer)

                # hide every layer first
                for l in lc:
                    sn = l.SerialNumber
                    if not activeViewFilter.LayerOverrides.Contains(sn):
                        activeViewFilter.LayerOverrides.Add(sn)
                    activeViewFilter.LayerOverrides[sn].Visible = False

                # hide every imitation type
                for ite in imicoll:
                    sn = ite.SerialNumber
                    if not activeViewFilter.ImitationTypeOverrides.Contains(sn):
                        activeViewFilter.ImitationTypeOverrides.Add(sn)
                    activeViewFilter.ImitationTypeOverrides[sn].Visible = False

            # unhide the collected layers (ensure override entry exists first)
            for lsn in layerserials:
                if not activeViewFilter.LayerOverrides.Contains(lsn):
                    activeViewFilter.LayerOverrides.Add(lsn)
                lo = activeViewFilter.LayerOverrides[lsn]
                if not lo.Visible:
                    lo.Visible = True

            # unhide the collected imitation types
            for itsn in imitypeserials:
                if not activeViewFilter.ImitationTypeOverrides.Contains(itsn):
                    activeViewFilter.ImitationTypeOverrides.Add(itsn)
                itoverride = activeViewFilter.ImitationTypeOverrides[itsn]
                if not itoverride.Visible:
                    itoverride.Visible = True

            ProgressBar.TBC_ProgressBar.Title = ""

            failGuard.Commit()
            currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
            UIEvents.RaiseAfterDataProcessing(cmd, UIEventArgs())

    except Exception as err:
        # EndMark MUST be set no matter what
        # otherwise TBC won't work anymore and needs to be restarted
        currentProject.TransactionManager.AddEndMark(CommandGranularity.Command)
        UIEvents.RaiseAfterDataProcessing(cmd, UIEventArgs())

    wv.PauseGraphicsCache(False)
  

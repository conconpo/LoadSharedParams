# -*- coding: utf-8 -*-
"""Create Bauteilliste from loaded shared parameters.
- Shows all Psets with their assigned categories
- For each selected Pset+category, scans elements for every unique value of
  "Typ in IFC exportieren als"
- Creates ONE schedule per unique value
- Schedule name: Pset_WallCommon.RvtWände.<TypInIFCExportierenAls value>
- Each schedule is filtered to show only that value (schedule filter)
- The IFC field "Typ in IFC exportieren als" is added but hidden in every schedule
- Column headers: part after dot, suffix stripped
- Handles special categories (Rooms, Areas, Spaces)
Compatible with Revit 2024, 2025 and 2026.
"""
__title__ = "Create\nSchedule"
__author__ = "conconpo"
__doc__ = "Creates Bauteillisten from loaded shared parameters."

from pyrevit import revit, DB, forms, script
from collections import OrderedDict

doc = revit.doc
app = doc.Application

# -----------------------------------------------------------
#  IFC field display names (German primary, English fallbacks)
# -----------------------------------------------------------
IFC_EXPORT_AS   = ("Typ in IFC exportieren als",
                   ["Type Export to IFC As", "Export to IFC As (Type)", "Export as (Type)"])

IFC_FIELD_NAMES = [IFC_EXPORT_AS]

# Fallback suffix when IFC-Typ value is empty
FALLBACK_IFC_SUFFIX = "OhneTyp"

# -----------------------------------------------------------
#  Helpers
# -----------------------------------------------------------
def col_header(param_name):
    """Pset_WallCommon.FireRating[Type] -> FireRating"""
    name = param_name
    for s in ["[Type]", "[Instance]", "[Typ]", "[Exemplar]"]:
        if name.endswith(s):
            name = name[:-len(s)]
    if "." in name:
        name = name.split(".")[-1]
    return name.strip()

def group_from_param(param_name):
    """Pset_WallCommon.FireRating[Type] -> Pset_WallCommon"""
    name = param_name
    for s in ["[Type]", "[Instance]", "[Typ]", "[Exemplar]"]:
        if name.endswith(s):
            name = name[:-len(s)]
    if "." in name:
        return name.rsplit(".", 1)[0]
    return name.strip()

def unique_name(base_name):
    """Append (2),(3)... if view name already exists."""
    existing = set(
        v.Name for v in DB.FilteredElementCollector(doc)
                           .OfClass(DB.ViewSchedule).ToElements()
    )
    if base_name not in existing:
        return base_name
    n = 2
    while "{} ({})".format(base_name, n) in existing:
        n += 1
    return "{} ({})".format(base_name, n)

def clean_for_name(text):
    """Remove characters Revit does not allow in view names."""
    invalid = r'{}[]:;|?/\<>'
    return "".join(c for c in text if c not in invalid)

def create_schedule_for_cat(cat_obj):
    """Create a ViewSchedule; handles Areas specially."""
    try:
        bic = DB.BuiltInCategory(cat_obj.Id.IntegerValue)
    except Exception:
        bic = None

    if bic == DB.BuiltInCategory.OST_Areas:
        area_schemes = DB.FilteredElementCollector(doc)\
            .OfClass(DB.AreaScheme).ToElements()
        if not area_schemes:
            raise Exception("Kein Flächenschema / No area scheme found")
        return DB.ViewSchedule.CreateAreaSchedule(doc, area_schemes[0].Id)
    return DB.ViewSchedule.CreateSchedule(doc, cat_obj.Id)

def find_schedulable_field(avail_fields, primary_name, fallbacks):
    """Find a schedulable field by display name (case-insensitive)."""
    lookup = {}
    for sf in avail_fields:
        try:
            lookup[sf.GetName(doc).lower()] = sf
        except Exception:
            pass
    for name in [primary_name] + list(fallbacks):
        match = lookup.get(name.lower())
        if match is not None:
            return match
    return None

def _read_param(p):
    """Extract string value from a Parameter object."""
    try:
        st = p.StorageType
        if st == DB.StorageType.String:
            return (p.AsString() or "").strip()
        if st == DB.StorageType.Integer:
            return str(p.AsInteger())
        if st == DB.StorageType.Double:
            return str(p.AsDouble())
        if st == DB.StorageType.ElementId:
            eid = p.AsElementId()
            if eid != DB.ElementId.InvalidElementId:
                e = doc.GetElement(eid)
                if e is not None and hasattr(e, "Name"):
                    return (e.Name or "").strip()
    except Exception:
        pass
    return ""

def get_param_value(element, param_names):
    """
    Read a parameter value from an element (instance then type).
    param_names: list of display names to try.
    Returns string or "".
    """
    for name in param_names:
        p = element.LookupParameter(name)
        if p is not None:
            val = _read_param(p)
            if val != "":
                return val
    type_id = element.GetTypeId()
    if type_id and type_id != DB.ElementId.InvalidElementId:
        elem_type = doc.GetElement(type_id)
        if elem_type is not None:
            for name in param_names:
                p = elem_type.LookupParameter(name)
                if p is not None:
                    val = _read_param(p)
                    if val != "":
                        return val
    return ""

def collect_ifc_export_values(cat_obj):
    """
    Scan all (non-type) elements of cat_obj and return a sorted list of unique
    "Typ in IFC exportieren als" values.
    Falls back to [""] so at least one schedule is always created.
    """
    export_as_names = [IFC_EXPORT_AS[0]] + list(IFC_EXPORT_AS[1])

    values = set()
    try:
        collector = DB.FilteredElementCollector(doc)\
                      .OfCategoryId(cat_obj.Id)\
                      .WhereElementIsNotElementType()
        for elem in collector:
            values.add(get_param_value(elem, export_as_names))
    except Exception:
        pass

    if not values:
        values.add("")
    return sorted(values)

# ===========================================================
#  STEP 1 — Read all shared parameters, group by Pset
# ===========================================================
binding_map = doc.ParameterBindings
iterator    = binding_map.ForwardIterator()
iterator.Reset()

pset_groups = OrderedDict()

while iterator.MoveNext():
    defn    = iterator.Key
    binding = iterator.Current
    name    = defn.Name
    grp     = group_from_param(name)

    if grp not in pset_groups:
        pset_groups[grp] = {"params": [], "cats": OrderedDict()}

    pset_groups[grp]["params"].append((name, defn, binding))

    try:
        for cat in binding.Categories:
            pset_groups[grp]["cats"][cat.Name] = cat
    except Exception:
        pass

if not pset_groups:
    forms.alert(
        "Keine gemeinsam genutzten Parameter gefunden.\n"
        "No shared parameters found.\n\n"
        "Bitte zuerst Parameter laden / Load parameters first.",
        title="Keine Parameter / No Parameters",
        exitscript=True)

# ===========================================================
#  STEP 2 — Selection dialog
# ===========================================================
entry_map = {}

for pset_name, data in pset_groups.items():
    params = data["params"]
    cats   = data["cats"]
    if cats:
        for cat_name, cat_obj in cats.items():
            display = "{}  [{}]  ({} Param.)".format(
                pset_name, cat_name, len(params))
            entry_map[display] = (pset_name, cat_name, cat_obj)
    else:
        display = "{}  [?]  ({} Param.)".format(pset_name, len(params))
        entry_map[display] = (pset_name, None, None)

selected_displays = forms.SelectFromList.show(
    sorted(entry_map.keys()),
    title="Psets auswählen / Select Psets  —  {} verfügbar / available".format(
        len(entry_map)),
    multiselect=True,
    button_name="Bauteillisten erstellen / Create Schedules"
)
if not selected_displays:
    script.exit()

# ===========================================================
#  STEP 3 — Expand selections into (cat, pset, params, pair) tuples
# ===========================================================
sched_queue = []   # list of dicts, one per schedule to create

for display in selected_displays:
    pset_name, cat_name, cat_obj = entry_map[display]
    params = pset_groups[pset_name]["params"]

    if cat_obj is None:
        # No category — still record an error entry
        sched_queue.append({
            "cat_obj":       None,
            "pset_name":     pset_name,
            "params":        params,
            "base_cat_title": pset_name,
            "val_export":    "",
        })
        continue

    clean_cat      = clean_for_name(cat_obj.Name)
    base_cat_title = "{}.Rvt{}".format(pset_name, clean_cat)

    values = collect_ifc_export_values(cat_obj)

    for val_export in values:
        sched_queue.append({
            "cat_obj":        cat_obj,
            "pset_name":      pset_name,
            "params":         params,
            "base_cat_title": base_cat_title,
            "val_export":     val_export,
        })

# ===========================================================
#  TX 1 — Create each schedule and add fields
# ===========================================================
sched_ids = []   # carry data from TX1 to TX2
created   = []

for q in sched_queue:
    cat_obj        = q["cat_obj"]
    params         = q["params"]
    base_cat_title = q["base_cat_title"]
    val_export     = q["val_export"]
    ok_cols        = []
    err_cols       = []

    # Build schedule name: base.SuffixFromIFCExportAsValue
    suffix     = clean_for_name(val_export.strip()) if val_export.strip() \
                 else FALLBACK_IFC_SUFFIX
    base_title = "{}.{}".format(base_cat_title, suffix)
    title      = unique_name(base_title)

    if cat_obj is None:
        err_cols.append("Keine Kategorie / No category")
        created.append({"title": title, "ok": [], "err": err_cols,
                        "headings": 0, "schedule": None,
                        "ifc_added": [], "val_export": val_export})
        continue

    with revit.Transaction("Erstelle / Create {}".format(title)):
        try:
            schedule = create_schedule_for_cat(cat_obj)
        except Exception as e:
            err_cols.append("Kategorie-Fehler / Category error: {}".format(str(e)))
            created.append({"title": title, "ok": [], "err": err_cols,
                            "headings": 0, "schedule": None,
                            "ifc_added": [], "val_export": val_export})
            continue

        schedule.Name = title

        try:
            schedule.Definition.IsItemized = True
        except Exception:
            pass

        avail_fields = schedule.Definition.GetSchedulableFields()

        field_lookup = {}
        for sf in avail_fields:
            try:
                field_lookup[sf.GetName(doc)] = sf
            except Exception:
                pass

        # Add Pset fields
        for param_name, defn, binding in params:
            header = col_header(param_name)
            sf = field_lookup.get(param_name) or field_lookup.get(header)

            if sf is None:
                for f in avail_fields:
                    try:
                        if f.GetName(doc) in (param_name, header):
                            sf = f
                            break
                    except Exception:
                        pass

            if sf is None:
                err_cols.append(param_name)
                continue

            try:
                schedule.Definition.AddField(sf)
                ok_cols.append(param_name)
            except Exception as e:
                err_cols.append("{} ({})".format(param_name, str(e)))

        # Add IFC fields
        ifc_added = []
        for primary, fallbacks in IFC_FIELD_NAMES:
            sf = find_schedulable_field(avail_fields, primary, fallbacks)
            if sf is None:
                err_cols.append(
                    "IFC-Feld nicht gefunden / IFC field not found: {}".format(primary))
                ifc_added.append((primary, None))
                continue
            try:
                added = schedule.Definition.AddField(sf)
                ifc_added.append((primary, added.FieldId))
            except Exception as e:
                err_cols.append(
                    "IFC-Feld Fehler / IFC field error: {} ({})".format(primary, str(e)))
                ifc_added.append((primary, None))

        sched_ids.append({
            "id":         schedule.Id,
            "title":      title,
            "ok_cols":    ok_cols,
            "err_cols":   err_cols,
            "ifc_added":  ifc_added,
            "val_export": val_export,
        })

# ===========================================================
#  TX 2 — Column headings + hide IFC fields + schedule filters
# ===========================================================
with revit.Transaction("Spaltenköpfe / Headings + IFC Filters"):
    for item in sched_ids:
        fresh = doc.GetElement(item["id"])
        if fresh is None:
            continue

        fresh_def  = fresh.Definition
        ok_cols    = item["ok_cols"]
        err_cols   = item["err_cols"]
        ifc_added  = item["ifc_added"]
        val_export = item["val_export"]

        # Column headings for Pset fields
        headers   = [col_header(n) for n in ok_cols]
        set_count = 0
        for idx in range(min(len(ok_cols), fresh_def.GetFieldCount())):
            try:
                fresh_def.GetField(idx).ColumnHeading = headers[idx]
                set_count += 1
            except Exception:
                pass

        # Hide all 3 IFC fields and build fid map
        ifc_fid_map = {}
        for primary, fid in ifc_added:
            if fid is None:
                continue
            ifc_fid_map[primary] = fid
            try:
                fresh_def.GetField(fid).IsHidden = True
            except Exception as e:
                err_cols.append("IFC hide error: {} ({})".format(primary, str(e)))

        # Add schedule filter:
        #   "Typ in IFC exportieren als"  == val_export
        fid = ifc_fid_map.get(IFC_EXPORT_AS[0])
        if fid is None:
            err_cols.append(
                "Filter-Feld fehlt / Filter field missing: {}".format(IFC_EXPORT_AS[0]))
        else:
            try:
                if val_export == "":
                    sched_filter = DB.ScheduleFilter(
                        fid, DB.ScheduleFilterType.HasNoValue)
                else:
                    sched_filter = DB.ScheduleFilter(
                        fid, DB.ScheduleFilterType.Equal, val_export)
                fresh_def.AddFilter(sched_filter)
            except Exception as e:
                err_cols.append(
                    "Filter-Fehler / Filter error: {} = '{}' ({})".format(
                        IFC_EXPORT_AS[0], val_export, str(e)))

        created.append({
            "title":      item["title"],
            "ok":         ok_cols,
            "err":        err_cols,
            "headings":   set_count,
            "schedule":   fresh,
            "ifc_added":  ifc_added,
            "val_export": val_export,
        })

# ===========================================================
#  STEP 4 — Open last created schedule
# ===========================================================
if created:
    for item in reversed(created):
        if item.get("schedule") is not None:
            try:
                revit.uidoc.ActiveView = item["schedule"]
            except Exception:
                pass
            break

# ===========================================================
#  STEP 5 — Results
# ===========================================================
output = script.get_output()
output.set_title("Bauteillisten erstellt / Created")
output.print_html("<h2>Ergebnis / Results</h2>")

for item in created:
    output.print_html("<h3>{}</h3>".format(item["title"]))
    output.print_html(
        "<p style='color:#555'>"
        "Filter: <b>Typ in IFC exportieren als</b> = '{}'</p>".format(
            item["val_export"] or "(leer/empty)"
        )
    )
    output.print_html(
        "<p style='color:green'>Spalten / Columns: {} | "
        "Köpfe / Headings: {}</p>".format(
            len(item["ok"]), item.get("headings", 0)))
    for c in item["ok"]:
        output.print_html(
            "<p style='color:green'>&#10003; {} &rarr; {}</p>".format(
                c, col_header(c)))

    ifc_added = item.get("ifc_added", [])
    if ifc_added:
        output.print_html(
            "<p style='color:blue'><b>IFC-Felder (versteckt / hidden):</b></p>")
        for primary, fid in ifc_added:
            if fid is not None:
                output.print_html(
                    "<p style='color:blue'>&#10003; {}</p>".format(primary))
            else:
                output.print_html(
                    "<p style='color:orange'>&#8594; {} &mdash; "
                    "nicht gefunden / not found</p>".format(primary))

    if item["err"]:
        output.print_html(
            "<p style='color:orange'>Warnungen / Warnings: {}</p>".format(
                len(item["err"])))
        for c in item["err"]:
            output.print_html(
                "<p style='color:orange'>&#8594; {}</p>".format(c))

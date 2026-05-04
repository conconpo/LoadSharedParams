# -*- coding: utf-8 -*-
"""Create Bauteilliste from loaded shared parameters.
- Shows all Psets with their assigned categories
- Creates one schedule per selected Pset+category combination
- Schedule name: Pset_WallCommon.RvtWände
- Column headers: part after dot, suffix stripped (FireRating not Pset_WallCommon.FireRating[Type])
- Handles special categories (Rooms, Areas, Spaces)
Compatible with Revit 2024, 2025 and 2026.
"""
__title__ = "Create\nBauteilliste"
__author__ = "conconpo"
__doc__ = "Creates Bauteillisten from loaded shared parameters."

from pyrevit import revit, DB, forms, script
from collections import OrderedDict

doc = revit.doc
app = doc.Application

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
    """Pset_WallCommon.FireRating[Type] -> Pset_WallCommon
       Aussenbauteil[Type]              -> Aussenbauteil"""
    name = param_name
    for s in ["[Type]", "[Instance]", "[Typ]", "[Exemplar]"]:
        if name.endswith(s):
            name = name[:-len(s)]
    if "." in name:
        return name.rsplit(".", 1)[0]
    return name.strip()

def unique_name(base_name):
    """Pset_WallCommon.RvtWände -> append (2),(3) if already exists."""
    existing = set()
    for v in DB.FilteredElementCollector(doc)\
              .OfClass(DB.ViewSchedule).ToElements():
        existing.add(v.Name)
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
    """
    Create a ViewSchedule for the given category.
    Handles special cases: Rooms, Areas, Spaces.
    Returns the schedule or raises an exception.
    """
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
    else:
        return DB.ViewSchedule.CreateSchedule(doc, cat_obj.Id)

# ===========================================================
#  STEP 1 — Read all parameters from project, group by Pset
# ===========================================================
binding_map = doc.ParameterBindings
iterator    = binding_map.ForwardIterator()
iterator.Reset()

# {pset_name: {"params": [(name, defn, binding)], "cats": {cat_name: cat_obj}}}
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
#  STEP 2 — Build display list: one entry per Pset per category
#  Format: "Pset_WallCommon  [Wände]  (10 Param.)"
# ===========================================================
entry_map = {}  # display -> (pset_name, cat_name, cat_obj)

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
#  STEP 3 — Create one schedule per selected entry
#           TX 1: create + add fields
#           TX 2: set column headings by index
# ===========================================================
sched_ids = []  # (ElementId, title, ok_cols, err_cols)
created   = []

# --- TX 1: create all schedules and add fields ---
for display in selected_displays:
    pset_name, cat_name, cat_obj = entry_map[display]
    params = pset_groups[pset_name]["params"]

    # Schedule name: Pset_WallCommon.RvtWände
    if cat_obj:
        clean = clean_for_name(cat_obj.Name)
        base_title = "{}.Rvt{}".format(pset_name, clean)
    else:
        base_title = pset_name

    title    = unique_name(base_title)
    ok_cols  = []
    err_cols = []

    if cat_obj is None:
        err_cols.append("Keine Kategorie / No category")
        created.append({"title": title, "ok": [], "err": err_cols,
                        "headings": 0, "schedule": None})
        continue

    with revit.Transaction("Erstelle / Create {}".format(title)):
        try:
            schedule = create_schedule_for_cat(cat_obj)
        except Exception as e:
            err_cols.append("Kategorie-Fehler / Category error: {}".format(str(e)))
            created.append({"title": title, "ok": [], "err": err_cols,
                            "headings": 0, "schedule": None})
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

        sched_ids.append((schedule.Id, title, ok_cols, err_cols))

# --- TX 2: set column headings by index on fresh objects ---
with revit.Transaction("Spaltenköpfe / Set Headings"):
    for sched_id, title, ok_cols, err_cols in sched_ids:
        fresh = doc.GetElement(sched_id)
        if fresh is None:
            continue

        fresh_def   = fresh.Definition
        field_count = fresh_def.GetFieldCount()
        headers     = [col_header(n) for n in ok_cols]
        set_count   = 0

        for idx in range(min(field_count, len(headers))):
            try:
                fresh_def.GetField(idx).ColumnHeading = headers[idx]
                set_count += 1
            except Exception:
                pass

        created.append({
            "title":    title,
            "ok":       ok_cols,
            "err":      err_cols,
            "headings": set_count,
            "schedule": fresh,
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
        "<p style='color:green'>Spalten / Columns: {} | "
        "Köpfe / Headings: {}</p>".format(
            len(item["ok"]), item.get("headings", 0)))
    for c in item["ok"]:
        output.print_html(
            "<p style='color:green'>&#10003; {} → {}</p>".format(
                c, col_header(c)))
    if item["err"]:
        output.print_html(
            "<p style='color:orange'>Nicht hinzugefügt / Not added: {}</p>".format(
                len(item["err"])))
        for c in item["err"]:
            output.print_html(
                "<p style='color:orange'>&#8594; {}</p>".format(c))

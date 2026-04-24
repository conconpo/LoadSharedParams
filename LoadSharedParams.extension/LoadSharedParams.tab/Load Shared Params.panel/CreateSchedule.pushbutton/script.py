# -*- coding: utf-8 -*-
"""Create Bauteilliste from loaded shared parameters.
No category selection needed — category is read from the parameter binding.
Schedule name  = part after underscore: Pset_WallCommon -> WallCommon
Column headers = part after dot, suffix stripped: Pset_WallCommon.FireRating[Type] -> FireRating
Name collision handled: appends (2), (3) etc. if name already exists.
Compatible with Revit 2024, 2025 and 2026.
"""
__title__ = "Create\nBauteilliste"
__author__ = "conconpo"
__doc__ = "Pick a Pset group and create a schedule. Category is taken from the parameter binding."

from pyrevit import revit, DB, forms, script
from collections import OrderedDict

doc = revit.doc
app = doc.Application

# -----------------------------------------------------------
#  Naming helpers
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

def sched_title(group_name):
    """Pset_WallCommon -> WallCommon"""
    if "_" in group_name:
        return group_name.split("_", 1)[1]
    return group_name

def group_from_param(param_name):
    """Pset_WallCommon.FireRating[Type] -> Pset_WallCommon"""
    name = param_name
    for s in ["[Type]", "[Instance]", "[Typ]", "[Exemplar]"]:
        if name.endswith(s):
            name = name[:-len(s)]
    if "." in name:
        return name.rsplit(".", 1)[0]
    return "ungrouped"

def unique_name(base_name):
    """Return base_name if unique, else base_name (2), (3) etc."""
    existing = set()
    collector = DB.FilteredElementCollector(doc)\
        .OfClass(DB.ViewSchedule)\
        .ToElements()
    for v in collector:
        existing.add(v.Name)
    if base_name not in existing:
        return base_name
    n = 2
    while "{} ({})".format(base_name, n) in existing:
        n += 1
    return "{} ({})".format(base_name, n)

# ===========================================================
#  STEP 1 - Read all parameters, group by Pset name
# ===========================================================
binding_map = doc.ParameterBindings
iterator    = binding_map.ForwardIterator()
iterator.Reset()

pset_groups = OrderedDict()  # {group_name: [(param_name, defn, binding)]}

while iterator.MoveNext():
    defn    = iterator.Key
    binding = iterator.Current
    name    = defn.Name
    grp     = group_from_param(name)
    if grp not in pset_groups:
        pset_groups[grp] = []
    pset_groups[grp].append((name, defn, binding))

if not pset_groups:
    forms.alert("No shared parameters found.\nLoad parameters first.",
                title="No Parameters", exitscript=True)

# ===========================================================
#  STEP 2 - User picks which groups to create schedules for
# ===========================================================
disp_map = {}
for grp, params in pset_groups.items():
    # Find which categories this group is bound to
    cats = set()
    for _, _, b in params:
        try:
            for c in b.Categories:
                cats.add(c.Name)
        except Exception:
            pass
    cat_str = ", ".join(sorted(cats)) if cats else "?"
    label = "{} ({} params | {})".format(
        sched_title(grp), len(params), cat_str)
    while label in disp_map:
        label += " "
    disp_map[label] = grp

selected = forms.SelectFromList.show(
    sorted(disp_map.keys()),
    title="Select Pset groups to create schedules for",
    multiselect=True,
    button_name="Create Schedule(s)"
)
if not selected:
    script.exit()

# ===========================================================
#  STEP 3 - For each selected group, determine category
#           from the binding and create the schedule
# ===========================================================
created   = []
sched_ids = []

for label in selected:
    grp_name = disp_map[label]
    params   = pset_groups[grp_name]
    title    = unique_name(sched_title(grp_name))
    ok_cols  = []
    err_cols = []

    # Determine category from the first parameter's binding
    cat_id = None
    for _, _, binding in params:
        try:
            for c in binding.Categories:
                cat_id = c.Id
                break
        except Exception:
            pass
        if cat_id:
            break

    if cat_id is None:
        err_cols.append("Could not determine category for group {}".format(grp_name))
        created.append({"title": title, "ok": [], "err": err_cols, "headings": 0, "schedule": None})
        continue

    # --- TX 1: create schedule and add fields ---
    with revit.Transaction("Create {}".format(title)):
        schedule  = DB.ViewSchedule.CreateSchedule(doc, cat_id)
        schedule.Name = title
        sched_def = schedule.Definition

        try:
            sched_def.IsItemized = True
        except Exception:
            pass

        avail_fields = sched_def.GetSchedulableFields()
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
                sched_def.AddField(sf)
                ok_cols.append(param_name)
            except Exception as e:
                err_cols.append("{} ({})".format(param_name, str(e)))

        sched_ids.append((schedule.Id, title, ok_cols, err_cols))

# --- TX 2: set headings by index on fresh objects ---
with revit.Transaction("Set Schedule Headings"):
    for sched_id, title, ok_cols, err_cols in sched_ids:
        fresh_sched = doc.GetElement(sched_id)
        if fresh_sched is None:
            continue

        fresh_def      = fresh_sched.Definition
        field_count    = fresh_def.GetFieldCount()
        exp_headers    = [col_header(n) for n in ok_cols]
        set_count      = 0

        for idx in range(min(field_count, len(exp_headers))):
            try:
                field = fresh_def.GetField(idx)
                field.ColumnHeading = exp_headers[idx]
                set_count += 1
            except Exception:
                pass

        created.append({
            "title":    title,
            "ok":       ok_cols,
            "err":      err_cols,
            "headings": set_count,
            "schedule": fresh_sched,
        })

# ===========================================================
#  STEP 4 - Open last created schedule
# ===========================================================
if created:
    for item in reversed(created):
        if item["schedule"] is not None:
            try:
                revit.uidoc.ActiveView = item["schedule"]
            except Exception:
                pass
            break

# ===========================================================
#  STEP 5 - Results
# ===========================================================
output = script.get_output()
output.set_title("Create Bauteilliste - Results")
output.print_html("<h2>Results</h2>")

for item in created:
    output.print_html("<h3>{}</h3>".format(item["title"]))
    output.print_html(
        "<p style='color:green'>Columns: {} | Headings set: {}</p>".format(
            len(item["ok"]), item["headings"]))
    for c in item["ok"]:
        output.print_html(
            "<p style='color:green'>&#10003; {} -> {}</p>".format(
                c, col_header(c)))
    if item["err"]:
        output.print_html(
            "<p style='color:orange'>Not added: {}</p>".format(len(item["err"])))
        for c in item["err"]:
            output.print_html(
                "<p style='color:orange'>&#8594; {}</p>".format(c))

# -*- coding: utf-8 -*-
"""Create Bauteilliste from loaded shared parameters.
- Step 1: shows all Psets with their assigned categories
- Creates one schedule per Pset per category
- Schedule name = full Pset name + category e.g. Pset_WallCommon.Wände / Walls
- Column headers = part after dot, suffix stripped
- Bilingual Deutsch / English throughout
- No "ungrouped" section — parameters without dot notation grouped by full name
Compatible with Revit 2024, 2025 and 2026.
"""
__title__ = "Create\nBauteilliste"
__author__ = "conconpo"
__doc__ = "Creates Bauteillisten from loaded shared parameters. One schedule per Pset per category."

from pyrevit import revit, DB, forms, script
from collections import OrderedDict

doc = revit.doc
app = doc.Application

# -----------------------------------------------------------
#  Bilingual category reverse lookup: BuiltInCategory.Id -> label
#  Used to display category names next to each Pset
# -----------------------------------------------------------
CAT_LABELS = {
    DB.BuiltInCategory.OST_Walls:               "Wände / Walls",
    DB.BuiltInCategory.OST_Floors:              "Böden / Floors",
    DB.BuiltInCategory.OST_Roofs:               "Dächer / Roofs",
    DB.BuiltInCategory.OST_Doors:               "Türen / Doors",
    DB.BuiltInCategory.OST_Windows:             "Fenster / Windows",
    DB.BuiltInCategory.OST_Columns:             "Stützen / Columns",
    DB.BuiltInCategory.OST_StructuralColumns:   "Tragende Stützen / Structural Columns",
    DB.BuiltInCategory.OST_StructuralFraming:   "Tragende Träger / Structural Framing",
    DB.BuiltInCategory.OST_Stairs:              "Treppen / Stairs",
    DB.BuiltInCategory.OST_Railings:            "Geländer / Railings",
    DB.BuiltInCategory.OST_Ceilings:            "Decken / Ceilings",
    DB.BuiltInCategory.OST_Rooms:               "Räume / Rooms",
    DB.BuiltInCategory.OST_MEPSpaces:           "HLK-Zonen / Spaces",
    DB.BuiltInCategory.OST_Areas:               "Flächen / Areas",
    DB.BuiltInCategory.OST_Furniture:           "Möbel / Furniture",
    DB.BuiltInCategory.OST_GenericModel:        "Allgemeines Modell / Generic Models",
    DB.BuiltInCategory.OST_Casework:            "Einbauschränke / Casework",
    DB.BuiltInCategory.OST_CurtainWallPanels:   "Vorhangfassaden-Paneele / Curtain Panels",
    DB.BuiltInCategory.OST_MechanicalEquipment: "Mechanische Ausrüstung / Mechanical Equipment",
    DB.BuiltInCategory.OST_PlumbingFixtures:    "Sanitärinstallationen / Plumbing Fixtures",
    DB.BuiltInCategory.OST_ElectricalFixtures:  "Elektrische Einrichtungen / Electrical Fixtures",
    DB.BuiltInCategory.OST_ElectricalEquipment: "Elektrische Ausrüstung / Electrical Equipment",
    DB.BuiltInCategory.OST_LightingFixtures:    "Leuchten / Lighting Fixtures",
    DB.BuiltInCategory.OST_PipeFitting:         "Rohrformstücke / Pipe Fittings",
    DB.BuiltInCategory.OST_PipeAccessory:       "Rohrzubehör / Pipe Accessories",
    DB.BuiltInCategory.OST_PipeCurves:          "Rohre / Pipes",
    DB.BuiltInCategory.OST_DuctCurves:          "Kanäle / Ducts",
    DB.BuiltInCategory.OST_DuctFitting:         "Kanalformstücke / Duct Fittings",
    DB.BuiltInCategory.OST_DuctAccessory:       "Kanalzubehör / Duct Accessories",
    DB.BuiltInCategory.OST_Sprinklers:          "Sprinkler / Sprinklers",
    DB.BuiltInCategory.OST_StructuralFoundation:"Tragende Fundamente / Structural Foundations",
}

def get_cat_label(cat_obj):
    """Return bilingual label for a category object, or its name if not in map."""
    try:
        for bic, label in CAT_LABELS.items():
            bic_id = doc.Settings.Categories.get_Item(bic)
            if bic_id and bic_id.Id == cat_obj.Id:
                return label
    except Exception:
        pass
    return cat_obj.Name

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
    """
    Pset_WallCommon.FireRating[Type] -> Pset_WallCommon
    Aussenbauteil[Type]              -> Aussenbauteil
    Aussenbauteil                    -> Aussenbauteil

    No "ungrouped" — parameters without a dot use their full name as group.
    This means every parameter always belongs to a group.
    """
    name = param_name
    for s in ["[Type]", "[Instance]", "[Typ]", "[Exemplar]"]:
        if name.endswith(s):
            name = name[:-len(s)]
    if "." in name:
        return name.rsplit(".", 1)[0]
    # No dot — the full parameter name IS the group name
    return name.strip()

def unique_name(base_name):
    """Return base_name if no schedule with that name exists, else append (2), (3)..."""
    existing = set()
    for v in DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule).ToElements():
        existing.add(v.Name)
    if base_name not in existing:
        return base_name
    n = 2
    while "{} ({})".format(base_name, n) in existing:
        n += 1
    return "{} ({})".format(base_name, n)

# ===========================================================
#  STEP 1 — Read all parameters, group by Pset name
#           Collect all categories per Pset
# ===========================================================
binding_map = doc.ParameterBindings
iterator    = binding_map.ForwardIterator()
iterator.Reset()

# pset_groups: {pset_name: {"params": [(name, defn, binding)], "cats": {cat_label: cat_obj}}}
pset_groups = OrderedDict()

while iterator.MoveNext():
    defn    = iterator.Key
    binding = iterator.Current
    name    = defn.Name
    grp     = group_from_param(name)

    if grp not in pset_groups:
        pset_groups[grp] = {"params": [], "cats": OrderedDict()}

    pset_groups[grp]["params"].append((name, defn, binding))

    # Collect all categories this parameter is bound to
    try:
        for cat in binding.Categories:
            label = get_cat_label(cat)
            pset_groups[grp]["cats"][label] = cat
    except Exception:
        pass

if not pset_groups:
    forms.alert(
        "Keine gemeinsam genutzten Parameter gefunden.\n"
        "No shared parameters found.\n\n"
        "Bitte zuerst Parameter laden.\n"
        "Load parameters first.",
        title="Keine Parameter / No Parameters",
        exitscript=True
    )

# ===========================================================
#  STEP 2 — Show Psets with their categories
#  User picks which Pset+category combinations to schedule
#
#  Display format:
#  "Pset_WallCommon  [Wände / Walls]  (10 Param.)"
#
#  If a Pset is bound to multiple categories, it appears
#  once per category so user can pick each independently.
# ===========================================================

# Build display entries — one per Pset per category
entry_map = {}  # display_label -> (pset_name, cat_label, cat_obj)

for pset_name, data in pset_groups.items():
    params = data["params"]
    cats   = data["cats"]

    if cats:
        for cat_label, cat_obj in cats.items():
            display = "{}  [{}]  ({} Param.)".format(
                pset_name, cat_label, len(params))
            entry_map[display] = (pset_name, cat_label, cat_obj)
    else:
        # No category detected — show without category label
        display = "{}  [?]  ({} Param.)".format(pset_name, len(params))
        entry_map[display] = (pset_name, None, None)

selected_displays = forms.SelectFromList.show(
    sorted(entry_map.keys()),
    title="Psets für Bauteillisten auswählen / Select Psets for Bauteillisten",
    multiselect=True,
    button_name="Bauteillisten erstellen / Create Schedules"
)

if not selected_displays:
    script.exit()

# ===========================================================
#  STEP 3 — Type or Instance
# ===========================================================
binding_choice = forms.CommandSwitchWindow.show(
    [
        "Typ (pro Familientyp) / Type (per family type)",
        "Exemplar (pro Element) / Instance (per element)",
    ],
    message="Parameter anzeigen als / Show parameters as:"
)
if not binding_choice:
    script.exit()

# ===========================================================
#  STEP 4 — Create one schedule per selected entry
#           Name = full Pset name + "." + category label
# ===========================================================
created   = []
sched_ids = []

for display in selected_displays:
    pset_name, cat_label, cat_obj = entry_map[display]
    params = pset_groups[pset_name]["params"]

    # Schedule name: Pset_WallCommon.Wände / Walls
    if cat_label:
        base_title = "{}.{}".format(pset_name, cat_label)
    else:
        base_title = pset_name

    title  = unique_name(base_title)
    ok_cols  = []
    err_cols = []

    # Use the category from the entry, or fall back to first available
    use_cat_id = cat_obj.Id if cat_obj else None

    if use_cat_id is None:
        err_cols.append("Keine Kategorie / No category for {}".format(pset_name))
        created.append({"title": title, "ok": [], "err": err_cols,
                        "headings": 0, "schedule": None})
        continue

    # --- TX 1: create schedule and add fields ---
    with revit.Transaction("Erstelle / Create {}".format(title)):
        schedule  = DB.ViewSchedule.CreateSchedule(doc, use_cat_id)
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

# --- TX 2: set column headings by index on fresh objects ---
with revit.Transaction("Spaltenköpfe setzen / Set Column Headings"):
    for sched_id, title, ok_cols, err_cols in sched_ids:
        fresh_sched = doc.GetElement(sched_id)
        if fresh_sched is None:
            continue

        fresh_def   = fresh_sched.Definition
        field_count = fresh_def.GetFieldCount()
        exp_headers = [col_header(n) for n in ok_cols]
        set_count   = 0

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
#  STEP 5 — Open last created schedule
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
#  STEP 6 — Results
# ===========================================================
output = script.get_output()
output.set_title("Bauteillisten erstellt / Bauteillisten Created")
output.print_html("<h2>Ergebnis / Results</h2>")

for item in created:
    output.print_html("<h3>{}</h3>".format(item["title"]))
    output.print_html(
        "<p style='color:green'>Spalten / Columns: {} | "
        "Köpfe gesetzt / Headings set: {}</p>".format(
            len(item["ok"]), item["headings"]))
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

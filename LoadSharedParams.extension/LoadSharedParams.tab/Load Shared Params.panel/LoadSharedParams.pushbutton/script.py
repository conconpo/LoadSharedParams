# -*- coding: utf-8 -*-
"""Load Shared Parameters — universal version.
Works with ANY shared parameter .txt file regardless of naming convention,
language, or structure. No hardcoded Pset names or classification rules.
Compatible with Revit 2024, 2025 and 2026.
"""
__title__ = "Load\nShared Params"
__author__ = "conconpo"
__doc__ = "Universal shared parameter loader. Works with any .txt file in any language."

from pyrevit import revit, DB, forms, script

import clr
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import OpenFileDialog, DialogResult

doc = revit.doc
app = doc.Application
revit_version = int(app.VersionNumber)

# -----------------------------------------------------------
#  Revit categories available for parameter binding
# -----------------------------------------------------------
CAT = {
    "Walls":                    DB.BuiltInCategory.OST_Walls,
    "Floors":                   DB.BuiltInCategory.OST_Floors,
    "Roofs":                    DB.BuiltInCategory.OST_Roofs,
    "Doors":                    DB.BuiltInCategory.OST_Doors,
    "Windows":                  DB.BuiltInCategory.OST_Windows,
    "Columns":                  DB.BuiltInCategory.OST_Columns,
    "Structural Columns":       DB.BuiltInCategory.OST_StructuralColumns,
    "Structural Framing":       DB.BuiltInCategory.OST_StructuralFraming,
    "Stairs":                   DB.BuiltInCategory.OST_Stairs,
    "Railings":                 DB.BuiltInCategory.OST_Railings,
    "Ceilings":                 DB.BuiltInCategory.OST_Ceilings,
    "Rooms":                    DB.BuiltInCategory.OST_Rooms,
    "Spaces":                   DB.BuiltInCategory.OST_MEPSpaces,
    "Areas":                    DB.BuiltInCategory.OST_Areas,
    "Furniture":                DB.BuiltInCategory.OST_Furniture,
    "Generic Models":           DB.BuiltInCategory.OST_GenericModel,
    "Casework":                 DB.BuiltInCategory.OST_Casework,
    "Curtain Panels":           DB.BuiltInCategory.OST_CurtainWallPanels,
    "Mechanical Equipment":     DB.BuiltInCategory.OST_MechanicalEquipment,
    "Plumbing Fixtures":        DB.BuiltInCategory.OST_PlumbingFixtures,
    "Electrical Fixtures":      DB.BuiltInCategory.OST_ElectricalFixtures,
    "Electrical Equipment":     DB.BuiltInCategory.OST_ElectricalEquipment,
    "Lighting Fixtures":        DB.BuiltInCategory.OST_LightingFixtures,
    "Pipe Fittings":            DB.BuiltInCategory.OST_PipeFitting,
    "Pipe Accessories":         DB.BuiltInCategory.OST_PipeAccessory,
    "Pipes":                    DB.BuiltInCategory.OST_PipeCurves,
    "Ducts":                    DB.BuiltInCategory.OST_DuctCurves,
    "Duct Fittings":            DB.BuiltInCategory.OST_DuctFitting,
    "Duct Accessories":         DB.BuiltInCategory.OST_DuctAccessory,
    "Sprinklers":               DB.BuiltInCategory.OST_Sprinklers,
    "Structural Foundations":   DB.BuiltInCategory.OST_StructuralFoundation,
}

def get_param_groups():
    """Return parameter group dict for current Revit version."""
    if revit_version <= 2024:
        return {
            "IFC Parameters":   DB.BuiltInParameterGroup.PG_IFC,
            "Identity Data":    DB.BuiltInParameterGroup.PG_IDENTITY_DATA,
            "Data":             DB.BuiltInParameterGroup.PG_DATA,
            "Construction":     DB.BuiltInParameterGroup.PG_CONSTRUCTION,
            "General":          DB.BuiltInParameterGroup.PG_GENERAL,
            "Structural":       DB.BuiltInParameterGroup.PG_STRUCTURAL,
            "Analysis":         DB.BuiltInParameterGroup.PG_ANALYSIS_RESULTS,
            "Fire Protection":  DB.BuiltInParameterGroup.PG_FIRE_PROTECTION,
            "Energy Analysis":  DB.BuiltInParameterGroup.PG_ENERGY_ANALYSIS,
        }
    else:
        candidates = {
            "IFC Parameters":   lambda: DB.GroupTypeId.Ifc,
            "Identity Data":    lambda: DB.GroupTypeId.IdentityData,
            "Data":             lambda: DB.GroupTypeId.Data,
            "Construction":     lambda: DB.GroupTypeId.Construction,
            "General":          lambda: DB.GroupTypeId.General,
            "Structural":       lambda: DB.GroupTypeId.Structural,
            "Analysis":         lambda: DB.GroupTypeId.Analysis,
            "Fire Protection":  lambda: DB.GroupTypeId.FireProtection,
            "Energy Analysis":  lambda: DB.GroupTypeId.EnergyAnalysis,
        }
        groups = {}
        for name, getter in candidates.items():
            try:
                groups[name] = getter()
            except Exception:
                pass
        if not groups:
            try:
                groups["Identity Data"] = DB.ForgeTypeId(
                    "autodesk.parameter.group:identityData-1.0.0")
            except Exception:
                pass
        return groups

# ===========================================================
#  STEP 1 — Pick the shared parameter .txt file
# ===========================================================
dialog = OpenFileDialog()
dialog.Title = "Select Shared Parameter File (.txt)"
dialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
dialog.InitialDirectory = r"C:\\"

result = dialog.ShowDialog()
if result != DialogResult.OK:
    script.exit()

spf_path = dialog.FileName

# ===========================================================
#  STEP 2 — Open file and read all groups exactly as written
# ===========================================================
original_spf = app.SharedParametersFilename
app.SharedParametersFilename = spf_path
spf = app.OpenSharedParameterFile()

if spf is None:
    app.SharedParametersFilename = original_spf
    forms.alert(
        "Could not open the selected file.\n\n"
        "Make sure it is a valid Revit shared parameter .txt file.",
        title="Invalid File",
        exitscript=True
    )

# Read all groups: {group_name: [definitions]}
# Group name is used exactly as written in the file — no classification.
file_groups = {}
for grp in spf.Groups:
    defs = list(grp.Definitions)
    if defs:
        file_groups[grp.Name] = defs

if not file_groups:
    app.SharedParametersFilename = original_spf
    forms.alert(
        "The selected file contains no parameters.",
        title="Empty File",
        exitscript=True
    )

total_groups = len(file_groups)
total_params = sum(len(d) for d in file_groups.values())

# ===========================================================
#  STEP 3 — User picks which groups to load
#
#  Groups are shown exactly as named in the file.
#  The search box at the top of the list lets the user
#  filter instantly — type any keyword in any language.
#  This works for IFC names, German names, custom names,
#  anything.
# ===========================================================

# Build display list: "GroupName (N params)"
group_display_map = {}  # display_string -> group_name
for gname, defs in file_groups.items():
    display = "{} ({} params)".format(gname, len(defs))
    group_display_map[display] = gname

selected_displays = forms.SelectFromList.show(
    sorted(group_display_map.keys()),
    title="Select groups to load  —  {} groups, {} params total".format(
        total_groups, total_params),
    multiselect=True,
    button_name="Load Selected Groups"
)

if not selected_displays:
    app.SharedParametersFilename = original_spf
    script.exit()

# Resolve back to group names
selected_group_names = [group_display_map[d] for d in selected_displays]
selected_param_count = sum(len(file_groups[g]) for g in selected_group_names)

confirmed = forms.alert(
    "You selected:\n"
    "  Groups:     {}\n"
    "  Parameters: {}\n"
    "  File:       {}\n"
    "  Revit:      {}\n\n"
    "Continue?".format(
        len(selected_group_names),
        selected_param_count,
        spf_path,
        revit_version),
    title="Confirm selection",
    yes=True, no=True
)
if not confirmed:
    app.SharedParametersFilename = original_spf
    script.exit()

# ===========================================================
#  STEP 4 — Pick Revit categories to assign parameters to
# ===========================================================
selected_revit_cats = forms.SelectFromList.show(
    sorted(CAT.keys()),
    title="Which Revit categories should receive these parameters?",
    multiselect=True,
    button_name="Apply to These Categories"
)
if not selected_revit_cats:
    app.SharedParametersFilename = original_spf
    script.exit()

# ===========================================================
#  STEP 5 — Instance or Type
# ===========================================================
binding_choice = forms.CommandSwitchWindow.show(
    ["Type (per family type)", "Instance (per element)"],
    message="Add parameters as:"
)
if not binding_choice:
    app.SharedParametersFilename = original_spf
    script.exit()

is_instance = "Instance" in binding_choice

# ===========================================================
#  STEP 6 — Properties panel section
# ===========================================================
PARAM_GROUPS = get_param_groups()

ui_group_name = forms.CommandSwitchWindow.show(
    sorted(PARAM_GROUPS.keys()),
    message="Place under which Properties panel section? (Revit {})".format(
        revit_version)
)
if not ui_group_name:
    app.SharedParametersFilename = original_spf
    script.exit()

param_group = PARAM_GROUPS[ui_group_name]

# ===========================================================
#  STEP 7 — Build Revit category set
# ===========================================================
cat_set = app.Create.NewCategorySet()
for cat_name in selected_revit_cats:
    if cat_name not in CAT:
        continue
    try:
        cat = doc.Settings.Categories.get_Item(CAT[cat_name])
        if cat:
            cat_set.Insert(cat)
    except Exception as e:
        print("WARNING - category {}: {}".format(cat_name, e))

# ===========================================================
#  STEP 8 — Collect existing parameter names (skip duplicates)
# ===========================================================
existing_names = set()
binding_map = doc.ParameterBindings
iterator = binding_map.ForwardIterator()
iterator.Reset()
while iterator.MoveNext():
    existing_names.add(iterator.Key.Name)

# ===========================================================
#  STEP 9 — Load parameters inside a transaction
# ===========================================================
if is_instance:
    binding = app.Create.NewInstanceBinding(cat_set)
else:
    binding = app.Create.NewTypeBinding(cat_set)

ok_list    = []
skip_list  = []
error_list = []

spf = app.OpenSharedParameterFile()

with revit.Transaction("Load Shared Parameters"):
    for grp in spf.Groups:
        if grp.Name not in selected_group_names:
            continue
        for defn in grp.Definitions:
            name = defn.Name
            if name in existing_names:
                skip_list.append("{} [{}]".format(name, grp.Name))
                continue
            try:
                doc.ParameterBindings.Insert(defn, binding, param_group)
                ok_list.append("{} [{}]".format(name, grp.Name))
            except Exception as e:
                error_list.append("{} - {}".format(name, str(e)))

# ===========================================================
#  STEP 10 — Restore original shared parameter file path
# ===========================================================
app.SharedParametersFilename = original_spf

# ===========================================================
#  STEP 11 — Results summary
# ===========================================================
output = script.get_output()
output.set_title("Load Shared Params (Revit {})".format(revit_version))

output.print_html("<h2>Results</h2>")
output.print_html("<p>File: <b>{}</b></p>".format(spf_path))
output.print_html("<p>Revit: <b>{}</b> | Binding: <b>{}</b> | Group: <b>{}</b></p>".format(
    revit_version,
    "Instance" if is_instance else "Type",
    ui_group_name))
output.print_html("<p>Revit categories: <b>{}</b></p>".format(
    ", ".join(selected_revit_cats)))
output.print_html("<p>Parameter groups loaded from: <b>{}</b></p>".format(
    ", ".join(selected_group_names)))

output.print_html("<h3 style='color:green'>Loaded: {}</h3>".format(len(ok_list)))
for n in ok_list:
    output.print_html("<p style='color:green'>&#10003; {}</p>".format(n))

output.print_html("<h3 style='color:orange'>Skipped (already exist): {}</h3>".format(
    len(skip_list)))
for n in skip_list:
    output.print_html("<p style='color:orange'>&#8594; {}</p>".format(n))

if error_list:
    output.print_html("<h3 style='color:red'>Errors: {}</h3>".format(len(error_list)))
    for n in error_list:
        output.print_html("<p style='color:red'>&#10007; {}</p>".format(n))

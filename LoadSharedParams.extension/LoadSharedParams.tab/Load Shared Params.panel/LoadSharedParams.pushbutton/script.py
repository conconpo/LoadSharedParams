# -*- coding: utf-8 -*-
"""Load Shared Parameters from .txt file into current Revit project.
Shows a file picker for the .txt file and a category checklist.
Paste this into a pyRevit button script (.py file).
"""
__title__ = "Load\nShared Params"
__author__ = "your name"
__doc__ = "Pick a shared parameter .txt file and load all its parameters into selected categories."

# -----------------------------------------------------------
#  pyRevit imports
# -----------------------------------------------------------
from pyrevit import revit, DB, forms, script

# -----------------------------------------------------------
#  Standard imports
# -----------------------------------------------------------
import clr
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import OpenFileDialog, DialogResult

# -----------------------------------------------------------
#  Revit API shortcuts
# -----------------------------------------------------------
doc = revit.doc
app = doc.Application

# ===========================================================
#  STEP 1 — Ask user to pick the .txt shared parameter file
# ===========================================================
dialog = OpenFileDialog()
dialog.Title = "Select Shared Parameter File (.txt)"
dialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
dialog.InitialDirectory = r"C:\\"

result = dialog.ShowDialog()

if result != DialogResult.OK:
    # User cancelled — exit silently
    script.exit()

spf_path = dialog.FileName

# ===========================================================
#  STEP 2 — Open the shared parameter file and read groups
# ===========================================================
original_spf = app.SharedParametersFilename
app.SharedParametersFilename = spf_path

spf = app.OpenSharedParameterFile()

if spf is None:
    forms.alert(
        "Could not open the selected file as a Shared Parameter File.\n\n"
        "Make sure it is a valid Revit shared parameter .txt file.",
        title="Invalid File",
        exitscript=True
    )

# Collect all parameter names from all groups for preview
all_param_names = []
for grp in spf.Groups:
    for defn in grp.Definitions:
        all_param_names.append(defn.Name)

if not all_param_names:
    forms.alert(
        "The selected file contains no parameters.",
        title="Empty File",
        exitscript=True
    )

# Show user what was found, ask to continue
confirmed = forms.alert(
    "Found {} parameters in:\n{}\n\nContinue?".format(
        len(all_param_names), spf_path
    ),
    title="Shared Parameters Found",
    yes=True, no=True
)

if not confirmed:
    app.SharedParametersFilename = original_spf
    script.exit()

# ===========================================================
#  STEP 3 — Let user pick which categories to assign to
# ===========================================================

# Build a list of common categories the user can choose from
# Add or remove entries here to suit your office workflow
COMMON_CATEGORIES = {
    "Walls":            DB.BuiltInCategory.OST_Walls,
    "Floors":           DB.BuiltInCategory.OST_Floors,
    "Ceilings":         DB.BuiltInCategory.OST_Ceilings,
    "Roofs":            DB.BuiltInCategory.OST_Roofs,
    "Doors":            DB.BuiltInCategory.OST_Doors,
    "Windows":          DB.BuiltInCategory.OST_Windows,
    "Columns":          DB.BuiltInCategory.OST_Columns,
    "Structural Columns": DB.BuiltInCategory.OST_StructuralColumns,
    "Structural Framing": DB.BuiltInCategory.OST_StructuralFraming,
    "Rooms":            DB.BuiltInCategory.OST_Rooms,
    "Areas":            DB.BuiltInCategory.OST_Areas,
    "Spaces":           DB.BuiltInCategory.OST_MEPSpaces,
    "Stairs":           DB.BuiltInCategory.OST_Stairs,
    "Railings":         DB.BuiltInCategory.OST_Railings,
    "Furniture":        DB.BuiltInCategory.OST_Furniture,
    "Generic Models":   DB.BuiltInCategory.OST_GenericModel,
    "Casework":         DB.BuiltInCategory.OST_Casework,
    "Curtain Panels":   DB.BuiltInCategory.OST_CurtainWallPanels,
    "Mechanical Equipment": DB.BuiltInCategory.OST_MechanicalEquipment,
    "Plumbing Fixtures":    DB.BuiltInCategory.OST_PlumbingFixtures,
    "Electrical Fixtures":  DB.BuiltInCategory.OST_ElectricalFixtures,
}

selected_cat_names = forms.SelectFromList.show(
    sorted(COMMON_CATEGORIES.keys()),
    title="Select Target Categories",
    multiselect=True,
    button_name="Apply to These Categories"
)

if not selected_cat_names:
    app.SharedParametersFilename = original_spf
    script.exit()

# ===========================================================
#  STEP 4 — Ask instance or type
# ===========================================================
param_binding_type = forms.CommandSwitchWindow.show(
    ["Instance (per element)", "Type (per family type)"],
    message="Add parameters as:"
)

if not param_binding_type:
    app.SharedParametersFilename = original_spf
    script.exit()

is_instance = "Instance" in param_binding_type

# ===========================================================
#  STEP 5 — Ask parameter group (UI section in Properties)
# ===========================================================
PARAM_GROUPS = {
    "Identity Data":        DB.BuiltInParameterGroup.PG_IDENTITY_DATA,
    "Data":                 DB.BuiltInParameterGroup.PG_DATA,
    "Analysis Results":     DB.BuiltInParameterGroup.PG_ANALYSIS_RESULTS,
    "Structural":           DB.BuiltInParameterGroup.PG_STRUCTURAL,
    "Construction":         DB.BuiltInParameterGroup.PG_CONSTRUCTION,
    "Energy Analysis":      DB.BuiltInParameterGroup.PG_ENERGY_ANALYSIS,
    "Fire Protection":      DB.BuiltInParameterGroup.PG_FIRE_PROTECTION,
    "General":              DB.BuiltInParameterGroup.PG_GENERAL,
}

selected_group_name = forms.CommandSwitchWindow.show(
    sorted(PARAM_GROUPS.keys()),
    message="Place parameters under which group?"
)

if not selected_group_name:
    app.SharedParametersFilename = original_spf
    script.exit()

param_group = PARAM_GROUPS[selected_group_name]

# ===========================================================
#  STEP 6 — Build category set
# ===========================================================
cat_set = app.Create.NewCategorySet()
for cat_name in selected_cat_names:
    bic = COMMON_CATEGORIES[cat_name]
    try:
        cat = doc.Settings.Categories.get_Item(bic)
        if cat:
            cat_set.Insert(cat)
    except Exception as e:
        print("WARNING: Could not add category {}: {}".format(cat_name, e))

# ===========================================================
#  STEP 7 — Get existing parameters to skip duplicates
# ===========================================================
existing_names = set()
binding_map = doc.ParameterBindings
iterator = binding_map.ForwardIterator()
iterator.Reset()
while iterator.MoveNext():
    existing_names.add(iterator.Key.Name)

# ===========================================================
#  STEP 8 — Load all parameters inside a transaction
# ===========================================================
if is_instance:
    binding = app.Create.NewInstanceBinding(cat_set)
else:
    binding = app.Create.NewTypeBinding(cat_set)

ok_list    = []
skip_list  = []
error_list = []

# Re-open the file (we already set the path above)
spf = app.OpenSharedParameterFile()

with revit.Transaction("Load Shared Parameters"):
    for grp in spf.Groups:
        for defn in grp.Definitions:
            name = defn.Name
            if name in existing_names:
                skip_list.append(name)
                continue
            try:
                doc.ParameterBindings.Insert(defn, binding, param_group)
                ok_list.append(name)
            except Exception as e:
                error_list.append("{} — {}".format(name, str(e)))

# ===========================================================
#  STEP 9 — Restore original shared parameter file
# ===========================================================
app.SharedParametersFilename = original_spf

# ===========================================================
#  STEP 10 — Show results summary
# ===========================================================
output = script.get_output()
output.set_title("Load Shared Parameters — Results")

output.print_html("<h2>Results</h2>")
output.print_html("<p>File: <b>{}</b></p>".format(spf_path))
output.print_html("<p>Categories: <b>{}</b></p>".format(", ".join(selected_cat_names)))
output.print_html("<p>Binding: <b>{}</b> | Group: <b>{}</b></p>".format(
    "Instance" if is_instance else "Type", selected_group_name))

output.print_html("<h3 style='color:green'>Loaded: {}</h3>".format(len(ok_list)))
for n in ok_list:
    output.print_html("<p style='color:green'>&#10003; {}</p>".format(n))

output.print_html("<h3 style='color:orange'>Skipped (already exist): {}</h3>".format(len(skip_list)))
for n in skip_list:
    output.print_html("<p style='color:orange'>&#8594; {}</p>".format(n))

if error_list:
    output.print_html("<h3 style='color:red'>Errors: {}</h3>".format(len(error_list)))
    for n in error_list:
        output.print_html("<p style='color:red'>&#10007; {}</p>".format(n))

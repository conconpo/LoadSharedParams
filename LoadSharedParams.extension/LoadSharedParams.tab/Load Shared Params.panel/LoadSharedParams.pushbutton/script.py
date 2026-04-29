# -*- coding: utf-8 -*-
"""Load Shared Parameters — universal version.
- Reads available categories directly from the current Revit project
- Displays them as Deutsch / English
- All selected groups go to all selected categories
- Always loads into IFC-Parameter section
- Same or individual categories per group
Compatible with Revit 2024, 2025 and 2026.
"""
__title__ = "Load\nShared Params"
__author__ = "conconpo"
__doc__ = "Universal shared parameter loader. Bilingual UI (DE/EN)."

from pyrevit import revit, DB, forms, script

import clr
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import OpenFileDialog, DialogResult

doc = revit.doc
app = doc.Application
revit_version = int(app.VersionNumber)

# Categories are read live from the Revit project (doc.Settings.Categories)
# and displayed using cat.Name — the exact name Revit shows in its UI.
# No translation needed.

def get_ifc_param_group():
    """Return IFC Parameters group for current Revit version."""
    if revit_version <= 2024:
        try:
            return DB.BuiltInParameterGroup.PG_IFC
        except Exception:
            return DB.BuiltInParameterGroup.PG_IDENTITY_DATA
    else:
        try:
            return DB.GroupTypeId.Ifc
        except Exception:
            try:
                return DB.GroupTypeId.IdentityData
            except Exception:
                try:
                    return DB.ForgeTypeId(
                        "autodesk.parameter.group:identityData-1.0.0")
                except Exception:
                    return None

# -----------------------------------------------------------
#  Read ALL bindable categories from the current project
#  and display them with bilingual labels where known,
#  or just the Revit name if unknown.
# -----------------------------------------------------------
def get_project_categories():
    """
    Returns a sorted list of display labels and a dict:
    {display_label: Category object}
    Only includes categories that allow bound parameters.
    """
    cat_map = {}  # display_label -> Category object
    cats = doc.Settings.Categories

    for cat in cats:
        # Only model categories that allow parameters
        if cat.CategoryType != DB.CategoryType.Model:
            continue
        if not cat.AllowsBoundParameters:
            continue

        # Build display label
        bic_name = cat.Id.ToString()
        # BuiltInCategory enum value is negative int — get name
        try:
            bic = DB.BuiltInCategory(cat.Id.IntegerValue)
            bic_str = bic.ToString()  # e.g. "OST_Walls"
        except Exception:
            bic_str = ""

        # Use cat.Name exactly as Revit shows it in the UI
        # No English translation needed — Revit name is authoritative
        label = cat.Name

        # Handle duplicate labels
        base_label = label
        counter = 2
        while label in cat_map:
            label = "{} ({})".format(base_label, counter)
            counter += 1

        cat_map[label] = cat

    return cat_map

# ===========================================================
#  SCHRITT 1 / STEP 1 — Parameterdatei auswählen / Pick file
# ===========================================================
dialog = OpenFileDialog()
dialog.Title = "Parameterdatei auswählen / Select Shared Parameter File (.txt)"
dialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
dialog.InitialDirectory = r"C:\\"

result = dialog.ShowDialog()
if result != DialogResult.OK:
    script.exit()

spf_path = dialog.FileName

# ===========================================================
#  SCHRITT 2 / STEP 2 — Datei öffnen / Open file
# ===========================================================
original_spf = app.SharedParametersFilename
app.SharedParametersFilename = spf_path
spf = app.OpenSharedParameterFile()

if spf is None:
    app.SharedParametersFilename = original_spf
    forms.alert(
        "Datei konnte nicht geöffnet werden.\n"
        "Could not open the selected file.\n\n"
        "Bitte eine gültige Revit-Parameterdatei (.txt) auswählen.\n"
        "Make sure it is a valid Revit shared parameter .txt file.",
        title="Ungültige Datei / Invalid File",
        exitscript=True)

file_groups = {}
for grp in spf.Groups:
    defs = list(grp.Definitions)
    if defs:
        file_groups[grp.Name] = defs

if not file_groups:
    app.SharedParametersFilename = original_spf
    forms.alert(
        "Die Datei enthält keine Parameter.\n"
        "The selected file contains no parameters.",
        title="Leere Datei / Empty File",
        exitscript=True)

total_groups = len(file_groups)
total_params = sum(len(d) for d in file_groups.values())

# ===========================================================
#  SCHRITT 3 / STEP 3 — Gruppen auswählen / Select groups
# ===========================================================
group_display_map = {}
for gname, defs in file_groups.items():
    display = "{} ({} Parameter)".format(gname, len(defs))
    group_display_map[display] = gname

selected_displays = forms.SelectFromList.show(
    sorted(group_display_map.keys()),
    title="Gruppen auswählen / Select groups  —  "
          "{} Gruppen, {} Parameter total".format(total_groups, total_params),
    multiselect=True,
    button_name="Ausgewählte laden / Load Selected"
)
if not selected_displays:
    app.SharedParametersFilename = original_spf
    script.exit()

selected_group_names = [group_display_map[d] for d in selected_displays]
selected_param_count = sum(len(file_groups[g]) for g in selected_group_names)

# ===========================================================
#  SCHRITT 4 / STEP 4 — Kategorie-Zuweisung
#  Read categories from project — universal, no hardcoded list
# ===========================================================
cat_map = get_project_categories()  # {display_label: Category object}

assignment_mode = forms.CommandSwitchWindow.show(
    [
        "Gleiche Kategorien für alle Gruppen / Same categories for all groups",
        "Individuelle Kategorien pro Gruppe / Individual categories per group",
    ],
    message="Kategorie-Zuweisung / Category assignment:"
)
if not assignment_mode:
    app.SharedParametersFilename = original_spf
    script.exit()

group_cat_map = {}  # {group_name: [Category object, ...]}

if "Gleiche" in assignment_mode or "Same" in assignment_mode:
    chosen_labels = forms.SelectFromList.show(
        sorted(cat_map.keys()),
        title="Kategorien für alle Gruppen / Categories for all groups",
        multiselect=True,
        button_name="Diese Kategorien verwenden / Use These Categories"
    )
    if not chosen_labels:
        app.SharedParametersFilename = original_spf
        script.exit()
    for gname in selected_group_names:
        group_cat_map[gname] = [cat_map[l] for l in chosen_labels]

else:
    for gname in selected_group_names:
        param_count = len(file_groups[gname])
        chosen_labels = forms.SelectFromList.show(
            sorted(cat_map.keys()),
            title="Kategorien für / Categories for:  '{}' ({} Param.)".format(
                gname, param_count),
            multiselect=True,
            button_name="Für diese Gruppe / Use for This Group"
        )
        if not chosen_labels:
            group_cat_map[gname] = []
        else:
            group_cat_map[gname] = [cat_map[l] for l in chosen_labels]

# ===========================================================
#  SCHRITT 5 / STEP 5 — Bestätigung / Confirmation
# ===========================================================
summary_lines = []
for gname in selected_group_names:
    cats = group_cat_map.get(gname, [])
    cat_labels = [c.Name for c in cats]
    if cats:
        summary_lines.append("  {} ->\n    {}".format(
            gname, "\n    ".join(cat_labels)))
    else:
        summary_lines.append("  {} -> (übersprungen / skipped)".format(gname))

confirmed = forms.alert(
    "Gruppen / Groups:  {}\n"
    "Parameter:         {}\n"
    "Revit:             {}\n\n"
    "Zuordnungen / Mappings:\n{}\n\n"
    "Fortfahren? / Continue?".format(
        len(selected_group_names),
        selected_param_count,
        revit_version,
        "\n".join(summary_lines)),
    title="Bestätigung / Confirm",
    yes=True, no=True
)
if not confirmed:
    app.SharedParametersFilename = original_spf
    script.exit()

# ===========================================================
#  SCHRITT 6 / STEP 6 — Typ oder Exemplar / Type or Instance
# ===========================================================
binding_choice = forms.CommandSwitchWindow.show(
    [
        "Typ (pro Familientyp) / Type (per family type)",
        "Exemplar (pro Element) / Instance (per element)",
    ],
    message="Parameter hinzufügen als / Add parameters as:"
)
if not binding_choice:
    app.SharedParametersFilename = original_spf
    script.exit()

is_instance = "Instance" in binding_choice or "Exemplar" in binding_choice

# ===========================================================
#  SCHRITT 7 / STEP 7 — IFC-Parameter (fest / fixed)
# ===========================================================
param_group = get_ifc_param_group()

if param_group is None:
    app.SharedParametersFilename = original_spf
    forms.alert(
        "IFC-Parameter-Gruppe konnte nicht aufgelöst werden.\n"
        "Could not resolve IFC Parameters group for Revit {}.".format(revit_version),
        title="Fehler / Error",
        exitscript=True)

# ===========================================================
#  SCHRITT 8 / STEP 8 — Vorhandene Parameter ermitteln
# ===========================================================
existing_names = set()
binding_map = doc.ParameterBindings
iterator = binding_map.ForwardIterator()
iterator.Reset()
while iterator.MoveNext():
    existing_names.add(iterator.Key.Name)

# ===========================================================
#  SCHRITT 9 / STEP 9 — Parameter laden / Load parameters
# ===========================================================
spf = app.OpenSharedParameterFile()
ok_list    = []
skip_list  = []
error_list = []

with revit.Transaction("Gemeinsam genutzte Parameter laden / Load Shared Parameters"):
    for grp in spf.Groups:
        if grp.Name not in selected_group_names:
            continue

        cat_objects = group_cat_map.get(grp.Name, [])
        if not cat_objects:
            continue

        cat_set = app.Create.NewCategorySet()
        for cat in cat_objects:
            try:
                cat_set.Insert(cat)
            except Exception as e:
                print("WARNUNG / WARNING {}: {}".format(cat.Name, e))

        if is_instance:
            binding = app.Create.NewInstanceBinding(cat_set)
        else:
            binding = app.Create.NewTypeBinding(cat_set)

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
#  SCHRITT 10 / STEP 10 — Originalpfad wiederherstellen
# ===========================================================
app.SharedParametersFilename = original_spf

# ===========================================================
#  SCHRITT 11 / STEP 11 — Ergebnis / Results
# ===========================================================
output = script.get_output()
output.set_title("Parameter laden / Load Shared Params (Revit {})".format(revit_version))
output.print_html("<h2>Ergebnis / Results</h2>")
output.print_html("<p>Datei / File: <b>{}</b></p>".format(spf_path))
output.print_html("<p>Revit: <b>{}</b> | Bindung / Binding: <b>{}</b> | "
                  "Gruppe / Group: <b>IFC-Parameter / IFC Parameters</b></p>".format(
    revit_version,
    "Exemplar / Instance" if is_instance else "Typ / Type"))

output.print_html("<h3>Kategorie-Zuordnungen / Category mappings:</h3>")
for gname in selected_group_names:
    cats = group_cat_map.get(gname, [])
    cat_labels = [c.Name for c in cats] if cats else ["(übersprungen / skipped)"]
    output.print_html("<p><b>{}</b> -> {}</p>".format(
        gname, ", ".join(cat_labels)))

output.print_html("<h3 style='color:green'>Geladen / Loaded: {}</h3>".format(
    len(ok_list)))
for n in ok_list:
    output.print_html("<p style='color:green'>&#10003; {}</p>".format(n))

output.print_html(
    "<h3 style='color:orange'>Übersprungen / Skipped: {}</h3>".format(len(skip_list)))
for n in skip_list:
    output.print_html("<p style='color:orange'>&#8594; {}</p>".format(n))

if error_list:
    output.print_html("<h3 style='color:red'>Fehler / Errors: {}</h3>".format(
        len(error_list)))
    for n in error_list:
        output.print_html("<p style='color:red'>&#10007; {}</p>".format(n))

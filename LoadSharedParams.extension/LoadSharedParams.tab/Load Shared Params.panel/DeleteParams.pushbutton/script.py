# -*- coding: utf-8 -*-
"""Delete selected Project Parameters from the current Revit project.
Shows a searchable checklist of all current parameters.
Compatible with Revit 2024, 2025 and 2026.
"""
__title__ = "Delete\nParams"
__author__ = "conconpo"
__doc__ = "Select which project parameters to delete from the current project."

from pyrevit import revit, DB, forms, script

doc = revit.doc
app = doc.Application

# -----------------------------------------------------------
#  STEP 1 - Collect all current project parameters
# -----------------------------------------------------------
all_params = {}  # {display_name: definition}

binding_map = doc.ParameterBindings
iterator = binding_map.ForwardIterator()
iterator.Reset()
while iterator.MoveNext():
    defn = iterator.Key
    binding = iterator.Current
    # Build a display name showing name + binding type
    binding_type = "Type" if isinstance(binding, DB.TypeBinding) else "Instance"
    display = "{} [{}]".format(defn.Name, binding_type)
    all_params[display] = defn

if not all_params:
    forms.alert("No project parameters found in this project.", title="Nothing to delete")
    script.exit()

# -----------------------------------------------------------
#  STEP 2 - Let user pick which ones to delete
# -----------------------------------------------------------
selected_display_names = forms.SelectFromList.show(
    sorted(all_params.keys()),
    title="Select Parameters to DELETE ({} total)".format(len(all_params)),
    multiselect=True,
    button_name="Delete Selected"
)

if not selected_display_names:
    script.exit()

# -----------------------------------------------------------
#  STEP 3 - Confirm before deleting
# -----------------------------------------------------------
confirmed = forms.alert(
    "You are about to DELETE {} parameters.\n\nThis cannot be undone.\n\nContinue?".format(
        len(selected_display_names)),
    title="Confirm Deletion",
    yes=True, no=True,
    warn_icon=True
)
if not confirmed:
    script.exit()

# -----------------------------------------------------------
#  STEP 4 - Delete selected parameters
# -----------------------------------------------------------
ok_list    = []
error_list = []

with revit.Transaction("Delete Project Parameters"):
    for display_name in selected_display_names:
        defn = all_params[display_name]
        try:
            doc.ParameterBindings.Remove(defn)
            ok_list.append(display_name)
        except Exception as e:
            error_list.append("{} - {}".format(display_name, str(e)))

# -----------------------------------------------------------
#  STEP 5 - Results
# -----------------------------------------------------------
output = script.get_output()
output.set_title("Delete Parameters - Results")
output.print_html("<h2>Results</h2>")
output.print_html("<h3 style='color:green'>Deleted: {}</h3>".format(len(ok_list)))
for n in ok_list:
    output.print_html("<p style='color:green'>&#10003; {}</p>".format(n))
if error_list:
    output.print_html("<h3 style='color:red'>Errors: {}</h3>".format(len(error_list)))
    for n in error_list:
        output.print_html("<p style='color:red'>&#10007; {}</p>".format(n))

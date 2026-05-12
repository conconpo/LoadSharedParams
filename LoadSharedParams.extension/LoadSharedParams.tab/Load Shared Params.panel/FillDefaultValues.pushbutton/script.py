# -*- coding: utf-8 -*-
"""Fill empty project/shared parameters with default values.

DEFAULT VALUES PER TYPE:
  YESNO   (bool stored as Integer) -> 0     (False / unchecked)
  TEXT    (String)                 -> "UNSET"
  DOUBLE  (Area, Length, Number,   -> 0.0   (zero, universal across all units)
           Temperature, Volume...)
  INTEGER (non-YESNO)              -> 999

FLOW:
  1. Read ParameterBindings -> group params by Pset/Mset prefix
  2. Show SelectFromList -> user picks which Psets to fill
  3. Show scope selector -> Active View / Selection / Whole Model
  4. Collect elements from categories bound to SELECTED params only
  5. Process instances + their unique types
  6. Fill empty params, report results

Compatible with Revit 2021 - 2026.
"""
__title__ = "Fill\nDefaults"
__author__ = "conconpo"
__doc__ = "Fill empty IFC/Pset parameter values with defaults. Select which Psets to fill first."

from collections import OrderedDict
from pyrevit import revit, DB, forms, script

doc   = revit.doc
uidoc = revit.uidoc

# Spatial categories need a different collector
SPATIAL_BICS = {
    int(DB.BuiltInCategory.OST_Rooms),
    int(DB.BuiltInCategory.OST_MEPSpaces),
    int(DB.BuiltInCategory.OST_Areas),
}

# ---------------------------------------------------------------------------
#  HELPERS -- grouping (same logic as Create Bauteilliste)
# ---------------------------------------------------------------------------
def _group_from_param(param_name):
    """
    Pset_WallCommon.FireRating  -> Pset_WallCommon
    Mset_SpaceStamp.NetArea     -> Mset_SpaceStamp
    Aussenbauteil               -> Aussenbauteil   (no dot -> whole name is group)
    Strips trailing [Type] / [Instance] suffixes first.
    """
    name = param_name
    for suffix in ["[Type]", "[Instance]", "[Typ]", "[Exemplar]"]:
        if name.endswith(suffix):
            name = name[:-len(suffix)]
    if "." in name:
        return name.rsplit(".", 1)[0]
    return name.strip()


# ---------------------------------------------------------------------------
#  STEP 1 -- Read ALL ParameterBindings and group by Pset/Mset
# ---------------------------------------------------------------------------
# pset_groups: { group_name -> { "param_names": set(), "cat_ids": set() } }
pset_groups = OrderedDict()

binding_map = doc.ParameterBindings
it = binding_map.ForwardIterator()
it.Reset()
while it.MoveNext():
    try:
        defn    = it.Key
        name    = defn.Name
        binding = it.Current
        grp     = _group_from_param(name)

        if grp not in pset_groups:
            pset_groups[grp] = {"param_names": set(), "cat_ids": set()}

        pset_groups[grp]["param_names"].add(name)

        if hasattr(binding, "Categories"):
            for cat in binding.Categories:
                pset_groups[grp]["cat_ids"].add(cat.Id)
    except Exception:
        pass

if not pset_groups:
    forms.alert(
        "No project parameters found.\nLoad shared parameters first.",
        title="No Parameters", exitscript=True
    )

# ---------------------------------------------------------------------------
#  STEP 2 -- Let user pick which Psets/Msets to fill
#            Format: "Pset_WallCommon   (8 params)"
# ---------------------------------------------------------------------------
display_to_group = OrderedDict()   # display string -> group name

for grp, data in sorted(pset_groups.items()):
    n = len(data["param_names"])
    display = "{}   ({} param{})".format(grp, n, "s" if n != 1 else "")
    display_to_group[display] = grp

selected_displays = forms.SelectFromList.show(
    sorted(display_to_group.keys()),
    title="Select Psets / Msets to fill   --   {} available".format(len(display_to_group)),
    multiselect=True,
    button_name="Fill Selected"
)

if not selected_displays:
    script.exit()

# Resolve selected display strings back to group names and build:
#   selected_param_names : set of param names to fill
#   selected_cat_ids     : set of category ElementIds to collect from
selected_param_names = set()
selected_cat_ids     = set()

for disp in selected_displays:
    grp  = display_to_group[disp]
    data = pset_groups[grp]
    selected_param_names.update(data["param_names"])
    selected_cat_ids.update(data["cat_ids"])

# ---------------------------------------------------------------------------
#  STEP 3 -- Scope selector
# ---------------------------------------------------------------------------
scope_choice = forms.CommandSwitchWindow.show(
    [
        "Active View -- elements visible in the active view",
        "Selection  -- currently selected elements only",
        "Whole Model -- every element in the project (slow!)",
    ],
    message="Which elements should be filled?"
)
if not scope_choice:
    script.exit()

# ---------------------------------------------------------------------------
#  STEP 4 -- Collect instances
#
#  Use selected_cat_ids (only categories relevant to chosen Psets).
#  Normal elements  : OfCategoryId per non-spatial category
#  Spatial elements : SpatialElementFilter (rooms / MEP spaces / areas)
# ---------------------------------------------------------------------------
has_spatial = any(cid.IntegerValue in SPATIAL_BICS for cid in selected_cat_ids)


def _collect_normal(view_id=None):
    normal_ids = [cid for cid in selected_cat_ids
                  if cid.IntegerValue not in SPATIAL_BICS]
    results = []
    for cat_id in normal_ids:
        try:
            col = (DB.FilteredElementCollector(doc, view_id)
                     .OfCategoryId(cat_id)
                     .WhereElementIsNotElementType()
                   if view_id else
                   DB.FilteredElementCollector(doc)
                     .OfCategoryId(cat_id)
                     .WhereElementIsNotElementType())
            results.extend(list(col))
        except Exception:
            pass
    return results


def _collect_spatial():
    if not has_spatial:
        return []
    try:
        col = (DB.FilteredElementCollector(doc)
                 .WherePasses(DB.SpatialElementFilter())
                 .WhereElementIsNotElementType())
        return [e for e in col
                if e and e.Category
                and e.Category.Id.IntegerValue in SPATIAL_BICS
                and e.Category.Id in selected_cat_ids]
    except Exception:
        results = []
        for bic in [DB.BuiltInCategory.OST_Rooms,
                    DB.BuiltInCategory.OST_MEPSpaces,
                    DB.BuiltInCategory.OST_Areas]:
            try:
                cid = DB.ElementId(bic)
                if cid in selected_cat_ids:
                    results.extend(list(
                        DB.FilteredElementCollector(doc)
                          .OfCategoryId(cid)
                          .WhereElementIsNotElementType()
                    ))
            except Exception:
                pass
        return results


if "Active View" in scope_choice:
    vid         = doc.ActiveView.Id
    normal      = _collect_normal(view_id=vid)
    spatial_all = _collect_spatial()
    try:
        lvl_id  = doc.ActiveView.GenLevel.Id
        spatial = [e for e in spatial_all
                   if hasattr(e, "LevelId") and e.LevelId == lvl_id] or spatial_all
    except Exception:
        spatial = spatial_all
    instances = normal + spatial

elif "Selection" in scope_choice:
    sel_ids = uidoc.Selection.GetElementIds()
    if not sel_ids:
        forms.alert("Nothing is selected.", title="No Selection", exitscript=True)
    all_sel   = [doc.GetElement(eid) for eid in sel_ids]
    instances = [e for e in all_sel
                 if e and e.Category and e.Category.Id in selected_cat_ids]
    if not instances:
        instances = [e for e in all_sel if e]

else:  # Whole Model
    instances = _collect_normal() + _collect_spatial()

if not instances:
    forms.alert(
        "No elements found for the selected Psets and scope.\n"
        "Selected categories: {}".format(len(selected_cat_ids)),
        exitscript=True
    )

# ---------------------------------------------------------------------------
#  STEP 5 -- Unique TYPE elements (type-bound params live on the type)
# ---------------------------------------------------------------------------
type_ids_seen = set()
type_elements = []
for inst in instances:
    try:
        tid = inst.GetTypeId()
        if tid and tid != DB.ElementId.InvalidElementId and tid not in type_ids_seen:
            type_ids_seen.add(tid)
            te = doc.GetElement(tid)
            if te:
                type_elements.append(te)
    except Exception:
        pass

all_targets = list(instances) + list(type_elements)

# ---------------------------------------------------------------------------
#  HELPERS -- emptiness and default value
# ---------------------------------------------------------------------------
def _should_process(param):
    """True only if this param belongs to the user-selected Psets."""
    try:
        if param.Definition.Name in selected_param_names:
            return True
    except Exception:
        pass
    return False


def _is_yesno(param):
    try:
        if hasattr(param.Definition, "GetDataType"):        # Revit 2025+
            ft = param.Definition.GetDataType()
            s  = str(ft.TypeId).lower() if ft else ""
            return "bool" in s or "yesno" in s
        else:                                               # Revit <=2024
            return param.Definition.ParameterType == DB.ParameterType.YesNo
    except Exception:
        return False


def _needs_default(param):
    """
    Returns (write, storage_type, value, skip_reason).

    STRING  -> "UNSET"  when null/empty
    YESNO   -> 0        when HasValue is False
    DOUBLE  -> 0.0      when AsDouble() == 0.0
    INTEGER -> 999      when AsInteger() == 0
    """
    if param.IsReadOnly:
        return False, None, None, "read-only"

    st = param.StorageType
    try:
        if st == DB.StorageType.String:
            v = param.AsString()
            if v is None or v == "":
                return True, st, "UNSET", None
            return False, None, None, "has value"

        elif st == DB.StorageType.Integer:
            if _is_yesno(param):
                if not param.HasValue:
                    return True, st, 0, None
                return False, None, None, "has value"
            else:
                if param.AsInteger() == 0:
                    return True, st, 999, None
                return False, None, None, "has value"

        elif st == DB.StorageType.Double:
            if param.AsDouble() == 0.0:
                return True, st, 0.0, None
            return False, None, None, "has value"

    except Exception as ex:
        return False, None, None, "error: {}".format(ex)

    return False, None, None, "unsupported"


# ---------------------------------------------------------------------------
#  STEP 6 -- Fill
# ---------------------------------------------------------------------------
filled_count   = 0
skipped_count  = 0
readonly_count = 0
error_count    = 0
error_details  = []

with revit.Transaction("Fill Default Parameter Values"):
    for elem in all_targets:
        try:
            params = elem.Parameters
        except Exception:
            continue

        for param in params:
            defn_name = ""
            try:
                defn_name = param.Definition.Name

                if not _should_process(param):
                    continue

                write, st, val, reason = _needs_default(param)

                if not write:
                    if reason == "read-only":
                        readonly_count += 1
                    else:
                        skipped_count += 1
                    continue

                if st == DB.StorageType.Integer:
                    param.Set(int(val))
                elif st == DB.StorageType.String:
                    param.Set(str(val))
                elif st == DB.StorageType.Double:
                    param.Set(float(val))

                filled_count += 1

            except Exception as e:
                error_count += 1
                try:
                    eid = elem.Id.IntegerValue
                except Exception:
                    eid = "?"
                error_details.append(
                    "ElemId {}: '{}' -- {}".format(eid, defn_name or "?", str(e))
                )

# ---------------------------------------------------------------------------
#  STEP 7 -- Report
# ---------------------------------------------------------------------------
n_spatial = len([e for e in instances
                 if e and e.Category
                 and e.Category.Id.IntegerValue in SPATIAL_BICS])
n_normal  = len(instances) - n_spatial

selected_group_names = sorted(
    display_to_group[d] for d in selected_displays
)

output = script.get_output()
output.set_title("Fill Default Values -- Results")
output.print_html("<h2>Fill Default Parameter Values</h2>")
output.print_html(
    "<p><b>Scope:</b> {}</p>"
    "<p><b>Psets filled:</b> {}</p>"
    "<p><b>Parameters targeted:</b> {}</p>"
    "<p><b>Normal instances:</b> {}  |  "
    "<b>Spatial elements:</b> {}  |  "
    "<b>Unique types:</b> {}</p>".format(
        scope_choice,
        ", ".join(selected_group_names),
        len(selected_param_names),
        n_normal, n_spatial, len(type_elements)
    )
)
output.print_html("<h3 style='color:green'>Filled: {}</h3>".format(filled_count))
output.print_html(
    "<p style='color:gray'>Already had a value (skipped): {}</p>".format(skipped_count)
)
output.print_html(
    "<p style='color:gray'>Read-only (skipped): {}</p>".format(readonly_count)
)
if error_count:
    output.print_html("<h3 style='color:red'>Errors: {}</h3>".format(error_count))
    for e in error_details[:50]:
        output.print_html("<p style='color:red'>X {}</p>".format(e))
    if len(error_details) > 50:
        output.print_html("<p style='color:red'>... and {} more</p>".format(
            len(error_details) - 50))
else:
    output.print_html("<p style='color:green'>No errors.</p>")

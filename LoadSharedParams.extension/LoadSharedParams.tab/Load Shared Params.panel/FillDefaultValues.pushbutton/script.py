# -*- coding: utf-8 -*-
"""Fill empty project/shared parameters with default values.

DEFAULT VALUES PER TYPE:
  YESNO   (bool stored as Integer) -> 0     (False / unchecked)
  TEXT    (String)                 -> "UNSET"
  DOUBLE  (Area, Length, Number,   -> 0.0   (zero, universal across all units)
           Temperature, Volume...)
  INTEGER (non-YESNO)              -> 999

DESIGN DECISIONS:
  DOUBLE -> 0.0, not 999:
    Dimensional doubles (Area, Length, Volume, Temperature...) are stored
    in Revit internal units (ft², ft, ft³, Kelvin offset...).  Converting
    999 to internal units requires knowing the spec type and project display
    unit for every param -- fragile, unit-system-dependent, and wrong for
    "calculated" params where a formula already drives the value.
    0.0 is always 0 in every unit system.  It is the universal sentinel for
    "not yet user-set" for dimensional params, and avoids all unit math.

  YESNO empty-check: HasValue == False
    HasValue becomes True after the first param.Set() call.  A second run
    therefore correctly sees "already set" and skips it (idempotent).

  DOUBLE/INTEGER empty-check: AsDouble()==0.0 / AsInteger()==0
    Revit initialises storage slots to 0/0.0.  A non-zero value means
    someone (user, script, formula) already set it -> skip.

Compatible with Revit 2021 - 2026.
"""
__title__ = "Fill\nDefaults"
__author__ = "conconpo"
__doc__ = "Fills empty IFC shared parameter values: YESNO=False, TEXT=UNSET, Double=0, Int=999."

from pyrevit import revit, DB, forms, script

doc   = revit.doc
uidoc = revit.uidoc

# ---------------------------------------------------------------------------
#  STEP 1 -- Scope selector
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
#  STEP 2 -- Read ParameterBindings
#            (a) all parameter names bound to this project
#            (b) all category IDs those parameters are bound to
# ---------------------------------------------------------------------------
project_param_names = set()
bound_category_ids  = set()

# Spatial elements need a different collector (rooms, MEP spaces, areas)
SPATIAL_BICS = {
    int(DB.BuiltInCategory.OST_Rooms),
    int(DB.BuiltInCategory.OST_MEPSpaces),
    int(DB.BuiltInCategory.OST_Areas),
}

binding_map = doc.ParameterBindings
it = binding_map.ForwardIterator()
it.Reset()
while it.MoveNext():
    try:
        project_param_names.add(it.Key.Name)
        binding = it.Current
        if hasattr(binding, "Categories"):
            for cat in binding.Categories:
                bound_category_ids.add(cat.Id)
    except Exception:
        pass

if not project_param_names:
    forms.alert(
        "No project parameters found.\nLoad shared parameters first.",
        title="No Parameters", exitscript=True
    )

has_spatial = bool(
    bound_category_ids and
    any(cid.IntegerValue in SPATIAL_BICS for cid in bound_category_ids)
)

# ---------------------------------------------------------------------------
#  STEP 3 -- Collect instances
#
#  Normal elements : OfCategoryId per bound non-spatial category (view-aware)
#  Spatial elements: SpatialElementFilter (always whole-model, then scoped)
#    Reason: OfCategoryId returns 0 for rooms in 3D/section views
# ---------------------------------------------------------------------------
def _collect_normal(view_id=None):
    normal_ids = [cid for cid in bound_category_ids
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
                if e and e.Category and e.Category.Id.IntegerValue in SPATIAL_BICS
                and e.Category.Id in bound_category_ids]
    except Exception:
        results = []
        for bic in [DB.BuiltInCategory.OST_Rooms,
                    DB.BuiltInCategory.OST_MEPSpaces,
                    DB.BuiltInCategory.OST_Areas]:
            try:
                cid = DB.ElementId(bic)
                if cid in bound_category_ids:
                    results.extend(list(
                        DB.FilteredElementCollector(doc)
                          .OfCategoryId(cid)
                          .WhereElementIsNotElementType()
                    ))
            except Exception:
                pass
        return results


if "Active View" in scope_choice:
    vid     = doc.ActiveView.Id
    normal  = _collect_normal(view_id=vid)
    spatial_all = _collect_spatial()
    # Scope spatial to the view's level when possible
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
                 if e and e.Category and e.Category.Id in bound_category_ids]
    if not instances:
        instances = [e for e in all_sel if e]

else:  # Whole Model
    instances = _collect_normal() + _collect_spatial()

if not instances:
    forms.alert(
        "No elements found in the relevant categories.\n"
        "Bound categories: {}  |  Spatial: {}".format(
            len(bound_category_ids), has_spatial),
        exitscript=True
    )

# ---------------------------------------------------------------------------
#  STEP 4 -- Unique TYPE elements (type-bound params live on the type)
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
#  HELPERS
# ---------------------------------------------------------------------------
def _should_process(param):
    """True for any param from our project binding set."""
    try:
        if param.IsShared:
            return True
    except Exception:
        pass
    try:
        if param.Definition.Name in project_param_names:
            return True
    except Exception:
        pass
    return False


def _is_yesno(param):
    """True when the parameter is a Yes/No boolean type."""
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

    STRING  -> "UNSET"  when AsString() is None or ""
    YESNO   -> 0        when HasValue is False  (idempotent: Set() flips HasValue)
    DOUBLE  -> 0.0      when AsDouble() == 0.0  (universal, no unit conversion needed)
    INTEGER -> 999      when AsInteger() == 0

    Why DOUBLE -> 0 instead of 999:
      Dimensional params store values in internal units (ft², ft, ft³...).
      0.0 internal = 0.0 in any display unit.  No conversion needed.
      It correctly signals "not reviewed" for calculated params too,
      since any formula-driven value will be non-zero and thus skipped.
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
                return True, st, 0.0, None     # write 0.0 -- universal, no unit math
            return False, None, None, "has value"

    except Exception as ex:
        return False, None, None, "error: {}".format(ex)

    return False, None, None, "unsupported"


# ---------------------------------------------------------------------------
#  STEP 5 -- Fill
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
                    param.Set(float(val))      # 0.0 -- no unit conversion needed

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
#  STEP 6 -- Report
# ---------------------------------------------------------------------------
n_spatial = len([e for e in instances
                 if e and e.Category and e.Category.Id.IntegerValue in SPATIAL_BICS])
n_normal  = len(instances) - n_spatial

output = script.get_output()
output.set_title("Fill Default Values -- Results")
output.print_html(
    "<h2>Fill Default Parameter Values</h2>"
    "<p><b>Scope:</b> {}</p>"
    "<p><b>Normal instances processed:</b> {}</p>"
    "<p><b>Spatial elements (rooms/spaces/areas):</b> {}</p>"
    "<p><b>Unique types processed:</b> {}</p>"
    "<p><b>Project parameters targeted:</b> {}</p>".format(
        scope_choice, n_normal, n_spatial,
        len(type_elements), len(project_param_names)
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

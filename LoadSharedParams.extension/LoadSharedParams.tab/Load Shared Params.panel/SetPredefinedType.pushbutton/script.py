# -*- coding: utf-8 -*-
"""Set IFC class and PredefinedType on elements.

WHAT THIS DOES:
  Step 1: Collects ALL elements in the whole model automatically.
  Step 2: Builds a list of unique element TYPES grouped by category
          for display. User picks which types to process -- not categories.
          Within one category (e.g. Tragwerkstützen) different types
          (Stütze 30x30, Pilaster 20x20) can get different PredefinedTypes.
  Step 3: For each selected type, shows a picker for the IFC class.
          If the type already has an IFC class, it is shown at the top.
  Step 4: Shows valid PredefinedType values for the chosen IFC class.
          User picks the value.
  Step 5: Writes BOTH IFC class and PredefinedType to the type element
          and to all its instances.

WHERE VALUES ARE WRITTEN:
  Revit has TWO sets of IFC params -- on the TYPE and on the INSTANCE.
  The instance-level fields ("In IFC exportieren als" and
  "Vordefinierter IFC-Typ") are what is visible in the properties panel
  for a selected element and what takes PRECEDENCE during IFC export
  when set to something other than "Nach Typ" / empty.

  This script writes to BOTH:
    Instance: IFC_EXPORT_ELEMENT_TYPE_AS  (In IFC exportieren als)
              IFC_EXPORT_PREDEFINEDTYPE   (Vordefinierter IFC-Typ)
    Type    : IFC_EXPORT_ELEMENT_AS       (Typ in IFC exportieren als)
              IFC_EXPORT_PREDEFINEDTYPE   (on type element)

  This guarantees the value is visible AND used during export
  regardless of whether the export setting is "Nach Typ" or "Instanz".

IFC CLASS SOURCE:
  Read from the element or its type. If missing, user assigns it.
  IFC classes are fixed by the IFC4 schema -- no new classes can be
  invented. Unknown classes default to IfcBuildingElementProxy.

Compatible with Revit 2021 - 2026.
"""
__title__ = "Set\nPredefined\nType"
__author__ = "conconpo"
__doc__ = (
    "Sets IFC class and PredefinedType on elements.\n\n"
    "Works even when no IFC class is currently assigned.\n"
    "Writes to both instance and type fields for reliable export."
)

from collections import OrderedDict
from pyrevit import revit, DB, forms, script

doc = revit.doc

# ---------------------------------------------------------------------------
#  IFC4 schema: valid PredefinedType values per IFC class
# ---------------------------------------------------------------------------
PREDEFINED_TYPES = {
    "IfcBeam":              ["BEAM","HOLLOWCORE","JOIST","LINTEL","SPANDREL","T_BEAM"],
    "IfcBuildingElementProxy": ["COMPLEX","ELEMENT","PARTIAL","PROVISIONFORVOID","PROVISIONFORSPACE"],
    "IfcColumn":            ["COLUMN","PILASTER","PILESPLICE"],
    "IfcCovering":          ["CEILING","CLADDING","FLOORING","INSULATION","MEMBRANE","MOLDING","ROOFING","SKIRTINGBOARD","SLEEVING","WRAPPING"],
    "IfcCurtainWall":       [],
    "IfcDoor":              ["DOOR","GATE","TRAPDOOR"],
    "IfcFooting":           ["CAISSON_FOUNDATION","FOOTING_BEAM","PAD_FOOTING","PILE_CAP","STRIP_FOOTING"],
    "IfcFurniture":         ["CHAIR","DESK","FILECABINET","SHELF","SOFA","TABLE","TECHNICALITEM","WARDROBE"],
    "IfcMember":            ["BRACE","CHORD","COLLAR","MEMBER","MULLION","PLATE","POST","PURLIN","RAFTER","STRINGER","STRUT","STUD"],
    "IfcPile":              ["BORED","COHESION","DRIVEN","JETGROUTING"],
    "IfcPlate":             ["CURTAIN_PANEL","SHEET"],
    "IfcRailing":           ["BALUSTRADE","FENCE","GUARDRAIL","HANDRAIL"],
    "IfcRamp":              ["HALF_TURN_RAMP","L_SHAPED_RAMP","SPIRAL_RAMP","STRAIGHT_RUN_RAMP","TWO_QUARTER_TURN_RAMP","TWO_STRAIGHT_RUN_RAMP","U_SHAPED_RAMP"],
    "IfcRoof":              ["BARREL_ROOF","BUTTERFLY_ROOF","DOME_ROOF","FLAT_ROOF","FREEFORM","GABLE_ROOF","GAMBREL_ROOF","HIPPED_GABLE_ROOF","HIP_ROOF","MANSARD_ROOF","PAVILION_ROOF","RAINBOW_ROOF","SHED_ROOF"],
    "IfcSanitaryTerminal":  ["BATH","BIDET","CISTERN","SHOWER","SINK","SANITARYFOUNTAIN","TOILETPAN","URINAL","WASHHANDBASIN","WCSEAT"],
    "IfcSlab":              ["BASESLAB","FLOOR","LANDING","ROOF"],
    "IfcSpace":             ["EXTERNAL","GFA","INTERNAL","PARKING","SPACE"],
    "IfcSpaceHeater":       ["CONVECTOR","RADIATOR"],
    "IfcStair":             ["CURVED_RUN_STAIR","DOUBLE_RETURN_STAIR","HALF_TURN_STAIR","HALF_WINDING_STAIR","L_SHAPED_TURN_STAIR","QUARTER_TURN_STAIR","QUARTER_WINDING_STAIR","SPIRAL_STAIR","STRAIGHT_RUN_STAIR","THREE_QUARTER_TURN_STAIR","THREE_QUARTER_WINDING_STAIR","TWO_CURVED_RUN_STAIR","TWO_QUARTER_TURN_STAIR","TWO_QUARTER_WINDING_STAIR","TWO_STRAIGHT_RUN_STAIR","U_SHAPED_TURN_STAIR"],
    "IfcStairFlight":       ["CURVED","SPIRAL","STRAIGHT","WINDER"],
    "IfcWall":              ["ELEMENTEDWALL","MOVABLE","PARAPET","PARTITIONING","PLUMBINGWALL","POLYGONAL","SHEAR","SOLIDWALL","STANDARD"],
    "IfcWindow":            ["FIXED","LIGHTDOME","SKYLIGHT","WINDOW"],
    # MEP
    "IfcAirTerminal":       ["DIFFUSER","GRILLE","LOUVRE","REGISTER"],
    "IfcDamper":            ["BACKDRAFTDAMPER","BALANCINGDAMPER","BLASTDAMPER","CONTROLDAMPER","FIREDAMPER","FIRESMOKEDAMPER","GRAVITYDAMPER","GRAVITYRELIEFDAMPER","RELIEFDAMPER","SMOKEDAMPER"],
    "IfcDuctFitting":       ["BEND","CONNECTOR","ENTRY","EXIT","JUNCTION","OBSTRUCTION","TRANSITION"],
    "IfcDuctSegment":       ["FLEXIBLESEGMENT","RIGIDSEGMENT"],
    "IfcFan":               ["CENTRIFUGALAIRFOIL","CENTRIFUGALBACKWARDINCLINEDCURVED","CENTRIFUGALFORWARDCURVED","CENTRIFUGALRADIAL","INLINE","MIXEDFLOW","PLUGFAN","PROPELLER","TUBEAXIAL","VANEAXIAL"],
    "IfcPipeFitting":       ["BEND","CONNECTOR","ENTRY","EXIT","JUNCTION","OBSTRUCTION","TRANSITION"],
    "IfcPipeSegment":       ["CULVERT","FLEXIBLESEGMENT","GUTTER","RIGIDSEGMENT","SPOOL"],
    "IfcPump":              ["CIRCULATOR","ENDSUCTION","SPLITCASE","SUBMERSIBLEPUMP","SUMPPUMP","VERTICALINLINE","VERTICALTURBINE"],
    "IfcTank":              ["BASIN","BREAKPRESSUREVESSEL","EXPANSION","FEEDANDEXPANSION","PRESSUREVESSEL","STORAGE","VESSEL"],
    "IfcValve":             ["AIRRELEASE","ANTIVACUUM","CHANGEOVER","CHECK","COMMISSIONING","DIVERTING","DRAWOFFCOCK","DOUBLECHECK","DOUBLEREGULATING","FLUSHING","GASCOCK","GASTAP","ISOLATING","MIXING","PRESSUREREDUCING","PRESSURERELIEF","REGULATING","SAFETYCUTOFF","STEAMTRAP","STOPCOCK"],
    "IfcSensor":            ["CONDUCTANCESENSOR","CONTACTSENSOR","COSENSOR","FIRESENSOR","FLOWSENSOR","FROSTSENSOR","GASSENSOR","HEATSENSOR","HUMIDITYSENSOR","LEVELSENSOR","LIGHTSENSOR","MOISTURESENSOR","MOVEMENTSENSOR","PRESSURESENSOR","SMOKESENSOR","SOUNDSENSOR","TEMPERATURESENSOR","WINDSENSOR"],
    "IfcOutlet":            ["AUDIOVISUALOUTLET","COMMUNICATIONSOUTLET","DATAOUTLET","POWEROUTLET","TELEPHONEOUTLET"],
    "IfcLightFixture":      ["DIRECTIONSOURCE","POINTSOURCE","SECURITYLIGHTING"],
    "IfcUnitaryEquipment":  ["AIRHANDLER","AIRCONDITIONINGUNIT","DEHUMIDIFIER","SPLITSYSTEM","ROOFTOPUNIT"],
}
ALWAYS_AVAILABLE = ["NOTDEFINED", "USERDEFINED"]

# All known IFC classes for the picker when no class is assigned
ALL_IFC_CLASSES = sorted(list(PREDEFINED_TYPES.keys()) + [
    "IfcDistributionElement",
    "IfcElectricAppliance",
    "IfcFlowSegment",
    "IfcFlowFitting",
    "IfcFlowTerminal",
    "IfcFlowController",
    "IfcFlowMovingDevice",
    "IfcFlowStorageDevice",
    "IfcEnergyConversionDevice",
    "IfcOpeningElement",
    "IfcGrid",
])

# Revit BIPs -- both instance and type variants
# Instance params (visible in properties panel for selected element)
BIP_INST_CLASS  = DB.BuiltInParameter.IFC_EXPORT_ELEMENT_TYPE_AS   # "In IFC exportieren als"
BIP_INST_PREDEF = DB.BuiltInParameter.IFC_EXPORT_PREDEFINEDTYPE    # "Vordefinierter IFC-Typ" (instance)
# Type params
BIP_TYPE_CLASS  = DB.BuiltInParameter.IFC_EXPORT_ELEMENT_AS        # "Typ in IFC exportieren als"
BIP_TYPE_PREDEF = DB.BuiltInParameter.IFC_EXPORT_PREDEFINEDTYPE    # same BIP, on type element

# ---------------------------------------------------------------------------
#  STEP 1 -- Collect ALL instances in the whole model
#
#  Two separate collectors are needed:
#  (a) Normal elements: WhereElementIsNotElementType() covers walls, floors,
#      columns, beams, doors, windows, families, etc.
#  (b) Spatial elements: rooms, MEP spaces, areas.
#      These are physical elements but FilteredElementCollector can miss
#      them or return them with Category=None depending on Revit version.
#      SpatialElementFilter guarantees they are found.
# ---------------------------------------------------------------------------
_normal = list(
    DB.FilteredElementCollector(doc)
      .WhereElementIsNotElementType()
)

_spatial = []
try:
    _spatial = list(
        DB.FilteredElementCollector(doc)
          .WherePasses(DB.SpatialElementFilter())
          .WhereElementIsNotElementType()
    )
except Exception:
    for _bic in [DB.BuiltInCategory.OST_Rooms,
                 DB.BuiltInCategory.OST_MEPSpaces,
                 DB.BuiltInCategory.OST_Areas]:
        try:
            _spatial.extend(list(
                DB.FilteredElementCollector(doc)
                  .OfCategoryId(DB.ElementId(_bic))
                  .WhereElementIsNotElementType()
            ))
        except Exception:
            pass

# Merge, deduplicate by ElementId
_seen_ids    = set()
all_instances = []
for _e in _normal + _spatial:
    try:
        if _e and _e.Id not in _seen_ids:
            _seen_ids.add(_e.Id)
            all_instances.append(_e)
    except Exception:
        pass

if not all_instances:
    forms.alert("No elements found in the model.", exitscript=True)

# ---------------------------------------------------------------------------
#  HELPERS
# ---------------------------------------------------------------------------
def _read_param_str(elem, bip):
    """Read a string BuiltInParameter from an element. Returns '' if missing."""
    try:
        p = elem.get_Parameter(bip)
        if p and p.HasValue:
            v = p.AsString()
            return v.strip() if v else ""
    except Exception:
        pass
    return ""


def _write_param_str(elem, bip, value, fallback_names=None):
    """Write value to a BIP or fallback parameter names. Returns True on success."""
    try:
        p = elem.get_Parameter(bip)
        if p and not p.IsReadOnly:
            p.Set(value)
            return True
    except Exception:
        pass
    for name in (fallback_names or []):
        try:
            p = elem.LookupParameter(name)
            if p and not p.IsReadOnly:
                p.Set(value)
                return True
        except Exception:
            pass
    return False


def _get_current_ifc_class(inst):
    """Read IFC class from instance first, then from its type."""
    v = _read_param_str(inst, BIP_INST_CLASS)
    if v and v.lower() not in ("", "nach typ", "by type"):
        return v
    try:
        te = doc.GetElement(inst.GetTypeId())
        if te:
            v = _read_param_str(te, BIP_TYPE_CLASS)
            if v:
                return v
    except Exception:
        pass
    return ""


def _get_elem_label(inst):
    """Family : Type label for display."""
    try:
        te = doc.GetElement(inst.GetTypeId())
        if te:
            fam = _read_param_str(te,
                DB.BuiltInParameter.ALL_MODEL_FAMILY_NAME)
            name = _read_param_str(te,
                DB.BuiltInParameter.SYMBOL_NAME_PARAM) or te.Name
            return "{} : {}".format(fam, name) if fam else name
    except Exception:
        pass
    return "Element #{}".format(inst.Id.IntegerValue)

# ---------------------------------------------------------------------------
#  STEP 3 -- Build list of unique ELEMENT TYPES grouped by category
#
#  Grouping unit = element type, NOT category.
#  "Tragwerkstützen / Stütze 30x30" and "Tragwerkstützen / Pilaster 20x20"
#  are two separate entries -- each gets its own IFC class + PredefinedType.
#
#  SPATIAL ELEMENTS (rooms, spaces, areas) have no type element in Revit --
#  GetTypeId() returns InvalidElementId. They are grouped by instance Name
#  (room name) under their category, and IFC params are written only to
#  the instance (no type write step).
#
#  Display format:
#    "Tragwerkstützen  |  Beton : Stütze 30x30  (10 insts)  [IFC: IfcColumn / COLUMN]"
#    "Räume  |  Raum: Büro  (1 inst)  [IFC: IfcSpace / not set]"
# ---------------------------------------------------------------------------

# Key: either (tid) for typed elements, or ("inst", inst.Id) for typeless.
# Value: dict with te, cat, fam, type, ifc_class, predef, instances, is_typeless
type_data  = OrderedDict()
seen_tids  = set()

# Sentinel used as key for typeless elements
_TYPELESS = "TYPELESS"

for inst in all_instances:
    try:
        # Get category name -- spatial elements always have a category
        if inst.Category:
            cat_name = inst.Category.Name
        else:
            # No category: internal Revit objects (views, annotations) -- skip
            continue

        tid = None
        try:
            tid = inst.GetTypeId()
        except Exception:
            pass

        is_typeless = (
            tid is None or
            tid == DB.ElementId.InvalidElementId
        )

        if is_typeless:
            # ---- Spatial / typeless element --------------------------------
            # Use instance element id as unique key so each room is its own entry
            key = (_TYPELESS, inst.Id.IntegerValue)

            # Display name: room Name parameter if available
            inst_name = ""
            try:
                p = inst.get_Parameter(DB.BuiltInParameter.ROOM_NAME)
                if p:
                    inst_name = p.AsString() or ""
            except Exception:
                pass
            if not inst_name:
                inst_name = getattr(inst, "Name", "") or "#{}" .format(
                    inst.Id.IntegerValue)

            ifc_class = _read_param_str(inst, BIP_INST_CLASS)
            if ifc_class.lower() in ("nach typ", "by type"):
                ifc_class = ""
            predef = _read_param_str(inst, BIP_INST_PREDEF)

            type_data[key] = {
                "te":         None,           # no type element
                "cat":        cat_name,
                "fam":        "",
                "type":       inst_name,
                "ifc_class":  ifc_class,
                "predef":     predef,
                "instances":  [inst],
                "is_typeless": True,
            }

        else:
            # ---- Normal typed element --------------------------------------
            key = tid
            if key not in seen_tids:
                seen_tids.add(key)
                te = doc.GetElement(tid)
                if not te:
                    continue

                fam_name  = _read_param_str(
                    te, DB.BuiltInParameter.ALL_MODEL_FAMILY_NAME)
                type_name = (
                    _read_param_str(te, DB.BuiltInParameter.SYMBOL_NAME_PARAM)
                    or getattr(te, "Name", str(tid.IntegerValue))
                )
                ifc_class = _get_current_ifc_class(inst)
                predef    = _read_param_str(te, BIP_TYPE_PREDEF)

                type_data[key] = {
                    "te":         te,
                    "cat":        cat_name,
                    "fam":        fam_name,
                    "type":       type_name,
                    "ifc_class":  ifc_class,
                    "predef":     predef,
                    "instances":  [],
                    "is_typeless": False,
                }

            if key in type_data:
                type_data[key]["instances"].append(inst)

    except Exception:
        pass

if not type_data:
    forms.alert("No element types found in the model.", exitscript=True)

# ---------------------------------------------------------------------------
#  STEP 3a -- CATEGORY picker
#
#  Build one entry per unique category, showing how many types are inside
#  and a hint about how many already have IFC class set.
#  "Tragwerkstützen  (3 types,  2 with IFC class set)"
#  "Räume  (5 rooms,  0 with IFC class set)"
# ---------------------------------------------------------------------------

# cat_name -> list of keys in type_data belonging to that category
cat_to_keys = OrderedDict()
for k, d in type_data.items():
    cat = d["cat"]
    if cat not in cat_to_keys:
        cat_to_keys[cat] = []
    cat_to_keys[cat].append(k)

def _cat_display(cat_name, keys):
    n_types   = len(keys)
    n_with_ifc = sum(1 for k in keys if type_data[k]["ifc_class"])
    type_word = "type" if n_types == 1 else "types"
    return "{}   ({} {},  {} with IFC class set)".format(
        cat_name, n_types, type_word, n_with_ifc)

cat_display_to_name = OrderedDict()
for cat_name in sorted(cat_to_keys.keys()):
    disp = _cat_display(cat_name, cat_to_keys[cat_name])
    cat_display_to_name[disp] = cat_name

selected_cat_displays = forms.SelectFromList.show(
    list(cat_display_to_name.keys()),
    title="Step 1 of 3 — Select categories   ({} found)".format(
        len(cat_display_to_name)),
    multiselect=True,
    button_name="Next: select types"
)
if not selected_cat_displays:
    script.exit()

# ---------------------------------------------------------------------------
#  STEP 3b -- TYPE picker per selected category
#
#  For each selected category, show a multiselect list of the types within
#  it. The user can pick all or a subset.
#  If there is only one type in the category the picker is still shown
#  so the user can confirm, but they can also just click OK.
#
#  Display per type:
#    "Stütze : 30x30  (10 insts)  [IFC: IfcColumn / COLUMN]"
#    "Stütze : Pilaster 20x20  (5 insts)  [IFC: not set]"
# ---------------------------------------------------------------------------

selected_keys = []   # flat list of type_data keys chosen by user

for cat_disp in selected_cat_displays:
    cat_name = cat_display_to_name[cat_disp]
    keys_in_cat = cat_to_keys[cat_name]

    # Sort by family + type name
    keys_sorted = sorted(
        keys_in_cat,
        key=lambda k: (type_data[k]["fam"], type_data[k]["type"])
    )

    type_disp_to_key = OrderedDict()
    for k in keys_sorted:
        d     = type_data[k]
        label = "{} : {}".format(d["fam"], d["type"]) if d["fam"] \
                else d["type"]
        n     = len(d["instances"])
        count = "({} inst{})".format(n, "s" if n != 1 else "")
        if d["ifc_class"] and d["predef"]:
            hint = "[IFC: {} / {}]".format(d["ifc_class"], d["predef"])
        elif d["ifc_class"]:
            hint = "[IFC: {} / not set]".format(d["ifc_class"])
        else:
            hint = "[IFC: not set]"
        disp = "{}  {}  {}".format(label, count, hint)
        type_disp_to_key[disp] = k

    chosen_type_disps = forms.SelectFromList.show(
        list(type_disp_to_key.keys()),
        title="Step 2 of 3 — [{}]  Select types  ({} found)".format(
            cat_name, len(type_disp_to_key)),
        multiselect=True,
        button_name="Next: assign IFC class"
    )
    if not chosen_type_disps:
        continue   # user skipped this category -- move to next

    for disp in chosen_type_disps:
        selected_keys.append(type_disp_to_key[disp])

if not selected_keys:
    script.exit()

# ---------------------------------------------------------------------------
#  STEP 4 -- Assign IFC class + PredefinedType
#
#  After the type selection, ask once per category batch:
#    "Same for all selected types"  → one picker, applied to all
#    "Set individually"             → loop through each type separately
#
#  The choice is asked once per category (not once globally) because
#  within one run you may want same-for-all for walls but individual
#  for columns.
# ---------------------------------------------------------------------------
assignments = []   # list of (instance, ifc_class, predef_type)

# Group selected_keys back by category for the per-category mode question
from collections import OrderedDict as _OD
cat_to_selected_keys = _OD()
for k in selected_keys:
    cat = type_data[k]["cat"]
    if cat not in cat_to_selected_keys:
        cat_to_selected_keys[cat] = []
    cat_to_selected_keys[cat].append(k)


def _pick_ifc_class(d, cat_label):
    """Show IFC class picker for a type entry. Returns chosen string or ''."""
    label = "{} : {}".format(d["fam"], d["type"]) if d["fam"] else d["type"]
    picker = []
    if d["ifc_class"]:
        picker.append(d["ifc_class"])
    for c in ALL_IFC_CLASSES:
        if c not in picker:
            picker.append(c)
    return forms.SelectFromList.show(
        picker,
        title="[{}  |  {}]  IFC class  ({} instance{})".format(
            cat_label, label,
            len(d["instances"]),
            "s" if len(d["instances"]) != 1 else ""),
        multiselect=False,
        button_name="Next: PredefinedType"
    ) or ""


def _pick_predef(chosen_class, d, cat_label):
    """Show PredefinedType picker. Returns chosen string or ''."""
    label = "{} : {}".format(d["fam"], d["type"]) if d["fam"] else d["type"]
    valid = list(PREDEFINED_TYPES.get(chosen_class, []))
    for v in ALWAYS_AVAILABLE:
        if v not in valid:
            valid.append(v)
    if d["predef"] and d["predef"] in valid:
        valid.remove(d["predef"])
        valid.insert(0, d["predef"])
    return forms.SelectFromList.show(
        valid,
        title="[{}  |  {}  /  {}]  PredefinedType".format(
            cat_label, label, chosen_class),
        multiselect=False,
        button_name="Assign"
    ) or ""


for cat_name, keys in cat_to_selected_keys.items():

    # Only ask the same/individual question when more than one type selected
    if len(keys) == 1:
        mode = "Set individually"
    else:
        n_total = sum(len(type_data[k]["instances"]) for k in keys)
        mode = forms.CommandSwitchWindow.show(
            [
                "Same for all -- one IFC class and PredefinedType for all {} selected types".format(len(keys)),
                "Set individually -- choose separately for each type",
            ],
            message="[{}]  {} types selected ({} instances total) -- assign how?".format(
                cat_name, len(keys), n_total)
        )
        if not mode:
            continue   # user cancelled this category

    # ---- SAME FOR ALL -------------------------------------------------------
    if "Same for all" in mode:
        # Use first type's current values as hint for the pickers
        d_first = type_data[keys[0]]

        # Collect all current IFC classes across selected types to build hint
        current_classes = list(dict.fromkeys(
            type_data[k]["ifc_class"] for k in keys
            if type_data[k]["ifc_class"]
        ))
        # Build picker: existing classes first (de-duped), then full list
        picker = list(current_classes)
        for c in ALL_IFC_CLASSES:
            if c not in picker:
                picker.append(c)

        n_total = sum(len(type_data[k]["instances"]) for k in keys)
        chosen_class = forms.SelectFromList.show(
            picker,
            title="[{}]  IFC class — applied to all {} types  ({} instances)".format(
                cat_name, len(keys), n_total),
            multiselect=False,
            button_name="Next: PredefinedType"
        )
        if not chosen_class:
            continue

        valid = list(PREDEFINED_TYPES.get(chosen_class, []))
        for v in ALWAYS_AVAILABLE:
            if v not in valid:
                valid.append(v)
        # Current predef hint from first type
        if d_first["predef"] and d_first["predef"] in valid:
            valid.remove(d_first["predef"])
            valid.insert(0, d_first["predef"])

        chosen_predef = forms.SelectFromList.show(
            valid,
            title="[{}  /  {}]  PredefinedType — applied to all {} types".format(
                cat_name, chosen_class, len(keys)),
            multiselect=False,
            button_name="Assign to all"
        )
        if not chosen_predef:
            continue

        for k in keys:
            for inst in type_data[k]["instances"]:
                assignments.append((inst, chosen_class, chosen_predef))

    # ---- SET INDIVIDUALLY ---------------------------------------------------
    else:
        for k in keys:
            d     = type_data[k]
            label = "{} : {}".format(d["fam"], d["type"]) if d["fam"] \
                    else d["type"]

            chosen_class = _pick_ifc_class(d, cat_name)
            if not chosen_class:
                continue

            chosen_predef = _pick_predef(chosen_class, d, cat_name)
            if not chosen_predef:
                continue

            for inst in d["instances"]:
                assignments.append((inst, chosen_class, chosen_predef))

if not assignments:
    forms.alert("No assignments made. Nothing was changed.", exitscript=True)

# ---------------------------------------------------------------------------
#  STEP 5 -- Write to instance AND type
# ---------------------------------------------------------------------------
inst_written  = 0
type_written  = 0
error_count   = 0
error_details = []
type_ids_done = set()

with revit.Transaction("Set IFC Class and PredefinedType"):
    for inst, ifc_class, predef in assignments:
        # --- Write to instance (always) ---
        try:
            ok_class  = _write_param_str(inst, BIP_INST_CLASS,
                ifc_class, ["In IFC exportieren als"])
            ok_predef = _write_param_str(inst, BIP_INST_PREDEF,
                predef, ["Vordefinierter IFC-Typ", "PredefinedType"])
            if ok_class or ok_predef:
                inst_written += 1
        except Exception as e:
            error_count += 1
            error_details.append(
                "Instance #{}: {}".format(inst.Id.IntegerValue, str(e)))

        # --- Write to type (once per unique type, skip for typeless) ---
        try:
            tid = inst.GetTypeId()
            # Rooms / spaces / areas return InvalidElementId -- skip type write
            if (tid and
                    tid != DB.ElementId.InvalidElementId and
                    tid not in type_ids_done):
                type_ids_done.add(tid)
                te = doc.GetElement(tid)
                if te:
                    _write_param_str(te, BIP_TYPE_CLASS,
                        ifc_class, ["Typ in IFC exportieren als"])
                    _write_param_str(te, BIP_TYPE_PREDEF,
                        predef, ["Vordefinierter IFC-Typ", "PredefinedType"])
                    type_written += 1
        except Exception as e:
            error_details.append(
                "Type of #{}: {}".format(inst.Id.IntegerValue, str(e)))

# ---------------------------------------------------------------------------
#  STEP 6 -- Report
# ---------------------------------------------------------------------------
output = script.get_output()
output.set_title("Set IFC Class + PredefinedType -- Results")
output.print_html(
    "<h2>Set IFC Class and PredefinedType</h2>"
    "<p><b>Scope:</b> Whole Model</p>"
    "<p><b>Categories selected:</b> {}</p>"
    "<p><b>Types updated:</b> {}</p>"
    "<p><b>Elements assigned:</b> {}</p>".format(
        len(selected_cat_displays),
        len(selected_keys),
        len(assignments)
    )
)
output.print_html(
    "<h3 style='color:green'>Written to instances: {}  |  "
    "Written to types: {}</h3>".format(inst_written, type_written)
)

# Summary table
if assignments:
    # Group by (category, ifc_class, predef) for compact display
    from collections import Counter
    summary = Counter(
        (ifc_class, predef)
        for (inst, ifc_class, predef) in assignments
    )
    output.print_html(
        "<table style='border-collapse:collapse;width:100%;margin-top:8px'>"
        "<tr style='border-bottom:2px solid #ccc'>"
        "<th style='text-align:left;padding:4px 8px'>IFC class</th>"
        "<th style='text-align:left;padding:4px 8px'>PredefinedType</th>"
        "<th style='text-align:left;padding:4px 8px'>Elements</th>"
        "</tr>"
    )
    for (ifc_class, predef), count in sorted(summary.items()):
        output.print_html(
            "<tr style='border-bottom:1px solid #eee'>"
            "<td style='padding:3px 8px'>{}</td>"
            "<td style='padding:3px 8px'><b>{}</b></td>"
            "<td style='padding:3px 8px'>{}</td>"
            "</tr>".format(ifc_class, predef, count)
        )
    output.print_html("</table>")

if error_count:
    output.print_html(
        "<h3 style='color:red'>Errors: {}</h3>".format(error_count))
    for e in error_details[:30]:
        output.print_html("<p style='color:red'>X {}</p>".format(e))
else:
    output.print_html("<p style='color:green'>No errors.</p>")

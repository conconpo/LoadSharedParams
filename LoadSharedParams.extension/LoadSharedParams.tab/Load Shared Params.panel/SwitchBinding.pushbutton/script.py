# -*- coding: utf-8 -*-
"""Switch Parameter Binding (mit Wertuebertragung) — Instance <-> Type.

Schaltet die Bindungsebene gebundener Parameter um UND uebertraegt die
vorhandenen Werte:

  TYP  ->  EXEMPLAR   (verlustfrei, automatisch)
      Der eine Typ-Wert wird auf ALLE Exemplare dieses Typs verteilt.

  EXEMPLAR  ->  TYP   (mit Konfliktaufloesung)
      Pro Familientyp werden die Exemplarwerte gesammelt. Sind sie
      eindeutig (oder nur 1 Exemplar), wird der Wert automatisch
      uebernommen. Bei UNTERSCHIEDLICHEN Werten erscheint ein
      Auswahldialog je Typ, in dem der zu uebertragende Wert gewaehlt
      wird. Leer ist als Option waehlbar. Anzeige ist formatiert
      (mit Einheiten, wo vorhanden).

Technik: Die Revit-API kann den Scope nicht direkt aendern. Werte werden
VOR dem Remove/Insert ausgelesen, zwischengespeichert und danach
zurueckgeschrieben.

Kompatibel mit Revit 2024, 2025 und 2026.
"""
__title__ = "Switch+Werte\nInst <-> Typ"
__author__ = "conconpo"
__doc__ = ("Schaltet die Bindungsebene um und uebertraegt die Werte. "
           "Typ->Exemplar verlustfrei; Exemplar->Typ mit Wertauswahl je "
           "Konflikt-Typ. Bilingual UI (DE/EN).")

from pyrevit import revit, DB, forms, script

doc = revit.doc
app = doc.Application
revit_version = int(app.VersionNumber)

# ===========================================================
#  Wert-Helfer / value helpers
# ===========================================================
def get_value(param):
    """Liest einen Parameterwert als dict:
    {has, st, raw, disp}. has=False -> leer."""
    if param is None:
        return None
    if not param.HasValue:
        return {"has": False, "st": str(param.StorageType), "disp": "(leer)"}

    st = param.StorageType
    if st == DB.StorageType.String:
        raw = param.AsString()
        disp = raw if raw not in (None, "") else "(leer)"
        return {"has": raw not in (None, ""), "st": "String",
                "raw": raw if raw is not None else "", "disp": disp}
    elif st == DB.StorageType.Integer:
        raw = param.AsInteger()
        disp = param.AsValueString() or str(raw)
        return {"has": True, "st": "Integer", "raw": raw, "disp": disp}
    elif st == DB.StorageType.Double:
        raw = param.AsDouble()
        disp = param.AsValueString() or str(raw)
        return {"has": True, "st": "Double", "raw": raw, "disp": disp}
    elif st == DB.StorageType.ElementId:
        raw = param.AsElementId()
        disp = param.AsValueString() or str(raw.IntegerValue)
        return {"has": True, "st": "ElementId", "raw": raw, "disp": disp}
    return None


def set_value(param, v):
    """Schreibt einen zwischengespeicherten Wert zurueck."""
    if param is None or v is None:
        return False
    if not v.get("has"):
        return False            # leer lassen
    if param.IsReadOnly:
        return False
    st = v["st"]
    try:
        if st == "String":
            param.Set(v["raw"] if v["raw"] is not None else "")
        elif st == "Integer":
            param.Set(v["raw"])
        elif st == "Double":
            param.Set(v["raw"])
        elif st == "ElementId":
            param.Set(v["raw"])
        else:
            return False
        return True
    except Exception:
        return False


def value_key(v):
    """Eindeutiger Schluessel zum Vergleichen von Werten."""
    if not v or not v.get("has"):
        return ("EMPTY",)
    st = v["st"]
    if st == "ElementId":
        return (st, v["raw"].IntegerValue)
    return (st, v["raw"])


def binding_is_instance(binding):
    return isinstance(binding, DB.InstanceBinding)


def copy_category_set(binding):
    cset = app.Create.NewCategorySet()
    for cat in binding.Categories:
        try:
            cset.Insert(cat)
        except Exception:
            pass
    return cset


def get_param_group(definition):
    try:
        return definition.GetGroupTypeId()
    except Exception:
        pass
    try:
        return definition.ParameterGroup
    except Exception:
        return None


def collect_instances(cats):
    """Alle Exemplare der gegebenen Kategorien."""
    res = []
    for cat in cats:
        try:
            col = (DB.FilteredElementCollector(doc)
                   .OfCategoryId(cat.Id)
                   .WhereElementIsNotElementType()
                   .ToElements())
            res.extend(col)
        except Exception:
            pass
    return res


def collect_types(cats):
    """Alle Familientypen der gegebenen Kategorien."""
    res = []
    for cat in cats:
        try:
            col = (DB.FilteredElementCollector(doc)
                   .OfCategoryId(cat.Id)
                   .WhereElementIsElementType()
                   .ToElements())
            res.extend(col)
        except Exception:
            pass
    return res


# ===========================================================
#  SCHRITT 1 — Gebundene Parameter lesen
# ===========================================================
binding_map = doc.ParameterBindings
all_bound = []
it = binding_map.ForwardIterator()
it.Reset()
while it.MoveNext():
    key_def = it.Key
    binding = it.Current
    all_bound.append((key_def.Name, key_def, binding,
                      binding_is_instance(binding)))

if not all_bound:
    forms.alert("Keine gebundenen Parameter im Projekt.\n"
                "No bound parameters in this project.",
                title="Nichts zu tun / Nothing to do", exitscript=True)

all_bound.sort(key=lambda x: x[0].lower())

# ===========================================================
#  SCHRITT 2 — Richtung
# ===========================================================
direction = forms.CommandSwitchWindow.show(
    [
        "Typ  ->  Exemplar / Type -> Instance",
        "Exemplar  ->  Typ / Instance -> Type",
        "Automatisch umschalten / Toggle each (auto)",
    ],
    message="Welche Richtung? / Which direction?"
)
if not direction:
    script.exit()

if "Typ  ->" in direction or "Type ->" in direction:
    mode = "to_instance"
elif "Exemplar  ->" in direction or "Instance ->" in direction:
    mode = "to_type"
else:
    mode = "toggle"


def eligible(is_inst):
    if mode == "to_instance":
        return not is_inst
    if mode == "to_type":
        return is_inst
    return True


# ===========================================================
#  SCHRITT 3 — Parameter waehlen
# ===========================================================
label_map = {}
for entry in all_bound:
    name, key_def, binding, is_inst = entry
    if not eligible(is_inst):
        continue
    scope = "Exemplar" if is_inst else "Typ"
    label_map["{}   [aktuell: {}]".format(name, scope)] = entry

if not label_map:
    forms.alert("Keine passenden Parameter fuer diese Richtung.\n"
                "No matching parameters for this direction.",
                title="Nichts zu tun / Nothing to do", exitscript=True)

chosen_labels = forms.SelectFromList.show(
    sorted(label_map.keys()),
    title="Parameter waehlen / Select parameters",
    multiselect=True,
    button_name="Weiter / Next"
)
if not chosen_labels:
    script.exit()

selected = [label_map[l] for l in chosen_labels]

# ===========================================================
#  SCHRITT 4 — Werte AUSLESEN (vor dem Umschalten)
#  plan[i] = dict mit allem, was zum Zurueckschreiben noetig ist
# ===========================================================
plan = []

for name, key_def, binding, is_inst in selected:
    cats = list(binding.Categories)

    # Zielrichtung bestimmen
    if mode == "to_instance":
        target_instance = True
    elif mode == "to_type":
        target_instance = False
    else:
        target_instance = not is_inst

    if target_instance == is_inst:
        continue   # schon im Zielzustand

    entry = {"name": name, "key_def": key_def, "binding": binding,
             "cats": cats, "from_instance": is_inst,
             "target_instance": target_instance,
             "type_values": {}, "skip": False}

    if not target_instance:
        # ----- EXEMPLAR -> TYP : Werte je Typ sammeln -----
        per_type = {}       # typeId.IntegerValue -> {"type_elem":te, "vals":[v,...]}
        for inst in collect_instances(cats):
            tid = inst.GetTypeId()
            if tid is None or tid == DB.ElementId.InvalidElementId:
                continue
            p = inst.LookupParameter(name)
            v = get_value(p)
            if v is None:
                continue
            slot = per_type.setdefault(tid.IntegerValue,
                                       {"tid": tid, "vals": []})
            slot["vals"].append(v)

        # Konflikte aufloesen
        chosen_per_type = {}     # tid.IntegerValue -> v
        for tid_int, slot in per_type.items():
            vals = slot["vals"]
            # distinkte Werte
            uniq = {}
            for v in vals:
                uniq.setdefault(value_key(v), v)

            if len(uniq) <= 1:
                # eindeutig -> automatisch
                chosen_per_type[tid_int] = list(uniq.values())[0]
            else:
                # Konflikt -> Auswahldialog je Typ
                te = doc.GetElement(slot["tid"])
                type_name = "{}".format(
                    DB.Element.Name.__get__(te) if te else "Typ {}".format(tid_int))
                # Optionen mit Anzahl
                counts = {}
                for v in vals:
                    counts[value_key(v)] = counts.get(value_key(v), 0) + 1
                opt_map = {}
                for k, v in uniq.items():
                    n = counts.get(k, 0)
                    lbl = "{}   ({} Exemplar(e))".format(v["disp"], n)
                    base = lbl
                    c = 2
                    while lbl in opt_map:
                        lbl = "{} #{}".format(base, c); c += 1
                    opt_map[lbl] = v

                pick = forms.SelectFromList.show(
                    sorted(opt_map.keys()),
                    title="Wert fuer Typ / Value for type:  '{}'  ({})".format(
                        type_name, name),
                    multiselect=False,
                    button_name="Diesen Wert in den Typ / Use for Type"
                )
                if not pick:
                    # Abbruch dieses Typs -> nichts uebertragen
                    continue
                chosen_per_type[tid_int] = opt_map[pick]

        entry["type_values"] = chosen_per_type

    else:
        # ----- TYP -> EXEMPLAR : Typ-Wert je Typ merken -----
        type_vals = {}      # tid.IntegerValue -> v
        for te in collect_types(cats):
            p = te.LookupParameter(name)
            v = get_value(p)
            if v is not None:
                type_vals[te.Id.IntegerValue] = v
        entry["type_values"] = type_vals

    plan.append(entry)

if not plan:
    forms.alert("Nichts umzuschalten (alle bereits im Zielzustand).\n"
                "Nothing to switch.", title="Fertig / Done", exitscript=True)

# ===========================================================
#  SCHRITT 5 — Bestaetigung
# ===========================================================
lines = []
for e in plan:
    frm = "Exemplar" if e["from_instance"] else "Typ"
    to = "Exemplar" if e["target_instance"] else "Typ"
    lines.append("  {} : {} -> {}".format(e["name"], frm, to))

confirmed = forms.alert(
    "Folgende Parameter werden umgeschaltet und die Werte uebertragen:\n\n"
    "{}\n\n"
    "Die Bindung wird intern neu angelegt; die ausgelesenen Werte werden "
    "danach zurueckgeschrieben.\n\n"
    "Fortfahren? / Continue?".format("\n".join(lines)),
    title="Bestaetigung / Confirm", yes=True, no=True)
if not confirmed:
    script.exit()

# ===========================================================
#  SCHRITT 6 — Umschalten + Werte zurueckschreiben (1 Transaction)
# ===========================================================
ok_list = []
val_ok = 0
val_skip = 0
err_list = []

with revit.Transaction("Bindung umschalten + Werte uebertragen"):
    for e in plan:
        name = e["name"]
        key_def = e["key_def"]
        try:
            cset = copy_category_set(e["binding"])
            grp = get_param_group(key_def)
            new_binding = (app.Create.NewInstanceBinding(cset)
                           if e["target_instance"]
                           else app.Create.NewTypeBinding(cset))

            doc.ParameterBindings.Remove(key_def)
            if grp is not None:
                doc.ParameterBindings.Insert(key_def, new_binding, grp)
            else:
                doc.ParameterBindings.Insert(key_def, new_binding)

            frm = "Exemplar" if e["from_instance"] else "Typ"
            to = "Exemplar" if e["target_instance"] else "Typ"
            ok_list.append("{} : {} -> {}".format(name, frm, to))

            # --- Werte zurueckschreiben ---
            if e["target_instance"]:
                # TYP -> EXEMPLAR : Typ-Wert auf jedes Exemplar
                tvals = e["type_values"]
                for inst in collect_instances(e["cats"]):
                    tid = inst.GetTypeId()
                    if tid is None:
                        continue
                    v = tvals.get(tid.IntegerValue)
                    if v is None:
                        continue
                    p = inst.LookupParameter(name)
                    if set_value(p, v):
                        val_ok += 1
                    else:
                        val_skip += 1
            else:
                # EXEMPLAR -> TYP : gewaehlten Wert auf den Typ
                tvals = e["type_values"]
                for te in collect_types(e["cats"]):
                    v = tvals.get(te.Id.IntegerValue)
                    if v is None:
                        continue
                    p = te.LookupParameter(name)
                    if set_value(p, v):
                        val_ok += 1
                    else:
                        val_skip += 1

        except Exception as ex:
            err_list.append("{} - {}".format(name, str(ex)))

# ===========================================================
#  SCHRITT 7 — Ergebnis
# ===========================================================
output = script.get_output()
output.set_title("Bindung + Werte (Revit {})".format(revit_version))
output.print_html("<h2>Ergebnis / Results</h2>")

output.print_html("<h3 style='color:green'>Umgeschaltet / Switched: {}</h3>".format(
    len(ok_list)))
for n in ok_list:
    output.print_html("<p style='color:green'>&#8644; {}</p>".format(n))

output.print_html(
    "<h3 style='color:#1e88e5'>Werte uebertragen / Values written: {}"
    "</h3>".format(val_ok))
if val_skip:
    output.print_html(
        "<p style='color:gray'>Uebersprungen (leer / schreibgeschuetzt): "
        "{}</p>".format(val_skip))

if err_list:
    output.print_html("<h3 style='color:red'>Fehler / Errors: {}</h3>".format(
        len(err_list)))
    for n in err_list:
        output.print_html("<p style='color:red'>&#10007; {}</p>".format(n))

output.print_html(
    "<p style='color:gray'>Tipp: Erscheinen Instanz-Parameter danach nicht "
    "in der Eigenschaftenpalette, Element neu selektieren. Instanz-Parameter "
    "in der Gruppe 'IFC-Parameter' werden teils nur eingeschraenkt "
    "angezeigt.</p>")

"""EWG3.0 Formulator Lite — input/output workbook conventions."""
from __future__ import annotations

import math
from typing import Dict, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from .ewg.gfc import EXAMPLE_WASTES
from .ewg.optimize import FULL_OXIDES

# Oxides that map to glass columns.  The engine's waste row is built over
# FULL_OXIDES; keys outside it (lanthanide/actinide lumps, analyte artifacts)
# are NOT silently dropped by the lite reader — they are passed through
# untouched because the full engine counts every key in the normalization
# denominator (`optimize._rows`: s = sum over ALL waste keys) while giving
# only FULL_OXIDES keys a glass column.  Filtering here would change results.
#
# "Others" is the engine's residual basis concept; accepted only from the
# engine path, rejected from user input with an explicit message.
WASTE_OXIDES = [o for o in FULL_OXIDES if o != "Others"]

_BOLD = Font(bold=True)
_HDR_FILL = PatternFill("solid", fgColor="D9E1F2")
_NOTE_FONT = Font(italic=True, size=9)


# ---------------------------------------------------------------- template
def build_template(path: str) -> str:
    """Write the default Excel input template.

    Sheets:
      README        — instructions
      Waste_Vector  — Oxide | Mass fraction rows (pre-filled with the report's
                      Table 3.4 "High Al2O3" example waste; overwrite freely).
                      Any positive scale is accepted — the engine normalizes.
      Settings      — waste mass (kg) for absolute outputs + confidence level.
    """
    wb = Workbook()

    ws0 = wb.active
    ws0.title = "README"
    lines = [
        "EWG3.0 Formulator Lite — waste vector input template",
        "",
        "1. On the 'Waste_Vector' sheet, enter the waste composition as oxide mass",
        "   fractions (or any mass unit — values are normalized to sum 1).",
        "   The pre-filled example is the PNNL-38953 Table 3.4 'High Al2O3' waste.",
        "   Only oxides listed in the sheet are recognized; delete or zero unused rows.",
        "",
        "2. On the 'Settings' sheet choose the waste mass (kg) used for absolute",
        "   GFC / glass quantities in the output workbook, and the prediction-",
        "   uncertainty confidence level (0.90 recommended for formulation;",
        "   0.99 reproduces the report's Table 3.7 examples).",
        "",
        "3. Run from a terminal:  ewg30-lite run <this_file.xlsx> -o <output.xlsx>",
        "",
        "The output workbook mirrors the full application's Excel export:",
        "Summary / GFC_Recipe / Glass_Composition / Glass_Properties / Constraints.",
    ]
    for i, ln in enumerate(lines, 1):
        c = ws0.cell(row=i, column=1, value=ln)
        if i == 1:
            c.font = Font(bold=True, size=13)
        elif ln.startswith("   "):
            c.font = _NOTE_FONT

    ws = wb.create_sheet("Waste_Vector")
    for j, h in enumerate(["Oxide", "Mass fraction"], 1):
        c = ws.cell(row=1, column=j, value=h)
        c.font = _BOLD
        c.fill = _HDR_FILL
        c.alignment = Alignment(horizontal="center")
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 16
    ws.freeze_panes = "A2"
    example = EXAMPLE_WASTES["High Al2O3"]
    # Prefill recognized oxide rows; skip non-oxide artifact keys (spaces).
    prefilled = [o for o in WASTE_OXIDES if example.get(o)]
    r = 2
    for ox in prefilled:
        ws.cell(row=r, column=1, value=ox)
        ws.cell(row=r, column=2, value=float(example[ox]))
        r += 1
    for ox in WASTE_OXIDES:
        if ox not in prefilled:
            ws.cell(row=r, column=1, value=ox)
            r += 1

    ws2 = wb.create_sheet("Settings")
    for j, h in enumerate(["Setting", "Value", "Units"], 1):
        c = ws2.cell(row=1, column=j, value=h)
        c.font = _BOLD
        c.fill = _HDR_FILL
    ws2.column_dimensions["A"].width = 22
    settings = [
        ("waste_mass_kg", 1000.0, "kg", "Total waste oxide mass; scales absolute output quantities. Ignored if 0."),
        ("confidence", 0.90, "-", "Upred confidence level: 0.90 recommended for formulation, 0.99 for the Table 3.7 examples."),
    ]
    for i, (k, v, u, note) in enumerate(settings, 2):
        ws2.cell(row=i, column=1, value=k).font = _BOLD
        ws2.cell(row=i, column=2, value=v)
        ws2.cell(row=i, column=3, value=u)
        ws2.cell(row=i, column=4, value=note).font = _NOTE_FONT
        ws2.column_dimensions["D"].width = 80

    wb.save(path)
    return path


# ---------------------------------------------------------------- read
def read_template(path: str) -> Tuple[Dict[str, float], Dict[str, float]]:
    """Read a waste-vector template → (waste dict, settings dict).

    Raises ValueError with a user-facing message on any malformed input.
    """
    wb = load_workbook(path, data_only=True)
    if "Waste_Vector" not in wb.sheetnames:
        raise ValueError("Input workbook has no 'Waste_Vector' sheet. "
                         "Generate the default template with: ewg30-lite template")
    ws = wb["Waste_Vector"]
    waste: Dict[str, float] = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        ox, val = row[0], row[1] if len(row) > 1 else None
        if ox is None or str(ox).strip() == "":
            continue
        ox = str(ox).strip()
        if ox == "Others":
            raise ValueError("'Others' is the engine's internal residual basis "
                             "concept and cannot be a waste constituent — remove "
                             "that row.")
        if val is None or val == "":
            continue
        try:
            v = float(val)
        except (TypeError, ValueError):
            raise ValueError(f"Non-numeric mass fraction for '{ox}': {val!r}")
        if v < 0:
            raise ValueError(f"Negative mass fraction for '{ox}'.")
        if v > 0:
            waste[ox] = v
    tot = sum(waste.values())
    if tot <= 0:
        raise ValueError("Waste vector is empty — enter at least one positive "
                         "mass fraction on the Waste_Vector sheet.")
    waste = {k: v for k, v in waste.items()}  # engine normalizes internally
    settings = dict(waste_mass_kg=1000.0, confidence=0.90)
    if "Settings" in wb.sheetnames:
        for row in wb["Settings"].iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None:
                continue
            k = str(row[0]).strip()
            if k == "waste_mass_kg" and row[1] is not None:
                settings["waste_mass_kg"] = max(float(row[1]), 0.0)
            elif k == "confidence" and row[1] is not None:
                c = float(row[1])
                if not (0.0 < c < 1.0):
                    raise ValueError(f"confidence must be in (0, 1), got {c}")
                settings["confidence"] = c
    return waste, settings


# ---------------------------------------------------------------- output
def _ws_title(ws, text, note=""):
    ws.cell(row=1, column=1, value=text).font = Font(bold=True, size=13)
    r = 2
    if note:
        ws.cell(row=r, column=1, value=note).font = _NOTE_FONT
        r += 1
    return r + 1


def _kv_table(ws, r, rows, headers=("Quantity", "Value", "")):
    for j, h in enumerate(headers, 1):
        c = ws.cell(row=r, column=j, value=h)
        c.font = _BOLD
        c.fill = _HDR_FILL
    r += 1
    for label, val, extra in rows:
        ws.cell(row=r, column=1, value=label)
        if isinstance(val, float):
            ws.cell(row=r, column=2, value=val).number_format = "General"
        else:
            ws.cell(row=r, column=2, value=val)
        ws.cell(row=r, column=3, value=extra)
        r += 1
    return r


def write_output(res: Dict, waste: Dict[str, float], settings: Dict[str, float],
                out_path: str) -> str:
    """Write the result workbook — the lite counterpart of the full
    application's Excel export (Campaign-detail + Glass Formers & Glass +
    Glass Properties & Constraints information, single formulation)."""
    if waste_mass_kg := settings.get("waste_mass_kg", 0.0):
        glass_kg = waste_mass_kg / res["waste_loading"]
    else:
        glass_kg = None
    conf = settings.get("confidence", 0.90)
    ev = res["evaluation"]
    wb = Workbook()

    # -------- Summary
    ws = wb.active
    ws.title = "Summary"
    r = _ws_title(
        ws, "EWG3.0 Glass Formulation — Summary",
        "Optimizer: maximize waste loading subject to the PNNL-38953 (EWG3.0) "
        "Table 3.1 constraint set, identical to the full application's "
        "ewg.optimize.formulate.")
    r = _kv_table(ws, r, [
        ("Optimizer success", "PASS" if res["success"] else "FAIL", res["message"]),
        ("Waste loading (max)", res["waste_loading"] * 100.0, "wt%"),
        ("Waste loading (physical: waste oxide / glass)", res["wl_physical"] * 100.0, "wt%"),
        ("Glass mass", glass_kg, "kg" if glass_kg else "(set waste_mass_kg > 0 in Settings)"),
        ("Waste mass (input)", waste_mass_kg or None, "kg"),
        ("Total GFC demand", sum(res["gfc_raw_mass_fraction"].values()) * glass_kg if glass_kg else None, "kg (as-fired)"),
        ("Glass input basis total", res["glass_input_total"], "-"),
        ("Upred confidence", conf, "-"),
        ("All glass constraints pass", "PASS" if ev.all_pass else "FAIL", ""),
    ], headers=("Quantity", "Value", "Units / note"))
    ws.column_dimensions["A"].width = 44
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 60

    # -------- Input waste
    ws = wb.create_sheet("Input_Waste")
    r = _ws_title(ws, "Waste vector (as entered, normalized)")
    tot = sum(waste.values())
    ws.append(["Oxide", "Mass fraction", "wt%"])
    for j in range(1, 4):
        ws.cell(row=r, column=j).font = _BOLD
        ws.cell(row=r, column=j).fill = _HDR_FILL
    r += 1
    for ox in sorted(waste, key=lambda o: -waste[o]):
        mf = waste[ox] / tot
        ws.cell(row=r, column=1, value=ox)
        ws.cell(row=r, column=2, value=mf).number_format = "0.000000"
        ws.cell(row=r, column=3, value=mf * 100.0).number_format = "0.0000"
        r += 1

    # -------- GFC recipe (mirrors 'Glass Formers & Glass')
    from .ewg.gfc import GFCS, GFC_NAMES
    ws = wb.create_sheet("GFC_Recipe")
    r = _ws_title(
        ws, "Glass-forming chemicals (GFC) — formulation",
        "x_k = oxide mass fraction of glass contributed by GFC k (optimizer "
        "decision variable).  Raw (as-fired) mass per unit glass = x_k / "
        "oxide_yield_k.  Absolute kg = raw fraction × glass mass — the same "
        "convention as the full application's Glass Formers & Glass export.")
    for j, h in enumerate(["GFC", "Glass oxide fraction x_k", "Raw mass fraction",
                           "Absolute kg"], 1):
        c = ws.cell(row=r, column=j, value=h)
        c.font = _BOLD
        c.fill = _HDR_FILL
    r += 1
    ordered = [n for n in GFC_NAMES if n in res["x"]]
    for name in ordered:
        xk = res["x"].get(name, 0.0)
        ws.cell(row=r, column=1, value=name)
        ws.cell(row=r, column=2, value=xk).number_format = "0.000000"
        raw = res["gfc_raw_mass_fraction"].get(name)
        if raw is not None:
            ws.cell(row=r, column=3, value=raw).number_format = "0.000000"
            if glass_kg:
                ws.cell(row=r, column=4, value=raw * glass_kg).number_format = "0.000"
        r += 1
    if glass_kg:
        ws.cell(row=r, column=1, value="Total GFC (kg as-fired)").font = _BOLD
        ws.cell(row=r, column=4, value=sum(res["gfc_raw_mass_fraction"].values()) * glass_kg)
    r += 2
    r = _kv_table(ws, r, [("Total glass (kg)", glass_kg, "")],
                  headers=("Quantity", "Value", "Units"))
    ws.column_dimensions["A"].width = 30

    # -------- Glass composition (mirrors Glass composition sections)
    ws = wb.create_sheet("Glass_Composition")
    r = _ws_title(
        ws, "Glass composition (full oxide mass fractions)",
        "Sums < 1.0 by the volatile fraction — H2O/halogens leave the melt "
        "(physical closure, not an error).")
    for j, h in enumerate(["Oxide", "Mass fraction", "wt%"], 1):
        c = ws.cell(row=r, column=j, value=h)
        c.font = _BOLD
        c.fill = _HDR_FILL
    r += 1
    glass = res["glass"]
    for ox in sorted(glass, key=lambda o: -glass[o]):
        if glass[ox] <= 0:
            continue
        ws.cell(row=r, column=1, value=ox)
        ws.cell(row=r, column=2, value=float(glass[ox])).number_format = "0.000000"
        ws.cell(row=r, column=3, value=float(glass[ox]) * 100.0).number_format = "0.0000"
        r += 1
    ws.cell(row=r, column=1, value="TOTAL").font = _BOLD
    ws.cell(row=r, column=2, value=float(sum(glass.values()))).number_format = "0.000000"

    # -------- Glass properties (mirrors 'Glass Properties & Constraints')
    ws = wb.create_sheet("Glass_Properties")
    r = _ws_title(
        ws, "Glass Properties (EWG3.0 property models)",
        "Each property on its own published basis; Upred = prediction-uncertainty "
        "margin (Appendix C covariances).  Same rows the full application's "
        "Glass Properties export tab carries.")
    for j, h in enumerate(["Property", "Value", "Units"], 1):
        c = ws.cell(row=r, column=j, value=h)
        c.font = _BOLD
        c.fill = _HDR_FILL
    r += 1

    def prow(label, val, unit=""):
        nonlocal r
        ws.cell(row=r, column=1, value=label)
        if isinstance(val, float):
            ws.cell(row=r, column=2, value=val).number_format = "0.000000"
        else:
            ws.cell(row=r, column=2, value=val)
        ws.cell(row=r, column=3, value=unit)
        r += 1

    def uprow(key):
        nonlocal r
        u = (ev.upred or {}).get(key)
        if u is not None:
            prow("  Upred", float(u))

    up = ev.upred or {}
    prow("Waste loading (wt%)", res["waste_loading"] * 100.0, "wt%")
    prow("Viscosity @1150 °C", ev.eta1150, "Pa·s"); uprow("eta1150")
    prow("Viscosity @1100 °C", ev.eta1100, "Pa·s")
    prow("Conductivity @1100 °C", ev.ec1100, "S/m")
    prow("Conductivity @1200 °C", ev.ec1200, "S/m")
    prow("PCT Ave ln[NCAve] (model value)", ev.ave_ln_nl, "ln(g/L)")
    prow("PCT NCAve", ev.nca_g_l, "g/L"); uprow("pct")
    prow("Sulfate solubility wSO3-3TS", ev.wso3_3ts, "wt%"); uprow("wso3")
    prow("wSO3-MT = 3TS − 0.33", ev.wso3_3ts - 0.33, "wt%")
    prow("Glass SO3", float(glass.get("SO3", 0.0)) * 100.0, "wt%")
    prow("K-3 corrosion k1208", math.exp(ev.ln_k1208), "in"); uprow("k1208")
    prow("K-3 ln[k1208] (model value)", ev.ln_k1208, "ln(in)")
    prow("Nepheline distance P (NP_P)", ev.np_p, "-"); uprow("np_p")
    prow("Immiscibility N_NaLi", ev.NNaLi, "-")
    prow("NALK = Na2O + 0.66 K2O + 2.07 Li2O", ev.NALK, "-")
    prow("NSiAl = SiO2 + 1.697 Al2O3", ev.NSiAl, "-")
    prow("NSiBi = SiO2 + 1.697 Bi2O3", ev.NSiBi, "-")
    prow("TLC limit P (gated, ≤1050 °C)", ev.tl_p, "°C"); uprow("tl_p")
    prow("TLC limit Zr (gated, ≤1050 °C)", ev.tl_zr, "°C"); uprow("tl_zr")
    prow("C950 spinel", ev.c950, "vol%"); uprow("c950")
    prow("TCLP ln[rB] (model value)", ev.ln_tclp_rB, "ln(mg/L)"); uprow("tclp")
    for el, key in (("Cd", "CCd"), ("Cr", "CCr"), ("V", "CV")):
        raw = (ev._tclp_raw or {}).get(el)
        if raw is not None:
            prow(f"TCLP C{el} (raw Upred-free)", float(raw), "mg/L")
    prow("BR glass fluoride gF", ev.gF_br, "wt% (BR basis)")
    prow("BR glass sulfate gSO3", ev.gSO3_br, "wt% (BR basis)")
    ws.column_dimensions["A"].width = 44

    # -------- Constraints (value, limit, pass/fail)
    ws = wb.create_sheet("Constraints")
    r = _ws_title(
        ws, "Property constraints (Upred margins included)",
        "value_used / limit / pass per constraint — the optimizer's own "
        "constraint evaluation, mirroring the full application export.")
    for j, h in enumerate(["Constraint", "Key", "Units", "Value",
                           "Limit", "Pass/Fail"], 1):
        c = ws.cell(row=r, column=j, value=h)
        c.font = _BOLD
        c.fill = _HDR_FILL
    r += 1
    cons = ev.constraints or {}
    for ck in sorted(cons):
        v, ok = cons[ck]
        limit = _cons_limit(ck)
        ws.cell(row=r, column=1, value=_CONS_HUMAN.get(ck, ck))
        ws.cell(row=r, column=2, value=ck)
        ws.cell(row=r, column=3, value=_CONS_UNITS.get(ck, ""))
        if isinstance(v, (int, float)):
            ws.cell(row=r, column=4, value=float(v)).number_format = "0.000000"
        else:
            ws.cell(row=r, column=4, value=str(v))
        if isinstance(limit, (int, float)):
            ws.cell(row=r, column=5, value=float(limit)).number_format = "0.000000"
        else:
            ws.cell(row=r, column=5, value=str(limit) if limit is not None else "per-value")
        ws.cell(row=r, column=6, value="PASS" if ok else "FAIL")
        r += 1
    c = ws.cell(row=r, column=1, value="ALL CONSTRAINTS")
    c.font = _BOLD
    ws.cell(row=r, column=6, value="PASS" if ev.all_pass else "FAIL").font = _BOLD
    for col, w in (("A", 40), ("B", 22), ("C", 10), ("D", 14), ("E", 14), ("F", 10)):
        ws.column_dimensions[col].width = w

    wb.save(out_path)
    return out_path


_CONS_UNITS = {
    "PCT_NL_g_m2": "g/m²", "PCT_NCAve_g_L": "g/L", "eta1150_Pa.s": "Pa·s",
    "eta1150_up": "Pa·s", "eta1150_lo": "Pa·s", "eta1100_Pa.s": "Pa·s",
    "eta1100_max": "Pa·s", "ec1100_S_cm": "S/cm", "ec1100_min": "S/m",
    "ec1200_S_cm": "S/cm", "ec1200_max": "S/m", "wSO3_MT_wt%": "wt%",
    "k1208_in": "in", "kneck_1208_in": "in", "p_phosphate": "-",
    "NP_P": "-", "N_NaLi": "-", "NALK_bounds": "-", "NSiAl_bounds": "-",
    "NSiBi_bounds": "-", "TCLP_rB_mg_L": "mg/L", "TCLP_CCd": "mg/L",
    "TCLP_CCr": "mg/L", "TCLP_CV": "mg/L", "TL-P_1050C": "°C",
    "TL-Zr_1050C": "°C", "C950_vol%": "vol%", "gF_br_2.5wt%": "wt%",
    "nonspinel_ThZr": "wt%",
}


_CONS_HUMAN = {
    "PCT_NCAve_g_L": "PCT NCAve ≤ 12.88 g/L", "PCT_NL_g_m2": "PCT NL ≤ 12.88 (legacy key)",
    "eta1150_Pa.s": "Viscosity 2–8 Pa·s @1150 °C", "eta1150_up": "η1150 ≤ 8 Pa·s",
    "eta1150_lo": "η1150 ≥ 2 Pa·s", "eta1100_Pa.s": "η1100 < 15 Pa·s",
    "eta1100_max": "η1100 < 15 Pa·s", "ec1100_S_cm": "ε1100 ≥ 10 (legacy S/cm key)",
    "ec1100_min": "ε1100 ≥ 10 S/m", "ec1200_S_cm": "ε1200 ≤ 70 (legacy S/cm key)",
    "ec1200_max": "ε1200 ≤ 70 S/m", "wSO3_MT_wt%": "Glass SO3 ≤ wSO3-3TS − 0.33 wt%",
    "k1208_in": "K-3 kneck ≤ 0.025 in @1208", "kneck_1208_in": "K-3 kneck ≤ 0.025 in @1208",
    "p_phosphate": "Phosphate P(fail) (DEPRECATED EWG2.5)",
    "NP_P": "Nepheline P ≥ 0.031", "N_NaLi": "Immiscibility NNaLi ≥ 0.20",
    "NALK_bounds": "NAlk range", "NSiAl_bounds": "NSiAl range",
    "NSiBi_bounds": "NSiBi range", "TCLP_rB_mg_L": "TCLP rB proxy",
    "TCLP_CCd": "TCLP Cd ≤ 0.48 mg/L", "TCLP_CCr": "TCLP Cr ≤ 4.95 mg/L",
    "TCLP_CV": "TCLP V ≤ 16.9 mg/L", "TL-P_1050C": "TLC P ≤ 1050 °C (gated)",
    "TL-Zr_1050C": "TLC Zr ≤ 1050 °C (gated)", "C950_vol%": "C950 spinel ≤ 1.0 vol%",
    "gF_br_2.5wt%": "BR fluoride ≤ 2.5 wt%", "nonspinel_ThZr": "ThO2+ZrO2 < 13 wt%",
}


def _cons_limit(ck: str):
    """Published limit for a constraint key — replicated verbatim from the full
    application's ewg/export_excel.py:_cons_limit (EWG3.0 Table 3.1/3.2)."""
    from .ewg import properties as P
    if ck == "PCT_NL_g_m2":   return 12.88
    if ck == "PCT_NCAve_g_L": return P.PCT_LIMIT_G_L
    if ck == "eta1150_Pa.s":  return "2.0 - 8.0"
    if ck == "eta1150_up":    return P.ETA1150_RANGE[1]
    if ck == "eta1150_lo":    return P.ETA1150_RANGE[0]
    if ck == "eta1100_Pa.s":  return P.ETA1100_MAX
    if ck == "eta1100_max":   return P.ETA1100_MAX
    if ck == "ec1100_S_cm":   return P.EC1100_MIN / 10.0
    if ck == "ec1100_min":    return P.EC1100_MIN
    if ck == "ec1200_S_cm":   return P.EC1200_MAX / 10.0
    if ck == "ec1200_max":    return P.EC1200_MAX
    if ck == "wSO3_MT_wt%":   return None
    if ck == "k1208_in":      return 0.025
    if ck == "kneck_1208_in": return P.KNECK_LIMIT_IN
    if ck == "p_phosphate":   return None
    if ck == "NP_P":          return P.NPH_P_MIN
    if ck == "N_NaLi":        return P.IMMIN_MIN
    if ck == "TCLP_rB_mg_L":  return None
    if ck == "TCLP_CCd":      return 0.48
    if ck == "TCLP_CCr":      return 4.95
    if ck == "TCLP_CV":       return 16.9
    if ck == "TL-P_1050C":    return P.TL_LIMIT_C
    if ck == "TL-Zr_1050C":   return P.TL_LIMIT_C
    if ck == "C950_vol%":     return P.C950_MAX
    if ck == "gF_br_2.5wt%":  return P.F_BR_MAX
    if ck == "nonspinel_ThZr": return P.THO2_ZRO2_MAX
    if ck == "NALK_bounds":   return f"{P.NALK_RANGE[0]} - {P.NALK_RANGE[1]}"
    if ck == "NSiAl_bounds":  return f"{P.NSIAL_RANGE[0]} - {P.NSIAL_RANGE[1]}"
    if ck == "NSiBi_bounds":  return f"{P.NSIBI_RANGE[0]} - {P.NSIBI_RANGE[1]}"
    if ck.startswith("MV_"):  return None
    return None

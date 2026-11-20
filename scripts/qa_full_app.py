#!/usr/bin/env python3
"""QA parity harness: lite CLI vs the FULL EWG3.0 application.

Each import context (full-repo `ewg` vs vendored lite engine) runs in its own
fresh subprocess — formulating after mutating module registries in one process
disturbs the engine and produces unphysical results.  The parent process only
compares JSON surfaces.

Checks:
  1. Vendored engine files byte-identical to the full application's, and
     vendored commit recorded.
  2. formulate() bit-identical parity on the report's Table 3.4 example wastes
     (full-repo worker vs lite worker, one subprocess each).
  3. Full application's LIVE campaign path (build_farm_mission with
     scenario_tag -> blend_vector -> formulate — the exact call chain the
     Streamlit app makes) compared bit-exactly to the lite engine and to a
     real `ewg30-lite run` subprocess on the same waste vector, including the
     result workbook's glass composition.

Usage:  python scripts/qa_full_app.py <full_app_repo_dir> [scenario_tag]
        (default scenario_tag: 135.120.150)
        python scripts/qa_full_app.py [--worker-full|--worker-lite] <dir> <json> <vectors-json>
"""
from __future__ import annotations
import hashlib
import json
import os
import subprocess
import sys

LITE_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LITE_PARENT = os.path.dirname(LITE_ROOT)


def _surface(res) -> dict:
    return {
        "success": res["success"], "message": res["message"],
        "waste_loading": res["waste_loading"], "wl_physical": res["wl_physical"],
        "glass_input_total": res["glass_input_total"],
        "x": res["x"], "glass": res["glass"],
        "gfc_raw_mass_fraction": res["gfc_raw_mass_fraction"],
        "all_pass": res["evaluation"].all_pass,
        "constraints": {k: [float(v), bool(ok)]
                        for k, (v, ok) in res["evaluation"].constraints.items()},
        "eta1150": res["evaluation"].eta1150, "eta1100": res["evaluation"].eta1100,
        "ec1100": res["evaluation"].ec1100, "ec1200": res["evaluation"].ec1200,
        "nca_g_l": res["evaluation"].nca_g_l, "ave_ln_nl": res["evaluation"].ave_ln_nl,
        "wso3_3ts": res["evaluation"].wso3_3ts, "ln_k1208": res["evaluation"].ln_k1208,
        "np_p": res["evaluation"].np_p, "NNaLi": res["evaluation"].NNaLi,
        "upred": res["evaluation"].upred,
    }


def _worker(full, out_json, vectors_json):
    """Subprocess: formulate with the FULL application's engine."""
    sys.path.insert(0, full)
    from ewg.optimize import formulate
    vectors = json.load(open(vectors_json))
    key = json.dumps(vectors["__order__"])
    surfaces = {}
    for name in vectors["__order__"]:
        surfaces[name] = _surface(formulate(dict(vectors[name])))
    json.dump(surfaces, open(out_json, "w"))


def _worker_lite(out_json, vectors_json):
    """Subprocess: formulate with the vendored lite engine."""
    sys.path.insert(0, LITE_PARENT)
    from ewg30lite.ewg import formulate
    vectors = json.load(open(vectors_json))
    surfaces = {}
    for name in vectors["__order__"]:
        surfaces[name] = _surface(formulate(dict(vectors[name])))
    json.dump(surfaces, open(out_json, "w"))


def _run_worker(script, *args):
    r = subprocess.run([sys.executable, script, *args],
                       capture_output=True, text=True, timeout=3600)
    if r.returncode != 0:
        sys.exit(f"worker failed:\n{r.stdout}\n{r.stderr}")
    return r


def _dump_vectors(vectors: dict, path: str) -> None:
    json.dump({"__order__": list(vectors), **vectors}, open(path, "w"))


def _diff(a, b):
    """Return list of (key-path, fa, fb) differences (exact float compare)."""
    out = []
    def walk(x, y, p=""):
        if isinstance(x, dict) and isinstance(y, dict):
            for k in sorted(set(x) | set(y)):
                walk(x.get(k), y.get(k), p + "." + str(k))
        elif x != y:
            out.append((p, x, y))
    walk(a, b)
    return out


def check_engine_bytes(full: str) -> None:
    print("== Check 1: engine file byte-identity")
    mods = ["optimize.py", "properties.py", "gfc.py", "tensor_backend.py"]
    for m in mods:
        a = open(os.path.join(LITE_ROOT, "ewg30lite", "ewg", m), "rb").read()
        b = open(os.path.join(full, "ewg", m), "rb").read()
        ha, hb = hashlib.sha256(a).hexdigest(), hashlib.sha256(b).hexdigest()
        print(f"   {'OK ' if ha == hb else 'DRIFT'} {m}: {ha[:12]}")
        assert ha == hb, f"engine drift in {m}"
    vendored = open(os.path.join(LITE_ROOT, "VENDORED_SOURCE_COMMIT.txt")).read().strip()
    head = subprocess.run(["git", "rev-parse", "HEAD"], cwd=full,
                          capture_output=True, text=True).stdout.strip()
    print(f"   {'OK ' if vendored == head else 'NOTE'} vendored commit {vendored[:12]} "
          f"vs full-repo HEAD {head[:12]}")
    if vendored != head:
        print("   (engine bytes still identical; note the full repo has moved)")
    data_ok = all(
        open(os.path.join(LITE_ROOT, "ewg30lite", "ewg", "data", "ewg30", f)).read()
        == open(os.path.join(full, "ewg", "data", "ewg30", f)).read()
        for f in os.listdir(os.path.join(LITE_ROOT, "ewg30lite", "ewg", "data", "ewg30")))
    print(f"   {'OK ' if data_ok else 'DRIFT'} coefficient data (ewg30 JSONs)")
    assert data_ok


def check_formulate_parity(full: str, vectors: dict, td: str) -> None:
    print(f"== Check 2: formulate() bit-identical parity on {len(vectors)} wastes")
    vj = os.path.join(td, "vectors.json")
    _dump_vectors(vectors, vj)
    fo, lo = os.path.join(td, "full.json"), os.path.join(td, "lite.json")
    _run_worker(__file__, "--worker-full", full, fo, vj)
    _run_worker(__file__, "--worker-lite", lo, vj)
    S_full = json.load(open(fo))
    S_lite = json.load(open(lo))
    for name in vectors:
        d = _diff(S_full[name], S_lite[name])
        ok = not d
        print(f"   {'OK ' if ok else 'DIFF'} {name}: "
              f"WL {S_lite[name]['waste_loading']*100:.4f} wt% "
              f"all_pass {S_lite[name]['all_pass']}, "
              f"{'bit-identical' if ok else 'MISMATCH'}")
        if not ok:
            json.dump({name: {"full": S_full[name], "lite": S_lite[name]}},
                      open("/tmp/qa_parity_failure.json", "w"), indent=1)
            sys.exit(f"   PARITY FAILURE on {name} — /tmp/qa_parity_failure.json "
                     f"({len(d)} differing fields, first: {d[0]})")


def check_full_app_campaign(full: str, tag: str, td: str) -> None:
    """Compare against the FULL APPLICATION's live campaign path — the same
    calls the Streamlit app makes (build_farm_mission + blend_vector)."""
    print(f"== Check 3: full application live path, scenario {tag}")
    camp_json = os.path.join(td, "campaign.json")
    wj = os.path.join(td, "campaign_vector.json")
    code = (
        "import sys, json, subprocess\n"
        f"sys.path.insert(0, {full!r})\n"
        # Rebuild this scenario's layer vector files exactly as the app does
        # (the app re-runs this subprocess each session; blend_vector then
        # reads the shared layer_vectors.json).
        f"subprocess.run([sys.executable, {full + '/build_layers_app_d.py'!r}, {tag!r}, '0.90'],\n"
        f"                cwd={full!r}, check=True, capture_output=True)\n"
        "from ewg.campaign_builder import build_farm_mission\n"
        "from ewg.optimize import formulate\n"
        "from ewg.farm_mission import evaluate_campaign_recipe\n"
        f"mission = build_farm_mission(wac_scalings=None, max_campaigns=250, scenario_tag={tag!r})\n"
        "camps = mission['campaigns']\n"
        "camp = camps[0]\n"
        "ev = evaluate_campaign_recipe(camp['recipe'])   # the app's own live path\n"
        "waste = {k: float(v) for k, v in ev['waste'].items() if float(v)}\n"
        "res = _surface(formulate(dict(waste)))\n"
        "json.dump({'n_campaigns': len(camps), 'recipe_len': len(camp['recipe']),"
        " 'waste_kg': camp.get('waste_kg'), 'waste': waste, 'surface': res},\n"
        f"  open({camp_json!r}, 'w'))\n"
    )
    # make the surface helper available inside the worker
    code = _surface_src() + "\n" + code
    wr = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True,
                        timeout=3600)
    if wr.returncode != 0:
        sys.exit(f"full-app campaign worker failed:\n{wr.stdout}\n{wr.stderr}")
    camp = json.load(open(camp_json))
    waste = camp["waste"]
    print(f"   mission built: {camp['n_campaigns']} campaigns; campaign 1 "
          f"({camp['recipe_len']} layers, waste_kg {camp['waste_kg']:.0f}, "
          f"{len(waste)} oxide keys)")

    # lite side: engine parity on the SAME live waste vector
    _dump_vectors({"campaign1": waste}, wj)
    lo = os.path.join(td, "campaign_lite.json")
    _run_worker(__file__, "--worker-lite", lo, wj)
    S_lite = json.load(open(lo))["campaign1"]
    d = _diff(camp["surface"], S_lite)
    ok = not d
    print(f"   {'OK ' if ok else 'DIFF'} lite engine vs full app campaign 1: "
          f"WL {S_lite['waste_loading']*100:.4f} wt%, all_pass {S_lite['all_pass']} "
          f"— {'bit-identical' if ok else 'MISMATCH'}")
    if not ok:
        json.dump({"full": camp["surface"], "lite": S_lite},
                  open("/tmp/qa_campaign_failure.json", "w"), indent=1)
        sys.exit(f"   PARITY FAILURE — /tmp/qa_campaign_failure.json: {d[:3]}")

    # real CLI subprocess on this exact vector
    opath = os.path.join(td, "out.xlsx")
    tpath = os.path.join(td, "in.xlsx")
    sys.path.insert(0, LITE_PARENT)
    from ewg30lite.io_excel import build_template
    build_template(tpath)
    from openpyxl import load_workbook
    wb = load_workbook(tpath)
    ws = wb["Waste_Vector"]
    n = 2
    while ws.cell(row=n, column=1).value is not None:
        ws.cell(row=n, column=2).value = None
        n += 1
    for i, (ox, v) in enumerate(sorted(waste.items()), start=2):
        ws.cell(row=i, column=1, value=ox)
        ws.cell(row=i, column=2, value=v)
    wb.save(tpath)
    r_ = subprocess.run(["ewg30-lite", "run", tpath, "-o", opath],
                        capture_output=True, text=True, timeout=1800)
    assert r_.returncode == 0, f"lite CLI failed: {r_.stderr}"
    print("   OK  lite CLI subprocess run on the full-app campaign waste vector")
    wl_out = [ln for ln in r_.stdout.splitlines() if ln.startswith("Waste loading")][0]
    wl_cli = float(wl_out.split(":")[1].strip().split()[0])
    assert abs(wl_cli - S_lite["waste_loading"] * 100) < 5e-6, "CLI stdout WL mismatch"
    ob = load_workbook(opath)["Summary"]
    for row in ob.iter_rows(min_row=2, values_only=True):
        if row[0] == "Waste loading (max)":
            assert abs(row[1] - S_lite["waste_loading"] * 100) < 1e-9
            break
    gb = load_workbook(opath)["Glass_Composition"]
    wb_rows = {row[0]: row[1] for row in gb.iter_rows(min_row=2, values_only=True)}
    for ox, mf in S_lite["glass"].items():
        if mf:
            assert abs(wb_rows[ox] - mf) < 1e-12, f"glass {ox} mismatch"
    print("   OK  result workbook glass composition matches the full app exactly")


def _surface_src() -> str:
    import inspect
    return inspect.getsource(_surface)


def _indent(s: str) -> str:
    return "\n".join("    " + ln for ln in s.splitlines())


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--worker-full":
        _worker(sys.argv[2], sys.argv[3], sys.argv[4]); raise SystemExit(0)
    if len(sys.argv) > 1 and sys.argv[1] == "--worker-lite":
        _worker_lite(sys.argv[2], sys.argv[3]); raise SystemExit(0)

    full = os.path.abspath(sys.argv[1] if len(sys.argv) > 1
                           else "/Users/aspen_mb/Documents/Hanford_Tank_BBI/ewg_app")
    if not os.path.isdir(os.path.join(full, "ewg")):
        sys.exit(f"ERROR: {full} does not look like the full application repo (no ewg/)")
    tag = sys.argv[2] if len(sys.argv) > 2 else "135.120.150"

    import tempfile
    with tempfile.TemporaryDirectory() as td:
        check_engine_bytes(full)
        # example wastes come from a short full-repo worker (import isolation)
        ex_json = os.path.join(td, "examples.json")
        r = subprocess.run([sys.executable, "-c",
                            f"import sys, json; sys.path.insert(0, {full!r}); "
                            "from ewg.gfc import EXAMPLE_WASTES; "
                            "json.dump(EXAMPLE_WASTES, open(sys.argv[1], 'w'))", ex_json],
                           capture_output=True, text=True)
        assert r.returncode == 0, r.stderr
        vectors = json.load(open(ex_json))
        check_formulate_parity(full, vectors, td)
        check_full_app_campaign(full, tag, td)
    print("\nQA PARITY: ALL CHECKS PASSED — lite results identical to the full application.")

"""EWG3.0 Formulator Lite — command-line glass formulation.

Pure CLI (no GUI/TUI).  Two commands:

  ewg30-lite template [-o waste_vector_template.xlsx]
      Write the default Excel input template (waste vector + settings).

  ewg30-lite run <input.xlsx> [-o output.xlsx]
      Formulate the user's waste vector with the full application's EWG3.0
      algorithm (ewg.optimize.formulate, PNNL-38953) and write the Excel
      result workbook mirroring the full application's export content.
      A summary is also printed to stdout.
"""
from __future__ import annotations

import argparse
import sys
import os

from .io_excel import build_template, read_template, write_output

DEFAULT_TEMPLATE = "waste_vector_template.xlsx"
DEFAULT_OUTPUT = "glass_formulation_output.xlsx"


def _summary_text(res, settings) -> str:
    ev = res["evaluation"]
    wm = settings.get("waste_mass_kg", 0.0)
    lines = [
        "=== EWG3.0 glass formulation (lite CLI) ===",
        f"Optimizer success : {'PASS' if res['success'] else 'FAIL'} ({res['message']})",
        f"Waste loading     : {res['waste_loading']*100:.6f} wt% "
        f"(physical {res['wl_physical']*100:.6f} wt%)",
        f"Upred confidence  : {settings.get('confidence', 0.90)}",
    ]
    if wm:
        gk = wm / res["waste_loading"]
        gfc_total = sum(res["gfc_raw_mass_fraction"].values()) * gk
        lines.append(f"Glass mass        : {gk:,.1f} kg (from {wm:,.0f} kg waste)")
        lines.append(f"Total GFC demand  : {gfc_total:,.1f} kg (as-fired)")
    lines.append(f"All constraints    : {'PASS' if ev.all_pass else 'FAIL'}")
    gfc = res["gfc_raw_mass_fraction"]
    if gfc:
        lines.append("GFC recipe (raw mass fraction of feed):")
        for name in sorted(gfc, key=lambda k: -gfc[k]):
            lines.append(f"  {name:<14} {gfc[name]:.6f}")
    return "\n".join(lines)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        prog="ewg30-lite",
        description="EWG3.0 glass formulation on a user-defined waste vector "
                    "(CLI; same algorithm as the full application).")
    sub = ap.add_subparsers(dest="cmd", required=True)

    ap_t = sub.add_parser("template", help="write the default Excel input template")
    ap_t.add_argument("-o", "--output", default=DEFAULT_TEMPLATE)
    ap_t.add_argument("--extended", action="store_true",
                      help="write the 60-oxide extended template (all user "
                           "oxides; lanthanides lumped to RE2O3, diluents "
                           "noted per row)")

    ap_r = sub.add_parser("run", help="formulate the waste vector in an input template")
    ap_r.add_argument("input", help="filled-in waste-vector .xlsx template")
    ap_r.add_argument("-o", "--output", default=None,
                      help=f"result .xlsx (default: {DEFAULT_OUTPUT})")

    args = ap.parse_args(argv)

    if args.cmd == "template":
        build_template(args.output, extended=args.extended)
        print(f"Template written: {os.path.abspath(args.output)}"
              + (" (60-oxide extended)" if args.extended else ""))
        print("Open it, edit Waste_Vector + Settings, then: "
              f"ewg30-lite run {args.output}")
        return 0

    in_path = args.input
    if not os.path.exists(in_path):
        print(f"ERROR: input file not found: {os.path.abspath(in_path)}", file=sys.stderr)
        return 2
    out_path = args.output or (
        os.path.splitext(os.path.basename(in_path))[0] + "_formulation.xlsx")

    try:
        waste, settings = read_template(in_path)
    except ValueError as e:
        print(f"ERROR in input workbook: {e}", file=sys.stderr)
        return 2

    import time
    t0 = time.time()
    from .ewg import formulate   # vendored full-application engine
    res = formulate(dict(waste), confidence=settings.get("confidence", 0.90))
    elapsed = time.time() - t0

    write_output(res, waste, settings, out_path)
    print(_summary_text(res, settings))
    print(f"Solve time: {elapsed:.1f} s")
    print(f"Workbook written: {os.path.abspath(out_path)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

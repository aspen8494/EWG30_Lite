# EWG3.0 Formulator Lite

Lightweight command-line glass formulation for **EWG3.0 / PNNL-38953** (DFHLW
HLW waste glass). You supply a waste-oxide vector in Excel; the tool solves the
identical waste-loading maximization used by the full application
(`HWG_mission_planning`, `ewg.optimize.formulate`) and writes an Excel result
workbook carrying the same information the full application exports.

No GUI: pure CLI.

## Install

```bash
pip install -e .          # provides the `ewg30-lite` command
```

## Use

```bash
# 1) Generate the default input template
ewg30-lite template                 # → waste_vector_template.xlsx

# 2) Fill in Waste_Vector (oxide mass fractions) and Settings (waste mass kg,
#    Upred confidence) in Excel. The template ships pre-filled with the
#    report's Table 3.4 "High Al2O3" example waste.

# 3) Formulate
ewg30-lite run waste_vector_template.xlsx -o result.xlsx
```

The run summary prints to stdout; the result workbook has five sheets:

| Sheet | Content (mirrors the full application's export) |
|---|---|
| `Summary` | optimizer status, waste loading (optimizer + physical), glass & GFC mass for the entered waste mass, confidence |
| `Input_Waste` | entered waste vector, normalized (mass fraction + wt%) |
| `GFC_Recipe` | glass-forming chemicals: optimizer oxide fraction x_k, raw (as-fired) mass fraction, absolute kg |
| `Glass_Composition` | full oxide composition of the glass (mass fraction + wt%; sums < 1 by the volatile fraction — physical) |
| `Glass_Properties` | η/EC (VFT), PCT NCAve, TCLP, K-3 kneck, wSO3-3TS, TL-P/Zr, C950, nepheline P, immiscibility/validity numbers, each with Upred |
| `Constraints` | every constraint as value / published limit / PASS-FAIL (Upred margins included) |

## Guarantees

The formulation engine (`ewg30lite/ewg/`) is a byte-copy of the full
application's EWG3.0 stack at the commit recorded in
`VENDORED_SOURCE_COMMIT.txt` — same optimizer, same property models, same
coefficients, same Upred conventions. `tests/test_parity.py` (when the full
repo is available as a sibling) asserts bit-identical formulation results
against the full application.

## Layout

```
ewg30lite/
  cli.py          argparse CLI (template / run)
  io_excel.py     template writer, template reader, result-workbook writer
  ewg/            vendored engine: optimize.py, properties.py, gfc.py,
                  tensor_backend.py + data/ewg30/*.json coefficients
scripts/          QA harness vs the full application
tests/            unit + parity tests
```

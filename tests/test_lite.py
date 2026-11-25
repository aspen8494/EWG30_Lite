"""Unit tests for the lite CLI: template write/read round-trip and output
workbook structure.  Run with: python -m pytest tests/ -q (from repo root)."""
import os
import subprocess
import sys

import pytest
from openpyxl import load_workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from ewg30lite.io_excel import build_template, read_template  # noqa: E402
from ewg30lite.ewg.gfc import EXAMPLE_WASTES  # noqa: E402
def test_template_round_trip(tmp_path):
    p = str(tmp_path / "tpl.xlsx")
    build_template(p)
    waste, settings = read_template(p)
    # prefilled with the Table 3.4 High Al2O3 example waste
    ex = EXAMPLE_WASTES["High Al2O3"]
    nz = {k: v for k, v in ex.items() if v}
    from ewg30lite.io_excel import WASTE_OXIDES
    nz = {k: v for k, v in nz.items() if k in WASTE_OXIDES}
    assert set(waste) == set(nz)
    assert sum(waste.values()) == pytest.approx(sum(nz.values()), rel=1e-12)
    assert settings["waste_mass_kg"] == 1000.0
    assert settings["confidence"] == 0.90


def test_reader_passes_through_engine_foldable_oxides(tmp_path):
    """Oxides outside FULL_OXIDES (lanthanides etc.) must be passed to the
    engine untouched — optimize._rows counts every key in the normalization
    denominator, so filtering would change results."""
    p = str(tmp_path / "tpl.xlsx")
    build_template(p)
    from openpyxl import load_workbook as lw
    wb = lw(p)
    ws = wb["Waste_Vector"]
    n = 2
    while ws.cell(row=n, column=1).value is not None:
        ws.cell(row=n, column=2).value = None
        n += 1
    ws.cell(row=2, column=1, value="La2O3")   # lanthanide, not a glass column
    ws.cell(row=2, column=2, value=0.4)
    ws.cell(row=3, column=1, value="Al2O3")
    ws.cell(row=3, column=2, value=0.3)
    wb.save(p)
    waste, _ = read_template(p)
    assert waste == {"La2O3": 0.4, "Al2O3": 0.3}


def test_template_rejects_others_row(tmp_path):
    p = str(tmp_path / "tpl.xlsx")
    build_template(p)
    from openpyxl import load_workbook as lw
    wb = lw(p)
    ws = wb["Waste_Vector"]
    n = 2
    while ws.cell(row=n, column=1).value is not None:
        n += 1
    ws.cell(row=n, column=1, value="Others")
    ws.cell(row=n, column=2, value=0.1)
    wb.save(p)
    with pytest.raises(ValueError, match="Others"):
        read_template(p)


def test_cli_end_to_end(tmp_path):
    tpl = str(tmp_path / "in.xlsx")
    out = str(tmp_path / "out.xlsx")
    build_template(tpl)
    r = subprocess.run([sys.executable, "-m", "ewg30lite.cli", "run", tpl, "-o", out],
                       capture_output=True, text=True, cwd=ROOT, timeout=1800)
    assert r.returncode == 0, r.stderr
    assert "Waste loading" in r.stdout and "PASS" in r.stdout
    wb = load_workbook(out)
    assert wb.sheetnames == ["Summary", "Input_Waste", "GFC_Recipe",
                             "Glass_Composition", "Glass_Properties", "Constraints"]
    # High Al2O3 example: B2O3 must be a top former (boric acid present)
    gfcs = [row[0] for row in wb["GFC_Recipe"].iter_rows(min_row=2, values_only=True)
            if row[0]]
    assert "Boric acid" in gfcs

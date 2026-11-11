"""Build the workbook, recalculate it with LibreOffice and compare every curve
cell with the Python reference."""
import copy

import openpyxl
import pytest

from tcc import build as b
from tcc import model
from tcc.recalc import recalc, soffice

pytestmark = pytest.mark.skipif(soffice() is None, reason="LibreOffice not installed")

VARIANT = copy.deepcopy(model.EXAMPLE)
VARIANT["plot_kv"] = 4.16
VARIANT["relay"] = [
    dict(en="Yes", name="R-A", kv=4.16, pickup=300, curve="IEC Extremely Inverse", td=0.3),
    dict(en="Yes", name="R-B", kv=13.8, pickup=120, curve="US U2 Inverse", td=4,
         inst=2000, instt=0),
    dict(en="Yes", name="R-G", kv=4.16, pickup=50, curve="Definite Time", td=0.5)]
VARIANT["breaker"][0].update(i2t="Off", instpu=8, ptol=5, ttol=20)
VARIANT["breaker"][1].update(en="No")
VARIANT["pt"][0].update(kv=4.16, scale=0.7)
VARIANT["xfmr"][0].update(cat="III", z=7.0, conn=0.58)
VARIANT["cable"][0].update(mat="Al", nper=2, kv=4.16, t1=90)
VARIANT["fault"][2].update(en="No")
CASES = {"example": model.EXAMPLE, "variant": VARIANT}


@pytest.fixture(scope="module", params=list(CASES))
def book(request, tmp_path_factory):
    tmp = tmp_path_factory.mktemp(request.param)
    data = CASES[request.param]
    layout = b.build(str(tmp / "t.xlsx"), data)
    out = recalc(tmp / "t.xlsx", tmp / "out")
    return data, layout, openpyxl.load_workbook(out, data_only=True)


def test_curve_cells_match_reference(book):
    data, layout, wb = book
    ws = wb["Calc"]
    checked = trips = 0
    for r in layout["fault_rows"] + layout["grid_rows"]:
        i_plot = ws.cell(r + 1, 1).value
        if not isinstance(i_plot, (int, float)):
            continue                                   # disabled fault row
        for key, want in model.evaluate(data, i_plot).items():
            got = ws.cell(r + 1, layout["cols"][key] + 1).value
            where = f"{key} at {i_plot} A (row {r + 1})"
            if want is None:
                assert not isinstance(got, (int, float)), where
            else:
                assert got == pytest.approx(want, rel=1e-6), where
                trips += 1
            checked += 1
    assert checked > 5000 and trips > 1000


def test_no_unexpected_errors(book):
    _, _, wb = book
    for name in ("Devices", "Coordination", "Marks", "Lib_Points"):
        for row in wb[name].iter_rows(values_only=True):
            for v in row:
                if isinstance(v, str) and v.startswith("#") and v != "#N/A":
                    raise AssertionError(f"{name}: {v}")
    for row in wb["Devices"].iter_rows(values_only=True):
        assert "#N/A" not in row
    for row in wb["Coordination"].iter_rows(values_only=True):
        assert "#N/A" not in row


def test_example_coordinates():
    """The worked example shipped in the workbook should pass its own check."""
    def t(amps):
        return model.evaluate(model.EXAMPLE, amps)
    for amps in (30000, 15000, 8000):
        r = t(amps)
        assert r["breaker", 0, "min"] >= r["breaker", 1, "max"]          # main over feeder
        assert r["relay", 0, "t"] - r["breaker", 0, "max"] >= 0.3 - 1e-9  # relay over main
    r = t(8000)
    assert r["breaker", 1, "min"] >= r["pt", 0, "max"]                    # feeder over motor CB


def test_coordination_sheet(book):
    data, layout, wb = book
    ws = wb["Coordination"]
    rows = list(ws.iter_rows(values_only=True))
    names = [r[0] for r in rows]
    if data is model.EXAMPLE:
        ev = model.evaluate(data, 30000)
        main = rows[names.index("CB-Main 2000A")]
        assert main[2] == pytest.approx(ev["breaker", 0, "min"])
        assert main[3] == pytest.approx(ev["breaker", 0, "max"])
        pairs = [r for r in rows if r[0] == "CB-Main 2000A" and r[1] == "R1 xfmr primary 51/50"]
        assert pairs[0][3] == pytest.approx(ev["relay", 0, "t"] - ev["breaker", 0, "max"])
        results = [v for r in rows for v in r if isinstance(v, str) and "LOW" in v]
        assert results == []
        assert pairs[0][4] == "OK"

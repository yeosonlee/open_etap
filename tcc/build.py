"""Generate OpenTCC.xlsx: a macro-free time-current curve workbook.

    python -m tcc.build [output.xlsx]

Every curve cell is an ordinary Excel formula, so the workbook recalculates
when settings on the Devices sheet change.
"""
from __future__ import annotations

import sys

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name as L

from . import curves as c
from .model import EXAMPLE, SLOTS, normalise

PER_DECADE = 100
I_MIN_EXP, I_MAX_EXP = 0, 5            # plot from 1 A to 100 kA
T_MIN, T_MAX = 0.01, 1000
LIB_RELAY_ROWS = 30
LIB_PT_TABLES = 8
LIB_PT_ROWS = 30

# Calc sheet rows (0-based)
R_NAME, R_DESC, R_FACTOR, R_AUX1, R_AUX2 = 0, 1, 2, 3, 4
R_FAULT0 = 6
R_GRID0 = R_FAULT0 + SLOTS["fault"] + 1
N_GRID = (I_MAX_EXP - I_MIN_EXP) * PER_DECADE + 1

PALETTE = ["#1F77B4", "#D62728", "#2CA02C", "#9467BD", "#FF7F0E", "#17BECF",
           "#8C564B", "#E377C2", "#BCBD22", "#393B79", "#AD494A", "#637939"]

# (key, header, is_input)
TABLES = {
    "relay": ("Overcurrent relays (51 / 50)", [
        ("en", "Enabled", 1), ("name", "Name", 1), ("kv", "kV", 1),
        ("pickup", "51 pickup (A primary)", 1), ("curve", "Curve", 1),
        ("td", "Time dial / TMS", 1), ("inst", "50 pickup (A primary, 0 = off)", 1),
        ("instt", "50 operating time (s)", 1),
        ("A", "A", 0), ("B", "B", 0), ("p", "p", 0)]),
    "breaker": ("LV breakers with electronic trip (LSI)", [
        ("en", "Enabled", 1), ("name", "Name", 1), ("kv", "kV", 1),
        ("rating", "Rating In (A)", 1), ("ltpu", "Long-time pickup (x In)", 1),
        ("tr", "Long-time delay (s at 6 x Ir)", 1),
        ("stpu", "Short-time pickup (x Ir, 0 = off)", 1),
        ("tsd", "Short-time delay (s)", 1), ("i2t", "Short-time I²t", 1),
        ("i2tref", "I²t ramp ends at (x Ir)", 1),
        ("instpu", "Instantaneous pickup (x In, 0 = off)", 1),
        ("tinst", "Instantaneous clearing time (s)", 1),
        ("ptol", "Pickup tolerance (± %)", 1), ("ttol", "Time tolerance (± %)", 1),
        ("Ir", "Ir (A)", 0), ("Isd", "Isd (A)", 0), ("Ii", "Ii (A)", 0)]),
    "pt": ("Point-table devices (fuses, thermal-magnetic breakers)", [
        ("en", "Enabled", 1), ("name", "Name", 1), ("kv", "kV", 1),
        ("table", "Table (see Lib_Points)", 1), ("scale", "Current multiplier", 1)]),
    "xfmr": ("Transformer damage curve (IEEE C57.109)", [
        ("en", "Enabled", 1), ("name", "Name", 1), ("kva", "kVA", 1),
        ("kv", "kV (side the curve is drawn for)", 1), ("z", "Impedance (%)", 1),
        ("cat", "Category", 1), ("conn", "Connection shift factor", 1),
        ("inrush", "Inrush (x FLA)", 1), ("inrusht", "Inrush time (s)", 1),
        ("Ib", "FLA (A)", 0), ("Mmax", "Max fault (x FLA)", 0),
        ("frac", "Mechanical from (x max fault)", 0)]),
    "cable": ("Cable short-circuit damage curve", [
        ("en", "Enabled", 1), ("name", "Name", 1), ("kv", "kV", 1),
        ("mat", "Conductor", 1), ("area", "Size (mm²)", 1),
        ("nper", "Conductors per phase", 1), ("t1", "Initial temp (°C)", 1),
        ("t2", "Final temp (°C)", 1),
        ("cmil", "Total area (cmil)", 0), ("K", "K", 0)]),
    "motor": ("Motor starting curve", [
        ("en", "Enabled", 1), ("name", "Name", 1), ("kv", "kV", 1),
        ("fla", "FLA (A)", 1), ("lram", "Locked rotor (x FLA)", 1),
        ("tacc", "Acceleration time (s)", 1),
        ("stall", "Safe stall time (s, 0 = hide)", 1), ("LRA", "LRA (A)", 0)]),
    "fault": ("Fault currents", [
        ("en", "Enabled", 1), ("name", "Name", 1), ("kv", "kV", 1),
        ("amps", "Fault current (A)", 1)]),
}
ORDER = ["relay", "breaker", "pt", "xfmr", "cable", "motor", "fault"]


class Tbl:
    """One settings table on the Devices sheet."""

    def __init__(self, kind, top):
        self.kind, self.top = kind, top
        self.title, self.cols = TABLES[kind]
        self.n = SLOTS[kind]
        self.keys = [k for k, _, _ in self.cols]

    def row(self, i):
        return self.top + 2 + i

    def ref(self, key, i):
        return f"Devices!${L(self.keys.index(key))}${self.row(i) + 1}"

    def rng(self, key):
        col = L(self.keys.index(key))
        return f"Devices!${col}${self.row(0) + 1}:${col}${self.row(self.n - 1) + 1}"


def build(path="OpenTCC.xlsx", data=None):
    """Write the workbook and return its layout (used by the tests)."""
    d = normalise(data or EXAMPLE)
    wb = xlsxwriter.Workbook(path)
    f = {
        "title": wb.add_format({"bold": True, "font_size": 14}),
        "h": wb.add_format({"bold": True, "bg_color": "#DDE3EA", "border": 1,
                            "text_wrap": True, "valign": "top"}),
        "in": wb.add_format({"bg_color": "#FFF8D6", "border": 1}),
        "calc": wb.add_format({"bg_color": "#EEEEEE", "border": 1, "num_format": "0.####"}),
        "sec": wb.add_format({"bold": True, "font_size": 12}),
        "note": wb.add_format({"italic": True, "font_color": "#555555"}),
        "wrap": wb.add_format({"text_wrap": True, "valign": "top"}),
        "t": wb.add_format({"num_format": "0.000", "border": 1}),
        "bad": wb.add_format({"bg_color": "#F8C9C9", "font_color": "#8B0000"}),
        "good": wb.add_format({"bg_color": "#D5EFD5", "font_color": "#1B5E20"}),
    }
    ws_dev = wb.add_worksheet("Devices")
    ws_tcc = wb.add_worksheet("TCC")
    ws_co = wb.add_worksheet("Coordination")
    ws_lr = wb.add_worksheet("Lib_Relay")
    ws_lp = wb.add_worksheet("Lib_Points")
    ws_lx = wb.add_worksheet("Lib_Xfmr")
    ws_calc = wb.add_worksheet("Calc")
    ws_mk = wb.add_worksheet("Marks")
    ws_about = wb.add_worksheet("About")

    # ---------------------------------------------------------------- libraries
    ws_lr.write(0, 0, "Relay curve library", f["title"])
    ws_lr.write(1, 0, "t = TD × (A / (M^p − 1) + B), M = current / pickup. "
                      "Add your own curves in the blank rows.", f["note"])
    for j, h in enumerate(["Curve", "Source", "A", "B", "p"]):
        ws_lr.write(2, j, h, f["h"])
    for i in range(LIB_RELAY_ROWS):
        ws_lr.write_row(3 + i, 0, [""] * 5, f["in"])
    for i, (name, (src, a, b, p)) in enumerate(c.RELAY_CURVES.items()):
        ws_lr.write_row(3 + i, 0, [name, src, a, b, p], f["in"])
    ws_lr.set_column(0, 1, 28)
    lr = lambda col: f"Lib_Relay!${col}$4:${col}${3 + LIB_RELAY_ROWS}"

    ws_lp.write(0, 0, "Point tables: current (A), minimum and maximum time (s). "
                      "Currents ascending, no repeated currents.", f["title"])
    last_col = L(3 * LIB_PT_TABLES - 1)
    lp_names = f"Lib_Points!$A$2:${last_col}$2"
    lp_data = f"Lib_Points!$A$4:${last_col}${3 + LIB_PT_ROWS}"
    idx_col = 3 * LIB_PT_TABLES + 1
    ws_lp.write(2, idx_col, "Table list", f["h"])
    for t in range(LIB_PT_TABLES):
        c0 = 3 * t
        pts = d["tables"][t]["points"] if t < len(d["tables"]) else []
        ws_lp.write(1, c0, d["tables"][t]["name"] if pts else "", f["in"])
        ws_lp.write_row(2, c0, ["Current (A)", "Min time (s)", "Max time (s)"], f["h"])
        for i in range(LIB_PT_ROWS):
            ws_lp.write_row(3 + i, c0, pts[i] if i < len(pts) else ["", "", ""], f["in"])
        ws_lp.write_formula(3 + t, idx_col, f'=IF({L(c0)}2="","",{L(c0)}2)')
    ws_lp.set_column(0, idx_col - 1, 12)
    ws_lp.set_column(idx_col, idx_col, 32)
    lp_list = f"Lib_Points!${L(idx_col)}$4:${L(idx_col)}${3 + LIB_PT_TABLES}"

    ws_lx.write(0, 0, "IEEE C57.109 thermal through-fault curve", f["title"])
    ws_lx.write_row(2, 0, ["x base current", "Time (s)", "Log-log slope"], f["h"])
    for i, ((m, t), s) in enumerate(zip(c.XFMR_THERMAL, c.xfmr_thermal_slopes())):
        ws_lx.write_row(3 + i, 0, [m, t, s])
    nx = len(c.XFMR_THERMAL)
    lx = lambda col: f"Lib_Xfmr!${col}$4:${col}${3 + nx}"
    ws_lx.set_column(0, 2, 16)

    # ------------------------------------------------------------------ Devices
    ws_dev.write(0, 0, "Open TCC: device settings", f["title"])
    ws_dev.write(1, 0, "Yellow cells are inputs; grey cells are calculated. "
                       "Currents are entered at each device's own kV.", f["note"])
    ws_dev.write(3, 0, "Plot reference kV", f["h"])
    ws_dev.write(3, 1, d["plot_kv"], f["in"])
    PLOT = "Devices!$B$4"
    T = {}
    top = 5
    for kind in ORDER:
        T[kind] = Tbl(kind, top)
        top += SLOTS[kind] + 3
    calc_formula = {
        ("relay", "A"): lambda t, i: f'=IFERROR(INDEX({lr("C")},MATCH({t.ref("curve", i)},{lr("A")},0)),0)',
        ("relay", "B"): lambda t, i: f'=IFERROR(INDEX({lr("D")},MATCH({t.ref("curve", i)},{lr("A")},0)),0)',
        ("relay", "p"): lambda t, i: f'=IFERROR(INDEX({lr("E")},MATCH({t.ref("curve", i)},{lr("A")},0)),1)',
        ("breaker", "Ir"): lambda t, i: f'={t.ref("ltpu", i)}*{t.ref("rating", i)}',
        ("breaker", "Isd"): lambda t, i: f'={t.ref("stpu", i)}*{t.ref("Ir", i)}',
        ("breaker", "Ii"): lambda t, i: f'={t.ref("instpu", i)}*{t.ref("rating", i)}',
        ("xfmr", "Ib"): lambda t, i: f'={t.ref("kva", i)}/(SQRT(3)*{t.ref("kv", i)})',
        ("xfmr", "Mmax"): lambda t, i: f'=100/{t.ref("z", i)}',
        ("xfmr", "frac"): lambda t, i: f'=IF({t.ref("cat", i)}="I",99,IF({t.ref("cat", i)}="II",0.7,0.5))',
        ("cable", "cmil"): lambda t, i: f'={t.ref("nper", i)}*{t.ref("area", i)}*{c.CMIL_PER_MM2}',
        ("cable", "K"): lambda t, i: (
            f'=IF({t.ref("mat", i)}="Cu",0.0297*LOG10(({t.ref("t2", i)}+234)/({t.ref("t1", i)}+234)),'
            f'0.0125*LOG10(({t.ref("t2", i)}+228)/({t.ref("t1", i)}+228)))'),
        ("motor", "LRA"): lambda t, i: f'={t.ref("fla", i)}*{t.ref("lram", i)}',
    }
    lists = {"en": ["Yes", "No"], "i2t": ["On", "Off"], "cat": ["I", "II", "III", "IV"],
             "mat": ["Cu", "Al"], "curve": "=" + lr("A"), "table": "=" + lp_list}
    for kind in ORDER:
        t = T[kind]
        ws_dev.write(t.top, 0, t.title, f["sec"])
        for j, (key, head, is_in) in enumerate(t.cols):
            ws_dev.write(t.top + 1, j, head, f["h"])
            for i in range(t.n):
                if is_in:
                    ws_dev.write(t.row(i), j, d[kind][i][key], f["in"])
                    if key in lists:
                        ws_dev.data_validation(t.row(i), j, t.row(i), j,
                                               {"validate": "list", "source": lists[key]})
                else:
                    ws_dev.write_formula(t.row(i), j, calc_formula[kind, key](t, i), f["calc"])
        ws_dev.set_row(t.top + 1, 45)
    ws_dev.set_column(0, 0, 9)
    ws_dev.set_column(1, 1, 26)
    ws_dev.set_column(2, 20, 13)
    ws_dev.write(top, 0, "Connection shift factor: 1.0 normally; 0.87 for a line-to-line "
                         "fault behind a delta-delta transformer; 0.58 for a ground fault "
                         "behind a delta-wye (solidly grounded) transformer.", f["note"])
    ws_dev.freeze_panes(5, 2)

    # --------------------------------------------------------------------- Calc
    data_rows = list(range(R_FAULT0, R_FAULT0 + SLOTS["fault"])) + \
        list(range(R_GRID0, R_GRID0 + N_GRID))
    ws_calc.write(R_DESC, 0, "I at plot kV (A)", f["h"])
    ws_calc.write(R_FACTOR, 0, "kV factor")
    ws_calc.write(R_FAULT0 - 1, 0, "Fault rows", f["note"])
    ws_calc.write(R_GRID0 - 1, 0, "Curve grid", f["note"])
    tf = T["fault"]
    for i in range(SLOTS["fault"]):
        ws_calc.write_formula(
            R_FAULT0 + i, 0,
            f'=IF({tf.ref("en", i)}="Yes",{tf.ref("amps", i)}*{tf.ref("kv", i)}/{PLOT},NA())')
    for k in range(N_GRID):
        ws_calc.write_number(R_GRID0 + k, 0, float(f"{10 ** (I_MIN_EXP + k / PER_DECADE):.6g}"))

    cols = {}
    state = {"col": 1}

    names = {}

    def add(key, desc, formula, name=None, aux=(), text=""):
        """Add a Calc column. formula(r) gets the 1-based row number."""
        col = state["col"]
        state["col"] += 1
        cols[key] = col
        names[key] = text
        if name:
            ws_calc.write_formula(R_NAME, col, name)
        ws_calc.write(R_DESC, col, desc, f["h"])
        for row, fx in aux:
            ws_calc.write_formula(row, col, fx)
        for r0 in data_rows:
            ws_calc.write_formula(r0, col, formula(r0 + 1))
        return L(col)

    def factor(t, i):
        return (R_FACTOR, f'={PLOT}/{t.ref("kv", i)}')

    for i in range(SLOTS["relay"]):
        t = T["relay"]
        g = lambda k, t=t, i=i: t.ref(k, i)
        cm = L(state["col"])
        add(("relay", i, "M"), "M = I / pickup",
            lambda r, g=g, cm=cm: f'=IF({g("pickup")}>0,$A{r}*${cm}${R_FACTOR + 1}/{g("pickup")},0)',
            aux=[factor(t, i)])

        def fx(r, g=g, cm=cm):
            m = f"${cm}{r}"
            tt = f'{g("td")}*({g("A")}/({m}^{g("p")}-1)+{g("B")})'
            return (f'=IFERROR(IF(OR({g("en")}<>"Yes",{m}<=1),NA(),'
                    f'IF(AND({g("inst")}>0,{m}*{g("pickup")}>={g("inst")}),'
                    f'MIN(MAX({g("instt")},{T_MIN}),{tt}),{tt})),NA())')
        add(("relay", i, "t"), "t (s)", fx, name=f'={g("name")}', text=d["relay"][i]["name"])

    for i in range(SLOTS["breaker"]):
        t = T["breaker"]
        g = lambda k, t=t, i=i: t.ref(k, i)
        ci = L(state["col"])
        add(("breaker", i, "I"), "I at device (A)",
            lambda r, ci=ci: f"=$A{r}*${ci}${R_FACTOR + 1}", aux=[factor(t, i)])

        def band(sign, tinst, g=g, ci=ci):
            pt = f'(1{sign}{g("ptol")}/100)'
            tt = f'(1{sign}{g("ttol")}/100)'

            def fx(r):
                x = f"${ci}{r}"
                lt = f'{g("tr")}*{tt}*(6*{g("Ir")}*{pt}/{x})^2'
                return (
                    f'=IFERROR(IF(OR({g("en")}<>"Yes",{x}<={g("Ir")}*{pt}),NA(),MIN('
                    f'IF(AND({g("Isd")}>0,{x}>={g("Isd")}*{pt}),MIN(MAX({lt},{g("tsd")}*{tt}),'
                    f'{g("tsd")}*{tt}*IF({g("i2t")}="On",MAX(1,({g("i2tref")}*{g("Ir")}*{pt}/{x})^2),1)),'
                    f'{lt}),'
                    f'IF(AND({g("Ii")}>0,{x}>={g("Ii")}*{pt}),{tinst},1E+99))),NA())')
            return fx
        add(("breaker", i, "min"), "t min (s)", band("-", str(T_MIN)), name=f'={g("name")}',
            text=d["breaker"][i]["name"])
        add(("breaker", i, "max"), "t max (s)", band("+", g("tinst")),
            name=f'={g("name")}&" (max)"')

    for i in range(SLOTS["pt"]):
        t = T["pt"]
        g = lambda k, t=t, i=i: t.ref(k, i)
        cx = L(state["col"])
        cidx = f"${cx}${R_AUX1 + 1}"
        npts = f"${cx}${R_AUX2 + 1}"
        add(("pt", i, "x"), "I / multiplier (A)",
            lambda r, g=g, cx=cx: f'=$A{r}*${cx}${R_FACTOR + 1}/{g("scale")}',
            aux=[factor(t, i),
                 (R_AUX1, f'=MATCH({g("table")},{lp_names},0)'),
                 (R_AUX2, f"=COUNT(INDEX({lp_data},0,{cidx}))")])
        ck = L(state["col"])
        add(("pt", i, "k"), "segment",
            lambda r, cx=cx, cidx=cidx, npts=npts: (
                f"=IFERROR(IF(OR(${cx}{r}<INDEX({lp_data},1,{cidx}),"
                f"${cx}{r}>INDEX({lp_data},{npts},{cidx})),NA(),"
                f"MIN(MATCH(${cx}{r},INDEX({lp_data},0,{cidx}),1),{npts}-1)),NA())"))

        def interp(off, g=g, cx=cx, ck=ck, cidx=cidx):
            def fx(r):
                k = f"${ck}{r}"
                y0 = f"LOG10(INDEX({lp_data},{k},{cidx}+{off}))"
                y1 = f"LOG10(INDEX({lp_data},{k}+1,{cidx}+{off}))"
                x0 = f"LOG10(INDEX({lp_data},{k},{cidx}))"
                x1 = f"LOG10(INDEX({lp_data},{k}+1,{cidx}))"
                return (f'=IFERROR(IF({g("en")}<>"Yes",NA(),'
                        f"10^({y0}+({y1}-{y0})*(LOG10(${cx}{r})-{x0})/({x1}-{x0}))),NA())")
            return fx
        add(("pt", i, "min"), "t min (s)", interp(1), name=f'={g("name")}', text=d["pt"][i]["name"])
        add(("pt", i, "max"), "t max (s)", interp(2), name=f'={g("name")}&" (max)"')

    for i in range(SLOTS["xfmr"]):
        t = T["xfmr"]
        g = lambda k, t=t, i=i: t.ref(k, i)
        cm = L(state["col"])
        add(("xfmr", i, "M"), "x base current",
            lambda r, g=g, cm=cm: f'=$A{r}*${cm}${R_FACTOR + 1}/({g("Ib")}*{g("conn")})',
            aux=[factor(t, i)])
        cth = L(state["col"])

        def th(r, cm=cm):
            k = f'MATCH(${cm}{r},{lx("A")},1)'
            return (f'=IFERROR(INDEX({lx("B")},{k})*(${cm}{r}/INDEX({lx("A")},{k}))'
                    f'^INDEX({lx("C")},{k}),NA())')
        add(("xfmr", i, "th"), "thermal (s)", th)
        add(("xfmr", i, "t"), "damage (s)",
            lambda r, g=g, cm=cm, cth=cth: (
                f'=IFERROR(IF(OR({g("en")}<>"Yes",${cm}{r}<2,${cm}{r}>{g("Mmax")}),NA(),'
                f'IF(${cm}{r}>={g("frac")}*{g("Mmax")},'
                f'MIN(${cth}{r},2*({g("Mmax")}/${cm}{r})^2),${cth}{r})),NA())'),
            name=f'={g("name")}&" damage"', text=d["xfmr"][i]["name"] + " damage")

    for i in range(SLOTS["cable"]):
        t = T["cable"]
        g = lambda k, t=t, i=i: t.ref(k, i)
        cc = L(state["col"])
        add(("cable", i, "raw"), "K (A/I)² (s)",
            lambda r, g=g, cc=cc: f'={g("K")}*({g("cmil")}/($A{r}*${cc}${R_FACTOR + 1}))^2',
            aux=[factor(t, i)])
        add(("cable", i, "t"), "damage (s)",
            lambda r, g=g, cc=cc: (
                f'=IFERROR(IF(OR({g("en")}<>"Yes",${cc}{r}<{T_MIN},${cc}{r}>10),NA(),'
                f"${cc}{r}),NA())"),
            name=f'={g("name")}&" damage"', text=d["cable"][i]["name"] + " damage")
    ws_calc.set_column(0, state["col"], 14)
    ws_calc.freeze_panes(R_GRID0, 1)

    # -------------------------------------------------------------------- Marks
    ws_mk.write_row(0, 0, ["Label", "I at plot kV (A)", "t (s)"], f["h"])
    ws_mk.set_column(0, 0, 36)
    ws_mk.set_column(1, 2, 16)
    mk = {"row": 1}

    mark_text = {}

    def mark(label, en, x, y, text=""):
        r = mk["row"]
        mark_text[r] = text
        mk["row"] += 1
        ws_mk.write_formula(r, 0, label)
        ws_mk.write_formula(r, 1, f'=IFERROR(IF({en}="Yes",{x},NA()),NA())')
        ws_mk.write_formula(r, 2, f'=IFERROR(IF({en}="Yes",{y},NA()),NA())')
        return r

    marks = {}
    for i in range(SLOTS["motor"]):
        g = lambda k, i=i: T["motor"].ref(k, i)
        ref = lambda a, g=g: f'{a}*{g("kv")}/{PLOT}'
        nm = d["motor"][i]["name"]
        rows = [mark(f'={g("name")}&" start"', g("en"), ref(g(a)), y, nm + " start") for a, y in
                [("fla", T_MAX), ("fla", g("tacc")), ("LRA", g("tacc")), ("LRA", T_MIN)]]
        marks["motor", i] = (rows[0], rows[-1])
        marks["stall", i] = mark(f'={g("name")}&" safe stall"',
                                 f'IF({g("stall")}>0,{g("en")},"No")',
                                 ref(g("LRA")), g("stall"), nm + " safe stall")
    for i in range(SLOTS["xfmr"]):
        g = lambda k, i=i: T["xfmr"].ref(k, i)
        ref = lambda a, g=g: f'{a}*{g("kv")}/{PLOT}'
        marks["fla", i] = mark(f'={g("name")}&" FLA"', g("en"), ref(g("Ib")), T_MAX,
                               d["xfmr"][i]["name"] + " FLA")
        marks["inrush", i] = mark(f'={g("name")}&" inrush"', g("en"),
                                  ref(f'{g("Ib")}*{g("inrush")}'), g("inrusht"),
                                  d["xfmr"][i]["name"] + " inrush")
    for i in range(SLOTS["fault"]):
        g = lambda k, i=i: T["fault"].ref(k, i)
        lab = f'={g("name")}&" "&TEXT({g("amps")}/1000,"0.0")&" kA"'
        x = f'{g("amps")}*{g("kv")}/{PLOT}'
        txt = f'{d["fault"][i]["name"]} {d["fault"][i]["amps"] / 1000:.1f} kA'
        r0 = mark(lab, g("en"), x, T_MAX, txt)
        mark(lab, g("en"), x, T_MIN, txt)
        marks["fault", i] = (r0, r0 + 1)

    # -------------------------------------------------------------------- chart
    ws_calc.write_formula(R_NAME, 0, f'="Current (A) at "&{PLOT}&" kV"')
    chart = wb.add_chart({"type": "scatter", "subtype": "straight"})
    g0, g1 = R_GRID0 + 1, R_GRID0 + N_GRID
    xs = f"=Calc!$A${g0}:$A${g1}"
    hidden, n_series, colour = [], 0, iter(PALETTE * 3)

    def series(key, col_hex, dash="solid", width=1.75, legend=True):
        nonlocal n_series
        cl = L(cols[key])
        chart.add_series({
            "name": f"=Calc!${cl}${R_NAME + 1}", "name_data": [names[key]],
            "categories": xs,
            "values": f"=Calc!${cl}${g0}:${cl}${g1}",
            "line": {"color": col_hex, "width": width, "dash_type": dash},
            "marker": {"type": "none"}})
        if not legend:
            hidden.append(n_series)
        n_series += 1

    def mark_series(rows, col_hex, line=None, marker=None, label_first=True, legend=False):
        nonlocal n_series
        r0, r1 = rows if isinstance(rows, tuple) else (rows, rows)
        s = {"name": f"=Marks!$A${r0 + 1}", "name_data": [mark_text[r0]],
             "categories": f"=Marks!$B${r0 + 1}:$B${r1 + 1}",
             "values": f"=Marks!$C${r0 + 1}:$C${r1 + 1}",
             "marker": marker or {"type": "none"}}
        s["line"] = line or {"none": True}
        if label_first:
            s["data_labels"] = {
                "value": True, "position": "right",
                "font": {"size": 8, "color": col_hex},
                "custom": [{"value": f"=Marks!$A${r0 + 1}", "data": [mark_text[r0]]}] +
                          [{"delete": True}] * (r1 - r0)}
        chart.add_series(s)
        if not legend:
            hidden.append(n_series)
        n_series += 1

    for i in range(SLOTS["relay"]):
        series(("relay", i, "t"), next(colour))
    for kind in ("breaker", "pt"):
        for i in range(SLOTS[kind]):
            col_hex = next(colour)
            series((kind, i, "min"), col_hex)
            series((kind, i, "max"), col_hex, legend=False)
    for i in range(SLOTS["xfmr"]):
        series(("xfmr", i, "t"), "#333333", dash="dash", width=2)
        mark_series(marks["fla", i], "#333333",
                    marker={"type": "triangle", "size": 7,
                            "fill": {"color": "#333333"}, "border": {"color": "#333333"}})
        mark_series(marks["inrush", i], "#333333",
                    marker={"type": "x", "size": 8, "border": {"color": "#333333"}})
    for i in range(SLOTS["cable"]):
        series(("cable", i, "t"), "#8C564B", dash="long_dash", width=2)
    for i in range(SLOTS["motor"]):
        mark_series(marks["motor", i], "#2E7D32",
                    line={"color": "#2E7D32", "width": 1.5, "dash_type": "dash_dot"},
                    label_first=False, legend=True)
        mark_series(marks["stall", i], "#2E7D32",
                    marker={"type": "diamond", "size": 7,
                            "fill": {"color": "#2E7D32"}, "border": {"color": "#2E7D32"}})
    for i in range(SLOTS["fault"]):
        mark_series(marks["fault", i], "#B71C1C",
                    line={"color": "#B71C1C", "width": 1, "dash_type": "square_dot"})

    grid_major = {"visible": True, "line": {"color": "#9E9E9E", "width": 0.75}}
    grid_minor = {"visible": True, "line": {"color": "#DDDDDD", "width": 0.5}}
    chart.set_x_axis({"name": "=Calc!$A$1", "data": [f"Current (A) at {d['plot_kv']} kV"],
                      "log_base": 10,
                      "min": 10 ** I_MIN_EXP, "max": 10 ** I_MAX_EXP, "crossing": 10 ** I_MIN_EXP,
                      "major_gridlines": grid_major, "minor_gridlines": grid_minor,
                      "num_format": "#,##0"})
    chart.set_y_axis({"name": "Time (s)", "log_base": 10, "min": T_MIN, "max": T_MAX,
                      "crossing": T_MIN, "major_gridlines": grid_major,
                      "minor_gridlines": grid_minor, "num_format": "General"})
    chart.set_title({"name": "Time-current curves"})
    chart.set_legend({"position": "right", "delete_series": hidden})
    chart.show_na_as_empty_cell()
    chart.set_size({"width": 1150, "height": 900})
    ws_tcc.insert_chart("A1", chart)
    ws_tcc.set_landscape()
    ws_tcc.fit_to_pages(1, 1)

    # ------------------------------------------------------------- Coordination
    nf = SLOTS["fault"]
    ws_co.write(0, 0, "Coordination check", f["title"])
    ws_co.write(1, 0, "Operating times at each fault current, calculated directly "
                      "(not read off the plotted grid).", f["note"])
    hr = 3
    ws_co.write_row(hr, 0, ["Device", "Type"], f["h"])
    for k in range(nf):
        nm = T["fault"].ref("name", k)
        ws_co.write_formula(hr, 2 + 2 * k, f'={nm}&" min (s)"', f["h"])
        ws_co.write_formula(hr, 3 + 2 * k, f'={nm}&" max (s)"', f["h"])
    ws_co.set_row(hr, 32)
    devs = [("relay", i, "t", "t", "Relay", "no trip") for i in range(SLOTS["relay"])] + \
           [("breaker", i, "min", "max", "LSI breaker", "no trip") for i in range(SLOTS["breaker"])] + \
           [("pt", i, "min", "max", "Point table", "off table") for i in range(SLOTS["pt"])]
    r = hr + 1
    for kind, i, kmin, kmax, label, none in devs:
        ws_co.write_formula(r, 0, f'={T[kind].ref("name", i)}')
        ws_co.write(r, 1, label)
        for k in range(nf):
            for off, key in ((0, kmin), (1, kmax)):
                src = f"Calc!${L(cols[kind, i, key])}${R_FAULT0 + k + 1}"
                ws_co.write_formula(r, 2 + 2 * k + off, f'=IFERROR({src},"{none}")', f["t"])
        r += 1
    d0, d1 = hr + 2, r                               # 1-based first and last device row
    names = f"$A${d0}:$A${d1}"
    ws_co.write(r + 1, 0, "Device pairs", f["sec"])
    ws_co.write(r + 2, 0, "Margin = upstream minimum time − downstream maximum time. "
                          "Typical requirement: 0.2 to 0.4 s between relays; 0 where "
                          "both devices are drawn as bands. \"Off table\" means the current is "
                          "outside the device's point table.", f["note"])
    pr = r + 3
    ws_co.write_row(pr, 0, ["Downstream device", "Upstream device", "Required margin (s)"], f["h"])
    for k in range(nf):
        nm = T["fault"].ref("name", k)
        ws_co.write_formula(pr, 3 + 2 * k, f'={nm}&" margin (s)"', f["h"])
        ws_co.write_formula(pr, 4 + 2 * k, f'={nm}&" result"', f["h"])
    ws_co.set_row(pr, 32)
    n_pairs = 8
    for j in range(n_pairs):
        row = pr + 1 + j
        dn, up, req = d["pairs"][j] if j < len(d["pairs"]) else ("", "", 0.3)
        ws_co.write(row, 0, dn, f["in"])
        ws_co.write(row, 1, up, f["in"])
        ws_co.write(row, 2, req, f["in"])
        ws_co.data_validation(row, 0, row, 1, {"validate": "list", "source": "=" + names})
        for k in range(nf):
            cmax = L(3 + 2 * k)
            cmin = L(2 + 2 * k)
            dmax = f"INDEX(${cmax}${d0}:${cmax}${d1},MATCH($A{row + 1},{names},0))"
            umin = f"INDEX(${cmin}${d0}:${cmin}${d1},MATCH($B{row + 1},{names},0))"
            blank = f'OR($A{row + 1}="",$B{row + 1}="")'
            ws_co.write_formula(
                row, 3 + 2 * k,
                f'=IF({blank},"",IFERROR(IF(AND(ISNUMBER({dmax}),ISNUMBER({umin})),'
                f'{umin}-{dmax},""),""))', f["t"])
            ws_co.write_formula(
                row, 4 + 2 * k,
                f'=IF({blank},"",IFERROR(IF(NOT(ISNUMBER({dmax})),"downstream does not operate",'
                f'IF(NOT(ISNUMBER({umin})),"OK (upstream does not operate)",'
                f'IF({umin}-{dmax}>=$C{row + 1}-0.000001,"OK","MARGIN LOW"))),"unknown device"))')
    last = pr + n_pairs
    ws_co.conditional_format(pr + 1, 3, last, 3 + 2 * nf, {
        "type": "text", "criteria": "containing", "value": "LOW", "format": f["bad"]})
    ws_co.conditional_format(pr + 1, 3, last, 3 + 2 * nf, {
        "type": "text", "criteria": "begins with", "value": "OK", "format": f["good"]})
    ws_co.set_column(0, 1, 28)
    ws_co.set_column(2, 3 + 2 * nf, 17)

    # -------------------------------------------------------------------- About
    about = [
        "Open TCC",
        "A free time-current curve (TCC) workbook for protective device coordination.",
        "",
        "How to use",
        "1. Enter devices on the Devices sheet (yellow cells). Set Enabled to Yes to plot a device.",
        "2. Set the plot reference kV. Every curve is referred to that voltage.",
        "3. Read the plot on the TCC sheet and the trip times and margins on the Coordination sheet.",
        "4. Add relay curves on Lib_Relay and fuse or breaker point tables on Lib_Points.",
        "",
        "Models and sources",
        "Relays: IEEE C37.112 and IEC 60255-151 inverse-time equations; US U1 to U5 from SEL manuals.",
        "Time-dial scales differ between makes: check the equation in the relay manual.",
        "LSI breakers: long-time I²t through the delay at 6 x Ir, short-time flat or I²t, instantaneous.",
        "Bands apply the pickup and time tolerances; the lower edge of the instantaneous region is the plot floor.",
        "Fuses and thermal-magnetic breakers: log-log interpolation of the point table.",
        "Transformer: IEEE C57.109 through-fault curve, thermal plus mechanical for categories II to IV.",
        "Cable: (I/A)² t = 0.0297 log10((T2+234)/(T1+234)) for copper, 0.0125 and 228 for aluminium.",
        "",
        "Limits",
        f"Curves are drawn on a grid of {PER_DECADE} points per decade, so vertical steps are drawn about 2 % wide.",
        "The two point tables supplied are made-up shapes for demonstration, not real devices.",
        "Ground-fault elements are not modelled separately; enter one as a relay with a Definite Time curve.",
        "This workbook is an aid. Check settings against manufacturer data before applying them.",
        "",
        "Generated by tcc/build.py. Not affiliated with ETAP or Operation Technology, Inc.",
    ]
    for i, line in enumerate(about):
        heads = ("Open TCC", "How to use", "Models and sources", "Limits")
        ws_about.write(i, 0, line, f["sec"] if line in heads else None)
    ws_about.set_column(0, 0, 110)

    ws_tcc.activate()
    wb.close()
    return {"cols": cols, "fault_rows": list(range(R_FAULT0, R_FAULT0 + SLOTS["fault"])),
            "grid_rows": list(range(R_GRID0, R_GRID0 + N_GRID))}


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "OpenTCC.xlsx"
    build(out)
    print(f"wrote {out}")

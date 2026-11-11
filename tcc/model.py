"""Study data: slot counts, defaults, the worked example, and a reference
evaluation that mirrors the workbook's Calc sheet."""
from __future__ import annotations

import copy

from . import curves as c

# Number of device slots of each kind in the workbook.
SLOTS = {"relay": 3, "breaker": 3, "pt": 3, "xfmr": 1, "cable": 1,
         "motor": 1, "fault": 3}

DEFAULTS = {
    "relay": dict(en="No", name="", kv=0.48, pickup=100, curve="IEC Standard Inverse",
                  td=0.1, inst=0, instt=0.02),
    "breaker": dict(en="No", name="", kv=0.48, rating=1000, ltpu=1.0, tr=5, stpu=0,
                    tsd=0.1, i2t="Off", i2tref=8, instpu=0, tinst=0.05,
                    ptol=10, ttol=10),
    "pt": dict(en="No", name="", kv=0.48, table="", scale=1.0),
    "xfmr": dict(en="No", name="", kva=1000, kv=0.48, z=5.75, cat="II", conn=1.0,
                 inrush=12, inrusht=0.1),
    "cable": dict(en="No", name="", kv=0.48, mat="Cu", area=240, nper=1, t1=75, t2=250),
    "motor": dict(en="No", name="", kv=0.48, fla=100, lram=6, tacc=5, stall=0),
    "fault": dict(en="No", name="", kv=0.48, amps=10000),
}


def _fuse_table():
    pts = []
    for m in (2, 2.5, 3, 4, 5, 6, 8, 10, 15, 20, 30, 50):
        tmin = 300 * (m / 2) ** -3.2
        pts.append((100.0 * m, float(f"{tmin:.3g}"), float(f"{tmin * 1.3 ** 3.2:.3g}")))
    return pts


# Point tables: (current A, minimum time s, maximum time s), ascending current.
# These two are made-up shapes for demonstration. Replace them with points
# read from the manufacturer's curve sheet.
POINT_TABLES = [
    {"name": "Fuse 100A (illustrative)", "points": _fuse_table()},
    {"name": "TM breaker 100A (illustrative)", "points": [
        (110, 600, 10000), (150, 60, 300), (200, 25, 100), (300, 9, 30),
        (400, 5, 15), (500, 3.2, 9), (500.5, 0.01, 8.98), (1000, 0.01, 2.5),
        (1001, 0.01, 0.025), (20000, 0.01, 0.025)]},
]

EXAMPLE = {
    "plot_kv": 0.48,
    "relay": [dict(en="Yes", name="R1 xfmr primary 51/50", kv=13.8, pickup=120,
                   curve="IEEE Extremely Inverse", td=1.5, inst=1400, instt=0.03)],
    "breaker": [
        dict(en="Yes", name="CB-Main 2000A", kv=0.48, rating=2000, ltpu=1.0, tr=2,
             stpu=2.5, tsd=0.3, i2t="On", i2tref=8, instpu=0, tinst=0.05,
             ptol=10, ttol=10),
        dict(en="Yes", name="CB-Feeder 800A", kv=0.48, rating=800, ltpu=1.0, tr=5,
             stpu=4, tsd=0.1, i2t="Off", i2tref=8, instpu=0, tinst=0.05,
             ptol=10, ttol=10)],
    "pt": [dict(en="Yes", name="CB-Motor 250A TM", kv=0.48,
                table="TM breaker 100A (illustrative)", scale=2.5),
           dict(en="Yes", name="FU-Lighting 200A", kv=0.48,
                table="Fuse 100A (illustrative)", scale=2.0)],
    "xfmr": [dict(en="Yes", name="T1 1500kVA", kva=1500, kv=13.8, z=5.75, cat="II",
                  conn=1.0, inrush=12, inrusht=0.1)],
    "cable": [dict(en="Yes", name="Feeder 500kcmil Cu", kv=0.48, mat="Cu",
                   area=253.4, nper=1, t1=75, t2=250)],
    "motor": [dict(en="Yes", name="M1 150hp", kv=0.48, fla=180, lram=6, tacc=3,
                   stall=12)],
    "fault": [dict(en="Yes", name="F1 480V bus", kv=0.48, amps=30000),
              dict(en="Yes", name="F2 feeder end", kv=0.48, amps=15000),
              dict(en="Yes", name="F3 MCC", kv=0.48, amps=8000)],
    "tables": POINT_TABLES,
    # (downstream device, upstream device, required margin in seconds)
    "pairs": [("CB-Motor 250A TM", "CB-Feeder 800A", 0),
              ("CB-Feeder 800A", "CB-Main 2000A", 0),
              ("CB-Main 2000A", "R1 xfmr primary 51/50", 0.3)],
}


def normalise(data):
    """Fill every slot, using defaults for the ones the study leaves out."""
    out = copy.deepcopy(data)
    out.setdefault("plot_kv", 0.48)
    out.setdefault("tables", POINT_TABLES)
    out.setdefault("pairs", [])
    for kind, n in SLOTS.items():
        rows = list(out.get(kind, []))
        if len(rows) > n:
            raise ValueError(f"too many {kind} devices: {len(rows)} > {n}")
        rows = [{**DEFAULTS[kind], **r} for r in rows]
        while len(rows) < n:
            d = dict(DEFAULTS[kind])
            d["name"] = f"{kind.capitalize()} {len(rows) + 1}"
            rows.append(d)
        out[kind] = rows
    return out


def evaluate(data, i_plot):
    """Operating time of every curve at a current given at the plot voltage.

    Returns {(kind, slot, column): seconds or None}, using the same keys as
    the Calc sheet layout returned by ``tcc.build.build``.
    """
    d = normalise(data)
    pkv = d["plot_kv"]
    res = {}
    for j, r in enumerate(d["relay"]):
        t = None
        if r["en"] == "Yes":
            _, a, b, p = c.RELAY_CURVES[r["curve"]]
            t = c.relay_time(c.refer(i_plot, pkv, r["kv"]), r["pickup"], a, b, p,
                             r["td"], r["inst"], r["instt"])
        res["relay", j, "t"] = t
    for j, b in enumerate(d["breaker"]):
        lo = hi = None
        if b["en"] == "Yes":
            ir = b["ltpu"] * b["rating"]
            lo, hi = c.lsig_band(c.refer(i_plot, pkv, b["kv"]), ir, b["tr"],
                                 b["stpu"] * ir, b["tsd"], b["i2t"] == "On",
                                 b["i2tref"], b["instpu"] * b["rating"], b["tinst"],
                                 b["ptol"], b["ttol"])
        res["breaker", j, "min"], res["breaker", j, "max"] = lo, hi
    tables = {t["name"]: t["points"] for t in d["tables"]}
    for j, f in enumerate(d["pt"]):
        lo = hi = None
        if f["en"] == "Yes" and f["table"] in tables and f["scale"] > 0:
            pts = tables[f["table"]]
            x = c.refer(i_plot, pkv, f["kv"]) / f["scale"]
            xs = [q[0] for q in pts]
            lo = c.loglog_interp(x, xs, [q[1] for q in pts])
            hi = c.loglog_interp(x, xs, [q[2] for q in pts])
        res["pt", j, "min"], res["pt", j, "max"] = lo, hi
    for j, x in enumerate(d["xfmr"]):
        t = None
        if x["en"] == "Yes":
            ib = x["kva"] / (c.SQRT3 * x["kv"])
            m = c.refer(i_plot, pkv, x["kv"]) / (ib * x["conn"])
            t = c.xfmr_damage(m, x["z"], x["cat"])
        res["xfmr", j, "t"] = t
    for j, k in enumerate(d["cable"]):
        t = None
        if k["en"] == "Yes":
            t = c.cable_damage(c.refer(i_plot, pkv, k["kv"]),
                               k["nper"] * k["area"] * c.CMIL_PER_MM2,
                               c.cable_k(k["mat"], k["t1"], k["t2"]))
        res["cable", j, "t"] = t
    return res

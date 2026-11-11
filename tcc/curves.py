"""Reference implementation of every curve model used by the workbook.

All functions take currents in amperes at the device's own voltage and return
a time in seconds, or ``None`` where the device does not operate (or the curve
is not defined). The Excel formulas written by ``tcc.build`` implement exactly
the same maths; ``tests/test_workbook.py`` checks the two against each other.
"""
from __future__ import annotations

import math

SQRT3 = math.sqrt(3.0)
T_FLOOR = 0.01          # bottom of the plot, seconds
CMIL_PER_MM2 = 1973.525

# Inverse-time relay curves: t = TD * (A / (M**p - 1) + B), M = I / pickup.
# name -> (source, A, B, p)
RELAY_CURVES = {
    "IEEE Moderately Inverse": ("IEEE C37.112", 0.0515, 0.114, 0.02),
    "IEEE Very Inverse": ("IEEE C37.112", 19.61, 0.491, 2.0),
    "IEEE Extremely Inverse": ("IEEE C37.112", 28.2, 0.1217, 2.0),
    "IEC Standard Inverse": ("IEC 60255-151", 0.14, 0.0, 0.02),
    "IEC Very Inverse": ("IEC 60255-151", 13.5, 0.0, 1.0),
    "IEC Extremely Inverse": ("IEC 60255-151", 80.0, 0.0, 2.0),
    "IEC Long-Time Inverse": ("IEC 60255-151", 120.0, 0.0, 1.0),
    "IEC Short-Time Inverse": ("IEC 60255-151", 0.05, 0.0, 0.04),
    "US U1 Moderately Inverse": ("SEL relay manuals", 0.0104, 0.0226, 0.02),
    "US U2 Inverse": ("SEL relay manuals", 5.95, 0.180, 2.0),
    "US U3 Very Inverse": ("SEL relay manuals", 3.88, 0.0963, 2.0),
    "US U4 Extremely Inverse": ("SEL relay manuals", 5.67, 0.0352, 2.0),
    "US U5 Short-Time Inverse": ("SEL relay manuals", 0.00342, 0.00262, 0.02),
    # Time dial is the operating time in seconds.
    "Definite Time": ("-", 0.0, 1.0, 1.0),
}

# IEEE C57.109 through-fault thermal curve: (times base current, seconds).
XFMR_THERMAL = [(2.0, 1800.0), (3.0, 300.0), (4.75, 60.0),
                (6.3, 30.0), (11.3, 10.0), (25.0, 2.0)]


def refer(current, kv_from, kv_to):
    """Refer a current from one voltage level to another (constant kVA)."""
    return current * kv_from / kv_to


def relay_time(i, pickup, a, b, p, td, inst_pickup=0.0, inst_time=0.02):
    """Overcurrent relay (51 with optional 50). Currents in primary amperes."""
    if pickup <= 0:
        return None
    m = i / pickup
    if m <= 1:
        return None
    t = td * (a / (m ** p - 1) + b)
    if inst_pickup > 0 and i >= inst_pickup:
        t = min(t, max(inst_time, T_FLOOR))
    return t


def lsig_time(i, ir, tr, isd, tsd, i2t, i2t_ref, ii, t_inst, pk=1.0, tk=1.0):
    """Electronic-trip (LSI) breaker.

    ir, isd, ii: long-time, short-time and instantaneous pickups in amperes
    (0 switches short-time or instantaneous off). tr is the long-time delay at
    6 x ir. With i2t on, the short-time delay ramps as I^2t until i2t_ref x ir.
    Above the short-time pickup the time never falls below the short-time delay.
    pk and tk scale the pickups and the delays; they are used for the band.
    """
    ir_, isd_, ii_ = ir * pk, isd * pk, ii * pk
    if i <= ir_:
        return None
    t = tr * tk * (6 * ir_ / i) ** 2
    if isd > 0 and i >= isd_:
        ramp = max(1.0, (i2t_ref * ir_ / i) ** 2) if i2t else 1.0
        t = min(max(t, tsd * tk), tsd * tk * ramp)
    if ii > 0 and i >= ii_:
        t = min(t, t_inst)
    return t


def lsig_band(i, ir, tr, isd, tsd, i2t, i2t_ref, ii, t_inst, ptol, ttol):
    """(minimum, maximum) operating time. Tolerances are in percent.

    In the instantaneous region the lower edge is drawn at the plot floor.
    """
    lo = lsig_time(i, ir, tr, isd, tsd, i2t, i2t_ref, ii, T_FLOOR,
                   1 - ptol / 100, 1 - ttol / 100)
    hi = lsig_time(i, ir, tr, isd, tsd, i2t, i2t_ref, ii, t_inst,
                   1 + ptol / 100, 1 + ttol / 100)
    return lo, hi


def loglog_interp(x, xs, ys):
    """Straight-line interpolation on log-log axes. None outside the table."""
    n = len(xs)
    if n < 2 or x < xs[0] or x > xs[-1]:
        return None
    k = 0
    for j in range(n):
        if xs[j] <= x:
            k = j
    k = min(k, n - 2)
    if ys[k] is None or ys[k + 1] is None or xs[k + 1] == xs[k]:
        return None
    lx0, lx1 = math.log10(xs[k]), math.log10(xs[k + 1])
    ly0, ly1 = math.log10(ys[k]), math.log10(ys[k + 1])
    return 10 ** (ly0 + (ly1 - ly0) * (math.log10(x) - lx0) / (lx1 - lx0))


def xfmr_thermal_slopes():
    """Log-log slope of each thermal segment; the last one continues I^2t = 1250."""
    s = []
    for (m0, t0), (m1, t1) in zip(XFMR_THERMAL, XFMR_THERMAL[1:]):
        s.append(math.log10(t1 / t0) / math.log10(m1 / m0))
    s.append(-2.0)
    return s


def xfmr_thermal(m):
    """Thermal through-fault limit for m times base current (m >= 2)."""
    if m < XFMR_THERMAL[0][0]:
        return None
    slopes = xfmr_thermal_slopes()
    k = max(j for j, (mm, _) in enumerate(XFMR_THERMAL) if mm <= m)
    m0, t0 = XFMR_THERMAL[k]
    return t0 * (m / m0) ** slopes[k]


def xfmr_damage(m, z_pct, category):
    """IEEE C57.109 through-fault damage curve.

    m is times base current, category is "I", "II", "III" or "IV". For
    categories II to IV the mechanical limit I^2t = 2 * Mmax^2 applies from
    70 % (II) or 50 % (III, IV) of the maximum fault current Mmax = 100 / Z%.
    """
    if z_pct <= 0:
        return None
    m_max = 100.0 / z_pct
    if m < 2 or m > m_max:
        return None
    th = xfmr_thermal(m)
    frac = {"I": None, "II": 0.7}.get(category, 0.5)
    if frac is not None and m >= frac * m_max:
        return min(th, 2 * (m_max / m) ** 2)
    return th


def cable_k(material, t1, t2):
    """Short-circuit constant: (I / A)^2 * t = K, A in circular mils."""
    if material == "Cu":
        return 0.0297 * math.log10((t2 + 234) / (t1 + 234))
    return 0.0125 * math.log10((t2 + 228) / (t1 + 228))


def cable_damage(i, area_cmil, k, t_lo=0.01, t_hi=10.0):
    """Conductor short-circuit heating limit, shown between t_lo and t_hi."""
    if i <= 0:
        return None
    t = k * (area_cmil / i) ** 2
    return t if t_lo <= t <= t_hi else None

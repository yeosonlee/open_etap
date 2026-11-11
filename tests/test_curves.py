import math

import pytest

from tcc import curves as c


def relay(name, m, td=1.0):
    _, a, b, p = c.RELAY_CURVES[name]
    return c.relay_time(m * 100, 100, a, b, p, td)


def test_relay_known_values():
    assert relay("IEC Standard Inverse", 10) == pytest.approx(2.971, abs=1e-3)
    assert relay("IEC Very Inverse", 10) == pytest.approx(1.5)
    assert relay("IEC Extremely Inverse", 10) == pytest.approx(80 / 99)
    assert relay("IEC Long-Time Inverse", 5) == pytest.approx(30.0)
    assert relay("IEEE Very Inverse", 5) == pytest.approx(1.308, abs=1e-3)
    assert relay("IEEE Extremely Inverse", 5, td=2) == pytest.approx(2 * (28.2 / 24 + 0.1217))
    assert relay("Definite Time", 3, td=0.4) == pytest.approx(0.4)


def test_relay_no_trip_and_instantaneous():
    assert relay("IEC Standard Inverse", 1.0) is None
    assert relay("IEC Standard Inverse", 0.5) is None
    _, a, b, p = c.RELAY_CURVES["IEC Standard Inverse"]
    assert c.relay_time(999, 100, a, b, p, 1, inst_pickup=1000, inst_time=0.03) > 2
    assert c.relay_time(1000, 100, a, b, p, 1, inst_pickup=1000, inst_time=0.03) == 0.03
    assert c.relay_time(1000, 100, a, b, p, 1, inst_pickup=1000, inst_time=0) == c.T_FLOOR


def test_lsig_segments():
    kw = dict(ir=1000, tr=5, isd=4000, tsd=0.2, i2t=False, i2t_ref=8, ii=12000, t_inst=0.05)
    assert c.lsig_time(1000, **kw) is None
    assert c.lsig_time(6000, **{**kw, "isd": 0}) == pytest.approx(5)
    assert c.lsig_time(2000, **kw) == pytest.approx(45)
    assert c.lsig_time(5000, **kw) == pytest.approx(0.2)
    assert c.lsig_time(12000, **kw) == pytest.approx(0.05)
    # the long-time I²t line must not pull the curve below the short-time delay
    assert c.lsig_time(200000, **{**kw, "ii": 0}) == pytest.approx(0.2)
    ramp = {**kw, "i2t": True}
    assert c.lsig_time(4000, **ramp) == pytest.approx(0.8)      # 0.2 * (8000/4000)^2
    assert c.lsig_time(8000, **ramp) == pytest.approx(0.2)
    assert c.lsig_time(10000, **ramp) == pytest.approx(0.2)


def test_lsig_band_brackets_nominal():
    kw = dict(ir=1000, tr=5, isd=4000, tsd=0.2, i2t=True, i2t_ref=8, ii=12000, t_inst=0.05)
    for i in (1200, 2500, 5000, 9000, 20000):
        lo, hi = c.lsig_band(i, ptol=10, ttol=10, **kw)
        nom = c.lsig_time(i, **kw)
        assert lo <= nom <= hi
    assert c.lsig_band(1050, ptol=10, ttol=10, **kw) == (pytest.approx(4.5 * (5400 / 1050) ** 2), None)


def test_loglog_interp():
    xs, ys = [10, 100, 1000], [100, 1, 0.01]
    assert c.loglog_interp(10, xs, ys) == pytest.approx(100)
    assert c.loglog_interp(1000, xs, ys) == pytest.approx(0.01)
    assert c.loglog_interp(math.sqrt(1000), xs, ys) == pytest.approx(10)
    assert c.loglog_interp(9.9, xs, ys) is None
    assert c.loglog_interp(1001, xs, ys) is None


def test_transformer_damage():
    for m, t in c.XFMR_THERMAL:
        assert c.xfmr_thermal(m) == pytest.approx(t)
    assert c.xfmr_thermal(40) == pytest.approx(1250 / 40 ** 2)
    # Category I follows the thermal curve to the maximum fault current.
    assert c.xfmr_damage(25, 4, "I") == pytest.approx(2)
    assert c.xfmr_damage(25.1, 4, "I") is None
    assert c.xfmr_damage(1.9, 4, "I") is None
    # Category II, Z = 5.75 %: mechanical part from 70 % of max fault, 2 s at max.
    m_max = 100 / 5.75
    assert c.xfmr_damage(m_max, 5.75, "II") == pytest.approx(2)
    assert c.xfmr_damage(0.7 * m_max, 5.75, "II") == pytest.approx(2 / 0.49)
    assert c.xfmr_damage(0.69 * m_max, 5.75, "II") == pytest.approx(c.xfmr_thermal(0.69 * m_max))
    assert c.xfmr_damage(0.5 * m_max, 5.75, "III") == pytest.approx(8)


def test_cable_damage():
    k = c.cable_k("Cu", 75, 250)
    assert k == pytest.approx(0.0297 * math.log10(484 / 309))
    # 500 kcmil copper at 30 kA
    assert c.cable_damage(30000, 500000, k) == pytest.approx(k * (500000 / 30000) ** 2)
    assert c.cable_damage(1000, 500000, k) is None       # beyond 10 s
    assert c.cable_k("Al", 75, 250) == pytest.approx(0.0125 * math.log10(478 / 303))


def test_refer():
    assert c.refer(1000, 13.8, 0.48) == pytest.approx(28750)

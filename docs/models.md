# Curve models

All currents are in amperes at the device's own voltage. To plot, a current is
referred to the plot voltage at constant kVA: `I_plot = I_device × kV_device / kV_plot`.

## Overcurrent relays

`M = I / pickup`. The relay operates for `M > 1`:

    t = TD × (A / (M^p − 1) + B)

| Curve | A | B | p | Source |
|---|---|---|---|---|
| IEEE Moderately Inverse | 0.0515 | 0.114 | 0.02 | IEEE C37.112 |
| IEEE Very Inverse | 19.61 | 0.491 | 2 | IEEE C37.112 |
| IEEE Extremely Inverse | 28.2 | 0.1217 | 2 | IEEE C37.112 |
| IEC Standard Inverse | 0.14 | 0 | 0.02 | IEC 60255-151 |
| IEC Very Inverse | 13.5 | 0 | 1 | IEC 60255-151 |
| IEC Extremely Inverse | 80 | 0 | 2 | IEC 60255-151 |
| IEC Long-Time Inverse | 120 | 0 | 1 | IEC 60255-151 |
| IEC Short-Time Inverse | 0.05 | 0 | 0.04 | IEC 60255-151 |
| US U1 Moderately Inverse | 0.0104 | 0.0226 | 0.02 | SEL manuals |
| US U2 Inverse | 5.95 | 0.180 | 2 | SEL manuals |
| US U3 Very Inverse | 3.88 | 0.0963 | 2 | SEL manuals |
| US U4 Extremely Inverse | 5.67 | 0.0352 | 2 | SEL manuals |
| US U5 Short-Time Inverse | 0.00342 | 0.00262 | 0.02 | SEL manuals |
| Definite Time | 0 | 1 | 1 | TD is the time in seconds |

Time-dial scales differ between makes. With the IEEE constants, a time dial of
1 is roughly a US-curve time dial of 5. Check the equation in the relay manual
and add the curve to `Lib_Relay` if it differs.

The instantaneous element (50) operates at its own time once the current
reaches its pickup. Times are floored at 0.01 s, the bottom of the plot.

## LV breakers with electronic trip (LSI)

Pickups: `Ir = long-time pickup × In`, `Isd = short-time pickup × Ir`,
`Ii = instantaneous pickup × In`.

- No trip for `I ≤ Ir`.
- Long-time: `t = tr × (6 Ir / I)²`, where `tr` is the delay at 6 × Ir.
- Short-time, for `I ≥ Isd`: `t = tsd`, or with I²t on,
  `t = tsd × max(1, (k Ir / I)²)` where `k` is the multiple of Ir at which the
  ramp ends. The time never falls below `tsd`.
- Instantaneous, for `I ≥ Ii`: the clearing time entered.

The band scales pickups by `1 ± pickup tolerance` and delays by
`1 ± time tolerance`. In the instantaneous region the lower edge is drawn at
0.01 s. Trip units differ in their I²t reference point and tolerances; set
them from the manufacturer's curve.

## Point-table devices

Fuses (minimum melt and total clearing) and thermal-magnetic breakers have no
equation. Each table lists current, minimum time and maximum time. Between
points the workbook interpolates in a straight line on log-log axes. Outside
the table there is no value. The current multiplier scales a table to another
rating of the same shape.

## Transformer through-fault damage (IEEE C57.109)

`M` is times base current, `Mmax = 100 / Z%`. The curve runs from `M = 2` to `Mmax`.

- Thermal: through (2, 1800 s), (3, 300 s), (4.75, 60 s), (6.3, 30 s),
  (11.3, 10 s), (25, 2 s), then `t = 1250 / M²`.
- Mechanical, categories II to IV: `t = 2 × (Mmax / M)²`, from 70 % of Mmax
  (category II) or 50 % (III and IV). The curve is the lower of the two.
- Category I uses the thermal curve only.

The step where the mechanical part begins is drawn as a vertical drop. The
connection shift factor moves the curve left for faults that the primary
device sees at reduced current (0.87 delta-delta line-to-line, 0.58 delta-wye
ground fault). Markers show full-load current and the inrush point.

## Cable short-circuit damage

    (I / A)² × t = 0.0297 × log10((T2 + 234) / (T1 + 234))    copper
    (I / A)² × t = 0.0125 × log10((T2 + 228) / (T1 + 228))    aluminium

`A` is the conductor area in circular mils (1 mm² = 1973.525 cmil), `T1` the
initial and `T2` the final conductor temperature in °C. Shown from 0.01 s to 10 s.

## Motor starting

Full-load current down to the acceleration time, across to locked-rotor
current, then down. A marker shows the safe stall time at locked-rotor current.

## Verification

`tests/test_curves.py` checks the maths against values from the standards.
`tests/test_workbook.py` builds the workbook for two sets of settings,
recalculates it in LibreOffice, and compares every curve cell (about 500
currents per curve) with the Python reference.

#!/usr/bin/env python3
"""Compare the fetched data in data/*.json with the figures quoted in the project brief.

The "expected" numbers below are the brief's claims, written down only so they can be checked.
The site never reads them; it reads data/*.json. Prints a table and flags every mismatch.
"""
import json
from pathlib import Path

D = Path(__file__).resolve().parent.parent / "data"
load = lambda n: json.loads((D / n).read_text())

rents, cmhc, own, fert = load("statcan_rents.json"), load("cmhc_turnover.json"), load("homeownership.json"), load("fertility.json")
two = {p["quarter"]: p for p in rents["series"]["2br"]}
peak = max(rents["series"]["2br"], key=lambda p: p["asking"])
peak_qs = [p["quarter"] for p in rents["series"]["2br"] if p["asking"] == peak["asking"]]
size = {r["size"]: r for r in cmhc["by_size"]}
ages = {g["age"]: g for g in own["groups"]}
ont = {y["year"]: y for y in fert["ontario"]}
ch = fert["asfr_20_24_change"]

rows = [
    ("A", "2-bed asking, peak (H2 2023)", 2920, peak["asking"], f"{' & '.join(peak_qs)}"),
    ("A", "2-bed asking, Q3 2025", 2720, two["2025 Q3"]["asking"], ""),
    ("A", "2-bed asking, Q1 2026", 2660, two["2026 Q1"]["asking"], ""),
    ("A", "2-bed asking, Q2 2026", 2650, two["2026 Q2"]["asking"], ""),
    ("A", "2-bed paid, Q2 2026", 2160, two["2026 Q2"]["paid"], "in 46-10-0092-01 itself"),
    ("A", "asking minus paid, Q2 2026 (site text)", 490, two["2026 Q2"]["asking"] - two["2026 Q2"]["paid"], ""),
]
for s, (n, st) in {"Studio": (1708, 1443), "1-bedroom": (2073, 1711), "2-bedroom": (2547, 1984), "3-bedroom+": (2893, 2263)}.items():
    rows.append(("B", f"{s} new tenant, Oct 2025", n, size[s]["new_tenant"], ""))
    rows.append(("B", f"{s} sitting tenant, Oct 2025", st, size[s]["sitting_tenant"], ""))
rows.append(("B", "1-bed gap (hero '$362')", 362, size["1-bedroom"]["gap"], ""))
for a, (x, y) in {"25 to 29": (44.1, 36.5), "30 to 34": (59.2, 52.3), "70 to 74": (75.5, 74.8)}.items():
    rows.append(("C", f"ages {a}, 2011", x, ages[a]["rate_2011"], "from Daily text"))
    rows.append(("C", f"ages {a}, 2021", y, ages[a]["rate_2021"], "from Daily text"))
rows.append(("D", "Ontario TFR 2025", 1.23, ont[2025]["tfr"], ""))
rows.append(("D", "20-24 rate change 2021-2025 ('about a fifth')", "~-20%", f"{ch['pct']}%",
             f"{ch['from']} -> {ch['to']} per 1,000 (a rate, not a birth count)"))

print(f"{'Set':<4}{'Figure':<46}{'Brief':>9}{'Fetched':>10}  Status  Note")
print("-" * 100)
bad = 0
for k, label, exp, got, note in rows:
    if isinstance(exp, str):
        ok = abs(float(str(got).rstrip("%")) + 20) <= 1
    else:
        ok = got == exp
    bad += not ok
    print(f"{k:<4}{label:<46}{str(exp):>9}{str(got):>10}  {'ok' if ok else 'MISMATCH':<7} {note}")
print("-" * 100)
print(f"{len(rows)} figures checked, {bad} mismatches")

#!/usr/bin/env python3
"""Fetch every dataset the Rent Reset site uses, straight from the official sources.

Usage:
    python scripts/fetch_data.py --sets all        # A B C D E F
    python scripts/fetch_data.py --sets F          # shelters only (daily job)
    python scripts/fetch_data.py --sets A,B,D,E    # weekly job

Datasets:
    A  StatCan 46-10-0092-01  asking and paid rent, Toronto CMA, 1- and 2-bedroom apartments
    B  CMHC Rental Market Survey, Table 6.0 turnover vs non-turnover rents, Toronto CMA
    C  StatCan The Daily, 2021 Census homeownership by age (numbers read from the release text)
    D  StatCan 13-10-0418-01  Ontario total fertility rate and age-specific rate 20-24
    E  Tribunals Ontario, LTB Applications Received by e-File (quarterly, all filing methods)
    F  Toronto Open Data, shelter system flow (monthly) and daily shelter occupancy

Rules this script follows:
    * Raw downloads are saved untouched in data/raw/. Processed files go in data/.
    * Nothing is written for a dataset unless the whole dataset fetched and parsed cleanly,
      so a failed fetch never wipes or half-overwrites the previous files.
    * Files are only rewritten when the processed numbers actually change, so the raw files
      in data/raw/ are always the ones that produced the processed files next to them.
    * Any failure is printed loudly and the script exits with status 1 after trying the
      remaining datasets.
"""

import argparse
import csv
import datetime as dt
import hashlib
import html
import io
import json
import os
import re
import sys
import time
import traceback
import zipfile
from pathlib import Path

import openpyxl
import pandas as pd
import requests

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
RAW = DATA / "raw"
MANIFEST = DATA / "manifest.json"

SESSION = requests.Session()
SESSION.headers["User-Agent"] = "Mozilla/5.0 (compatible; rent-gap-site-data-fetch/1.0)"

STATCAN_WDS = "https://www150.statcan.gc.ca/t1/wds/rest/"
CMHC_PAGE = ("https://www.cmhc-schl.gc.ca/professionals/housing-markets-data-and-research/"
             "housing-data/data-tables/rental-market/rental-market-report-data-tables")
CMHC_API = "https://www.cmhc-schl.gc.ca/api/Sitecore/PubsAndReports/"
CMHC_DATASOURCE = ("/sitecore/content/CMHC/Sites/Main/Home/professionals/housing-markets-data-and-research/"
                   "housing-data/data-tables/rental-market/rental-market-report-data-tables")
CMHC_CANADA_ID = "{9EE6E91C-4719-412C-909E-A4B27F3FB16E}"  # "Canada" in the page's geography dropdown
CENSUS_DAILY = "https://www150.statcan.gc.ca/n1/daily-quotidien/220921/dq220921b-eng.htm"
LTB_OPEN_DATA = "https://tribunalsontario.ca/en/about/open-data/"
CKAN = "https://ckan0.cf.opendata.inter.prod-toronto.ca/api/3/action/"
SHELTER_FLOW_RESOURCE = "fb32967d-ee16-4d46-a880-0614c18f2137"
OCCUPANCY_PACKAGE = "daily-shelter-overnight-service-occupancy-capacity"


class FetchError(Exception):
    pass


# ---------------------------------------------------------------- helpers

def get(url, **kw):
    """GET with three tries. Raises FetchError on a non-200 answer."""
    last = None
    for attempt in range(3):
        try:
            r = SESSION.get(url, timeout=120, **kw)
            if r.status_code == 200:
                return r
            last = f"HTTP {r.status_code}"
        except requests.RequestException as e:
            last = str(e)
        time.sleep(3 * (attempt + 1))
    raise FetchError(f"GET {url} failed: {last}")


def post_json(url, payload):
    r = SESSION.post(url, json=payload, timeout=120)
    if r.status_code != 200:
        raise FetchError(f"POST {url} failed: HTTP {r.status_code}")
    return r.json()


def today():
    return dt.date.today().isoformat()


def num(v):
    """'1,708' -> 1708, 2405 -> 2405, '**' or '' or None -> None."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return v
    s = str(v).replace(",", "").replace("$", "").strip()
    try:
        return float(s) if "." in s else int(s)
    except ValueError:
        return None


def page_text(raw_html):
    t = re.sub(r"<script.*?</script>|<style.*?</style>", " ", raw_html, flags=re.S)
    t = html.unescape(re.sub(r"<[^>]+>", " ", t))
    return re.sub(r"\s+", " ", t)


def statcan_table(pid):
    """Download a full StatCan table through the WDS. Returns (raw files, DataFrame, release date)."""
    meta = post_json(STATCAN_WDS + "getCubeMetadata", [{"productId": pid}])[0]
    if meta.get("status") != "SUCCESS":
        raise FetchError(f"getCubeMetadata {pid}: {meta}")
    link = get(f"{STATCAN_WDS}getFullTableDownloadCSV/{pid}/en").json()
    if link.get("status") != "SUCCESS":
        raise FetchError(f"getFullTableDownloadCSV {pid}: {link}")
    zbytes = get(link["object"]).content
    with zipfile.ZipFile(io.BytesIO(zbytes)) as z:
        df = pd.read_csv(z.open(f"{pid}.csv"), dtype=str)
    obj = meta["object"]
    raw = {f"{pid}-eng.zip": zbytes,
           f"{pid}_cube_metadata.json": json.dumps(meta, indent=1).encode()}
    info = {"product_id": pid, "title": obj["cubeTitleEn"], "release_date": obj["releaseTime"][:10],
            "zip_url": link["object"], "start": obj["cubeStartDate"], "end": obj["cubeEndDate"]}
    return raw, df, info


# ---------------------------------------------------------------- A. StatCan asking + paid rents

def fetch_A():
    pid = 46100092
    raw, df, info = statcan_table(pid)
    geo = "Toronto, Census metropolitan area (CMA)"
    units = {"1br": "Apartment - 1 bedroom", "2br": "Apartment - 2 bedrooms"}
    t = df[(df["GEO"] == geo) & (df["REF_DATE"] >= "2019-01")]
    if t.empty:
        raise FetchError(f"46-10-0092-01: no rows for {geo}")
    series, rows = {}, []
    for key, unit in units.items():
        u = t[t["Rental unit type"] == unit]
        pv = u.pivot(index="REF_DATE", columns="Estimates", values="VALUE").sort_index()
        for col in ("Average asking rent", "Average paid rent"):
            if col not in pv.columns:
                raise FetchError(f"46-10-0092-01: '{col}' missing for {unit}")
        pts = []
        for ref, r in pv.iterrows():
            q = (int(ref[5:7]) - 1) // 3 + 1
            p = {"ref_date": ref, "quarter": f"{ref[:4]} Q{q}",
                 "asking": num(r["Average asking rent"]), "paid": num(r["Average paid rent"])}
            pts.append(p)
            rows.append({"unit": unit, **p})
        series[key] = pts
    info.update(geography=geo, unit_types=list(units.values()), from_period="2019-01")
    out = {"meta": info, "series": series}
    return raw, {"statcan_rents.json": out, "statcan_rents.csv": rows}, info


# ---------------------------------------------------------------- B. CMHC turnover vs non-turnover

def cmhc_editions():
    """Ask the CMHC page's own API which Canada workbooks exist. Newest first."""
    eds = get(CMHC_API + "GetEditionList", params=dict(
        cityId=CMHC_CANADA_ID, dataSource=CMHC_DATASOURCE, contextLanguage="en")).json()
    out = []
    for e in eds:
        f = get(CMHC_API + "GetFileDetails", params=dict(
            cityId=CMHC_CANADA_ID, edition="{" + e["Itemid"].upper() + "}",
            dataSource=CMHC_DATASOURCE, contextLanguage="en")).json()
        out.append({"edition": e["Editions"], "published": f.get("DatePublished"), "url": f.get("DocumentUrl")})
    return sorted(out, key=lambda x: x["edition"], reverse=True)


def period_label(v):
    """Header cell -> 'YYYY-10'. Cells are 'Oct-25' strings or datetime objects."""
    if isinstance(v, dt.datetime):
        return v.strftime("%Y-%m")
    m = re.fullmatch(r"([A-Z][a-z]{2})-(\d{2})", str(v).strip())
    if m:
        return dt.datetime.strptime(m.group(0), "%b-%y").strftime("%Y-%m")
    return None


def parse_table_60(xbytes):
    """Return {(group, size, period): (value, quality)} for the Toronto CMA row of Table 6.0."""
    wb = openpyxl.load_workbook(io.BytesIO(xbytes), read_only=True, data_only=True)
    sheet = None
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        title = " ".join(str(c) for r in rows[:4] for c in r if c)
        if "Turnover and Non-turnover Units" in title and "Average" in title:
            sheet = (ws.title, title.strip(), rows)
            break
    if not sheet:
        raise FetchError("CMHC workbook: no 'Average Rent ... Turnover and Non-turnover Units' sheet")
    name, title, rows = sheet
    hdr_i = next(i for i, r in enumerate(rows) if r and r[0] == "Centre")
    hdr = rows[hdr_i]
    tor = [r for r in rows if r and isinstance(r[0], str) and r[0].strip() == "Toronto CMA"]
    if len(tor) != 1:
        raise FetchError(f"CMHC {name}: expected one 'Toronto CMA' row, found {len(tor)}")
    tor = tor[0]
    pcols = [(c, period_label(v)) for c, v in enumerate(hdr) if c > 0 and period_label(v)]
    sizes_row = rows[hdr_i - 2] if hdr_i >= 2 else ()
    labels = [(c, str(v).strip()) for c, v in enumerate(sizes_row) if v]
    cells = {}
    if "Two Bedrooms Only" in title:
        half = len(pcols) // 2
        for k, (c, p) in enumerate(pcols):
            g = "turnover" if k < half else "non_turnover"
            cells[(g, "2 Bedroom", p)] = (num(tor[c]), tor[c + 1] if c + 1 < len(tor) else None)
    else:
        if len(labels) != 10:
            raise FetchError(f"CMHC {name}: expected 10 bedroom-type labels, found {labels}")
        for c, p in pcols:
            j = min(range(len(labels)), key=lambda k: abs(labels[k][0] - c))
            g = "turnover" if j < 5 else "non_turnover"
            cells[(g, labels[j][1], p)] = (num(tor[c]), tor[c + 1] if c + 1 < len(tor) else None)
    return name, title, cells


def fetch_B():
    eds = cmhc_editions()
    if not eds:
        raise FetchError("CMHC: edition list is empty")
    raw, parsed, without = {}, [], []
    for i, e in enumerate(eds):  # every edition that has the table (2022 edition onward so far)
        xb = get(e["url"]).content
        fname = f"rmr-canada-{e['edition']}-en.xlsx"
        try:
            sheet, title, cells = parse_table_60(xb)
        except FetchError:
            if i == 0:
                raise  # the newest edition must have it
            without.append(e["edition"])
            continue
        raw[fname] = xb
        parsed.append((e, fname, sheet, title, cells))
    raw["cmhc_editions.json"] = json.dumps(eds, indent=1).encode()

    latest, fname, sheet, title, cells = parsed[0]
    periods = sorted({p for (_, _, p) in cells})
    cur = periods[-1]
    sizes = [("Studio", "Studio"), ("1 Bedroom", "1-bedroom"), ("2 Bedroom", "2-bedroom"), ("3 Bedroom +", "3-bedroom+")]
    by_size = []
    for src, label in sizes:
        n, ql_n = cells.get(("turnover", src, cur), (None, None))
        s, ql_s = cells.get(("non_turnover", src, cur), (None, None))
        if n is None or s is None:
            raise FetchError(f"CMHC {latest['edition']}: missing {src} for {cur}")
        by_size.append({"size": label, "new_tenant": n, "sitting_tenant": s, "gap": n - s,
                        "quality_new": ql_n, "quality_sitting": ql_s})

    # 2-bedroom series across editions; the newest edition wins for any October it covers
    two = {}
    for e, f, _, _, c in reversed(parsed):
        for (g, size, p), (v, _) in c.items():
            if size == "2 Bedroom" and v is not None:
                two.setdefault(p, {})[g] = (v, e["edition"])
    series = [{"period": p, "new_tenant": d["turnover"][0], "sitting_tenant": d["non_turnover"][0],
               "edition": d["turnover"][1]}
              for p, d in sorted(two.items()) if "turnover" in d and "non_turnover" in d]

    info = {"edition": latest["edition"], "release_date": latest["published"], "period": cur,
            "workbook_url": latest["url"], "sheet": sheet, "table_title": title,
            "geography": "Toronto CMA",
            "universe": "Privately initiated apartment structures of three units and over",
            "editions_used": [{"edition": e["edition"], "published": e["published"], "url": e["url"],
                               "sheet": s, "table_title": t} for e, _, s, t, _ in parsed],
            "editions_without_table": without,
            "quality_codes": "a Excellent, b Very good, c Good, d Poor (use with caution)"}
    rows = [{"period": cur, **r} for r in by_size]
    out = {"meta": info, "by_size": by_size, "two_bedroom_series": series}
    return raw, {"cmhc_turnover.json": out, "cmhc_turnover.csv": rows}, info


# ---------------------------------------------------------------- C. Census homeownership by age

def fetch_C():
    body = get(CENSUS_DAILY).text
    txt = page_text(body)
    pats = {
        "25 to 29": r"aged 25 to 29 years fell from ([\d.]+)% in 2011 to ([\d.]+)% in 2021",
        "30 to 34": r"aged 30 to 34 years, homeownership rates fell from ([\d.]+)% in 2011 to ([\d.]+)% in 2021",
        "70 to 74": r"aged 70 to 74 years, for example, the homeownership rate fell from ([\d.]+)% to ([\d.]+)%",
    }
    groups = []
    for age, pat in pats.items():
        m = re.search(pat, txt)
        if not m:
            raise FetchError(f"Census Daily: sentence for ages {age} not found; the page may have changed")
        groups.append({"age": age, "rate_2011": float(m.group(1)), "rate_2021": float(m.group(2)),
                       "quote": m.group(0)})
    info = {"url": CENSUS_DAILY, "release_date": "2022-09-21", "geography": "Canada",
            "note": ("No data table with these rates is linked from the release (it links 46-10-0053-01, "
                     "property ownership, and 34-10-0066-01, building permits), so the rates are read "
                     "from the release text.")}
    out = {"meta": info, "groups": groups}
    return {"dq220921b-eng.htm": body.encode("utf-8")}, {"homeownership.json": out}, info


# ---------------------------------------------------------------- D. Fertility

def fetch_D():
    pid = 13100418
    raw, df, info = statcan_table(pid)
    geo = "Ontario, place of residence of mother"
    gcol = [c for c in df.columns if c.startswith("GEO")][0]
    o = df[(df[gcol] == geo) & (df["REF_DATE"] >= "2021")]
    want = {"tfr": "Total fertility rate per female",
            "asfr_20_24": "Age-specific fertility rate, females 20 to 24 years"}
    pv = o.pivot(index="REF_DATE", columns="Characteristics", values="VALUE").sort_index()
    for k, col in want.items():
        if col not in pv.columns:
            raise FetchError(f"13-10-0418-01: '{col}' missing for Ontario")
    years = [{"year": int(y), "tfr": num(r[want["tfr"]]), "asfr_20_24": num(r[want["asfr_20_24"]])}
             for y, r in pv.iterrows()]
    first, last = years[0], years[-1]
    change = round((last["asfr_20_24"] - first["asfr_20_24"]) / first["asfr_20_24"] * 100, 1)
    info.update(geography=geo, from_year=first["year"], to_year=last["year"],
                units="TFR: children per female; age-specific rate: live births per 1,000 females aged 20 to 24")
    out = {"meta": info, "ontario": years,
           "asfr_20_24_change": {"from_year": first["year"], "to_year": last["year"],
                                 "from": first["asfr_20_24"], "to": last["asfr_20_24"], "pct": change}}
    return raw, {"fertility.json": out, "fertility.csv": years}, info


# ---------------------------------------------------------------- E. Landlord and Tenant Board

def parse_header_dates(s):
    m = re.search(r"\(([^)]*?)\s*-\s*([^)]*?)\)", s)
    if not m:
        return None, None
    out = []
    for part in m.groups():
        part = part.replace("Sept ", "Sep ")
        d = pd.to_datetime(part, errors="coerce")
        out.append(None if pd.isna(d) else d.date().isoformat())
    return out


def fetch_E():
    page = get(LTB_OPEN_DATA).text
    links = []
    for href, inner in re.findall(r'<a[^>]*href="([^"]+)"[^>]*>(.*?)</a>', page, flags=re.S):
        text = re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", "", inner))).strip()
        m = re.fullmatch(r"LTB Applications Received by e-File\s*[–-]\s*(.+)", text)
        if m and href.lower().endswith(".xlsx"):
            links.append((href, m.group(1)))
    if not links:
        raise FetchError("Tribunals Ontario open data page: no 'LTB Applications Received by e-File' links")
    raw, quarters, skipped = {"tribunals_open_data.htm": page.encode("utf-8")}, [], []
    for href, label in links:
        href = html.unescape(href)
        url = requests.compat.urljoin("https://tribunalsontario.ca/", href.replace(" ", "%20"))
        fname = requests.utils.unquote(href.rsplit("/", 1)[-1])
        dates = re.search(r"(\d{4}-\d{1,2}-\d{1,2})_(\d{4}-\d{1,2}-\d{1,2})", fname)
        xb = get(url).content
        raw["ltb_efile/" + fname] = xb
        ws = openpyxl.load_workbook(io.BytesIO(xb), read_only=True).worksheets[0]
        rows = [r for r in ws.iter_rows(values_only=True) if any(c is not None for c in r)]
        head = str(rows[0][0])
        hdr_i = next((i for i, r in enumerate(rows) if any(str(c).strip().lower().startswith("total apps received") for c in r if c)), None)
        totals_cols = [c for c, v in enumerate(rows[hdr_i]) if v and str(v).strip().lower().startswith("total apps received")] if hdr_i is not None else []
        if len(totals_cols) != 1 or not dates:
            skipped.append({"label": label, "file": fname, "reason": "multi-quarter or unrecognised layout; kept in raw only"})
            continue
        tc = totals_cols[0]
        counts = {}
        for r in rows[hdr_i + 1:]:
            k = str(r[0]).strip() if r[0] else ""
            if k in ("L1", "L2", "T2", "T6", "Other"):
                counts[k] = num(r[tc])
        if counts.get("L1") is None:
            raise FetchError(f"LTB {label}: no L1 count")
        start = pd.to_datetime(dates.group(1)).date()
        end = pd.to_datetime(dates.group(2)).date()
        h_start, h_end = parse_header_dates(head)
        notes = []
        if h_end and h_end < end.isoformat():
            notes.append(f"partial quarter: sheet header covers {h_start} to {h_end}")
        elif (h_start and h_start != start.isoformat()) or (h_end and h_end != end.isoformat()):
            notes.append(f"sheet header says '{head}', file name says {start} to {end}; file name used")
        fy = start.year if start.month >= 4 else start.year - 1
        fq = ((start.month - 4) % 12) // 3 + 1
        quarters.append({"fiscal_label": f"Q{fq} {fy}-{str(fy + 1)[2:]}", "source_label": label,
                         "start": start.isoformat(), "end": end.isoformat(),
                         "calendar": f"{start.strftime('%b')}–{end.strftime('%b')} {end.year}",
                         "L1": counts.get("L1"), "L2": counts.get("L2"), "T2": counts.get("T2"),
                         "T6": counts.get("T6"), "Other": counts.get("Other"),
                         "header": head, "note": "; ".join(notes) or None, "file": fname})
    quarters.sort(key=lambda q: q["start"])
    latest = quarters[-1]
    info = {"page": LTB_OPEN_DATA, "dataset": "LTB – Applications Received by e-File",
            "geography": "Ontario (province-wide; no regional split in this dataset)",
            "granularity": "quarterly (Ontario government fiscal quarters, April–March year)",
            "measure": "total applications received, all filing methods (column 'total apps received - all filing methods')",
            "first_quarter": quarters[0]["start"], "latest_quarter_end": latest["end"],
            "latest_fiscal_label": latest["fiscal_label"], "skipped_files": skipped,
            "bill60_effective": "2026-09-21",
            "notes": ["'Other' application types are only reported from Jan–Mar 2022 onward, so a total of all "
                      "rows is not comparable over time. L1 is reported in every quarter.",
                      "Numbers are subject to change (Tribunals Ontario note in every file)."]}
    rows = [{k: q[k] for k in ("fiscal_label", "start", "end", "L1", "L2", "T2", "T6", "Other", "note")} for q in quarters]
    out = {"meta": info, "quarters": quarters}
    return raw, {"ltb_applications.json": out, "ltb_applications.csv": rows}, info


# ---------------------------------------------------------------- F. Toronto shelters

def ckan(action, **params):
    r = get(CKAN + action, params=params).json()
    if not r.get("success"):
        raise FetchError(f"CKAN {action} failed: {r.get('error')}")
    return r["result"]


def fetch_F():
    # Monthly system flow, every record, paginated
    pages, records, offset = [], [], 0
    while True:
        res = ckan("datastore_search", resource_id=SHELTER_FLOW_RESOURCE, limit=500, offset=offset)
        pages.append(res)
        records += res["records"]
        offset += 500
        if offset >= res["total"]:
            break
    if len(records) != res["total"]:
        raise FetchError(f"Shelter flow: got {len(records)} of {res['total']} records")
    flow_pkg = ckan("package_show", id="toronto-shelter-system-flow")
    months = []
    for r in records:
        if r["population_group"] in ("All Population", "Families"):
            m = dt.datetime.strptime(r["date(mmm-yy)"], "%b-%y").strftime("%Y-%m")
            months.append({"month": m, "population_group": r["population_group"],
                           **{k: v for k, v in r.items() if k not in ("_id", "date(mmm-yy)", "population_group")}})
    months.sort(key=lambda x: (x["month"], x["population_group"]))

    # Daily occupancy: the current-year datastore resource, latest date only
    pkg = ckan("package_show", id=OCCUPANCY_PACKAGE)
    cur = [x for x in pkg["resources"] if x.get("datastore_active") and not re.search(r"\d{4}$", x["name"])]
    if len(cur) != 1:
        raise FetchError(f"Occupancy: expected one current-year datastore resource, found {[x['name'] for x in cur]}")
    rid = cur[0]["id"]
    newest = ckan("datastore_search", resource_id=rid, limit=1, sort="OCCUPANCY_DATE desc")
    day = newest["records"][0]["OCCUPANCY_DATE"][:10]
    res = ckan("datastore_search", resource_id=rid, filters=json.dumps({"OCCUPANCY_DATE": day}), limit=5000)
    if len(res["records"]) != res["total"]:
        raise FetchError("Occupancy: did not get every record for the latest date")
    o = pd.DataFrame(res["records"])
    s = lambda c: int(pd.to_numeric(o[c], errors="coerce").sum())
    beds_occ, beds_cap = s("OCCUPIED_BEDS"), s("CAPACITY_ACTUAL_BED")
    rooms_occ, rooms_cap = s("OCCUPIED_ROOMS"), s("CAPACITY_ACTUAL_ROOM")
    occ = {"date": day, "programs": len(o),
           "beds": {"occupied": beds_occ, "actual_capacity": beds_cap, "pct": round(beds_occ / beds_cap * 100, 1)},
           "rooms": {"occupied": rooms_occ, "actual_capacity": rooms_cap, "pct": round(rooms_occ / rooms_cap * 100, 1)},
           "combined_pct": round((beds_occ + rooms_occ) / (beds_cap + rooms_cap) * 100, 1),
           "resource_id": rid, "resource_name": cur[0]["name"]}

    info = {"flow_resource_id": SHELTER_FLOW_RESOURCE, "flow_latest_month": months[-1]["month"],
            "flow_last_refreshed": flow_pkg.get("last_refreshed"),
            "occupancy_resource_id": rid, "occupancy_date": day,
            "occupancy_last_refreshed": pkg.get("last_refreshed")}
    raw = {"toronto_shelter_system_flow_pages.json": json.dumps(pages).encode(),
           "daily_shelter_occupancy_latest_day.json": json.dumps(res).encode(),
           "daily_shelter_occupancy_package.json": json.dumps(pkg).encode()}
    flow_out = {"meta": {k: info[k] for k in ("flow_resource_id", "flow_latest_month", "flow_last_refreshed")},
                "months": months}
    occ_out = {"meta": {k: info[k] for k in ("occupancy_resource_id", "occupancy_date", "occupancy_last_refreshed")},
               **occ}
    return raw, {"shelter_flow.json": flow_out, "shelter_flow.csv": months,
                 "shelter_occupancy.json": occ_out}, info


# ---------------------------------------------------------------- writing

FETCHERS = {"A": fetch_A, "B": fetch_B, "C": fetch_C, "D": fetch_D, "E": fetch_E, "F": fetch_F}


def render(obj):
    """Processed object -> bytes (JSON or CSV by type)."""
    if isinstance(obj, list):
        buf = io.StringIO()
        keys = list(dict.fromkeys(k for r in obj for k in r))
        w = csv.DictWriter(buf, fieldnames=keys, lineterminator="\n")
        w.writeheader()
        w.writerows(obj)
        return buf.getvalue().encode()
    return (json.dumps(obj, indent=1, ensure_ascii=False) + "\n").encode()


def write_atomic(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_bytes(data)
    os.replace(tmp, path)


def run_set(key, manifest):
    raw, processed, info = FETCHERS[key]()
    rendered = {name: render(obj) for name, obj in processed.items()}
    digest = hashlib.sha256(b"".join(rendered[n] for n in sorted(rendered))).hexdigest()
    prev = manifest.get(key, {})
    if prev.get("processed_sha256") == digest and all((DATA / n).exists() for n in rendered):
        print(f"[{key}] unchanged since {prev.get('retrieved')}")
        return
    for name, data in raw.items():
        write_atomic(RAW / name, data)
    for name, data in rendered.items():
        write_atomic(DATA / name, data)
    manifest[key] = {"retrieved": today(), "processed_sha256": digest,
                     "raw_files": sorted(raw), "processed_files": sorted(rendered), "info": info}
    print(f"[{key}] updated: {', '.join(sorted(rendered))}")


def write_sources(m):
    """Regenerate data/SOURCES.md from the manifest so retrieval dates are always the real ones."""
    def g(k, *path, default="(not fetched yet)"):
        v = m.get(k, {})
        for p in ("info",) + path:
            v = v.get(p, {}) if isinstance(v, dict) else {}
        return v if v not in ({}, None, "") else default
    r = lambda k: m.get(k, {}).get("retrieved", "(not fetched yet)")
    f = lambda k: ", ".join(f"`{x}`" for x in m.get(k, {}).get("processed_files", []))
    eds = g("B", "editions_used", default=[])
    ed_lines = "\n".join(f"  - {e['edition']} edition (published {e['published']}): {e['url']} — sheet `{e['sheet']}`, \"{e['table_title']}\""
                         for e in eds) if eds else "  - (not fetched yet)"
    skipped = g("E", "skipped_files", default=[])
    md = f"""# Data sources

Generated by `scripts/fetch_data.py` from `data/manifest.json`. Do not edit by hand.
"Retrieved" is the date the current numbers were first downloaded; a re-fetch that finds the
same numbers leaves the date (and the files) alone. Raw downloads are in `data/raw/`, untouched.

## A. Asking and paid rent — Statistics Canada table 46-10-0092-01

- Table: "{g('A', 'title')}" (product ID 46100092)
- Endpoint: `GET https://www150.statcan.gc.ca/t1/wds/rest/getFullTableDownloadCSV/46100092/en` → `{g('A', 'zip_url')}`; release date from `POST .../getCubeMetadata`
- Latest StatCan release: {g('A', 'release_date')} · Retrieved: {r('A')}
- Filters: GEO = "Toronto, Census metropolitan area (CMA)"; Rental unit type = "Apartment - 1 bedroom" and "Apartment - 2 bedrooms"; Estimates = "Average asking rent" and "Average paid rent"; REF_DATE ≥ 2019-01 (quarterly)
- Paid rent is in this same table (Estimates = "Average paid rent"), so no second table is needed. The Daily of Sept 9, 2026 (https://www150.statcan.gc.ca/n1/daily-quotidien/260909/dq260909c-eng.htm) quotes the same Q2 2026 Toronto figures.
- Raw: `data/raw/46100092-eng.zip`, `data/raw/46100092_cube_metadata.json` · Processed: {f('A')}
- Used by: "But aren't rents falling?" text (peak year, asking minus paid gap) and the asking vs. paid chart (2-bedroom). 1-bedroom is fetched but not charted.

## B. Turnover vs. non-turnover rents — CMHC Rental Market Survey, Table 6.0

- Page: {CMHC_PAGE}
- The page's geography/edition dropdown calls `GET {CMHC_API}GetEditionList` and `GetFileDetails` (cityId = Canada). Table 6.0 is only in the Canada ("all centres") workbook; the Toronto workbook has turnover *rates* but not these rents.
- Workbooks used (every edition that has the table):
{ed_lines}
- Latest edition: {g('B', 'edition')}, released {g('B', 'release_date')}, period {g('B', 'period')} · Retrieved: {r('B')}
- Filters: row "Toronto CMA"; universe: {g('B', 'universe')}. Quality codes kept in the processed file ({g('B', 'quality_codes')}).
- The 2025 edition reports all unit sizes (Oct 2024 and Oct 2025). The 2022–2024 editions report 2-bedroom units only, so the multi-year series (`two_bedroom_series`) is 2-bedroom only, Oct 2021–Oct 2025. Editions without Table 6.0: {', '.join(g('B', 'editions_without_table', default=[])) or 'none'}.
- Raw: `data/raw/rmr-canada-*.xlsx`, `data/raw/cmhc_editions.json` · Processed: {f('B')}
- Used by: hero number, the two stat cards, the gap chart, the calculator and the MPP letter.

## C. Homeownership rate by age, 2011 vs. 2021 — Statistics Canada, The Daily, Sept 21, 2022

- URL: {CENSUS_DAILY} · Retrieved: {r('C')}
- Geography: Canada. Ages 25–29, 30–34, 70–74.
- {g('C', 'note')} The script finds the three sentences in the page and stops if any is missing.
- Raw: `data/raw/dq220921b-eng.htm` · Processed: {f('C')}
- Used by: the homeownership chart.

## D. Fertility — Statistics Canada table 13-10-0418-01

- Table: "{g('D', 'title')}" (product ID 13100418)
- Endpoint: `GET https://www150.statcan.gc.ca/t1/wds/rest/getFullTableDownloadCSV/13100418/en` → `{g('D', 'zip_url')}`
- Latest StatCan release: {g('D', 'release_date')} · Retrieved: {r('D')}
- Filters: GEO = "Ontario, place of residence of mother"; Characteristics = "Total fertility rate per female" and "Age-specific fertility rate, females 20 to 24 years"; years {g('D', 'from_year')}–{g('D', 'to_year')}
- Units: {g('D', 'units')}. The 20–24 change is a change in the *rate*, not in the number of births.
- Raw: `data/raw/13100418-eng.zip`, `data/raw/13100418_cube_metadata.json` · Processed: {f('D')}
- Used by: the fertility callout.

## E. Landlord and Tenant Board applications — Tribunals Ontario open data

- Page: {LTB_OPEN_DATA} (dataset "LTB – Applications Received by e-File"; one .xlsx per quarter, links read from the page)
- Granularity found: {g('E', 'granularity')}; geography: {g('E', 'geography')}
- Measure: {g('E', 'measure')}, rows L1, L2, T2, T6 and (from Jan–Mar 2022) Other
- Coverage: {g('E', 'first_quarter')} to {g('E', 'latest_quarter_end')} (latest file: {g('E', 'latest_fiscal_label')}) · Retrieved: {r('E')}
- Not parsed (kept in raw): {"; ".join(s['file'] + ' — ' + s['reason'] for s in skipped) or 'none'}
- Data quirks: no file for Apr–Jun 2018 or Jul–Sep 2021; the Oct–Dec 2021 file only covers Oct 1–Dec 7, 2021; two older sheets have header dates that contradict their file names (file names used; each case is in the `note` field). {' '.join(g('E', 'notes', default=[]))}
- After Bill 60 (in force {g('E', 'bill60_effective')}): no post-change quarter is published yet. Shown on the site as pending.
- Raw: `data/raw/tribunals_open_data.htm`, `data/raw/ltb_efile/*.xlsx` · Processed: {f('E')}
- Used by: "Evictions for arrears: before and after Bill 60".

## F. Toronto shelters — City of Toronto Open Data (CKAN)

- Base: `{CKAN}`
- Toronto Shelter System Flow: `datastore_search?resource_id={SHELTER_FLOW_RESOURCE}` (all records, limit 500 + offset). Kept population_group "All Population" and "Families". Latest month: {g('F', 'flow_latest_month')}; dataset refreshed {g('F', 'flow_last_refreshed')}.
- Daily Shelter & Overnight Service Occupancy & Capacity: `package_show?id={OCCUPANCY_PACKAGE}` → current-year datastore resource `{g('F', 'occupancy_resource_id')}`; records for the latest OCCUPANCY_DATE ({g('F', 'occupancy_date')}). Occupancy % = (sum OCCUPIED_BEDS + sum OCCUPIED_ROOMS) ÷ (sum CAPACITY_ACTUAL_BED + sum CAPACITY_ACTUAL_ROOM); beds and rooms are also reported separately.
- Retrieved: {r('F')} (re-fetched daily)
- Raw: `data/raw/toronto_shelter_system_flow_pages.json`, `data/raw/daily_shelter_occupancy_latest_day.json`, `data/raw/daily_shelter_occupancy_package.json` · Processed: {f('F')}
- Used by: "Toronto's shelters, updated automatically".

## Figures on the page that are not from these datasets

These are legal facts written into the page text, not data series:
- 2026 rent increase guideline 2.1% and the Nov 15, 2018 exemption: https://www.ontario.ca/page/rent-increase-guideline (checked 2026-11-14; the page also lists 1.9% for 2027).
- Bill 60, N4 payment window 14 → 7 days, in force Sept 21, 2026: https://www.ola.org/en/legislative-business/bills/parliament-44/session-1/bill-60 (ola.org refused automated access; not machine-checked).
- New York HSTPA 2019 (S6458): https://www.nysenate.gov/legislation/bills/2019/S6458 (not machine-checked).
"""
    write_atomic(DATA / "SOURCES.md", md.encode())


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--sets", default="all", help="comma list of A-F, or 'all'")
    args = ap.parse_args()
    keys = list(FETCHERS) if args.sets == "all" else [k.strip().upper() for k in args.sets.split(",")]
    manifest = json.loads(MANIFEST.read_text()) if MANIFEST.exists() else {}
    failed = []
    for k in keys:
        try:
            run_set(k, manifest)
        except Exception as e:  # keep going, keep the old files, fail at the end
            failed.append(k)
            print(f"::error::[{k}] FETCH FAILED, previous files kept: {e}", file=sys.stderr)
            traceback.print_exc()
    write_atomic(MANIFEST, (json.dumps(manifest, indent=1, ensure_ascii=False) + "\n").encode())
    write_sources(manifest)
    if failed:
        print(f"FAILED: {', '.join(failed)}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()

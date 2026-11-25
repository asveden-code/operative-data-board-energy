#!/usr/bin/env python3
"""Rebuild data/series.json from the publishers of the UK energy indicators that have no browser-readable feed.

Deterministic: every series is a function that downloads the current edition of a published file (located through
the gov.uk content API or a stable landing page, because gov.uk asset URLs change with every edition), reads the
cells it needs with openpyxl and returns rows [[period, value], ...]. A series that fails keeps its previous rows
from the existing data/series.json and is marked "stale" with the error, so the page never loses a series silently.

Usage:  python scripts/fetch_series.py [--out data/series.json] [--only key,key]
Period formats: "2025" (annual), "2026 Q2" (quarterly), "2026-08" (monthly), "2026-11-25" (daily / step series).
"""
import argparse, datetime as dt, io, json, re, sys, time, traceback, urllib.request, html
from pathlib import Path
import openpyxl

UA = {"User-Agent": "Mozilla/5.0 (compatible; operative-data-board-energy/1.0)"}
ROOT = Path(__file__).resolve().parent.parent
TWH_PER_MTOE = 11.63
_cache = {}

def get(url, tries=3, timeout=60, binary=False):
    if url in _cache: return _cache[url]
    for i in range(tries):
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=timeout) as r:
                b = r.read()
            _cache[url] = b if binary else b.decode("utf-8", "replace"); return _cache[url]
        except Exception as e:
            if i == tries - 1: raise
            time.sleep(3 * (i + 1))

def govuk_attachment(slug, pattern):
    """URL and page metadata of the attachment of a gov.uk statistics page whose file name or title matches pattern."""
    d = json.loads(get(f"https://www.gov.uk/api/content/government/statistics/{slug}"))
    rx = re.compile(pattern, re.I)
    for a in d["details"].get("attachments", []):
        if rx.search(a.get("url", "")) or rx.search(a.get("title", "")):
            return a["url"], {"title": a.get("title"), "page": f"https://www.gov.uk/government/statistics/{slug}", "published": d.get("public_updated_at", "")[:10]}
    raise RuntimeError(f"no attachment matching {pattern} on {slug}")

def govuk_latest_in_collection(collection_slug):
    d = json.loads(get(f"https://www.gov.uk/api/content/government/collections/{collection_slug}"))
    docs = d["links"].get("documents", [])
    if not docs: raise RuntimeError("empty collection " + collection_slug)
    return docs[0]["base_path"].rsplit("/", 1)[1]

def workbook(url):
    return openpyxl.load_workbook(io.BytesIO(get(url, binary=True)), data_only=True, read_only=True)

def norm(s):
    """Normalise a cell label: collapse whitespace, drop [note n]/[provisional] tags."""
    s = re.sub(r"\[[^\]]*\]", "", str(s if s is not None else ""))
    return re.sub(r"\s+", " ", s).strip()

def sheet_rows(wb, name):
    ws = wb[name]; return [list(r) for r in ws.iter_rows(values_only=True)]

def find_row(rows, label, col=0, start=0):
    for i in range(start, len(rows)):
        if rows[i] and norm(rows[i][col]).lower() == label.lower(): return i
    raise RuntimeError(f"row '{label}' not found")

def find_col(header, label):
    for j, h in enumerate(header):
        if norm(h).lower() == label.lower(): return j
    raise RuntimeError(f"column '{label}' not found in {[norm(h) for h in header][:30]}")

def period_from_header(h):
    """'2026\\n2nd quarter [provisional]' -> '2026 Q2'; '2025 [provisional]' -> '2025'; 'Quarter 2 2026' -> '2026 Q2'; 'August 2026' -> '2026-08'."""
    s = norm(h)
    m = re.match(r"^(\d{4})\s+(\d)(?:st|nd|rd|th) quarter", s)
    if m: return f"{m[1]} Q{m[2]}"
    m = re.match(r"^Quarter (\d) (\d{4})", s)
    if m: return f"{m[2]} Q{m[1]}"
    m = re.match(r"^(\d{4}) Q(\d)", s)
    if m: return f"{m[1]} Q{m[2]}"
    m = re.match(r"^([A-Z][a-z]+) (\d{4})$", s)
    if m:
        mon = ["January","February","March","April","May","June","July","August","September","October","November","December"].index(m[1]) + 1
        return f"{m[2]}-{mon:02d}"
    m = re.match(r"^(\d{4})$", s)
    if m: return m[1]
    return None

def num(v):
    if v is None or (isinstance(v, str) and v.strip() in ("", "[x]", "[c]", "[z]", "..", "-")): return None
    try: return float(v)
    except (TypeError, ValueError): return None

def wide_series(rows, hdr_row, label, scale=1.0, last=None):
    """Periods across columns (ET annual/quarter sheets): returns [[period, value]] for one labelled row."""
    r = rows[find_row(rows, label, start=hdr_row)]
    out = []
    for j, h in enumerate(rows[hdr_row]):
        p = period_from_header(h) if j > 0 else None
        if p and num(r[j]) is not None: out.append([p, round(num(r[j]) * scale, 4)])
    return out[-last:] if last else out

def long_series(rows, hdr_row, col_label, period_col=0, scale=1.0, last=None, col_index=None):
    """Periods down rows (ET 4.2 style): returns [[period, value]] for one labelled column."""
    j = col_index if col_index is not None else find_col(rows[hdr_row], col_label)
    out = []
    for r in rows[hdr_row + 1:]:
        if not r or r[period_col] is None: continue
        p = period_from_header(re.sub(r"\[provisional\]\s*", "", str(r[period_col])))
        if p and num(r[j]) is not None: out.append([p, round(num(r[j]) * scale, 4)])
    return out[-last:] if last else out

# ---------------------------------------------------------------- series definitions -----
REG = {}
def series(key, **meta):
    def deco(f): REG[key] = (f, meta); return f
    return deco

def et(slug, pattern):
    url, meta = govuk_attachment(slug, pattern); return workbook(url), meta, url

# --- III. Oil and gas supply (Energy Trends 3.1 quarterly, 4.1/4.2, 1.3; NSTA) ---
@series("crude", name="Crude oil production", unit="Mt", cad="Q")
def s_crude():
    wb, meta, url = et("oil-and-oil-products-section-3-energy-trends", r"ET_3\.1_[A-Z]{3}_\d{2}\.xlsx$")
    q = sheet_rows(wb, "Quarter"); a = sheet_rows(wb, "Annual")
    return {"rows": wide_series(q, 4, "Crude oil", 1/1000, last=10), "annual": wide_series(a, 4, "Crude oil", 1/1000, last=4), "src": "DESNZ, Energy Trends table 3.1", "edition": meta}

@series("ngl", name="NGLs production", unit="Mt", cad="Q")
def s_ngl():
    wb, meta, url = et("oil-and-oil-products-section-3-energy-trends", r"ET_3\.1_[A-Z]{3}_\d{2}\.xlsx$")
    q = sheet_rows(wb, "Quarter"); a = sheet_rows(wb, "Annual")
    return {"rows": wide_series(q, 4, "NGLs", 1/1000, last=10), "annual": wide_series(a, 4, "NGLs", 1/1000, last=4), "src": "DESNZ, Energy Trends table 3.1", "edition": meta}

@series("gasprod", name="Natural gas production", unit="TWh", cad="Q")
def s_gasprod():
    wb, meta, url = et("gas-section-4-energy-trends", r"ET_4\.1_[A-Z]{3}_\d{2}\.xlsx$")
    q = sheet_rows(wb, "Quarter (GWh)"); a = sheet_rows(wb, "Annual (GWh)")
    return {"rows": wide_series(q, 2, "Indigenous production", 1/1000, last=10), "annual": wide_series(a, 2, "Indigenous production", 1/1000, last=4), "src": "DESNZ, Energy Trends table 4.1", "edition": meta}

def _et13():
    wb, meta, url = et("total-energy-section-1-energy-trends", r"ET_1\.3_[A-Z]{3}_\d{2}\.xlsx$")
    rows = sheet_rows(wb, "Quarter supply"); hdr = 4
    # column groups: Total imports (7) with Primary oil 10, Petroleum 11, Natural gas 12; Total exports (15) with 18, 19, 20
    H = [norm(h).lower() for h in rows[hdr]]
    ti = H.index("total imports"); te = H.index("total exports")
    assert H[ti+3].startswith("primary oil") and H[ti+4].startswith("petroleum") and H[ti+5].startswith("natural gas"), H[ti:ti+6]
    assert H[te+3].startswith("primary oil") and H[te+4].startswith("petroleum") and H[te+5].startswith("natural gas"), H[te:te+6]
    out = {"oil": [], "gas": []}
    for r in rows[hdr+1:]:
        p = period_from_header(r[0]) if r and r[0] else None
        if not p: continue
        vals = [num(r[j]) for j in (ti+3, ti+4, te+3, te+4, ti+5, te+5)]
        if None in vals: continue
        out["oil"].append([p, round((vals[0]+vals[1]+vals[2]+vals[3]) / 1000, 3)]); out["gas"].append([p, round((vals[4]+vals[5]) / 1000, 3)])
    return out, meta

@series("oilimp", name="Net imports of oil", unit="mtoe", cad="Q")
def s_oilimp():
    d, meta = _et13(); return {"rows": d["oil"][-10:], "src": "DESNZ, Energy Trends table 1.3 (primary oil and petroleum products, imports less exports)", "edition": meta}

@series("gasimp", name="Net imports of natural gas", unit="mtoe", cad="Q")
def s_gasimp():
    d, meta = _et13(); return {"rows": d["gas"][-10:], "src": "DESNZ, Energy Trends table 1.3 (imports less exports)", "edition": meta}

@series("opex", name="Average unit operating cost, UKCS", unit="£/boe", cad="A")
def s_opex():
    page = "https://www.nstauthority.co.uk/data-and-insights/insights-and-analysis/production-and-expenditure-projections/"
    h = get(page)
    m = re.search(r"href=['\"]([^'\"]*projections-of-uk-oil-and-gas-production-and-expenditure\.xlsx)['\"]", h, re.I)
    if not m: raise RuntimeError("NSTA projections workbook link not found")
    url = m[1] if m[1].startswith("http") else "https://www.nstauthority.co.uk" + m[1]
    wb = workbook(url); rows = sheet_rows(wb, "Public Table")
    yr = None
    for r in rows:
        if r and norm(r[0]).lower() == "production": yr = r; break
    ri = next(i for i, r in enumerate(rows) if r and norm(r[0]) == "Average Unit" and norm(r[1]) == "(£/boe)")
    hist, proj = [], []
    for j in range(3, len(yr)):
        y = norm(yr[j]); v = num(rows[ri][j])
        if not re.match(r"^\d{4}$", y) or v is None: continue
        (proj if int(y) >= dt.date.today().year else hist).append([y, round(v, 2)])
    ed = re.search(r"nsta-([a-z]+-\d{4})-projections", url)
    return {"rows": hist, "projection": proj, "src": "North Sea Transition Authority, production and expenditure projections", "edition": {"title": ed[1].replace("-", " ").title() if ed else url, "page": page}}

# --- IV. Gas demand (Energy Trends 4.2 monthly/quarterly/annual) ---
def _et42(col, last=10):
    wb, meta, url = et("gas-section-4-energy-trends", r"ET_4\.2_[A-Z]{3}_\d{2}\.xlsx$")
    q = sheet_rows(wb, "Quarter (GWh)"); a = sheet_rows(wb, "Annual (GWh)"); m = sheet_rows(wb, "Month (GWh)")
    return {"rows": long_series(q, 7, col, scale=1/1000, last=last), "annual": long_series(a, 7, col, scale=1/1000, last=4), "monthly": long_series(m, 7, col, scale=1/1000, last=24), "src": "DESNZ, Energy Trends table 4.2", "edition": meta}

@series("gasdem", name="Natural gas demand", unit="TWh", cad="Q")
def s_gasdem(): return _et42("Demand")
@series("gasdom", name="Domestic gas demand", unit="TWh", cad="Q")
def s_gasdom(): return _et42("Domestic")
@series("gasgen", name="Gas for electricity generation", unit="TWh", cad="Q")
def s_gasgen(): return _et42("Electricity generation")
@series("gascom", name="Commercial and services gas demand", unit="TWh", cad="Q")
def s_gascom(): return _et42("Services")
@series("gasind", name="Industrial gas demand", unit="TWh", cad="Q")
def s_gasind(): return _et42("Industrial")

@series("gasnet", name="Net imports of natural gas (ET 4.2)", unit="TWh", cad="Q")
def s_gasnet(): return _et42("Net imports")

@series("hp", name="Heat pump installations", unit="thousand", cad="Q")
def s_hp():
    slug = govuk_latest_in_collection("heat-pump-deployment-statistics")
    url, meta = govuk_attachment(slug, r"heat-pump-deployment-tables-q\d-\d{4}\.xlsx$")
    rows = sheet_rows(workbook(url), "Table 1.1")
    hdr = next(i for i, r in enumerate(rows) if r and norm(r[0]).lower().startswith("installation quarter"))
    j = next(k for k, h in enumerate(rows[hdr]) if norm(h).lower().startswith("total heat"))
    q, byyear = [], {}
    for r in rows[hdr+1:]:
        if not r or not r[0]: continue
        m = re.match(r"^(\d{4}) Q(\d)", norm(r[0]))
        if not m or num(r[j]) is None: continue
        q.append([f"{m[1]} Q{m[2]}", round(num(r[j]) / 1000, 3)]); byyear.setdefault(m[1], []).append(num(r[j]))
    annual = [[y, round(sum(v) / 1000, 3)] for y, v in byyear.items() if len(v) == 4]
    return {"rows": q[-10:], "annual": annual[-4:], "src": "DESNZ, heat pump deployment statistics (MCS installation database, retrofit, up to 45 kW)", "edition": meta}

# ---------------------------------------------------------------- main -----
def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--out", default="data/series.json"); ap.add_argument("--only", default="")
    a = ap.parse_args(); out = ROOT / a.out
    prev = json.loads(out.read_text())["series"] if out.exists() else {}
    keys = [k for k in a.only.split(",") if k] or list(REG)
    result = {k: v for k, v in prev.items()}
    now = dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%MZ"); ok = 0
    for k in keys:
        f, meta = REG[k]
        try:
            d = f(); d.update(meta); d["fetched"] = now; d["stale"] = False; d.pop("error", None); result[k] = d; ok += 1
            print(f"ok   {k:9s} {d['rows'][-1] if d.get('rows') else '-'}  {d.get('edition',{}).get('title','')[:60]}")
        except Exception as e:
            old = result.get(k, {}); old.update(meta); old["stale"] = True; old["error"] = f"{type(e).__name__}: {e}"[:200]; result[k] = old
            print(f"FAIL {k:9s} {old['error']}", file=sys.stderr)
    out.parent.mkdir(parents=True, exist_ok=True)
    json.dump({"generated_at": now, "series": result}, open(out, "w"), ensure_ascii=False, indent=1)
    print(f"wrote {a.out}: {ok}/{len(keys)} series refreshed, {sum(1 for v in result.values() if v.get('stale'))} stale")
    return 0 if ok == len(keys) else 1

if __name__ == "__main__":
    sys.exit(main())

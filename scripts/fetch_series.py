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
    return {"rows": wide_series(q, 4, "Crude oil", 1/1000, last=10), "annual": wide_series(a, 4, "Crude oil", 1/1000, last=4), "src": "DESNZ, Energy Trends table 3.1 (quarterly)", "edition": meta}

@series("ngl", name="NGLs production", unit="Mt", cad="Q")
def s_ngl():
    wb, meta, url = et("oil-and-oil-products-section-3-energy-trends", r"ET_3\.1_[A-Z]{3}_\d{2}\.xlsx$")
    q = sheet_rows(wb, "Quarter"); a = sheet_rows(wb, "Annual")
    return {"rows": wide_series(q, 4, "NGLs", 1/1000, last=10), "annual": wide_series(a, 4, "NGLs", 1/1000, last=4), "src": "DESNZ, Energy Trends table 3.1 (quarterly)", "edition": meta}

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


# --- V–VII. Ofgem everviz charts, gov.uk content API, IMF ---
import csv as _csv
def everviz(cid):
    """Header, rows and freshness stamp of the CSV behind an Ofgem chart hosted on everviz."""
    js = get(f"https://app.everviz.com/inject/{cid}")
    m = re.search(r'"data":\{"csv":"((?:[^"\\]|\\.)*)"', js, re.S)
    if not m: raise RuntimeError("no csv in everviz " + cid)
    text = json.loads('"' + m[1] + '"'); first = text.split("\n", 1)[0]
    rows = list(_csv.reader(io.StringIO(text), delimiter=";" if first.count(";") > first.count(",") else ","))
    asof = re.search(r"correct as of:?\s*([A-Za-z]+ \d{4})", js)
    return [h.strip() for h in rows[0]], [r for r in rows[1:] if r and r[0].strip()], {"title": "Ofgem chart " + cid, "page": None, "asof": asof[1] if asof else None}

MON3 = {m: i for i, m in enumerate(["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"], 1)}
def p_mdy(s):  # "09/01/2026" -> "2026-09"
    mm, dd, yy = s.strip().strip('"').split("/"); return f"{yy}-{mm}"
def p_monyy(s):  # "Jun-26" -> "2026 Q2"
    mm, yy = s.strip().split("-"); return f"20{yy} Q{(MON3[mm]+2)//3}"
def p_qlabel(s):  # "Q2 2026" or "2026 Q2" -> "2026 Q2"
    m = re.match(r"^Q(\d) (\d{4})$", s.strip()) or re.match(r"^(\d{4}) Q(\d)$", s.strip())
    if not m: return None
    return f"{m[2]} Q{m[1]}" if s.strip().startswith("Q") else f"{m[1]} Q{m[2]}"
def p_caplabel(s):  # "Oct - Dec 2026" / "Apr - Sep 2017" -> "2026 Q4" (start month decides the quarter)
    m = re.match(r"^([A-Z][a-z]{2})[a-z]*\s*-\s*[A-Z][a-z]+\s+(\d{2}|\d{4})$", s.strip())
    if not m: return None
    y = m[2] if len(m[2]) == 4 else "20" + m[2]
    return f"{y} Q{(MON3[m[1]]+2)//3}"

OFGEM_WMI = "https://www.ofgem.gov.uk/news-and-insight/data/data-portal/wholesale-market-indicators"
OFGEM_RMI = "https://www.ofgem.gov.uk/news-and-insight/data/data-portal/retail-market-indicators"
OFGEM_DEBT = "https://www.ofgem.gov.uk/data/debt-and-arrears-indicators"

@series("power", name="Wholesale electricity price (GB)", unit="£/MWh", cad="M")
def s_power():
    h, rows, meta = everviz("n9WSDfgxc"); meta["page"] = OFGEM_WMI
    return {"rows": [[p_mdy(r[0]), round(float(r[1]), 2)] for r in rows if num(r[1]) is not None][-24:], "src": "Ofgem wholesale market indicators: day-ahead baseload contracts, monthly average", "edition": meta}

@series("nbp", name="NBP gas price (GB)", unit="p/therm", cad="M")
def s_nbp():
    h, rows, meta = everviz("YFsS4g98P"); meta["page"] = OFGEM_WMI
    return {"rows": [[p_mdy(r[0]), round(float(r[1]), 2)] for r in rows if num(r[1]) is not None][-24:], "src": "Ofgem wholesale market indicators: day-ahead contracts, monthly average", "edition": meta}

@series("cap", name="Ofgem price cap, typical dual-fuel direct debit", unit="£/year", cad="Q")
def s_cap():
    h, rows, meta = everviz("lVeuoASKI"); meta["page"] = OFGEM_RMI
    H = [x.lower() for x in h]
    jt = H.index("total"); jw = next(i for i, x in enumerate(H) if x.startswith("wholesale")); jd = next(i for i, x in enumerate(H) if x.startswith("debt related")); je = next(i for i, x in enumerate(H) if x.startswith("earnings before"))
    out = []
    for r in rows:
        p = p_caplabel(r[0])
        if p and num(r[jt]) is not None: out.append([p, round(num(r[jt])), round(num(r[jw]) or 0), round(num(r[jd]) or 0), round(num(r[je]) or 0)])
    return {"rows": out[-10:], "cols": ["Total", "Wholesale costs", "Debt-related costs", "EBIT allowance"], "src": "Ofgem retail market indicators: costs included in the price cap level (direct debit, typical consumption, GB average, incl. VAT)", "edition": meta}

@series("capparams", name="Price cap EBIT allowance parameters", unit="", cad="A")
def s_capparams():
    idx = get("https://www.ofgem.gov.uk/energy-regulation/domestic-and-non-domestic/energy-pricing-rules/energy-price-cap/energy-price-cap-default-tariff-levels")
    m = re.search(r'href="([^"]*[Dd]efault-tariff-cap-level[^"]*\.xlsx)"', idx)
    if not m: raise RuntimeError("Default tariff cap level xlsx link not found")
    url = m[1] if m[1].startswith("http") else "https://www.ofgem.gov.uk" + m[1]
    rows = sheet_rows(workbook(url), "3k EBIT")
    hdr = next(i for i, r in enumerate(rows) if r and any(isinstance(c, str) and re.search(r"Oct(ober)? 20\d\d", c) for c in r if c))
    def rowv(label):
        r = rows[next(i for i, r in enumerate(rows) if r and any(isinstance(c, str) and label.lower() in c.lower() for c in r[:3]))]
        return [[norm(rows[hdr][j]), round(num(r[j]), 4)] for j in range(len(r)) if num(r[j]) is not None and rows[hdr][j]]
    fx, var = rowv("EBIT Fixed"), rowv("EBIT Variable")
    return {"rows": [[fx[-1][0], fx[-1][1]]], "fixed": fx[-4:], "variable": var[-4:], "src": "Ofgem default tariff cap model, sheet 3k EBIT", "edition": {"title": url.rsplit("/", 1)[1], "page": "https://www.ofgem.gov.uk/energy-regulation/domestic-and-non-domestic/energy-pricing-rules/energy-price-cap/energy-price-cap-default-tariff-levels"}}

@series("debt", name="Domestic customer debt and arrears", unit="£bn", cad="Q")
def s_debt():
    h, rows, meta = everviz("0IwJatYrR"); meta["page"] = OFGEM_DEBT
    out = [[p_qlabel(r[0]), round(float(r[3]), 2), round(float(r[1]), 2), round(float(r[2]), 2)] for r in rows if p_qlabel(r[0]) and num(r[3]) is not None]
    return {"rows": out[-10:], "cols": ["Total", "Debt (repayment arrangement)", "Arrears (no arrangement)"], "src": "Ofgem debt and arrears indicators: domestic debt and arrears older than 91 days, gas and electricity", "edition": meta}

@series("suppliers", name="Active domestic suppliers (GB)", unit="suppliers", cad="Q")
def s_suppliers():
    h, rows, meta = everviz("-7wWeiqwF"); meta["page"] = OFGEM_RMI
    jt = [x.lower() for x in h].index("total")
    return {"rows": [[p_monyy(r[0]), int(float(r[jt]))] for r in rows if num(r[jt]) is not None][-10:], "src": "Ofgem retail market indicators: number of active domestic suppliers by fuel type", "edition": meta}

def _shares(cid):
    h, rows, meta = everviz(cid); meta["page"] = OFGEM_RMI
    out = []
    for r in rows:
        p = p_qlabel(r[0])
        if not p: continue
        d = {h[j]: round(float(r[j]), 2) for j in range(1, len(h)) if j < len(r) and num(r[j]) not in (None, 0.0)}
        out.append([p, d])
    return out[-9:], meta

@series("shares_e", name="Domestic electricity supply market shares", unit="%", cad="Q", kind="shares")
def s_shares_e():
    rows, meta = _shares("_o5LqZ5Dk"); return {"rows": rows, "src": "Ofgem retail market indicators: electricity supply market shares by company, domestic (GB)", "edition": meta}

@series("shares_g", name="Domestic gas supply market shares", unit="%", cad="Q", kind="shares")
def s_shares_g():
    rows, meta = _shares("4GHD8gVq8"); return {"rows": rows, "src": "Ofgem retail market indicators: gas supply market shares by company, domestic (GB)", "edition": meta}

@series("switch", name="Domestic customers switching supplier", unit="thousand", cad="M")
def s_switch():
    h, rows, meta = everviz("eGQWqcSLx"); meta["page"] = OFGEM_RMI
    out = [[p_mdy(r[0]), round(float(r[1]) / 1000, 1), round(float(r[2]) / 1000, 1)] for r in rows if num(r[1]) is not None]
    return {"rows": out[-24:], "cols": ["Electricity", "Gas"], "src": "Ofgem retail market indicators: domestic customers switching supplier by fuel type (Electralink and Xoserve data)", "edition": meta}

@series("switchrate", name="Switching rate, domestic electricity", unit="% per month", cad="M")
def s_switchrate():
    d = json.loads(get("https://www.gov.uk/api/content/government/statistical-data-sets/quarterly-domestic-energy-switching-statistics"))
    att = next(a for a in d["details"]["attachments"] if a.get("url", "").endswith(".xlsx"))
    rows = sheet_rows(workbook(att["url"]), "2.7.1 (Monthly)")
    hdr = next(i for i, r in enumerate(rows) if r and norm(r[0]).lower() == "year")
    H = [norm(x).lower() for x in rows[hdr]]
    je, jg = H.index("electricity transfers"), H.index("gas transfers"); jce, jcg = H.index("total electricity customers"), H.index("total gas customers")
    out = []
    for r in rows[hdr+1:]:
        if not r or num(r[0]) is None or num(r[je]) is None or num(r[jce]) is None: continue
        y = int(float(r[0])); mo = r[1]
        m = mo.month if hasattr(mo, "month") else MON3.get(str(mo).strip()[:3].title())
        if not m: continue
        out.append([f"{y}-{m:02d}", round(100 * num(r[je]) / num(r[jce]), 2), round(100 * num(r[jg]) / num(r[jcg]), 2) if num(r[jcg]) else None, round(num(r[jce]) / 1e6, 2)])
    return {"rows": out[-24:], "cols": ["Electricity", "Gas", "Electricity customers (m)"], "src": "DESNZ quarterly domestic energy switching statistics, table 2.7.1 (transfers as a share of customer accounts)", "edition": {"title": att.get("title"), "page": "https://www.gov.uk/government/statistical-data-sets/quarterly-domestic-energy-switching-statistics", "published": d.get("public_updated_at", "")[:10]}}

def govuk_tables(path):
    d = json.loads(get("https://www.gov.uk/api/content" + path)); body = d["details"]["body"]
    tabs = []
    for t in re.findall(r"<table.*?</table>", body, re.S):
        tabs.append([[html.unescape(re.sub(r"<[^>]+>", "", c)).strip() for c in re.findall(r"<t[hd][^>]*>(.*?)</t[hd]>", r, re.S)] for r in re.findall(r"<tr.*?</tr>", t, re.S)])
    return tabs, d

@series("uka", name="UK ETS allowance price (UKA)", unit="£/tCO2", cad="M")
def s_uka():
    path = "/government/publications/taking-part-in-the-uk-emissions-trading-scheme-markets/cost-containment-mechanism-ccm-trigger-prices-and-average-monthly-prices-full-table"
    tabs, d = govuk_tables(path); ser = {}; trig = {}
    for t in tabs:
        for cells in t:
            if not cells or not re.match(r"^[A-Z][a-z]{2}-\d\d$", cells[0]): continue
            mm, yy = cells[0].split("-"); y = 2000 + int(yy); m = MON3[mm]
            n = len(cells) - 5   # columns: Month, reference period, trigger, monitoring period, n monthly averages, triggered?
            if cells[2].startswith("£"): trig[f"{y}-{m:02d}"] = float(cells[2][1:].replace(",", ""))
            for k, p in enumerate(cells[4:4+n]):
                ym = y * 12 + m - 1 - (n - k); yy2, mm2 = divmod(ym, 12)
                if p.startswith("£"): ser[f"{yy2}-{mm2+1:02d}"] = float(p[1:].replace(",", ""))
    rows = sorted(ser.items()); tr = sorted(trig.items())
    return {"rows": [list(r) for r in rows[-24:]], "trigger": [list(r) for r in tr[-3:]], "src": "UK ETS Authority: cost containment mechanism tables, average monthly UKA price", "edition": {"title": d.get("title"), "page": "https://www.gov.uk" + path, "published": d.get("public_updated_at", "")[:10]}}

def page_text(path):
    d = json.loads(get("https://www.gov.uk/api/content" + path)); body = d["details"].get("body") or json.dumps(d["details"])
    return re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", body))), d

@series("cps", name="Carbon Price Support rate", unit="£/tCO2", cad="S")
def s_cps():
    txt, d = page_text("/government/statistics/environmental-taxes-bulletin/environmental-taxes-bulletin-historical-rates")
    m = re.search(r"1 April 2016\s+£(\d+\.\d\d)\s+£0\.00331", txt)
    if not m: raise RuntimeError("CPS £/tCO2 row not found in the environmental taxes bulletin")
    ccl, d2 = page_text("/guidance/climate-change-levy-rates")
    end = re.search(r"1 April 2016 to (\d{1,2} [A-Z][a-z]+ \d{4})\s+0\.00331", ccl)
    return {"rows": [["2016-04-01", float(m[1])]], "ends": end[1] if end else None, "statement": f"Carbon Price Support £{m[1]}/tCO2 (gas 0.00331 £/kWh) from 1 April 2016" + (f" to {end[1]}" if end else ""), "src": "HMRC climate change levy rates; environmental taxes bulletin", "edition": {"title": d2.get("title"), "page": "https://www.gov.uk/guidance/climate-change-levy-rates", "published": d2.get("public_updated_at", "")[:10]}}

@series("epl", name="Energy Profits Levy rate", unit="%", cad="S")
def s_epl():
    txt, d = page_text("/government/publications/energy-profits-levy-reforms-2024/energy-profits-levy-reforms-2024")
    m = re.search(r"increases the rate of Energy Profits Levy by \d+ percentage points to (\d+)%, extends the end date of the Energy Profits Levy to (\d{1,2} [A-Z][a-z]+ \d{4})", txt)
    if not m: raise RuntimeError("EPL rate sentence not found")
    return {"rows": [["2022-05-26", 25], ["2023-01-01", 35], ["2024-11-01", int(m[1])]], "ends": m[2], "statement": f"Energy Profits Levy {m[1]}% from 1 November 2024 to {m[2]} (or earlier if the Energy Security Investment Mechanism is triggered); 25% from 26 May 2022, 35% from 1 January 2023", "src": "HM Treasury / HMRC policy paper, Energy Profits Levy reforms 2024", "edition": {"title": d.get("title"), "page": "https://www.gov.uk/government/publications/energy-profits-levy-reforms-2024/energy-profits-levy-reforms-2024", "published": d.get("public_updated_at", "")[:10]}}

@series("egl", name="Electricity Generator Levy rate", unit="%", cad="S")
def s_egl():
    txt, d = page_text("/government/publications/increase-in-the-rate-of-the-electricity-generator-levy/amended-rate-of-the-electricity-generator-levy-2026-to-2027")
    m = re.search(r"increases the rate at which the Electricity Generator Levy \(\s*EGL\s*\) is charged from (\d+)% to (\d+)% from (\d{1,2} [A-Z][a-z]+ \d{4})", txt)
    if not m: raise RuntimeError("EGL rate sentence not found")
    bm = re.search(r"benchmark price[^.]*?set at £(\d+(?:\.\d\d)?)", txt)
    return {"rows": [["2023-01-01", int(m[1])], ["2026-07-01", int(m[2])]], "benchmark": float(bm[1]) if bm else None, "statement": f"Electricity Generator Levy {m[2]}% from {m[3]} ({m[1]}% from 1 January 2023); charged on receipts above a benchmark price" + (f" of £{bm[1]}/MWh" if bm else ""), "src": "HMRC policy paper, amended rate of the Electricity Generator Levy", "edition": {"title": d.get("title"), "page": "https://www.gov.uk/government/publications/increase-in-the-rate-of-the-electricity-generator-levy", "published": d.get("public_updated_at", "")[:10]}}

@series("upstream", name="Headline tax rate on upstream oil and gas profits", unit="%", cad="S")
def s_upstream():
    txt, d = page_text("/government/publications/oil-and-gas-taxation-oil-and-gas-revenue-levy/oil-and-gas-revenue-levy-ogrl")
    m = re.search(r"Ring Fence Corporation Tax \(RFCT\) which is charged at (\d+)% and the Supplementary Charge \(SC\) which is charged at (\d+)%", txt)
    if not m: raise RuntimeError("RFCT/SC sentence not found")
    epl = s_epl(); rate = int(m[1]) + int(m[2]) + epl["rows"][-1][1]
    return {"rows": [["2022-05-26", int(m[1]) + int(m[2]) + 25], ["2023-01-01", int(m[1]) + int(m[2]) + 35], ["2024-11-01", rate]], "statement": f"Ring Fence Corporation Tax {m[1]}% + Supplementary Charge {m[2]}% + Energy Profits Levy {epl['rows'][-1][1]}% = {rate}%", "src": "HM Treasury policy papers (Oil and Gas Revenue Levy; Energy Profits Levy reforms 2024)", "edition": {"title": d.get("title"), "page": "https://www.gov.uk/government/publications/oil-and-gas-taxation-oil-and-gas-revenue-levy/oil-and-gas-revenue-levy-ogrl", "published": d.get("public_updated_at", "")[:10]}}

def imf(indicator, transform):
    start = (dt.date.today() - dt.timedelta(days=760)).strftime("%Y-%m")
    req = urllib.request.Request(f"https://api.imf.org/external/sdmx/2.1/data/IMF.RES,PCPS/G001.{indicator}.{transform}.M?startPeriod={start}", headers={**UA, "Accept": "application/vnd.sdmx.data+csv"})
    with urllib.request.urlopen(req, timeout=60) as r: text = r.read().decode()
    rows = list(_csv.DictReader(io.StringIO(text)))
    return sorted([[r["TIME_PERIOD"].replace("-M", "-"), round(float(r["OBS_VALUE"]), 2)] for r in rows if r.get("OBS_VALUE")])

@series("brent", name="Brent crude", unit="$/bbl", cad="M")
def s_brent(): return {"rows": imf("POILBRE", "USD")[-24:], "src": "IMF Primary Commodity Price System, POILBRE (monthly average)", "edition": {"title": "IMF PCPS", "page": "https://data.imf.org/en/datasets/IMF.RES:PCPS"}}
@series("ttf", name="TTF gas (EU)", unit="$/MMBtu", cad="M")
def s_ttf(): return {"rows": imf("PNGASEU", "USD")[-24:], "src": "IMF Primary Commodity Price System, PNGASEU (monthly average)", "edition": {"title": "IMF PCPS", "page": "https://data.imf.org/en/datasets/IMF.RES:PCPS"}}


# --- I–II. Electricity (DUKES chapter 5/6 annual; Energy Trends 5.1, 5.2, 5.6, 6.1 quarterly; NESO) ---
DUKES5 = "electricity-chapter-5-digest-of-united-kingdom-energy-statistics-dukes"
DUKES6 = "renewable-sources-of-energy-chapter-6-digest-of-united-kingdom-energy-statistics-dukes"
ET5 = "electricity-section-5-energy-trends"; ET6 = "energy-trends-section-6-renewables"

def year_cols(hdr):
    out = {}
    for i, h in enumerate(hdr):
        if isinstance(h, (int, float)) and 1900 < h < 2100: out[str(int(h))] = i
        elif isinstance(h, str) and re.fullmatch(r"\s*(\d{4})\s*(\[.*\])?\s*", h): out[re.match(r"\s*(\d{4})", h)[1]] = i
    return out

def block(rows, title_frag):
    """index of the row whose first cell contains title_frag (a DUKES sub-table title)"""
    return next(i for i, r in enumerate(rows) if r and r[0] and title_frag in str(r[0]))

def _d512():
    url, meta = govuk_attachment(DUKES5, r"DUKES_5\.12\.xlsx$"); rows = sheet_rows(workbook(url), "5.12")
    ci = block(rows, "5.12.C"); yc = year_cols(rows[ci + 1]); years = sorted(yc)[-4:]
    def row(label):
        r = next(r for r in rows[ci+2:] if r and norm(r[0]) == "All" and norm(r[1]).lower() == label.lower())
        return {y: num(r[yc[y]]) or 0 for y in years}
    labels = {"total": "Total", "ccgt": "CCGT gas", "ocgt": "OCGT and conventional thermal gas", "oil": "Oil", "coal": "Coal", "pwr": "Nuclear - PWR", "agr": "Nuclear - AGR", "magnox": "Nuclear - Magnox", "hydro": "Hydro (natural flow)", "bio": "Bioenergy", "other": "Other fuels", "pumped": "Pumped storage"}
    d = {k: row(v) for k, v in labels.items()}
    return d, years, meta

def _d516():
    url, meta = govuk_attachment(DUKES5, r"DUKES_5\.16\.xlsx$"); rows = sheet_rows(workbook(url), "5.16")
    out = {}
    for frag in ("5.16.A", "5.16.C"):
        i = block(rows, frag); yc = year_cols(rows[i + 1]); t = {}
        for r in rows[i+2:]:
            if not r or r[0] is None: break
            t[norm(r[0])] = {y: num(r[yc[y]]) for y in yc}
        out[frag] = t
    return out, meta

def annual_rows(d, years, keys, scale=1/1000, dec=2):
    return [[y, round(sum(d[k][y] or 0 for k in keys) * scale, dec)] for y in years]

def _et61():
    url, meta = govuk_attachment(ET6, r"ET_6\.1_[A-Z]{3}_\d{2}\.xlsx$"); rows = sheet_rows(workbook(url), "Quarter")
    out = {}
    for frag, key in (("CUMULATIVE INSTALLED CAPACITY", "cap"), ("ELECTRICITY GENERATED", "gen")):
        i = block(rows, frag); hdr = rows[i]
        r = next(r for r in rows[i+1:] if r and norm(r[0]) == "TOTAL")
        out[key] = [[period_from_header(h), num(r[j])] for j, h in enumerate(hdr) if j and h and period_from_header(h) and num(r[j]) is not None]
    return out, meta

@series("capren", name="Renewable capacity", unit="GW", cad="Q")
def s_capren():
    d, meta = _et61(); url, m2 = govuk_attachment(DUKES6, r"DUKES_6\.2\.xlsx$"); rows = sheet_rows(workbook(url), "6.2")
    i = block(rows, "Installed Capacity (MW)"); yc = year_cols(rows[i]); r = next(r for r in rows[i+1:] if r and norm(r[0]) == "Total Installed Capacity")
    annual = [[y, round(num(r[yc[y]]) / 1000, 2)] for y in sorted(yc)[-4:] if num(r[yc[y]]) is not None]
    return {"rows": [[p, round(v / 1000, 2)] for p, v in d["cap"]][-10:], "annual": annual, "src": "DESNZ, Energy Trends table 6.1 (quarterly, cumulative installed capacity)", "edition": meta}

def _d56():
    url, meta = govuk_attachment(DUKES5, r"DUKES_5\.6\.xlsx$"); rows = sheet_rows(workbook(url), "5.6")
    bi = block(rows, "5.6.B"); ei = block(rows, "5.6.C"); yc = year_cols(rows[bi + 1]); years = sorted(yc)[-4:]
    g = {}
    for r in rows[bi+2:ei]:
        if r and r[0] and str(r[0]).startswith("All generating companies"): g[norm(r[1])] = {y: num(r[yc[y]]) for y in years}
    return g, years, meta

def _et51():
    url, meta = govuk_attachment(ET5, r"ET_5\.1_[A-Z]{3}_\d{2}\.xlsx$"); rows = sheet_rows(workbook(url), "Quarter")
    bi = block(rows, "Table 5.1b"); ci = block(rows, "Table 5.1c"); hdr = rows[bi + 1]
    qc = [(j, period_from_header(h)) for j, h in enumerate(hdr) if j > 1 and h and period_from_header(h)]
    g = {}
    for r in rows[bi+2:ci]:
        if r and r[0] and str(r[0]).startswith("All generating companies"): g[norm(r[1])] = [[q, num(r[j])] for j, q in qc if num(r[j]) is not None]
    return g, meta

def sum_by_period(series_list, scale=1.0, dec=2, require_all=False):
    """Sum several [[period, value]] lists period by period (a series that starts later counts as 0 before it starts)."""
    tot = {}; cnt = {}
    for ser in series_list:
        for p, v in ser: tot[p] = tot.get(p, 0) + (v or 0); cnt[p] = cnt.get(p, 0) + 1
    ps = sorted(tot); first = max(min(p for p, _ in ser) for ser in series_list if ser) if require_all else ps[0]
    return [[p, round(tot[p] * scale, dec)] for p in ps if p >= first]

def gen_series(qkeys, akeys, dec=2):
    q, meta = _et51(); a, years, m2 = _d56()
    rows = sum_by_period([q[k] for k in qkeys], dec=dec)[-10:]
    annual = [[y, round(sum((a[k][y] or 0) for k in akeys) / 1000, 1)] for y in years]
    return rows, annual, meta

@series("gentotal", name="Total electricity generation", unit="TWh", cad="Q")
def s_gentotal():
    r, a, m = gen_series(["Total all generating companies"], ["Total all generating companies"]); return {"rows": r, "annual": a, "src": "DESNZ, Energy Trends table 5.1 (quarterly), all generating companies", "edition": m}
@series("genbase", name="Baseload generation (nuclear and bioenergy)", unit="TWh", cad="Q")
def s_genbase():
    r, a, m = gen_series(["Nuclear", "Bioenergy"], ["Nuclear", "Thermal renewables"]); return {"rows": r, "annual": a, "src": "DESNZ, Energy Trends table 5.1: nuclear plus bioenergy (thermal renewables)", "edition": m}
@series("gennuc", name="Nuclear generation", unit="TWh", cad="Q")
def s_gennuc():
    r, a, m = gen_series(["Nuclear"], ["Nuclear"]); return {"rows": r, "annual": a, "src": "DESNZ, Energy Trends table 5.1", "edition": m}
@series("gendisp", name="Dispatchable generation", unit="TWh", cad="Q")
def s_gendisp():
    r, a, m = gen_series(["Gas", "Oil", "Coal", "Nuclear", "Hydro (natural flow)", "Bioenergy", "Other fuels", "Pumped storage", "Battery storage"], ["Gas", "Oil", "Coal", "Nuclear", "Hydro (natural flow)", "Thermal renewables", "Other fuels", "Energy storage"])
    return {"rows": r, "annual": a, "src": "DESNZ, Energy Trends table 5.1: gas, oil, coal, nuclear, hydro, bioenergy, other fuels and storage", "edition": m}
@series("gengas", name="Gas-fired generation", unit="TWh", cad="Q")
def s_gengas():
    r, a, m = gen_series(["Gas"], ["Gas"]); return {"rows": r, "annual": a, "src": "DESNZ, Energy Trends table 5.1", "edition": m}
@series("genren", name="Renewable generation", unit="TWh", cad="Q")
def s_genren():
    d, meta = _et61(); a, years, m2 = _d56()
    return {"rows": [[p, round(v / 1000, 2)] for p, v in d["gen"]][-10:], "annual": [[y, round((a["Total renewable generation"][y] or 0) / 1000, 1)] for y in years], "src": "DESNZ, Energy Trends table 6.1 (quarterly)", "edition": meta}
@series("genstor", name="Electricity storage output", unit="TWh", cad="Q")
def s_genstor():
    r, a, m = gen_series(["Pumped storage", "Battery storage"], ["Energy storage"]); return {"rows": r, "annual": a, "src": "DESNZ, Energy Trends table 5.1: pumped storage and grid batteries", "edition": m}

def _et56():
    url, meta = govuk_attachment(ET5, r"ET_5\.6_[A-Z]{3}_\d{2}\.xlsx$"); rows = sheet_rows(workbook(url), "Quarterly imports and exports")
    hi = next(i for i, r in enumerate(rows) if r and norm(r[0]) == "Year"); H = [norm(h).lower() for h in rows[hi]]
    ji, je = H.index("uk total imports"), H.index("uk total exports"); out = {"imp": [], "exp": []}
    for r in rows[hi+1:]:
        if not r or num(r[0]) is None: continue
        m = re.match(r"Quarter (\d)", norm(r[1]))
        if not m or num(r[ji]) is None: continue
        p = f"{int(num(r[0]))} Q{m[1]}"; out["imp"].append([p, round(num(r[ji]) / 1000, 2)]); out["exp"].append([p, round(num(r[je]) / 1000, 2)])
    url2, m2 = govuk_attachment(DUKES5, r"DUKES_5\.13\.xlsx$"); rows = sheet_rows(workbook(url2), "5.13")
    hi = next(i for i, r in enumerate(rows) if r and norm(r[0]) == "Year"); H = [norm(h) for h in rows[hi]]
    ai, ae = H.index("Total Imports (to UK)"), H.index("Total Exports (from UK)"); ann = {"imp": [], "exp": []}
    for r in rows[hi+1:]:
        if r and num(r[0]) is not None and num(r[ai]) is not None: ann["imp"].append([str(int(num(r[0]))), round(num(r[ai]) / 1000, 1)]); ann["exp"].append([str(int(num(r[0]))), round(num(r[ae]) / 1000, 1)])
    return out, ann, meta

@series("imports", name="Electricity imports", unit="TWh", cad="Q")
def s_imports():
    q, a, m = _et56(); return {"rows": q["imp"][-10:], "annual": a["imp"][-4:], "src": "DESNZ, Energy Trends table 5.6 (quarterly)", "edition": m}
@series("exports", name="Electricity exports", unit="TWh", cad="Q")
def s_exports():
    q, a, m = _et56(); return {"rows": q["exp"][-10:], "annual": a["exp"][-4:], "src": "DESNZ, Energy Trends table 5.6 (quarterly)", "edition": m}

def _et52():
    url, meta = govuk_attachment(ET5, r"ET_5\.2_[A-Z]{3}_\d{2}\.xlsx$"); wb = workbook(url); out = {}
    for sh in ("Quarter", "Annual"):
        rows = sheet_rows(wb, sh); hi = next(i for i, r in enumerate(rows) if r and r[0] and str(r[0]).startswith("Components of supply")); hdr = rows[hi]
        cols = [(j, period_from_header(h)) for j, h in enumerate(hdr) if j and h and period_from_header(h)]
        t = {}
        for r in rows[hi+1:]:
            if r and r[0]: t[norm(r[0])] = [[p, num(r[j])] for j, p in cols if num(r[j]) is not None]
        out[sh] = t
    return out, meta

def dem_series(keys, dec=2):
    d, meta = _et52()
    def comb(t):
        n = min(len(t[k]) for k in keys); return [[t[keys[0]][i][0], round(sum(t[k][i][1] for k in keys) / 1000, dec)] for i in range(n)]
    return comb(d["Quarter"])[-10:], comb(d["Annual"])[-4:], meta

@series("demtotal", name="Total electricity demand", unit="TWh", cad="Q")
def s_demtotal():
    r, a, m = dem_series(["Total demand"]); return {"rows": r, "annual": a, "src": "DESNZ, Energy Trends table 5.2: total demand (final consumption plus energy-industry use and losses)", "edition": m}
@series("demdom", name="Domestic electricity consumption", unit="TWh", cad="Q")
def s_demdom():
    r, a, m = dem_series(["Domestic (excluding electric vehicles)"]); return {"rows": r, "annual": a, "src": "DESNZ, Energy Trends table 5.2 (domestic, excluding home charging of electric vehicles)", "edition": m}
@series("demind", name="Industrial electricity consumption", unit="TWh", cad="Q")
def s_demind():
    r, a, m = dem_series(["Iron and steel", "Other industries"]); return {"rows": r, "annual": a, "src": "DESNZ, Energy Trends table 5.2: iron and steel plus other industries", "edition": m}
@series("demcom", name="Commercial and other final users", unit="TWh", cad="Q")
def s_demcom():
    r, a, m = dem_series(["Other final users (excluding electric vehicles)"]); return {"rows": r, "annual": a, "src": "DESNZ, Energy Trends table 5.2: other final users (commercial, public administration, agriculture; excluding electric vehicles and transport)", "edition": m}
@series("losses", name="Transmission and distribution losses", unit="TWh", cad="Q")
def s_losses():
    r, a, m = dem_series(["Losses"]); return {"rows": r, "annual": a, "src": "DESNZ, Energy Trends table 5.2", "edition": m}

@series("dc", name="Data-centre electricity consumption (GB)", unit="TWh", cad="A")
def s_dc():
    d = json.loads(get("https://www.gov.uk/api/search.json?q=data+centre+electricity+consumption&filter_organisations=department-for-energy-security-and-net-zero&order=-public_timestamp&count=10"))
    hits = [r for r in d.get("results", []) if "data centre electricity consumption" in r.get("title", "").lower()]
    res = hits[0] if hits else {"link": "/government/publications/energy-trends-june-2026-special-feature-article-data-centre-electricity-consumption-in-great-britain-2020-to-2024"}
    meta_d = json.loads(get("https://www.gov.uk/api/content" + res["link"]))
    att = next(a for a in meta_d["details"]["attachments"] if a.get("url", "").endswith(".xlsx"))
    rows = sheet_rows(workbook(att["url"]), "Table_2"); hi = next(i for i, r in enumerate(rows) if r and norm(r[0]) == "Area"); yc = year_cols(rows[hi])
    r = next(r for r in rows[hi+1:] if r and norm(r[0]) == "Great Britain")
    hist = [[y, round(num(r[yc[y]]), 2)] for y in sorted(yc)]
    # NESO Ten Year Outlook / FES electricity demand summary (ED1): the regular forward estimate
    pk = json.loads(get("https://api.neso.energy/api/3/action/package_show?id=fes-electricity-demand-summary-data-table-ed1"))
    rsrc = sorted([x for x in pk["result"]["resources"] if x.get("url", "").endswith(".csv")], key=lambda x: x.get("last_modified") or "", reverse=True)[0]
    rr = list(_csv.reader(io.StringIO(get(rsrc["url"])))); hdr = rr[0]
    row = next(x for x in rr if x[1] == "Data Centres" and x[2] == "GWh")
    proj = [[hdr[j], round(float(row[j]) / 1000, 2)] for j in range(6, len(hdr)) if row[j] and int(hdr[j]) <= dt.date.today().year + 5]
    return {"rows": hist[-5:], "projection": proj, "projection_src": f"NESO {rsrc.get('name') or 'electricity demand summary (ED1)'}", "src": "DESNZ, Energy Trends special feature: data-centre electricity consumption in Great Britain (metered colocation, hosting and hyperscale sites); projection from NESO's electricity demand summary", "edition": {"title": meta_d.get("title"), "page": "https://www.gov.uk" + res["link"], "published": meta_d.get("public_updated_at", "")[:10]}}


# --- Capacity in use and peak demand: Elexon half-hourly generation by fuel (monthly maxima) and NESO demand data ---
ELEXON = "https://data.elexon.co.uk/bmrs/api/v1"
FUEL_GROUPS = {"NUCLEAR": ["NUCLEAR"], "GAS": ["CCGT", "OCGT"], "BASE": ["NUCLEAR", "BIOMASS"], "DISP": ["CCGT", "OCGT", "OIL", "COAL", "NUCLEAR", "BIOMASS", "NPSHYD", "PS", "OTHER"], "PS": ["PS"], "WIND": ["WIND"], "BIOMASS": ["BIOMASS"],
               "TOTAL": ["CCGT", "OCGT", "OIL", "COAL", "NUCLEAR", "BIOMASS", "NPSHYD", "PS", "OTHER", "WIND"]}
def elexon_monthly(months_back=24):
    """Monthly maximum half-hourly output (MW) and energy (TWh) per fuel group from Elexon's FUELHH dataset, with a
    cache in data/elexon_fuel_monthly.json so that only the current and previous month are re-read each day."""
    cache_p = ROOT / "data/elexon_fuel_monthly.json"; cache = json.loads(cache_p.read_text()) if cache_p.exists() else {}
    today = dt.date.today(); first = (today.replace(day=1) - dt.timedelta(days=31 * months_back)).replace(day=1)
    months = []; m = first
    while m <= today.replace(day=1): months.append(m); m = (m.replace(day=28) + dt.timedelta(days=4)).replace(day=1)
    for m in months:
        key = m.strftime("%Y-%m"); nxt = (m.replace(day=28) + dt.timedelta(days=4)).replace(day=1); last = nxt - dt.timedelta(days=1)
        if cache.get(key, {}).get("complete") and key not in (today.strftime("%Y-%m"), (today.replace(day=1) - dt.timedelta(days=1)).strftime("%Y-%m")): continue
        per = {}  # (date, sp) -> {fuel: MW}
        d0 = m
        while d0 <= min(last, today):
            d1 = min(d0 + dt.timedelta(days=6), last, today)
            data = json.loads(get(f"{ELEXON}/datasets/FUELHH?settlementDateFrom={d0}&settlementDateTo={d1}&format=json", timeout=120)).get("data", [])
            for r in data:
                k = (r["settlementDate"], r["settlementPeriod"]); per.setdefault(k, {})[r["fuelType"]] = r.get("generation") or 0
            d0 = d1 + dt.timedelta(days=1)
        if not per: continue
        mx = {}; twh = {}
        for g, fuels in FUEL_GROUPS.items():
            vals = [sum(v.get(f, 0) for f in fuels) for v in per.values()]
            mx[g] = round(max(vals) / 1000, 2); twh[g] = round(sum(vals) / 2 / 1e6, 3)
        cache[key] = {"max_gw": mx, "twh": twh, "periods": len(per), "complete": last < today}
    cache = {k: cache[k] for k in sorted(cache)}
    cache_p.write_text(json.dumps(cache, indent=0)); return cache

def tec_snapshot():
    """Built transmission-connected capacity by plant type from NESO's TEC register (a current snapshot, MW)."""
    pk = json.loads(get("https://api.neso.energy/api/3/action/package_show?id=transmission-entry-capacity-tec-register"))["result"]
    r = pk["resources"][0]; rows = list(_csv.DictReader(io.StringIO(get(r["url"], timeout=120).lstrip("﻿"))))
    agg = {}
    for x in rows:
        if x.get("Project Status", "").strip() != "Built": continue
        try: mw = float((x.get("MW Connected") or "0").replace(",", ""))
        except ValueError: continue
        agg[x.get("Plant Type", "")] = agg.get(x.get("Plant Type", ""), 0) + mw
    def grp(pred): return round(sum(v for k, v in agg.items() if pred(k)) / 1000, 2)
    snap = {"nuclear": grp(lambda k: k == "Nuclear"), "gas": grp(lambda k: k.startswith("CCGT") or k.startswith("OCGT")), "storage": grp(lambda k: "Energy Storage System" in k or "Pump Storage" in k), "pumped": grp(lambda k: k == "Pump Storage"), "total": grp(lambda k: "Demand" not in k and "Reactive" not in k and "Hydrogen" not in k)}
    return snap, {"title": r.get("name"), "page": "https://www.neso.energy/data-portal/transmission-entry-capacity-tec-register", "published": (r.get("last_modified") or "")[:10]}

def cap_series(group, name, unit="GW"):
    c = elexon_monthly(); rows = [[k, v["max_gw"][group]] for k, v in c.items() if v["periods"] > 100][-24:]
    return {"rows": rows, "twh": [[k, v["twh"][group]] for k, v in c.items()][-24:], "src": "Elexon Insights, half-hourly generation by fuel type (FUELHH): maximum half-hourly output in the month, transmission-metered GB", "edition": {"title": "Elexon FUELHH", "page": "https://bmrs.elexon.co.uk/generation-by-fuel-type", "published": dt.date.today().isoformat()}}

@series("captotal", name="Generation capacity in use (all transmission-metered fuels)", unit="GW", cad="M")
def s_captotal():
    d = cap_series("TOTAL", "total"); snap, meta = tec_snapshot(); d["tec"] = [["Built transmission-connected capacity, TEC register " + meta["published"], snap["total"]]]; return d
@series("capbase", name="Baseload capacity in use (nuclear and biomass)", unit="GW", cad="M")
def s_capbase(): return cap_series("BASE", "base")
@series("capnuc", name="Nuclear capacity in use", unit="GW", cad="M")
def s_capnuc():
    d = cap_series("NUCLEAR", "nuclear"); snap, meta = tec_snapshot(); d["tec"] = [["Connected nuclear capacity, TEC register " + meta["published"], snap["nuclear"]]]; return d
@series("capdisp", name="Dispatchable capacity in use", unit="GW", cad="M")
def s_capdisp(): return cap_series("DISP", "dispatchable")
@series("capgas", name="Gas-fired capacity in use", unit="GW", cad="M")
def s_capgas():
    d = cap_series("GAS", "gas"); snap, meta = tec_snapshot(); d["tec"] = [["Connected CCGT and OCGT capacity, TEC register " + meta["published"], snap["gas"]]]; return d
@series("capstor", name="Storage capacity in use (pumped storage output)", unit="GW", cad="M")
def s_capstor():
    d = cap_series("PS", "pumped"); snap, meta = tec_snapshot(); d["tec"] = [["Connected storage (batteries and pumped), TEC register " + meta["published"], snap["storage"]], ["of which pumped storage", snap["pumped"]]]; return d

@series("peak", name="Peak transmission demand (GB)", unit="GW", cad="M")
def s_peak():
    """Monthly maximum half-hourly transmission system demand (TSD) and national demand (ND) from NESO's historic demand data."""
    pk = json.loads(get("https://api.neso.energy/api/3/action/package_show?id=historic-demand-data"))["result"]
    today = dt.date.today(); years = {today.year, today.year - 1, today.year - 2}; mx = {}
    for r in pk["resources"]:
        yr = re.search(r"(20\d\d)", r.get("name", "") or ""); 
        if not yr or int(yr[1]) not in years or not r.get("url", "").endswith(".csv"): continue
        rows = list(_csv.DictReader(io.StringIO(get(r["url"], timeout=180).lstrip("﻿"))))
        for x in rows:
            if x.get("FORECAST_ACTUAL_INDICATOR", "A").strip() not in ("A", ""): continue
            ds = x.get("SETTLEMENT_DATE", "").strip(); m = None
            for fmt in ("%Y-%m-%d", "%d-%b-%y", "%d-%b-%Y", "%d/%m/%Y"):
                try: m = dt.datetime.strptime(ds[:10] if fmt == "%Y-%m-%d" else ds, fmt).strftime("%Y-%m"); break
                except ValueError: pass
            if not m: continue
            try: tsd = float(x.get("TSD") or 0); nd = float(x.get("ND") or 0)
            except ValueError: continue
            cur = mx.setdefault(m, [0, 0]); cur[0] = max(cur[0], tsd); cur[1] = max(cur[1], nd)
    rows = [[k, round(v[0] / 1000, 2), round(v[1] / 1000, 2)] for k, v in sorted(mx.items())][-24:]
    pk2 = json.loads(get("https://api.neso.energy/api/3/action/package_show?id=fes-electricity-demand-summary-data-table-ed1"))
    rsrc = sorted([x for x in pk2["result"]["resources"] if x.get("url", "").endswith(".csv")], key=lambda x: x.get("last_modified") or "", reverse=True)[0]
    rr = list(_csv.reader(io.StringIO(get(rsrc["url"])))); hdr = rr[0]; row = next(x for x in rr if x[1] == "Total Underlying Peak")
    und = [[hdr[j], round(float(row[j]), 2)] for j in range(6, len(hdr)) if row[j] and int(hdr[j]) <= today.year + 1]
    return {"rows": rows, "cols": ["Transmission system demand", "National demand"], "underlying": und, "src": "NESO historic demand data: monthly maximum half-hourly transmission system demand and national demand, Great Britain", "edition": {"title": "NESO historic demand data", "page": "https://www.neso.energy/data-portal/historic-demand-data", "published": today.isoformat()}}

@series("mid", name="Wholesale electricity price, daily (APX market index)", unit="£/MWh", cad="D")
def s_mid():
    """Daily volume-weighted average of Elexon's APX Market Index Data (the exchange price used for imbalance pricing),
    cached in data/elexon_mid_daily.json; only the last 10 days are re-read each run."""
    cache_p = ROOT / "data/elexon_mid_daily.json"; cache = json.loads(cache_p.read_text()) if cache_p.exists() else {}
    today = dt.date.today(); start = today - dt.timedelta(days=740)
    d0 = max(start, today - dt.timedelta(days=10)) if cache else start
    d0 = min(d0, min((dt.date.fromisoformat(k) for k in cache), default=today)) if cache and dt.date.fromisoformat(min(cache)) > start else d0
    while d0 <= today:
        d1 = min(d0 + dt.timedelta(days=6), today)
        data = json.loads(get(f"{ELEXON}/balancing/pricing/market-index?from={d0}T00:00Z&to={d1 + dt.timedelta(days=1)}T00:00Z&dataProviders=APXMIDP", timeout=120)).get("data", [])
        acc = {}
        for r in data:
            k = r["settlementDate"]; v = r.get("volume") or 0; pr = r.get("price")
            if pr is None or v <= 0: continue
            a = acc.setdefault(k, [0.0, 0.0, 0]); a[0] += pr * v; a[1] += v; a[2] += 1
        for k, (pv, v, n) in acc.items():
            if v > 0 and (k != today.isoformat()): cache[k] = {"vwap": round(pv / v, 2), "volume_mwh": round(v), "periods": n}
        d0 = d1 + dt.timedelta(days=1)
    cache = {k: cache[k] for k in sorted(cache) if dt.date.fromisoformat(k) >= start}
    cache_p.write_text(json.dumps(cache, separators=(",", ":")))
    rows = [[k, v["vwap"]] for k, v in cache.items() if v["periods"] >= 40]
    return {"rows": rows, "src": "Elexon Insights, APX market index data (MID): daily volume-weighted average price of exchange trades, GB", "edition": {"title": "Elexon MID, APXMIDP", "page": "https://bmrs.elexon.co.uk/market-index-price", "published": today.isoformat()}}

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
            d = f(); d.pop("annual", None); d.update(meta); d["fetched"] = now; d["stale"] = False; d.pop("error", None); result[k] = d; ok += 1
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

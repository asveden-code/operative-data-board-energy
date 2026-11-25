#!/usr/bin/env python3
"""Assemble index.html from template.html, app.js, data/series.json, quotes.json and the bundled fallbacks in src/.
Run from the repository root:  python scripts/build.py [--built-at "25 Nov 2026"]"""
import json, re, sys, datetime as dt
from pathlib import Path
root = Path(__file__).resolve().parent.parent
built_at = sys.argv[sys.argv.index("--built-at") + 1] if "--built-at" in sys.argv else dt.date.today().strftime("%-d %b %Y")
app = (root / "app.js").read_text(); t = (root / "template.html").read_text()
con = json.loads((root / "src/con_min.json").read_text()) if (root / "src/con_min.json").exists() else {"pm": {}, "mf": {}}
keep_pm = set(re.findall(r'\{pm:"([^"]+)"\}', app)); keep_mf = set(re.findall(r'\{mf:"([^"]+)"\}', app))
con["pm"] = {k: v for k, v in con["pm"].items() if k in keep_pm}; con["mf"] = {k: v for k, v in con["mf"].items() if k in keep_mf}
snap = json.loads((root / "src/snap_min.json").read_text()) if (root / "src/snap_min.json").exists() else {}
series = json.loads((root / "data/series.json").read_text())
q = json.loads((root / "quotes.json").read_text())
data = (f'const BUILT_AT="{built_at}";\nconst SNAP={json.dumps(snap,separators=(",",":"))};\nconst CON={json.dumps(con,separators=(",",":"))};\n'
        f'const SERIES={json.dumps(series,separators=(",",":"),ensure_ascii=False)};\nconst QUOTES={json.dumps(q,separators=(",",":"),ensure_ascii=False)};')
html = t.replace("/*__DATA__*/", data).replace("/*__APP__*/", app)
(root / "index.html").write_text(html); print("built index.html", len(html))

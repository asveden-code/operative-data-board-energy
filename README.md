# Operative Data and Prediction Markets — UK energy markets

Single-page board: for each UK energy-market indicator, the official data with a two-to-three-year trend, the
prediction-market contracts that price it or its global drivers (Polymarket, Manifold, Metaculus), and recent passages
from BBC Business, CNBC Economy, The Economist and the think tanks and specialist analysts that follow these markets
(Carbon Brief, Cornwall Insight, E3G, Energy UK, Climate Change Committee, Carbon Tracker, NIESR, Resolution
Foundation, Institute for Government, Oxford Economics, Bruegel), linked with W3C text fragments to the exact passage.

Sister board for the key economic factors: https://asveden-code.github.io/operative-data-board/

- `index.html` — the whole page. Works opened as a local file (with the data and quotes bundled when it was built) or hosted.
- `data/series.json` — the operative data for every series that has no browser-readable feed (DESNZ DUKES and Energy
  Trends spreadsheets, Ofgem, NESO, NSTA, MCS, HMRC): rebuilt daily by `scripts/fetch_series.py`; the hosted page
  fetches it on load. Series with an open feed (IMF Brent, Ofgem's NBP chart data, Elexon) are fetched live by the browser.
- `quotes.json` — the daily media selection, served at https://asveden-code.github.io/operative-data-board-energy/quotes.json.
- `scripts/fetch_series.py` — downloads the publisher files and writes `data/series.json` (deterministic; no judgement).
- `scripts/fetch_candidates.py` — gathers candidate articles for the daily quote review.
- `scripts/verify_quotes.py` — verifies every quote against its source and builds the text-fragment links.
- `scripts/build.py` — rebuilds `index.html` from `template.html`, `app.js`, the data files and the fallbacks in `src/`.
- `JOB.md` — instructions the daily scheduled refresh follows.

Hosting: GitHub Pages, source "Deploy from a branch" (main, root). Nothing else to configure.

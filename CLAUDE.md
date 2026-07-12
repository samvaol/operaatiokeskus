# Operaatiokeskus — project guide

A single-page **TV dashboard** for the **Kaiku 2026** scout-camp operations centre, plus a
small **Node ticket-server** that feeds it SharePoint tickets. Everything is Finnish-first
and uses the **Kaiku 2026 V1** visual identity. Sibling projects: `../Leirilukkari`,
`../Leirikartta` (both embeddable HTML).

## ⚠️ Read first
- The dashboard is **one self-contained `index.html`** — HTML + CSS + vanilla JS, **no
  build step, no framework, no dependencies**. Keep it that way. Edit the file directly.
- It targets a **1920×1080 TV**. Sizing is rem-based via `html{font-size:clamp(9px,calc(.42vw+.60vh),24px)}`
  — size things in `rem`, not `px`.
- **Bilingual context but UI is Finnish.** Match the surrounding Finnish copy.

## Structure
```
index.html              # the whole dashboard (styles + markup + script in one file)
README.md               # user-facing setup (incl. Node-on-Windows)
progress.md             # running dev log
ticket-server/
  server.js             # Express + Playwright: login → read all tickets (grouped) → serve JSON
  public/index.html     # Kaiku-styled all-statuses ticket board served by the server
  package.json          # express + playwright
  .gitignore            # node_modules/ .auth/   (NEVER commit .auth — login session)
```

## Data sources (all reached from the browser except tickets)
- **Weather + forest-fire**: FMI WFS `fmi::forecast::edited::weather::scandinavia::point::timevaluepair`,
  `latlon=61.208,25.128`, params incl. `SmartSymbol` (night = code+100) and `ForestFireWarning`
  (NaN/1 = none, ≥2 = active). CORS `*`.
- **Pääuutiset**: `https://api.ww-api.com/front/get_items/4554399/78074354/` → `{items:[…]}`,
  CORS `*`. Section **78074354 = Pääuutiset** (not the widget 78074355). Content URL discovered
  via `kaiku2026.coregoapp.com/apiv4/getSettings?platform=webapp`
  (`gbsettings.sections.<id>.contentSource.url`).
- **Sadetutka (radar)**: Leaflet map on `61.204934767500795,25.1210434592283`; FMI radar WMS
  `openwms.fmi.fi/geoserver/wms` layer `Radar:suomi_dbz_eureffin` (EPSG:3857, CORS `*`) over a
  CARTO light base. Leaflet loaded from unpkg CDN. `.radarwrap` uses `isolation:isolate` so
  Leaflet's z-index 200–700 panes don't paint over the ticket modal (z 900).
- **Schedule**: embedded whole-camp snapshot + optional live `kaiku2026.fi/api/schedules`
  (usually CORS-blocked → snapshot used).
- **Työvuorot**: embedded `WORKSHIFTS` object parsed from `Operaatiokeskuksen työvuorolista.xlsx`.
- **Tiketit**: via the local `ticket-server` (see below), never SharePoint directly.
- **Konfetti**: `tickTimer` fires `celebrate(pct)` (canvas confetti + toast) on each whole-percent
  advance of the camp progress; `#confetti`/`#celebrateToast` at z 1000/1100.

Every source has an **embedded fallback** so the dashboard never goes blank.

## Ticket server
- SharePoint **can't be read from the browser**: REST sends `Access-Control-Allow-Origin: *`
  but no `Access-Control-Allow-Credentials` (cookie blocked cross-origin), and list pages set
  `X-Frame-Options` (no iframe). So the Node server logs in via a real Playwright browser
  window and reads REST **from inside** the authenticated page (same-origin).
- List referenced **by URL** `/sites/Tiketin/Lists/OpkeOspa` via `_api/web/getList(@l)?@l='…'`
  (not the GUID — an authed call to a wrong GUID 404s). Status field = **`Status`** (choice;
  7 values Uusi…Ei käsitellä). Fetches **all** items, groups by status.
- Extraction uses **`context.request.get`** (shares the logged-in context cookies; immune to
  which tab is open), with a server-owned background page (`requestJsonViaPage`) as fallback.
  Do **not** poll via `page.evaluate` on a user-visible tab — it breaks on SPA nav / tab switch.
- Serves `GET /api/tickets` (CORS `*`) → `{status:'ok'|'awaiting-login'|'starting', buckets:[…],
  uusi:[…], count}`. Dashboard's Uusi panel reads `j.uusi`; popup reads `j.buckets`. Override host
  with `?ticketApi=`.
- Run: `cd ticket-server && npm install && npm start`. `HEADLESS=1` runs headless (only works
  once `.auth` is warm). Session persists in `ticket-server/.auth` (git-ignored, secret).

## Conventions
- **Visual identity**: Bricolage Grotesque; colors metsä `#005448`, meri `#00445E`, rusko
  `#542337`, savu `#F9F3E6`, punainen `#FF633A`, oranssi `#FF8940`, kulta `#FFAE40`.
  Each `.card` sets a `--accent` (timer=kulta, weather=meri, shift=rusko, sched=oranssi,
  news=punainen) driving its top stripe + icon chip.
- The header has a **Kaiku-1 gradient edge** (`header::after`). The old stretched aaltoviiva
  was removed on purpose — **do not re-add the header wave.**
- The **kuosi** background SVG must use a wide viewBox (`0 0 1920 220`) or it scales ~5× too
  large on a TV.
- Comments explain *why*, in Finnish, matching the file.

## Verify
- Run a static server and open the dashboard: `python3 -m http.server 8133`.
- Preview **screenshots render at ~½ size** here (devicePixelRatio 2) — trust
  `getComputedStyle` / bounding-rect measurements over screenshots.
- Check `getElementById` targets still exist after edits; no console errors.

## Deploy / workflow
- Repo: <https://github.com/samvaol/operaatiokeskus> (branch `main`). Commit/push only when asked.
- No CI. The dashboard is static; the ticket-server runs on the TV's machine next to it.

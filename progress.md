# Progress — Kaiku 2026 Operaatiokeskus

A running log of what's built, the key decisions, and what's left.

## Status: working

The dashboard (`index.html`) and the `ticket-server` are functional and verified.
Hosted at <https://github.com/samvaol/operaatiokeskus>.

## Done

### Dashboard (`index.html`, single self-contained file)
- **Layout** — 1920×1080 TV, **4-column grid**, rem-scaled via `html{font-size:clamp(vw+vh)}`.
  Col 1: Leirikello + Työvuoro + Uudet tiketit · Col 2: Sää + Sadetutka · Col 3: Ohjelma · Col 4: Pääuutiset.
- **Greeting + clock** — Helsinki-time greeting (huomenta/päivää/iltaa/yötä) + live clock
  + date, plus the animated **kaiku equalizer / LIVE** signature.
- **Leirikello** — elapsed camp timer from `2026-07-10T09:00+03:00` to `18.7. 16:30`,
  with progress bar. **Confetti + "Leiri N % takana!" toast** fires each whole-percent
  advance (canvas confetti, no lib; respects `prefers-reduced-motion`).
- **Sää · Evo** — live FMI WFS forecast (temp, feels-like, wind, rain, humidity, day/night
  SmartSymbol strip) + `ForestFireWarning`. CORS-open (`*`).
- **Sadetutka · Evo** — Leaflet map centred on `61.204934767500795, 25.1210434592283`,
  FMI radar WMS (`Radar:suomi_dbz_eureffin`, CORS `*`) over a CARTO light base, refreshed
  every 5 min. Non-interactive. `.radarwrap` needs `isolation:isolate` so Leaflet's internal
  z-indexes (200–700) don't paint over the ticket modal.
- **Uudet tiketit** — Uusi tickets always on screen (col 1), 60 s refresh from ticket-server.
- **Päivän ohjelma** — today's whole-camp events (nyt/seuraava), embedded snapshot +
  optional live `kaiku2026.fi/api/schedules`.
- **Työvuorossa nyt** — current 1./2. shift from `Operaatiokeskuksen työvuorolista.xlsx`
  (embedded), carried forward to now; overnight falls back to previous evening.
  "Operaatiokeskuksen päiväpalaveri" pinned to **16:00** every day.
- **Pääuutiset** — live from `api.ww-api.com/front/get_items/4554399/78074354/`
  (section 78074354 = Pääuutiset), 5-article embedded fallback.
- **Tiketit** — 🎫 header button opens a Kaiku popup modal of the **full board (all status
  columns grouped)**, refreshed every 60 s from the `ticket-server`.

### Ticket server (`ticket-server/`, Node + Express + Playwright)
- Opens a login window to the SharePoint Tiketin site; session persisted in `.auth`.
- Reads **all tickets** from the *Opke/Ospa* list (referenced by URL
  `/sites/Tiketin/Lists/OpkeOspa`, not the GUID) every 60 s, groups them by the `Status`
  field into the 7 status buckets, and picks display fields by their SharePoint column titles.
- Extraction uses **`context.request`** (carries the logged-in cookies; immune to tab
  navigation), with a server-owned background page as fallback. This replaced the earlier
  `page.evaluate`-in-a-visible-tab approach, which broke on SPA navigation / tab switches.
- Serves CORS-open `GET /api/tickets` → `{status, buckets, uusi, count}` (+ `/api/health`)
  and a Kaiku all-statuses board at `/`.
- **Needs live verification** against a logged-in SharePoint session (restart the server to
  load new code). Verified locally: syntax, endpoints, `awaiting-login` fallback; REST list
  paths return 403 unauth (valid).

### Design
- Kaiku 2026 V1 identity throughout (Bricolage Grotesque; metsä/savu + punainen/oranssi/
  kulta). Coordinated per-panel accent system, accent icon-chips, layered shadows,
  Kaiku-1 gradient header edge (replaced the disliked stretched aaltoviiva — do not re-add).

## Key decisions / gotchas
- **SharePoint can't be read from the browser directly** — `Access-Control-Allow-Origin: *`
  without `Access-Control-Allow-Credentials` blocks the auth cookie cross-origin, and the
  list page sets `X-Frame-Options` (no iframe). Hence the separate `ticket-server`.
- **kuosi pattern** must use a wide viewBox (`0 0 1920 220`) or it scales ~5× on a TV.
- Preview screenshots render at ~½ size here due to devicePixelRatio 2 — trust DOM
  measurements over screenshots.

## Possible next steps
- Auto-launch `ticket-server` on the TV machine at boot (e.g. `pm2` / a login item).
- Optional: run the server headless after the first login (`HEADLESS=1`) once `.auth` is warm.
- Wire live `kaiku2026.fi/api/schedules` if/when CORS allows, instead of the snapshot.
- Refresh embedded snapshots (news/schedule/shifts) if the source data changes.

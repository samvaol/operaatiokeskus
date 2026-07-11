# Progress — Kaiku 2026 Operaatiokeskus

A running log of what's built, the key decisions, and what's left.

## Status: working

The dashboard (`index.html`) and the `ticket-server` are functional and verified.
Hosted at <https://github.com/samvaol/operaatiokeskus>.

## Done

### Dashboard (`index.html`, single self-contained file)
- **Layout** — 1920×1080 TV, 3-column grid, rem-scaled via `html{font-size:clamp(vw+vh)}`.
- **Greeting + clock** — Helsinki-time greeting (huomenta/päivää/iltaa/yötä) + live clock
  + date, plus the animated **kaiku equalizer / LIVE** signature.
- **Leirikello** — elapsed camp timer from `2026-07-10T09:00+03:00` to `18.7. 16:30`,
  with progress bar.
- **Sää · Evo** — live FMI WFS forecast (temp, feels-like, wind, rain, humidity, day/night
  SmartSymbol strip) + `ForestFireWarning`. CORS-open (`*`).
- **Päivän ohjelma** — today's whole-camp events (nyt/seuraava), embedded snapshot +
  optional live `kaiku2026.fi/api/schedules`.
- **Työvuorossa nyt** — current 1./2. shift from `Operaatiokeskuksen työvuorolista.xlsx`
  (embedded), carried forward to now; overnight falls back to previous evening.
  "Operaatiokeskuksen päiväpalaveri" pinned to **16:00** every day.
- **Pääuutiset** — live from `api.ww-api.com/front/get_items/4554399/78074354/`
  (section 78074354 = Pääuutiset), 5-article embedded fallback.
- **Tiketit** — 🎫 header button opens a Kaiku popup modal of the new tickets, refreshed
  every 60 s from the `ticket-server`.

### Ticket server (`ticket-server/`, Node + Express + Playwright)
- Opens a login window to the SharePoint Tiketin site; session persisted in `.auth`.
- Reads **Status = "Uusi"** items from list `df73229b-1f4b-4e2a-b342-c91b7dbd8a12`
  (`/sites/Tiketin`) every 60 s via same-origin REST fetch inside the logged-in page.
- Serves CORS-open `GET /api/tickets` (+ `/api/health`) and a Kaiku board at `/`.
- Verified: Express + endpoints + poll pipeline all work (headless test → `awaiting-login`
  when not signed in; full Chromium installed for the real headed login).

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

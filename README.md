# Kaiku 2026 · Operaatiokeskus

A single-page TV dashboard for the **Kaiku 2026** scout camp operations centre
(*operaatiokeskus*), built in the Kaiku 2026 visual identity. Everything lives in one
self-contained `index.html` (HTML + CSS + vanilla JS, no build step) sized for a
1920×1080 TV.

## Panels

- **Tervehdys + kello** — time-of-day greeting, live clock, and an animated "kaiku"
  equalizer / LIVE indicator.
- **Leirikello** — elapsed camp time since 10.7.2026 09:00 with a progress bar to 18.7.
- **Sää · Evo, Hämeenlinna** — live forecast from the Finnish Meteorological Institute
  (Ilmatieteen laitos), including the forest-fire warning (*metsäpalovaroitus*).
- **Päivän ohjelma** — today's whole-camp programme with *nyt* / *seuraava* markers.
- **Työvuorossa nyt** — who is on shift right now (from the työvuorolista).
- **Pääuutiset** — the latest articles from the Kaiku 2026 app.
- **Tiketit** (🎫 button) — a popup of the new SharePoint tickets, refreshed every 60 s.

## Data sources

| Panel | Source |
|-------|--------|
| Weather + forest-fire | Ilmatieteen laitos open data (WFS) |
| Pääuutiset | Kaiku 2026 app content API (Corego / GoodBarber) |
| Päivän ohjelma | Leirilukkari camp schedule + embedded snapshot |
| Työvuorot | `Operaatiokeskuksen työvuorolista.xlsx` (embedded) |
| Tiketit | SharePoint list *Opke/Ospa* (REST, status = "Uusi") |

Each source has an embedded fallback so the dashboard keeps working offline.

### SharePoint tickets note

The tickets popup calls the SharePoint REST API for items whose status is **Uusi**.
Because SharePoint returns `Access-Control-Allow-Origin: *` **without**
`Access-Control-Allow-Credentials`, a browser will not send the login cookie
cross-origin, so a standalone page cannot read the list directly. To make it live,
serve this page **same-origin on SharePoint**, or put a small proxy / Microsoft Graph
app-token in front of it. Until then the popup shows a "Kirjaudu" prompt.

## Run it

Any static file server works, e.g.:

```bash
python3 -m http.server 8133   # then open http://localhost:8133
```

On the TV, open `index.html` and go full-screen (F11).

## Visual identity

Kaiku 2026 V1 — Bricolage Grotesque, metsä `#005448`, savu `#F9F3E6`, and the
punainen / oranssi / kulta accent trio. Each panel is colour-coded with a Kaiku accent.

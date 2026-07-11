'use strict';
/*
 * Operaatiokeskus · Tiketti-palvelin
 * ----------------------------------
 * Avaa selainikkunan SharePointin Tiketin-sivustolle, jossa kirjaudut kerran
 * partio-tunnuksilla. Sen jälkeen palvelin lukee 60 s välein "Uusi"-tilaiset
 * tiketit Opke/Ospa-listalta (SharePoint REST -rajapinta ajetaan sisältä
 * kirjautuneelta sivulta, joten evästeet kulkevat samasta originista) ja
 * tarjoaa ne CORS-avoimena JSON:ina + Kaiku-tyylisenä sivuna.
 *
 * Käyttö:
 *   npm install          (asentaa Expressin + Playwrightin selaimineen)
 *   npm start            (avaa kirjautumisikkunan, käynnistää palvelimen)
 * Ympäristömuuttujat:
 *   PORT=8137            palvelimen portti
 *   HEADLESS=1           aja selain piilossa (toimii vain jos jo kirjautunut)
 */
const path = require('path');
const express = require('express');
const { chromium } = require('playwright');

const PORT = Number(process.env.PORT || 8137);
const HEADLESS = process.env.HEADLESS === '1';
const USER_DATA_DIR = path.join(__dirname, '.auth');

// --- SharePointin Opke/Ospa-lista (Tiketin-sivusto) ---
const SITE = 'https://partio.sharepoint.com/sites/Tiketin';
const LIST_GUID = 'df73229b-1f4b-4e2a-b342-c91b7dbd8a12';
const VIEW_ID = '6ab91000-1635-4e47-b2a8-cb248f338cff';
const STATUS_URL = `${SITE}/Lists/OpkeOspa/AllItems.aspx?viewid=${VIEW_ID}`;
const REST_URL =
  `${SITE}/_api/web/lists(guid'${LIST_GUID}')/items` +
  `?$select=*,Author/Title&$expand=Author&$filter=Status eq 'Uusi'` +
  `&$top=200&$orderby=Created desc`;
const REST_URL_PLAIN =
  `${SITE}/_api/web/lists(guid'${LIST_GUID}')/items` +
  `?$filter=Status eq 'Uusi'&$top=200&$orderby=Created desc`;

// --- palvelimen tila (välimuisti dashboardille) ---
const state = { status: 'starting', tickets: [], updatedAt: null, error: null, lastHttp: null };

// --- yhden tiketin kentät → dashboardin muoto ---
const first = (it, keys) => { for (const k of keys) if (it[k] != null && it[k] !== '') return it[k]; return ''; };
const stripHtml = (h) => String(h || '').replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/\s+/g, ' ').trim();
function mapItem(it) {
  return {
    id: it.Id ?? it.ID ?? '',
    title: first(it, ['Title', 'Otsikko', 'LinkTitle']) || '(nimetön tiketti)',
    desc: stripHtml(first(it, ['Kuvaus', 'Description', 'Selite', 'Viesti', 'Ongelma', 'Lisatiedot'])).slice(0, 400),
    topic: first(it, ['Aihe', 'Tapahtuma', 'Category', 'Luokka']),
    location: first(it, ['Sijainti', 'Location', 'Paikka']),
    reporter: first(it, ['Ilmoittaja', 'Ilmoittaja0']) || (it.Author && it.Author.Title) || '',
    safety: it.Vaikuttaakovikaturvallisuuteen === true || /kyll/i.test(String(first(it, ['Vaikuttaakovikaturvallisuuteen']))),
    created: it.Created || null
  };
}

// --- lue tiketit kirjautuneen sivun kontekstissa (same-origin fetch) ---
async function fetchTickets(page) {
  return page.evaluate(async ([urlExpand, urlPlain]) => {
    const get = async (u) => {
      const r = await fetch(encodeURI(u), { headers: { Accept: 'application/json;odata=nometadata' }, credentials: 'include' });
      const ct = r.headers.get('content-type') || '';
      return { ok: r.ok, status: r.status, json: ct.includes('json') ? await r.json().catch(() => null) : null };
    };
    let r = await get(urlExpand);
    if (!r.ok || !r.json) r = await get(urlPlain); // Author-expand voi kaatua → yritä ilman
    if (r.ok && r.json && Array.isArray(r.json.value)) return { ok: true, value: r.json.value };
    return { ok: false, status: r.status };
  }, [REST_URL, REST_URL_PLAIN]);
}

async function poll(page) {
  try {
    const res = await fetchTickets(page);
    if (res.ok) {
      state.tickets = res.value.map(mapItem);
      state.status = 'ok';
      state.updatedAt = new Date().toISOString();
      state.error = null;
      state.lastHttp = 200;
      console.log(`[poll] ${new Date().toLocaleTimeString('fi-FI')} — ${state.tickets.length} uutta tikettiä`);
    } else {
      state.status = 'awaiting-login';
      state.lastHttp = res.status || null;
      console.log(`[poll] ei kirjautunut (HTTP ${res.status || '?'}) — kirjaudu selainikkunassa`);
    }
  } catch (e) {
    state.status = 'awaiting-login';
    state.error = String(e && e.message || e);
    console.log('[poll] virhe:', state.error);
  }
}

// --- HTTP-palvelin ---
const app = express();
app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Cache-Control', 'no-store');
  if (req.method === 'OPTIONS') return res.sendStatus(204);
  next();
});
app.get('/api/tickets', (_req, res) => {
  res.json({ status: state.status, updatedAt: state.updatedAt, count: state.tickets.length, tickets: state.tickets, error: state.error });
});
app.get('/api/health', (_req, res) => res.json({ status: state.status, updatedAt: state.updatedAt, count: state.tickets.length }));
app.use(express.static(path.join(__dirname, 'public')));

(async () => {
  console.log('Käynnistetään tiketti-palvelin…');
  const context = await chromium.launchPersistentContext(USER_DATA_DIR, {
    headless: HEADLESS,
    viewport: { width: 1280, height: 900 },
    args: ['--no-first-run']
  });
  const page = context.pages()[0] || await context.newPage();
  await page.goto(STATUS_URL, { waitUntil: 'domcontentloaded', timeout: 60000 }).catch(() => {});
  if (!HEADLESS) { try { await page.bringToFront(); } catch (_) {} }

  app.listen(PORT, () => {
    console.log(`\n  Tiketti-palvelin:  http://localhost:${PORT}`);
    console.log(`  JSON-rajapinta:    http://localhost:${PORT}/api/tickets`);
    console.log(`  → Kirjaudu avautuneessa selainikkunassa partio-tunnuksilla.\n`);
  });

  await poll(page);
  setInterval(() => poll(page), 60000); // 60 s

  const shutdown = async () => { try { await context.close(); } catch (_) {} process.exit(0); };
  process.on('SIGINT', shutdown);
  process.on('SIGTERM', shutdown);
})().catch((e) => { console.error('Palvelin ei käynnistynyt:', e); process.exit(1); });

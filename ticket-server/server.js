'use strict';
/*
 * Operaatiokeskus · Tiketti-palvelin
 * ----------------------------------
 * Avaa selainikkunan SharePointin Tiketin-sivustolle, jossa kirjaudut kerran
 * partio-tunnuksilla. Sen jälkeen palvelin lukee 60 s välein "Uusi"-tilaiset
 * tiketit Opke/Ospa-listalta ja tarjoaa ne CORS-avoimena JSON:ina + Kaiku-
 * tyylisenä sivuna.
 *
 * Tiketit haetaan Playwrightin context.request-rajapinnalla, joka käyttää
 * kirjautuneen selainkontekstin evästeitä. Näin luku ei riipu siitä, millä
 * välilehdellä käyttäjä on, eikä SharePointin sivunavigointi katkaise sitä.
 *
 * Käyttö:  npm install  &&  npm start
 * Ympäristö:  PORT=8137   HEADLESS=1 (aja piilossa – vaatii valmiin .auth-istunnon)
 */
const path = require('path');
const express = require('express');
const { chromium } = require('playwright');

const PORT = Number(process.env.PORT || 8137);
const HEADLESS = process.env.HEADLESS === '1';
const USER_DATA_DIR = path.join(__dirname, '.auth');

// --- SharePointin Opke/Ospa-lista (Tiketin-sivusto) ---
// Viitataan listaan sen palvelinpolun kautta (varma) eikä GUIDilla.
const SITE = 'https://partio.sharepoint.com/sites/Tiketin';
const VIEW_ID = '6ab91000-1635-4e47-b2a8-cb248f338cff';
const LIST_URL = '/sites/Tiketin/Lists/OpkeOspa';
const LIST_ENC = encodeURIComponent(LIST_URL); // %2Fsites%2FTiketin%2FLists%2FOpkeOspa
const STATUS_URL = `${SITE}/Lists/OpkeOspa/AllItems.aspx?viewid=${VIEW_ID}`;
// Haetaan KAIKKI tiketit (kaikki tilat) ja ryhmitellään tilan mukaan.
const ITEMS_URL = `${SITE}/_api/web/getList(@l)/items`
  + `?@l='${LIST_ENC}'&$orderby=Created%20desc&$top=500`;
const FIELDS_URL = `${SITE}/_api/web/getList(@l)/fields`
  + `?@l='${LIST_ENC}'&$select=InternalName,Title,Hidden&$top=500`;
// Statusnäkymän sarakkeet järjestyksessä
const STATUS_ORDER = ['Uusi', 'Käsittelemättömät', 'Odottaa myöhempää käsittelyä',
  'Käsittelyssä operaatiokeskuksessa', 'Käsittelyssä muilla', 'Valmis', 'Ei käsitellä'];

// --- palvelimen tila (välimuisti dashboardille) ---
const state = { status: 'starting', buckets: [], uusi: [], count: 0, updatedAt: null, error: null, lastHttp: null };
let context = null;
let fieldMap = null; // { internalName: displayTitle }

// --- kentän arvon poiminta näyttönimen perusteella (Suomi-sarakkeet) ---
const stripHtml = (h) => String(h == null ? '' : h).replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ')
  .replace(/&amp;/g, '&').replace(/&#160;/g, ' ').replace(/\s+/g, ' ').trim();
function pick(item, ...titleSubs) {
  if (fieldMap) {
    for (const sub of titleSubs)
      for (const [intl, title] of Object.entries(fieldMap))
        if (String(title).toLowerCase().includes(sub) && item[intl] != null && item[intl] !== '') return item[intl];
  }
  for (const sub of titleSubs) {
    const key = sub.replace(/\s/g, '');
    for (const k of Object.keys(item))
      if (k.toLowerCase().includes(key) && item[k] != null && item[k] !== '') return item[k];
  }
  return '';
}
function mapItem(it) {
  return {
    id: it.Id ?? it.ID ?? '',
    status: it.Status || pick(it, 'tila') || 'Muu',
    title: stripHtml(it.Title || pick(it, 'lyhyt kuvaus', 'otsikko')) || '(nimetön tiketti)',
    desc: stripHtml(pick(it, 'lisätiet', 'kuvaus', 'selite', 'viesti')).slice(0, 500),
    location: stripHtml(pick(it, 'leirialue', 'sijainti', 'paikka')),
    reporter: stripHtml(pick(it, 'ilmoittaj', 'yhteystied')),
    topic: stripHtml(pick(it, 'aihe', 'tapahtuma', 'luokka')),
    safety: /kyll/i.test(String(pick(it, 'turvallisuu'))),
    created: it.Created || null
  };
}
// ryhmittele tiketit statuksen mukaan STATUS_ORDER-järjestyksessä
function groupByStatus(items) {
  const by = new Map();
  for (const t of items) { if (!by.has(t.status)) by.set(t.status, []); by.get(t.status).push(t); }
  const buckets = STATUS_ORDER.map((name) => ({ name, count: (by.get(name) || []).length, tickets: by.get(name) || [] }));
  // mahdolliset tuntemattomat statukset loppuun
  for (const [name, tickets] of by) if (!STATUS_ORDER.includes(name)) buckets.push({ name, count: tickets.length, tickets });
  return buckets;
}

// --- REST-kutsu kirjautuneen kontekstin evästeillä (kaksi tapaa) ---
let spPage = null; // palvelimen oma taustasivu SharePointin originissa (varareitti)
async function ensureSpPage() {
  if (spPage && !spPage.isClosed()) return spPage;
  spPage = await context.newPage();
  await spPage.goto(SITE, { waitUntil: 'domcontentloaded', timeout: 60000 }).catch(() => {});
  return spPage;
}
async function requestJsonViaPage(url) {
  const p = await ensureSpPage();
  if (!p.url().startsWith('https://partio.sharepoint.com')) await p.goto(SITE, { waitUntil: 'domcontentloaded' }).catch(() => {});
  return p.evaluate(async (u) => {
    try {
      const r = await fetch(u, { headers: { Accept: 'application/json;odata=nometadata' }, credentials: 'include' });
      const ct = r.headers.get('content-type') || '';
      if (r.status === 200 && ct.includes('json')) return { ok: true, status: 200, json: await r.json() };
      return { ok: false, status: r.status };
    } catch (e) { return { ok: false, status: 0 }; }
  }, url);
}
async function requestJson(url) {
  // 1) context.request — jakaa evästeet selainkontekstista, ei riipu välilehdistä
  let lastStatus = 0;
  try {
    const r = await context.request.get(url, { headers: { Accept: 'application/json;odata=nometadata' }, timeout: 20000 });
    const ct = String(r.headers()['content-type'] || '');
    if (r.status() === 200 && ct.includes('json')) return { ok: true, status: 200, json: await r.json() };
    lastStatus = r.status();
  } catch (_) { /* jatka varareittiin */ }
  // 2) varareitti: palvelimen oma kirjautunut taustasivu (same-origin fetch)
  try {
    const viaPage = await requestJsonViaPage(url);
    if (viaPage.ok) return viaPage;
    return { ok: false, status: viaPage.status || lastStatus };
  } catch (_) { return { ok: false, status: lastStatus }; }
}

async function poll() {
  if (!context) return;
  try {
    const res = await requestJson(ITEMS_URL);
    if (!res.ok) {
      state.lastHttp = res.status;
      if (res.status === 404) { state.status = 'error'; state.error = 'Listaa ei löytynyt (404). Tarkista lista-osoite.'; }
      else { state.status = 'awaiting-login'; state.error = null; }
      console.log(`[poll] ${res.status === 404 ? 'lista puuttuu' : 'ei kirjautunut'} (HTTP ${res.status}) — kirjaudu selainikkunassa`);
      return;
    }
    if (!fieldMap) {
      const f = await requestJson(FIELDS_URL);
      if (f.ok && Array.isArray(f.json.value)) fieldMap = Object.fromEntries(f.json.value.map((x) => [x.InternalName, x.Title]));
    }
    const all = (res.json.value || []).map(mapItem);
    state.buckets = groupByStatus(all);
    state.uusi = (state.buckets.find((b) => b.name === 'Uusi') || { tickets: [] }).tickets;
    state.count = all.length;
    state.status = 'ok';
    state.updatedAt = new Date().toISOString();
    state.error = null;
    state.lastHttp = 200;
    console.log(`[poll] ${new Date().toLocaleTimeString('fi-FI')} — ${all.length} tikettiä (Uusi: ${state.uusi.length})`);
  } catch (e) {
    state.status = 'awaiting-login';
    state.error = String((e && e.message) || e);
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
app.get('/api/tickets', (_req, res) => res.json({
  status: state.status, updatedAt: state.updatedAt, count: state.count,
  buckets: state.buckets, uusi: state.uusi, tickets: state.uusi, error: state.error
}));
app.get('/api/health', (_req, res) => res.json({ status: state.status, updatedAt: state.updatedAt, count: state.count, uusi: state.uusi.length, lastHttp: state.lastHttp }));
app.use(express.static(path.join(__dirname, 'public')));

(async () => {
  console.log('Käynnistetään tiketti-palvelin…');
  context = await chromium.launchPersistentContext(USER_DATA_DIR, {
    headless: HEADLESS,
    viewport: { width: 1280, height: 900 },
    args: ['--no-first-run']
  });
  // kirjautumissivu (käyttäjää varten) – pollausta ei ajeta tästä
  const page = context.pages()[0] || await context.newPage();
  await page.goto(STATUS_URL, { waitUntil: 'domcontentloaded', timeout: 60000 }).catch(() => {});
  if (!HEADLESS) { try { await page.bringToFront(); } catch (_) {} }
  context.on('close', () => { console.log('Selainkonteksti suljettiin — käynnistä palvelin uudelleen (npm start).'); process.exit(0); });

  app.listen(PORT, () => {
    console.log(`\n  Tiketti-palvelin:  http://localhost:${PORT}`);
    console.log(`  JSON-rajapinta:    http://localhost:${PORT}/api/tickets`);
    console.log(`  → Kirjaudu avautuneessa selainikkunassa partio-tunnuksilla.\n`);
  });

  await poll();
  setInterval(poll, 60000); // 60 s

  const shutdown = async () => { try { await context.close(); } catch (_) {} process.exit(0); };
  process.on('SIGINT', shutdown);
  process.on('SIGTERM', shutdown);
})().catch((e) => { console.error('Palvelin ei käynnistynyt:', e); process.exit(1); });

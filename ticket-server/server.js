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
const zlib = require('zlib');
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

// --- Osallistujaviestintä-lomake (Excel-työkirja toisella sivustolla) ---
// Sama partio-kirjautuminen kattaa koko tenantin, joten samat evästeet toimivat.
// Viitataan tiedostoon sen uniikilla ID:llä (sourcedoc-GUID linkistä).
const FORM_SITE = 'https://partio.sharepoint.com/sites/UudenmaanPiirileiri2026';
const FORM_DOC_ID = process.env.FORM_DOC_ID || '2fdda7e3-7548-48b4-b6fe-2cfbce158960';
const FORM_URL = `${FORM_SITE}/_api/web/GetFileById('${FORM_DOC_ID}')/$value`;
const FORM_COUNT = 3; // montako viimeisintä viestiä näytetään

// --- palvelimen tila (välimuisti dashboardille) ---
const state = { status: 'starting', buckets: [], uusi: [], count: 0, updatedAt: null, error: null, lastHttp: null };
const formState = { status: 'starting', entries: [], updatedAt: null, error: null, lastHttp: null };
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
// SharePoint Onlinessa FedAuth-eväste on SITE COLLECTION -kohtainen: Tiketin-kirjautuminen
// ei yksin riitä toiselle sivustokokoelmalle. Käydään lomakkeen sivustolla kerran, jolloin
// selain suorittaa SSO-kättelyn (rtFa → FedAuth) ja eväste tallentuu kontekstiin.
let formPage = null;
async function ensureFormPage() {
  if (!formPage || formPage.isClosed()) formPage = await context.newPage();
  if (!formPage.url().startsWith(FORM_SITE)) {
    await formPage.goto(FORM_SITE, { waitUntil: 'domcontentloaded', timeout: 60000 }).catch(() => {});
  }
  return formPage;
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

// --- binäärilataus kirjautuneilla evästeillä (Excel-työkirja) ---
// HUOM: kirjautumaton SharePoint palauttaa sign-in-HTML:n HTTP 200:lla, joten pelkkä
// status ei riitä — hylätään text/html-vastaukset (ne eivät ole työkirja).
async function requestBuffer(url) {
  let lastStatus = 0;
  // varmista lomakkeen sivustokokoelman kirjautuminen ennen latausta
  await ensureFormPage().catch(() => {});
  // 1) context.request — jakaa selainkontekstin evästeet
  try {
    const r = await context.request.get(url, { timeout: 30000 });
    const ct = String(r.headers()['content-type'] || '');
    if (r.status() === 200 && !ct.includes('text/html')) return { ok: true, status: 200, buf: await r.body() };
    lastStatus = r.status() === 200 ? 401 : r.status(); // 200+HTML ≈ kirjautumissivu
  } catch (_) { /* jatka varareittiin */ }
  // 2) varareitti: palvelimen kirjautunut taustasivu LOMAKKEEN sivustolla (same-origin fetch → base64)
  try {
    const p = await ensureFormPage();
    const out = await p.evaluate(async (u) => {
      try {
        const r = await fetch(u, { credentials: 'include' });
        const ct = r.headers.get('content-type') || '';
        if (r.status !== 200 || ct.includes('text/html')) return { ok: false, status: r.status === 200 ? 401 : r.status };
        const bytes = new Uint8Array(await r.arrayBuffer());
        let bin = ''; for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
        return { ok: true, status: 200, b64: btoa(bin) };
      } catch (e) { return { ok: false, status: 0 }; }
    }, url);
    if (out && out.ok) return { ok: true, status: 200, buf: Buffer.from(out.b64, 'base64') };
    return { ok: false, status: (out && out.status) || lastStatus };
  } catch (_) { return { ok: false, status: lastStatus }; }
}

// --- pieni xlsx-lukija (zip + XML), ei riippuvuuksia ---
// Purkaa keskushakemiston kautta vain tarvittavat entryt (worksheet + sharedStrings + styles).
function unzip(buf) {
  const files = {};
  let eocd = -1;
  for (let i = buf.length - 22; i >= 0 && i >= buf.length - 22 - 65536; i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) return files;
  const count = buf.readUInt16LE(eocd + 10);
  let off = buf.readUInt32LE(eocd + 16);
  for (let n = 0; n < count && off + 46 <= buf.length; n++) {
    if (buf.readUInt32LE(off) !== 0x02014b50) break;
    const method = buf.readUInt16LE(off + 10);
    const compSize = buf.readUInt32LE(off + 20);
    const nameLen = buf.readUInt16LE(off + 28);
    const extraLen = buf.readUInt16LE(off + 30);
    const commentLen = buf.readUInt16LE(off + 32);
    const localOff = buf.readUInt32LE(off + 42);
    const name = buf.toString('utf8', off + 46, off + 46 + nameLen);
    if (buf.readUInt32LE(localOff) === 0x04034b50) {
      const lNameLen = buf.readUInt16LE(localOff + 26);
      const lExtraLen = buf.readUInt16LE(localOff + 28);
      const dataStart = localOff + 30 + lNameLen + lExtraLen;
      const comp = buf.subarray(dataStart, dataStart + compSize);
      try { files[name] = method === 0 ? comp : zlib.inflateRawSync(comp); } catch (_) { /* ohita */ }
    }
    off += 46 + nameLen + extraLen + commentLen;
  }
  return files;
}
function decodeXml(s) {
  return String(s)
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(+d))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&amp;/g, '&');
}
function colToNum(letters) { let n = 0; for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64); return n - 1; }
// Excel-sarjanumero → luettava pvm (sarjanumero on "seinäkellon" aika ilman aikavyöhykettä → UTC-getterit)
function excelDate(n) {
  const d = new Date(Math.round((n - 25569) * 86400000));
  const p = (x) => String(x).padStart(2, '0');
  const base = `${d.getUTCDate()}.${d.getUTCMonth() + 1}.${d.getUTCFullYear()}`;
  return Math.abs(n - Math.floor(n)) > 1e-6 ? `${base} klo ${p(d.getUTCHours())}.${p(d.getUTCMinutes())}` : base;
}
function parseSharedStrings(buf) {
  if (!buf) return [];
  const s = buf.toString('utf8'); const out = [];
  const re = /<si\b[^>]*>([\s\S]*?)<\/si>/g; let m;
  while ((m = re.exec(s))) {
    let text = ''; const tre = /<t\b[^>]*>([\s\S]*?)<\/t>/g; let tm;
    while ((tm = tre.exec(m[1]))) text += tm[1];
    out.push(decodeXml(text));
  }
  return out;
}
// styles.xml → funktio joka kertoo onko tyyli-indeksi pvm-muotoinen
function parseStyles(buf) {
  if (!buf) return () => false;
  const s = buf.toString('utf8');
  const custom = {};
  const nfRe = /<numFmt\b[^>]*numFmtId="(\d+)"[^>]*formatCode="([^"]*)"/g; let m;
  while ((m = nfRe.exec(s))) custom[+m[1]] = decodeXml(m[2]);
  const styleFmt = [];
  const block = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(s);
  if (block) { const xfRe = /<xf\b([^>]*?)\/?>/g; let x; while ((x = xfRe.exec(block[1]))) styleFmt.push(+((/numFmtId="(\d+)"/.exec(x[1]) || [])[1] || 0)); }
  const builtin = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47]);
  const isDateId = (id) => {
    if (builtin.has(id)) return true;
    const fc = custom[id]; if (!fc) return false;
    const stripped = fc.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '');
    return /[dy]/.test(stripped) || /:mm|:ss|hh?\b/.test(stripped);
  };
  return (idx) => { const id = styleFmt[idx]; return id != null && isDateId(id); };
}
function parseSheet(buf, shared, isDate) {
  const s = buf.toString('utf8'); const rows = [];
  const rowRe = /<row\b[^>]*>([\s\S]*?)<\/row>/g; let rm;
  while ((rm = rowRe.exec(s))) {
    const cells = [];
    const cRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g; let cm;
    while ((cm = cRe.exec(rm[1]))) {
      const attrs = cm[1] || ''; const body = cm[2] || '';
      const rref = (/r="([A-Z]+)\d+"/.exec(attrs) || [])[1];
      const ci = rref ? colToNum(rref) : cells.length;
      const t = (/t="([^"]+)"/.exec(attrs) || [])[1] || 'n';
      const sIdx = +((/s="(\d+)"/.exec(attrs) || [])[1] || 0);
      const vRaw = (/<v>([\s\S]*?)<\/v>/.exec(body) || [])[1];
      let v = '', date = null;
      if (t === 's') { v = vRaw != null ? (shared[+vRaw] || '') : ''; }
      else if (t === 'inlineStr') { let tx = ''; const tre = /<t\b[^>]*>([\s\S]*?)<\/t>/g; let tm; while ((tm = tre.exec(body))) tx += tm[1]; v = decodeXml(tx); }
      else if (t === 'str') { v = vRaw != null ? decodeXml(vRaw) : ''; }
      else if (vRaw != null && vRaw !== '') {
        const num = parseFloat(vRaw);
        if (!isNaN(num) && isDate(sIdx)) { date = num; v = excelDate(num); } else v = vRaw;
      }
      cells[ci] = { v, date };
    }
    rows.push(cells);
  }
  return rows;
}
function roleOf(h) {
  const s = h.toLowerCase();
  if (/aika|pvm|päiv|time|luotu|lähet|klo|completion|start/.test(s)) return 'when';
  if (/nimi|name|yhteys|puhelin|sähköp|email|osasto|leiri|ryhmä|joukkue/.test(s)) return 'who';
  if (/aihe|otsikko|\basia\b|kysymys|subject/.test(s)) return 'subject';
  if (/viesti|kommentti|palaute|kuvaus|vastaus|terveis|message/.test(s)) return 'message';
  return 'extra';
}
// Poimii työkirjasta FORM_COUNT viimeisintä lomakevastausta jäsenneltynä.
function extractForm(files) {
  const shared = parseSharedStrings(files['xl/sharedStrings.xml']);
  const isDate = parseStyles(files['xl/styles.xml']);
  const sheetName = Object.keys(files)
    .filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n))
    .sort((a, b) => (+a.match(/(\d+)/)[1]) - (+b.match(/(\d+)/)[1]))[0];
  if (!sheetName) return [];
  const rows = parseSheet(files[sheetName], shared, isDate);
  if (rows.length < 2) return [];
  const headers = (rows[0] || []).map((c) => (c ? String(c.v).trim() : ''));
  const roles = headers.map(roleOf);
  const entries = [];
  for (let r = 1; r < rows.length; r++) {
    const cells = rows[r] || [];
    if (!cells.some((c) => c && String(c.v).trim() !== '')) continue;
    let id = '', when = '', who = '', subject = '', message = '', whenTs = 0;
    const extras = [];
    for (let ci = 0; ci < headers.length; ci++) {
      const h = headers[ci]; if (!h) continue;
      const cell = cells[ci]; const v = cell ? String(cell.v).trim() : '';
      if (v === '') continue;
      if (/^id$|^vastaustunnus|response id/i.test(h)) { id = v; continue; }
      const role = roles[ci];
      if (role === 'when' && !when) { when = v; whenTs = cell && cell.date != null ? cell.date : (Date.parse(v) || 0); continue; }
      if (role === 'who' && !who) { who = v; continue; }
      if (role === 'subject' && !subject) { subject = v; continue; }
      if (role === 'message' && !message) { message = v; continue; }
      extras.push({ k: h, v: v.length > 160 ? v.slice(0, 160) + '…' : v });
    }
    if (!message) { // pisin extra-kenttä pääviestiksi
      let li = -1, ll = 0;
      extras.forEach((e, i) => { if (e.v.length > ll) { ll = e.v.length; li = i; } });
      if (li >= 0) { message = extras[li].v; extras.splice(li, 1); }
    }
    if (!whenTs) whenTs = r; // varajärjestys: rivijärjestys (Forms lisää uusimman loppuun)
    entries.push({ id, when, who, subject, message, extras: extras.slice(0, 4), whenTs });
  }
  entries.sort((a, b) => b.whenTs - a.whenTs);
  return entries.slice(0, FORM_COUNT).map(({ whenTs, ...e }) => e);
}

async function pollForm() {
  if (!context) return;
  try {
    const res = await requestBuffer(FORM_URL);
    if (!res.ok) {
      formState.lastHttp = res.status;
      if (res.status === 404) { formState.status = 'error'; formState.error = 'Tiedostoa ei löytynyt (404).'; }
      else { formState.status = 'awaiting-login'; formState.error = null; }
      console.log(`[form] ${res.status === 404 ? 'tiedosto puuttuu' : 'ei kirjautunut'} (HTTP ${res.status}) — kirjaudu UudenmaanPiirileiri2026-sivustolle`);
      return;
    }
    // Varmista että lataus on aito xlsx (zip: PK\x03\x04) — muutoin se on esim. kirjautumissivu.
    if (!(res.buf.length >= 4 && res.buf.readUInt32BE(0) === 0x504b0304)) {
      formState.status = 'awaiting-login'; formState.error = null; formState.lastHttp = 401;
      console.log(`[form] vastaus ei ollut xlsx-tiedosto (${res.buf.length} tavua) — todennäk. kirjautumissivu; kirjaudu UudenmaanPiirileiri2026-sivustolle`);
      return;
    }
    const entries = extractForm(unzip(res.buf));
    formState.entries = entries;
    formState.status = 'ok';
    formState.updatedAt = new Date().toISOString();
    formState.error = null;
    formState.lastHttp = 200;
    console.log(`[form] ${new Date().toLocaleTimeString('fi-FI')} — ${entries.length} viimeisintä viestiä`);
  } catch (e) {
    formState.status = 'error';
    formState.error = String((e && e.message) || e);
    console.log('[form] virhe:', formState.error);
  }
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
app.get('/api/form', (_req, res) => res.json({
  status: formState.status, updatedAt: formState.updatedAt, entries: formState.entries, error: formState.error
}));
app.get('/api/health', (_req, res) => res.json({
  status: state.status, updatedAt: state.updatedAt, count: state.count, uusi: state.uusi.length, lastHttp: state.lastHttp,
  form: { status: formState.status, updatedAt: formState.updatedAt, entries: formState.entries.length, lastHttp: formState.lastHttp }
}));
app.use(express.static(path.join(__dirname, 'public')));

const start = async () => {
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
  setInterval(poll, 60000); // 60 s — tiketit
  await pollForm();
  setInterval(pollForm, 60000); // 60 s — osallistujaviestintä päivittyy jatkuvasti (live)

  const shutdown = async () => { try { await context.close(); } catch (_) {} process.exit(0); };
  process.on('SIGINT', shutdown);
  process.on('SIGTERM', shutdown);
};

// Käynnistä vain suorana ajona; require() (testit) saa apufunktiot ilman selainta.
if (require.main === module) start().catch((e) => { console.error('Palvelin ei käynnistynyt:', e); process.exit(1); });
module.exports = { unzip, decodeXml, parseSharedStrings, parseStyles, parseSheet, extractForm, excelDate, colToNum, roleOf };

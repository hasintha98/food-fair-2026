/**
 * Reads the shared planning workbook and turns its five tabs into the single
 * payload the dashboard renders. Owned by the backend: the resulting JSON is
 * cached under backend/data and only ever leaves the server through an
 * authenticated API call.
 */
import fs from 'node:fs'
import path from 'node:path'
import ExcelJS from 'exceljs'
import { config, DATA_DIR, PLAN_FILE } from './config.js'

export const SHEET_ID = config.sheetId
export const SHEET_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=xlsx`
export const SHEET_EDIT_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/edit`

const OUT = config.exportDir
const OUT_BASE = 'Foodfair 2026 - Delivery Plan'
const PLAN_JSON = PLAN_FILE
const CACHE = path.join(DATA_DIR, 'sheet-cache.xlsx')

// ------------------------------------------------------------------ cells
const clean = (s) => String(s == null ? '' : s).replace(/\s+/g, ' ').trim();

function txt(c) {
  const v = c == null ? null : c.value;
  if (v == null) return '';
  if (v instanceof Date) return v.toISOString().slice(0, 16).replace('T', ' ');
  if (typeof v !== 'object') return clean(v);
  if (Array.isArray(v.richText)) return clean(v.richText.map((r) => r.text).join(''));
  if ('result' in v) {
    const r = v.result;
    if (r == null || (typeof r === 'object' && !(r instanceof Date))) return '';
    return txt({ value: r });
  }
  if ('text' in v) return txt({ value: v.text });
  return '';
}

const num = (v) => {
  const n = parseFloat(String(v == null ? '' : v).replace(/[^0-9.\-]/g, ''));
  return Number.isFinite(n) ? n : 0;
};

const phone = (s) => clean(s).replace(/^'+/, '');

function hhmm(c) {
  const v = c == null ? null : c.value;
  if (v instanceof Date) {
    const p = (x) => String(x).padStart(2, '0');
    return `${p(v.getUTCHours())}:${p(v.getUTCMinutes())}`;
  }
  if (typeof v === 'number' && v > 0 && v < 1) {
    const mins = Math.round(v * 24 * 60);
    return `${String(Math.floor(mins / 60)).padStart(2, '0')}:${String(mins % 60).padStart(2, '0')}`;
  }
  return txt(c);
}

const rowCells = (row, n) => Array.from({ length: n }, (_, i) => txt(row.getCell(i + 1)));

// Banner rows are merged across the tab, so every cell holds the same string.
const isBanner = (cells) => {
  const set = new Set(cells.filter(Boolean));
  return set.size === 1 && cells.filter(Boolean).length > 1;
};

// ------------------------------------------------------------------ loading
async function loadWorkbook({ file, signal } = {}) {
  const wb = new ExcelJS.Workbook();
  if (file) {
    await wb.xlsx.readFile(file);
    return { wb, from: file };
  }
  const res = await fetch(SHEET_URL, { redirect: 'follow', signal });
  if (!res.ok) throw new Error(`Sheet download failed (HTTP ${res.status}). Is the link still shared?`);
  const buf = Buffer.from(await res.arrayBuffer());
  if (buf.slice(0, 2).toString() !== 'PK') {
    throw new Error('Google returned a sign-in page, not a workbook. Check the sheet is still shared with "anyone with the link".');
  }
  try { fs.mkdirSync(DATA_DIR, { recursive: true }); fs.writeFileSync(CACHE, buf); } catch { /* cache is a convenience */ }
  await wb.xlsx.load(buf);
  return { wb, from: SHEET_URL, bytes: buf.length };
}

// ------------------------------------------------------------------ extract
function extract(wb) {
  const sheet = (want) => {
    const ws = wb.worksheets.find((w) => w.name.toLowerCase().startsWith(want.toLowerCase()));
    if (!ws) throw new Error(`Tab not found: ${want}. Tabs present: ${wb.worksheets.map((w) => w.name).join(', ')}`);
    return ws;
  };

  // ---- Overview ----
  const ov = sheet('Overview');
  const plan = { facts: {}, decisions: [], assumptions: [] };
  {
    const rows = [];
    ov.eachRow({ includeEmpty: false }, (r) => rows.push(rowCells(r, 6)));
    plan.title = rows.length ? rows[0].find(Boolean) || '' : '';
    let section = 'facts';
    for (const c of rows.slice(1)) {
      if (isBanner(c)) {
        const b = c.find(Boolean) || '';
        if (/timing assumption/i.test(b)) { section = 'assumptions'; continue; }
        if (section === 'assumptions' && b && !plan.assumptions.includes(b)) plan.assumptions.push(b);
        continue;
      }
      const [k, v] = [c[0], c[1]];
      const decision = c[3];
      if (decision && decision !== 'Operational decision' && !plan.decisions.includes(decision)) plan.decisions.push(decision);
      if (section === 'assumptions' && k) { if (!plan.assumptions.includes(k)) plan.assumptions.push(k); continue; }
      if (k && v && k !== 'Plan item') plan.facts[k] = v;
    }
  }

  // ---- Route Summary ----
  const rsWs = sheet('Route Summary');
  const routes = [];
  {
    const rows = [];
    rsWs.eachRow({ includeEmpty: false }, (r, i) => rows.push({ i, c: rowCells(r, 17), row: r }));
    let head = -1;
    for (const { i, c } of rows) if (c[0] === 'Rank' && c[1] === 'Route') { head = i; break; }
    if (head < 0) throw new Error('Could not find the header row on Route Summary');
    for (const { i, c, row } of rows) {
      if (i <= head || !c[1]) continue;
      routes.push({
        id: c[1],
        rank: num(c[0]),
        area: c[2],
        slot: c[3],
        driver: c[4] || c[3],
        orders: num(c[5]),
        uniqueStops: num(c[6]),
        suburbs: c[7] ? c[7].split(',').map((s) => s.trim()).filter(Boolean) : [],
        depotDrive: num(c[8]),
        betweenStops: num(c[9]),
        stopTime: num(c[10]),
        generalBuffer: num(c[11]),
        eventBuffer: num(c[12]),
        totalMin: num(c[13]),
        leaveBy: hhmm(row.getCell(15)),
        finishBy: hhmm(row.getCell(16)),
        note: c[16],
      });
    }
  }
  const routeById = new Map(routes.map((r) => [r.id, r]));

  // ---- Master Data ----
  const md = wb.worksheets.find((w) => /^master/i.test(w.name)) || wb.worksheets[0];
  const master = new Map();
  const voids = [];
  md.eachRow({ includeEmpty: false }, (row, i) => {
    if (i === 1) return;
    const orderId = txt(row.getCell(3));
    if (!orderId) return;
    const customer = txt(row.getCell(4));
    const rec = {
      orderId,
      sheetRow: i,
      orderDate: txt(row.getCell(1)),
      takenBy: txt(row.getCell(2)),
      customer,
      email: txt(row.getCell(5)),
      phone: phone(txt(row.getCell(6))),
      address: txt(row.getCell(7)),
      mode: txt(row.getCell(8)),
      instructions: txt(row.getCell(9)),
      chicken: num(txt(row.getCell(10))),
      veg: num(txt(row.getCell(11))),
      total: num(txt(row.getCell(12))),
      paid: txt(row.getCell(13)),
      banked: num(txt(row.getCell(14))),
      paymentStatus: txt(row.getCell(15)),
      suburb: txt(row.getCell(16)),
      notes: txt(row.getCell(17)),
    };
    if (/^void/i.test(customer)) { voids.push(rec); return; }
    master.set(orderId, rec);
  });

  // ---- Detailed Stops ----
  const dsWs = sheet('Detailed Stops');
  const stops = [];
  {
    const rows = [];
    dsWs.eachRow({ includeEmpty: false }, (r, i) => rows.push({ i, c: rowCells(r, 13), row: r }));
    let head = -1;
    for (const { i, c } of rows) if (c[0] === 'Route' && c[4] === 'Order ID') { head = i; break; }
    if (head < 0) throw new Error('Could not find the header row on Detailed Stops');
    for (const { i, c, row } of rows) {
      if (i <= head) continue;
      const routeId = c[0];
      const orderId = c[4];
      if (!routeId || !orderId) continue;
      const r = routeById.get(routeId);
      const m = master.get(orderId);
      stops.push({
        orderId,
        route: routeId,
        routeRank: r ? r.rank : 0,
        area: r ? r.area : '',
        slot: c[1] || (r ? r.slot : ''),
        driver: r ? r.driver : '',
        leaveBy: hhmm(row.getCell(3)) || (r ? r.leaveBy : ''),
        finishBy: r ? r.finishBy : '',
        routeNote: r ? r.note : '',
        stopNo: num(c[3]),
        customer: c[5] || (m ? m.customer : ''),
        phone: phone(c[6]) || (m ? m.phone : ''),
        address: c[7] || (m ? m.address : ''),
        suburb: c[8] || (m ? m.suburb : ''),
        instructions: c[9] || (m ? m.instructions : ''),
        chicken: num(c[10]),
        veg: num(c[11]),
        packs: num(c[12]),
        email: m ? m.email : '',
        total: m ? m.total : 0,
        paid: m ? m.paid : '',
        paymentStatus: m ? m.paymentStatus : '',
        banked: m ? m.banked : 0,
        takenBy: m ? m.takenBy : '',
        orderDate: m ? m.orderDate : '',
        notes: m ? m.notes : '',
        sheetRow: m ? m.sheetRow : 0,
        mode: 'Delivery',
      });
    }
  }

  const shareCount = new Map();
  for (const s of stops) {
    const k = s.route + '|' + s.stopNo;
    shareCount.set(k, (shareCount.get(k) || 0) + 1);
  }
  for (const s of stops) s.sharedStop = shareCount.get(s.route + '|' + s.stopNo);

  // ---- pickups: in Master, deliberately not routed ----
  const routed = new Set(stops.map((s) => s.orderId));
  const pickups = [...master.values()]
    .filter((m) => !routed.has(m.orderId))
    .map((m) => ({
      ...m,
      route: '', routeRank: 999, area: 'Pickup', slot: '', driver: '',
      leaveBy: '', finishBy: '', routeNote: '', stopNo: 0,
      packs: m.chicken + m.veg, sharedStop: 1,
    }))
    .sort((a, b) => a.orderId.localeCompare(b.orderId, undefined, { numeric: true }));

  // ---- Staffing Checks ----
  const stWs = sheet('Staffing');
  const staffing = { checks: [], sizes: [] };
  stWs.eachRow({ includeEmpty: false }, (r) => {
    const c = rowCells(r, 7);
    if (isBanner(c)) return;
    if (c[0] && c[0] !== 'Check') staffing.checks.push({ check: c[0], result: c[1], action: c[2] });
    if (c[4] && c[4] !== 'Route size') staffing.sizes.push({ size: num(c[4]), routes: num(c[5]), status: c[6] });
  });

  // ---- order the way the day runs ----
  const all = [...stops, ...pickups].sort((a, b) => {
    if (a.routeRank !== b.routeRank) return a.routeRank - b.routeRank;
    if (a.route !== b.route) return a.route.localeCompare(b.route);
    if (a.stopNo !== b.stopNo) return a.stopNo - b.stopNo;
    return a.orderId.localeCompare(b.orderId, undefined, { numeric: true });
  });
  all.forEach((o, i) => { o.seq = i + 1; });
  routes.sort((a, b) => a.rank - b.rank);

  // ---- cross-tab checks ----
  const warnings = [];
  const noMaster = stops.filter((s) => !master.has(s.orderId));
  if (noMaster.length) warnings.push(`${noMaster.length} routed order(s) are not in Master: ${noMaster.map((s) => s.orderId).join(', ')}`);

  const unrouted = [...master.values()].filter((m) => m.mode === 'Delivery' && !routed.has(m.orderId));
  if (unrouted.length) warnings.push(`${unrouted.length} Master delivery(s) have no route: ${unrouted.map((m) => m.orderId).join(', ')}`);

  const notPickup = pickups.filter((p) => p.mode !== 'Pickup');
  if (notPickup.length) warnings.push(`${notPickup.length} unrouted order(s) are not marked Pickup: ${notPickup.map((p) => p.orderId).join(', ')}`);

  for (const r of routes) {
    const got = stops.filter((s) => s.route === r.id).length;
    if (r.orders && got !== r.orders) warnings.push(`${r.id}: Route Summary says ${r.orders} orders, Detailed Stops has ${got}`);
  }
  const orphan = stops.filter((s) => !routeById.has(s.route));
  if (orphan.length) warnings.push(`${orphan.length} stop(s) reference a route missing from Route Summary`);

  const byDriver = {};
  for (const r of routes) (byDriver[r.driver] ||= []).push(r.id);
  for (const [d, ids] of Object.entries(byDriver)) {
    if (ids.length > 1) warnings.push(`driver "${d}" is on ${ids.length} routes: ${ids.join(', ')}`);
  }

  const payload = {
    plan: {
      ...plan,
      depot: plan.facts['Depot'] || '11 Pukeora Road, Otahuhu',
      deadline: plan.facts['Delivery deadline'] || '',
      source: SHEET_EDIT_URL,
      syncedAt: new Date().toISOString(),
    },
    routes,
    orders: all,
    voids: voids.map((v) => ({ orderId: v.orderId, customer: v.customer, notes: v.notes, sheetRow: v.sheetRow })),
    staffing,
  };

  const totals = {
    routes: routes.length,
    areas: new Set(routes.map((r) => r.area)).size,
    drivers: new Set(routes.map((r) => r.driver)).size,
    deliveries: stops.length,
    doors: shareCount.size,
    pickups: pickups.length,
    voids: voids.length,
    packs: stops.reduce((a, b) => a + b.packs, 0),
    value: all.reduce((a, b) => a + b.total, 0),
  };

  return { payload, warnings, totals };
}

// ------------------------------------------------------------------ outputs
function writePlanJson(payload) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
  fs.writeFileSync(PLAN_JSON, JSON.stringify(payload));
  return PLAN_JSON;
}

const CSV_COLS = [
  ['seq', 'Seq'], ['route', 'Route'], ['routeRank', 'Rank'], ['area', 'Area'],
  ['driver', 'Driver'], ['slot', 'Driver slot'], ['leaveBy', 'Leave by'], ['finishBy', 'Finish by'],
  ['stopNo', 'Stop'], ['sharedStop', 'Orders at stop'], ['orderId', 'Order ID'],
  ['customer', 'Customer'], ['phone', 'Phone'], ['email', 'Email'],
  ['address', 'Address'], ['suburb', 'Suburb'], ['mode', 'Type'],
  ['chicken', 'Chicken'], ['veg', 'Veg'], ['packs', 'Packs'],
  ['total', 'Total $'], ['paid', 'Paid'], ['paymentStatus', 'Payment'], ['banked', 'Banked $'],
  ['takenBy', 'Order taken by'], ['instructions', 'Instructions'], ['notes', 'Notes'],
  ['routeNote', 'Route note'], ['orderDate', 'Ordered'], ['sheetRow', 'Master row'],
];

function writeCsv(payload) {
  const esc = (v) => {
    const s = v == null ? '' : String(v);
    return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  const csv = [CSV_COLS.map((c) => esc(c[1])).join(',')]
    .concat(payload.orders.map((o) => CSV_COLS.map((c) => esc(o[c[0]])).join(',')))
    .join('\r\n');
  const p = path.join(OUT, OUT_BASE + '.csv');
  fs.writeFileSync(p, '﻿' + csv, 'utf8');
  return p;
}

async function writeXlsx(payload) {
  const { routes, orders } = payload;
  const stops = orders.filter((o) => o.mode === 'Delivery');
  const pickups = orders.filter((o) => o.mode !== 'Delivery');

  const out = new ExcelJS.Workbook();
  out.creator = 'Food Fair 2026 delivery plan';
  out.created = new Date();
  const HEAD = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
  const FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E5F' } };

  const dress = (ws, cols) => {
    ws.columns = cols.map((c) => ({ header: c[1], key: c[0], width: c[2] }));
    const h = ws.getRow(1);
    h.font = HEAD; h.fill = FILL; h.height = 26;
    h.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    ws.views = [{ state: 'frozen', ySplit: 1 }];
    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: cols.length } };
  };
  const paint = (ws, list, cols) => {
    const has = new Set(cols.map((c) => c[0]));
    for (const o of list) {
      const row = ws.addRow(Object.fromEntries(cols.map((c) => [c[0], o[c[0]]])));
      if (has.has('paymentStatus')) {
        if (o.paymentStatus === 'Outstanding') row.getCell('paymentStatus').font = { color: { argb: 'FFC00000' } };
        if (o.paymentStatus === 'Paid') row.getCell('paymentStatus').font = { color: { argb: 'FF1E7B34' } };
      }
      for (const k of ['total', 'banked']) if (has.has(k)) row.getCell(k).numFmt = '#,##0.00';
    }
    ws.eachRow({ includeEmpty: false }, (r, i) => {
      if (i > 1 && i % 2 === 0) r.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF4F7F9' } };
    });
  };

  const sum = out.addWorksheet('Routes', { properties: { tabColor: { argb: 'FF1F4E5F' } } });
  dress(sum, [
    ['rank', 'Rank', 7], ['id', 'Route', 9], ['area', 'Area', 22], ['slot', 'Slot', 10],
    ['driver', 'Driver', 16], ['orders', 'Orders', 9], ['uniqueStops', 'Stops', 8],
    ['leaveBy', 'Leave by', 10], ['finishBy', 'Finish by', 10], ['totalMin', 'Total min', 10],
    ['packs', 'Packs', 8], ['value', 'Value $', 11], ['owing', 'Owing $', 11],
    ['suburbList', 'Suburbs', 52], ['note', 'Traffic / sequencing note', 54],
  ]);
  for (const r of routes) {
    const mine = stops.filter((s) => s.route === r.id);
    sum.addRow({
      ...r,
      packs: mine.reduce((a, b) => a + b.packs, 0),
      value: mine.reduce((a, b) => a + b.total, 0),
      owing: mine.filter((s) => s.paymentStatus === 'Outstanding').reduce((a, b) => a + b.total, 0),
      suburbList: r.suburbs.join(', '),
    });
  }
  sum.eachRow((r, i) => { if (i > 1) ['value', 'owing'].forEach((k) => (r.getCell(k).numFmt = '#,##0.00')); });

  const ALL = CSV_COLS.map(([k, l]) => [k, l, Math.min(46, Math.max(9, l.length + 4))]);
  const allWs = out.addWorksheet('All stops', { properties: { tabColor: { argb: 'FF2E7D32' } } });
  dress(allWs, ALL);
  paint(allWs, orders, ALL);

  const RUN = [
    ['stopNo', 'Stop', 7], ['orderId', 'Order ID', 11], ['customer', 'Customer', 26],
    ['phone', 'Phone', 15], ['address', 'Address', 48], ['suburb', 'Suburb', 18],
    ['chicken', 'Chicken', 9], ['veg', 'Veg', 7], ['packs', 'Packs', 8],
    ['total', 'Total $', 10], ['paymentStatus', 'Payment', 13], ['instructions', 'Instructions', 34],
  ];
  for (const r of routes) {
    const ws = out.addWorksheet((r.id + ' ' + r.driver).replace(/[\\/?*[\]:]/g, '-').slice(0, 31));
    dress(ws, RUN);
    const mine = stops.filter((s) => s.route === r.id).sort((a, b) => a.stopNo - b.stopNo);
    paint(ws, mine, RUN);
    const t = ws.addRow({
      customer: `${r.area} · leave ${r.leaveBy} · finish by ${r.finishBy}`,
      chicken: mine.reduce((a, b) => a + b.chicken, 0),
      veg: mine.reduce((a, b) => a + b.veg, 0),
      packs: mine.reduce((a, b) => a + b.packs, 0),
      total: mine.reduce((a, b) => a + b.total, 0),
    });
    t.font = { bold: true };
    t.border = { top: { style: 'double' } };
  }

  const pk = out.addWorksheet('Pickup');
  dress(pk, RUN);
  paint(pk, pickups, RUN);

  const p = path.join(OUT, OUT_BASE + '.xlsx');
  await out.xlsx.writeFile(p);
  return p;
}

/**
 * Full refresh: fetch, extract, write plan.json (always) and the CSV/XLSX
 * (best effort — Excel locks an open workbook on Windows, which must not fail
 * a refresh triggered from the browser).
 */
async function refresh({ file, signal, writeFiles = true } = {}) {
  const { wb, from, bytes } = await loadWorkbook({ file, signal });
  const { payload, warnings, totals } = extract(wb);
  const written = [];
  const softErrors = [];

  written.push(writePlanJson(payload));
  if (writeFiles && OUT) {
    try { written.push(writeCsv(payload)); }
    catch (e) { softErrors.push('CSV not written: ' + e.message); }
    try { written.push(await writeXlsx(payload)); }
    catch (e) {
      softErrors.push(e.code === 'EBUSY'
        ? 'XLSX not written — it is open in Excel. Close it and refresh again.'
        : 'XLSX not written: ' + e.message);
    }
  }
  return { payload, warnings, totals, written, softErrors, from, bytes };
}

export { loadWorkbook, extract, refresh, writePlanJson, writeCsv, writeXlsx, readPlan, PLAN_JSON }

/** The cached plan on disk, or null if nothing has been synced yet. */
function readPlan() {
  try { return JSON.parse(fs.readFileSync(PLAN_JSON, 'utf8')) } catch { return null }
}

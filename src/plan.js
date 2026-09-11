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

const deMacron = (s) => String(s).normalize('NFD').replace(/[\u0300-\u036f]/g, '');
const same = (a, b) => deMacron(a).toLowerCase().replace(/[^a-z0-9]/g, '') === deMacron(b).toLowerCase().replace(/[^a-z0-9]/g, '');

/**
 * A map-search string that always lands in Auckland. Raw addresses from the
 * sheet are sometimes missing the city or country, or carry stray commas and
 * spaces ("2/41,, Frederic Street", "5/ 32, Beulah Avenue"), and Google will
 * happily guess a street of the same name in another country.
 */
function mapQuery(address, suburb) {
  let parts = clean(address)
    .replace(/(\d)\s*\/\s*(\d)/g, '$1/$2')      // "5/ 32" -> "5/32"
    .replace(/^(\S+)\.\s+/, '$1 ')              // "18. Basra Drive" -> "18 Basra Drive"
    .replace(/\s*,\s*/g, ',')                   // tidy around commas
    .split(',')
    .map((p) => p.replace(/[.\s]+$/g, '').trim())
    .filter(Boolean);

  // drop consecutive repeats ("Auckland, Auckland") and generic/country tails; we re-add those
  parts = parts.filter((p, i) => i === 0 || !same(p, parts[i - 1]));
  parts = parts.filter((p) => !/^(new zealand|nz)$/i.test(p));
  const city = /^auckland(\s+\d{4})?$/i;
  const postcode =
    parts.find((p) => city.test(p) && /\d{4}/.test(p))?.match(/\d{4}/)?.[0] ||
    parts.find((p, i) => i > 0 && /^\d{4}$/.test(p));                        // bare "1023" part
  parts = parts.filter((p, i) => !city.test(p) && !(i > 0 && /^\d{4}$/.test(p)));

  // "10,Subritzky Avenue" -> "10 Subritzky Avenue": a bare number part belongs to the next part
  for (let i = 0; i < parts.length - 1; i++) {
    if (/^[\w/-]+$/.test(parts[i]) && /\d/.test(parts[i]) && !/\d/.test(parts[i + 1][0] || '')) {
      parts.splice(i, 2, parts[i] + ' ' + parts[i + 1]);
    }
  }

  const sub = clean(suburb);
  const subIsReal = sub && !/^auckland$/i.test(sub) && !/avenue|street|road|drive|place|lane$/i.test(sub);
  if (subIsReal && !parts.some((p) => same(p, sub) || deMacron(p).toLowerCase().includes(deMacron(sub).toLowerCase()))) {
    parts.push(sub);
  }
  parts.push(postcode ? `Auckland ${postcode}` : 'Auckland', 'New Zealand');
  return parts.join(', ');
}

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
// "7.15" in the Area List means 07:15 (hours.minutes), not 7.15 hours.
function hDotMm(v) {
  const s = clean(v);
  if (!s) return '';
  const m = s.match(/^(\d{1,2})(?:\.(\d{1,2}))?$/);
  if (!m) return s;
  const h = Number(m[1]);
  const mm = m[2] ? Number(m[2].length === 1 ? m[2] + '0' : m[2]) : 0;
  if (h > 23 || mm > 59) return s;
  return `${String(h).padStart(2, '0')}:${String(mm).padStart(2, '0')}`;
}

const UNASSIGNED = /^(driver\s+)?not\s+assigned|^tba$|^tbc$|^unassigned$/i;

/**
 * Customers type everything into one "Instructions" box. Notes about the food
 * itself belong to the packers, not the driver, so they are split out at sync.
 * A note that talks about both (rare) is kept for both.
 */
const PACKING_WORDS = /\b(eggs?|soya?|sambol|mild|hot|spicy|chill?i|vegan|vegetarian|halal|gluten|dairy|nuts?|allerg\w*|sauce|gravy|rice|curry|no onion|onions?|less oil|extra)\b/i;
const DELIVERY_WORDS = /\b(call|ring|text|notify|knock|door|gate|leave|arriv\w*|deliver\w*|drive\w*|house|unit|flat|building|park\w*|contact|number|mobile|pay|paid|cash|behind|near|front|floor|lane|road|street|home|address)\b/i;

function splitInstructions(raw, packingRaw) {
  const text = clean(raw);
  let delivery = text;
  let packing = clean(packingRaw);
  let moved = false;
  if (text && PACKING_WORDS.test(text)) {
    const alsoDelivery = DELIVERY_WORDS.test(text);
    if (!packing) packing = text;
    else if (!packing.toLowerCase().includes(text.toLowerCase())) packing = packing + ' · ' + text;
    if (!alsoDelivery) { delivery = ''; moved = true; }
  }
  return { delivery, packing, moved };
}

// ------------------------------------------------------------------ extract
function extract(wb) {
  const sheet = (want, { optional = false } = {}) => {
    const ws = wb.worksheets.find((w) => w.name.toLowerCase().startsWith(want.toLowerCase()));
    if (!ws && !optional) {
      throw new Error(`Tab not found: ${want}. Tabs present: ${wb.worksheets.map((w) => w.name).join(', ')}`);
    }
    return ws || null;
  };
  const findHeader = (ws, width, test) => {
    let head = -1;
    const rows = [];
    ws.eachRow({ includeEmpty: false }, (r, i) => rows.push({ i, c: rowCells(r, width), row: r }));
    for (const { i, c } of rows) if (test(c)) { head = i; break; }
    return { head, rows };
  };

  // ---- Overview ----
  const plan = { facts: {}, decisions: [], assumptions: [] };
  {
    const rows = [];
    sheet('Overview').eachRow({ includeEmpty: false }, (r) => rows.push(rowCells(r, 6)));
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

  // ---- Area List: packing priority and estimated dispatch start per area ----
  const areas = [];
  {
    const ws = sheet('Area List', { optional: true });
    if (ws) {
      const { head, rows } = findHeader(ws, 9, (c) => c[0] === 'Area' && /priority/i.test(c[1]));
      if (head > 0) {
        for (const { i, c } of rows) {
          if (i <= head || !c[0] || !c[1]) continue;
          areas.push({
            name: c[0],
            priority: num(c[1]),
            label: c[2] || `(${String(num(c[1])).padStart(2, '0')}) ${c[0]}`,
            startTime: hDotMm(c[4]),
            cumulativePacks: num(c[5]),
            packs: num(c[6]),
          });
        }
      }
    }
    areas.sort((a, b) => a.priority - b.priority);
  }
  const areaByName = new Map(areas.map((a) => [a.name, a]));

  // ---- Route Summary ----
  const routes = [];
  {
    const ws = sheet('Route Summary');
    const { head, rows } = findHeader(ws, 17, (c) => c[0] === 'Rank' && c[1] === 'Route');
    if (head < 0) throw new Error('Could not find the header row on Route Summary');
    for (const { i, c, row } of rows) {
      if (i <= head || !c[1]) continue;
      const driverRaw = c[3];
      const unassigned = !driverRaw || UNASSIGNED.test(driverRaw);
      routes.push({
        id: c[1],
        rank: num(c[0]),
        area: c[2],
        areaPriority: areaByName.get(c[2])?.priority ?? 99,
        driver: unassigned ? '' : driverRaw,
        driverPhone: phone(c[4]),
        unassigned,
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

  // ---- Stops: the "(Master)" tab is one row per delivery with driver phone and area priority ----
  const stops = [];
  const movedNotes = [];   // instructions reclassified from delivery to packing
  {
    const ws = sheet('Filter by Area (Master)', { optional: true }) || sheet('Detailed Stops');
    const width = Math.max(20, ws.getRow(1).cellCount, ws.getRow(3).cellCount);
    const { head, rows } = findHeader(ws, width, (c) => c[0] === 'Route' && c.includes('Order ID'));
    if (head < 0) throw new Error(`Could not find the header row on ${ws.name}`);
    const H = rows.find((r) => r.i === head).c;
    const col = (name, re) => {
      const i = H.findIndex((h) => h.toLowerCase() === name.toLowerCase());
      return i >= 0 || !re ? i : H.findIndex((h) => re.test(h));
    };
    const ix = {
      route: col('Route'), area: col('Area'), driver: col('Driver'), driverPhone: col('Driver phone'),
      stop: col('Stop'), orderId: col('Order ID'), customer: col('Customer'),
      phone: col('Customer phone') >= 0 ? col('Customer phone') : col('Phone'),
      address: col('Address'), suburb: col('Suburb'),
      // "Delivery Instructions" on the Master tab, plain "Instructions" on the older views
      instructions: col('Delivery Instructions', /^(delivery\s+)?instructions?$/i),
      chicken: col('Chicken'), veg: col('Vegetable'), packs: col('Total packs'),
      areaPriority: col('Area by Priority'), leaveBy: col('Leave by'),
      packing: col('Packing Instructions', /^packing\s+instr/i),   // the sheet spells it "Instrustions"
    };
    const get = (c, k) => (ix[k] >= 0 ? c[ix[k]] : '');
    const required = ['route', 'stop', 'orderId', 'customer', 'address', 'suburb', 'chicken', 'veg', 'instructions'];
    const missing = required.filter((k) => ix[k] < 0);
    if (missing.length) {
      throw new Error(`${ws.name}: cannot find column(s) ${missing.join(', ')}. Headers present: ${H.filter(Boolean).join(' | ')}`);
    }

    for (const { i, c, row } of rows) {
      if (i <= head) continue;
      const routeId = get(c, 'route');
      const orderId = get(c, 'orderId');
      if (!routeId || !orderId) continue;
      const r = routeById.get(routeId);
      const areaName = get(c, 'area') || (r ? r.area : '');
      const prioLabel = get(c, 'areaPriority');
      const prio = prioLabel.match(/^\((\d+)\)/)?.[1];
      const driverRaw = get(c, 'driver');
      const driver = r ? r.driver : (UNASSIGNED.test(driverRaw) ? '' : driverRaw);
      const notes = splitInstructions(get(c, 'instructions'), get(c, 'packing'));
      if (notes.moved) movedNotes.push({ orderId, text: get(c, 'instructions') });
      stops.push({
        orderId,
        route: routeId,
        routeRank: r ? r.rank : 0,
        area: areaName,
        areaPriority: prio ? Number(prio) : (areaByName.get(areaName)?.priority ?? 99),
        driver,
        driverPhone: (r && r.driverPhone) || phone(get(c, 'driverPhone')),
        leaveBy: (ix.leaveBy >= 0 ? hhmm(row.getCell(ix.leaveBy + 1)) : '') || (r ? r.leaveBy : ''),
        finishBy: r ? r.finishBy : '',
        routeNote: r ? r.note : '',
        stopNo: num(get(c, 'stop')),
        customer: get(c, 'customer'),
        phone: phone(get(c, 'phone')),
        address: get(c, 'address'),
        mapQuery: mapQuery(get(c, 'address'), get(c, 'suburb')),
        suburb: get(c, 'suburb'),
        instructions: notes.delivery,
        packingInstructions: notes.packing,
        chicken: num(get(c, 'chicken')),
        veg: num(get(c, 'veg')),
        packs: num(get(c, 'packs')) || num(get(c, 'chicken')) + num(get(c, 'veg')),
        mode: 'Delivery',
      });
    }
  }

  // Orders that share one physical door (same route + stop number).
  const shareCount = new Map();
  for (const s of stops) {
    const k = s.route + '|' + s.stopNo;
    shareCount.set(k, (shareCount.get(k) || 0) + 1);
  }
  for (const s of stops) s.sharedStop = shareCount.get(s.route + '|' + s.stopNo);

  // ---- Staffing Checks ----
  const staffing = { checks: [], sizes: [] };
  sheet('Staffing').eachRow({ includeEmpty: false }, (r) => {
    const c = rowCells(r, 7);
    if (isBanner(c)) return;
    if (c[0] && c[0] !== 'Check') staffing.checks.push({ check: c[0], result: c[1], action: c[2] });
    if (c[4] && c[4] !== 'Route size') staffing.sizes.push({ size: num(c[4]), routes: num(c[5]), status: c[6] });
  });

  // ---- things a coordinator has to deal with before dispatch ----
  const attention = [];
  for (const r of routes) {
    if (r.unassigned) attention.push({ kind: 'driver', route: r.id, text: `${r.id} ${r.area} (leave ${r.leaveBy}) has no driver assigned.` });
    else if (!r.driverPhone) attention.push({ kind: 'phone', route: r.id, text: `${r.id} ${r.driver} has no phone number.` });
  }
  for (const c of staffing.checks) {
    if (c.action && !/^none$/i.test(c.action) && !/duplicate-address|hibiscus/i.test(c.action) && !/exception/i.test(c.check)) {
      // skip a check that just restates a route we have already flagged (e.g. "Assign a driver to Route R09")
      const dup = attention.some((a) => a.route && c.action.includes(a.route));
      if (!dup) attention.push({ kind: 'check', route: '', text: `${c.check}: ${c.action}` });
    }
  }

  // ---- order the way the day runs ----
  stops.sort((a, b) =>
    (a.routeRank - b.routeRank) || a.route.localeCompare(b.route) || (a.stopNo - b.stopNo) ||
    a.orderId.localeCompare(b.orderId, undefined, { numeric: true }));
  stops.forEach((o, i) => { o.seq = i + 1; });
  routes.sort((a, b) => a.rank - b.rank);

  // ---- cross-tab checks ----
  const warnings = [];
  for (const r of routes) {
    const got = stops.filter((s) => s.route === r.id).length;
    if (r.orders && got !== r.orders) warnings.push(`${r.id}: Route Summary says ${r.orders} orders, stop list has ${got}`);
  }
  const orphan = stops.filter((s) => !routeById.has(s.route));
  if (orphan.length) warnings.push(`${orphan.length} stop(s) reference a route missing from Route Summary: ${[...new Set(orphan.map((s) => s.route))].join(', ')}`);
  const ds = sheet('Detailed Stops', { optional: true });
  if (ds) {
    const ids = new Set();
    ds.eachRow({ includeEmpty: false }, (r) => { const c = rowCells(r, 6); if (/^WEB-/i.test(c[4])) ids.add(c[4]); });
    const mine = new Set(stops.map((s) => s.orderId));
    const onlyDs = [...ids].filter((x) => !mine.has(x));
    const onlyMaster = [...mine].filter((x) => !ids.has(x));
    if (onlyDs.length) warnings.push(`${onlyDs.length} order(s) in Detailed Stops but not in the Master tab: ${onlyDs.slice(0, 10).join(', ')}`);
    if (onlyMaster.length) warnings.push(`${onlyMaster.length} order(s) in the Master tab but not in Detailed Stops: ${onlyMaster.slice(0, 10).join(', ')}`);
  }
  const claimed = num(plan.facts['Delivery orders']);
  if (claimed && claimed !== stops.length) warnings.push(`Overview says ${claimed} delivery orders, found ${stops.length}`);
  const claimedDoors = num(plan.facts['Unique delivery stops']);
  if (claimedDoors && claimedDoors !== shareCount.size) warnings.push(`Overview says ${claimedDoors} unique stops, found ${shareCount.size}`);
  for (const a of attention) warnings.push(a.text);

  const payload = {
    plan: {
      ...plan,
      depot: plan.facts['Depot'] || '11 Pukeora Road, Otahuhu',
      deadline: plan.facts['Delivery deadline'] || '',
      source: SHEET_EDIT_URL,
      syncedAt: new Date().toISOString(),
    },
    routes,
    areas,
    orders: stops,
    staffing,
    attention,
  };

  const totals = {
    routes: routes.length,
    areas: new Set(routes.map((r) => r.area)).size,
    drivers: new Set(routes.filter((r) => r.driver).map((r) => r.driver)).size,
    unassigned: routes.filter((r) => r.unassigned).length,
    deliveries: stops.length,
    doors: shareCount.size,
    packs: stops.reduce((a, b) => a + b.packs, 0),
    packingNotes: stops.filter((s) => s.packingInstructions).length,
  };

  return { payload, warnings, totals, movedNotes };
}

// ------------------------------------------------------------------ outputs
function writePlanJson(payload) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
  fs.writeFileSync(PLAN_JSON, JSON.stringify(payload));
  return PLAN_JSON;
}

const CSV_COLS = [
  ['seq', 'Seq'], ['route', 'Route'], ['routeRank', 'Rank'], ['areaPriority', 'Area priority'], ['area', 'Area'],
  ['driver', 'Driver'], ['driverPhone', 'Driver phone'], ['leaveBy', 'Leave by'], ['finishBy', 'Finish by'],
  ['stopNo', 'Stop'], ['sharedStop', 'Orders at stop'], ['orderId', 'Order ID'],
  ['customer', 'Customer'], ['phone', 'Customer phone'],
  ['address', 'Address'], ['suburb', 'Suburb'],
  ['chicken', 'Chicken'], ['veg', 'Veg'], ['packs', 'Packs'],
  ['instructions', 'Delivery instructions'], ['packingInstructions', 'Packing instructions'],
  ['routeNote', 'Route note'], ['mapQuery', 'Map search'],
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
  const { routes, orders, areas } = payload;

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
    for (const o of list) ws.addRow(Object.fromEntries(cols.map((c) => [c[0], o[c[0]]])));
    ws.eachRow({ includeEmpty: false }, (r, i) => {
      if (i > 1 && i % 2 === 0) r.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF4F7F9' } };
    });
  };

  const sum = out.addWorksheet('Routes', { properties: { tabColor: { argb: 'FF1F4E5F' } } });
  dress(sum, [
    ['rank', 'Rank', 7], ['id', 'Route', 9], ['areaPriority', 'Area #', 8], ['area', 'Area', 22],
    ['driver', 'Driver', 18], ['driverPhone', 'Driver phone', 14],
    ['orders', 'Orders', 9], ['uniqueStops', 'Stops', 8],
    ['leaveBy', 'Leave by', 10], ['finishBy', 'Finish by', 10], ['totalMin', 'Total min', 10],
    ['packs', 'Packs', 8], ['suburbList', 'Suburbs', 52], ['note', 'Traffic / sequencing note', 54],
  ]);
  for (const r of routes) {
    const mine = orders.filter((s) => s.route === r.id);
    sum.addRow({ ...r, driver: r.driver || 'DRIVER NOT ASSIGNED', packs: mine.reduce((a, b) => a + b.packs, 0), suburbList: r.suburbs.join(', ') });
    if (r.unassigned) sum.lastRow.getCell('driver').font = { color: { argb: 'FFC00000' }, bold: true };
  }

  if (areas.length) {
    const aw = out.addWorksheet('Area schedule');
    dress(aw, [['priority', 'Priority', 9], ['name', 'Area', 26], ['startTime', 'Est. start', 11], ['packs', 'Packs', 8], ['cumulativePacks', 'Cumulative', 11]]);
    paint(aw, areas, [['priority'], ['name'], ['startTime'], ['packs'], ['cumulativePacks']]);
  }

  const packing = orders.filter((o) => o.packingInstructions);
  if (packing.length) {
    const pw = out.addWorksheet('Packing notes', { properties: { tabColor: { argb: 'FFB26A00' } } });
    dress(pw, [['areaPriority', 'Area #', 8], ['area', 'Area', 22], ['route', 'Route', 8], ['orderId', 'Order ID', 11],
      ['customer', 'Customer', 26], ['chicken', 'Chicken', 9], ['veg', 'Veg', 7], ['packingInstructions', 'Packing instructions', 48]]);
    paint(pw, packing.slice().sort((a, b) => a.areaPriority - b.areaPriority || a.seq - b.seq),
      [['areaPriority'], ['area'], ['route'], ['orderId'], ['customer'], ['chicken'], ['veg'], ['packingInstructions']]);
  }

  const ALL = CSV_COLS.map(([k, l]) => [k, l, Math.min(46, Math.max(9, l.length + 4))]);
  const allWs = out.addWorksheet('All stops', { properties: { tabColor: { argb: 'FF2E7D32' } } });
  dress(allWs, ALL);
  paint(allWs, orders, ALL);

  const RUN = [
    ['stopNo', 'Stop', 7], ['orderId', 'Order ID', 11], ['customer', 'Customer', 26],
    ['phone', 'Phone', 15], ['address', 'Address', 48], ['suburb', 'Suburb', 18],
    ['chicken', 'Chicken', 9], ['veg', 'Veg', 7], ['packs', 'Packs', 8],
    ['instructions', 'Delivery instructions', 34], ['packingInstructions', 'Packing', 24],
  ];
  for (const r of routes) {
    const ws = out.addWorksheet((r.id + ' ' + (r.driver || 'UNASSIGNED')).replace(/[\\/?*[\]:]/g, '-').slice(0, 31));
    dress(ws, RUN);
    const mine = orders.filter((s) => s.route === r.id).sort((a, b) => a.stopNo - b.stopNo);
    paint(ws, mine, RUN);
    const t = ws.addRow({
      customer: `${r.area} · ${r.driver || 'no driver'}${r.driverPhone ? ' · ' + r.driverPhone : ''} · leave ${r.leaveBy} · finish by ${r.finishBy}`,
      chicken: mine.reduce((a, b) => a + b.chicken, 0),
      veg: mine.reduce((a, b) => a + b.veg, 0),
      packs: mine.reduce((a, b) => a + b.packs, 0),
    });
    t.font = { bold: true };
    t.border = { top: { style: 'double' } };
  }

  const p = path.join(OUT, OUT_BASE + '.xlsx');
  await out.xlsx.writeFile(p);
  return p;
}

async function refresh({ file, signal, writeFiles = true } = {}) {
  const { wb, from, bytes } = await loadWorkbook({ file, signal });
  const { payload, warnings, totals, movedNotes } = extract(wb);
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
  return { payload, warnings, totals, movedNotes, written, softErrors, from, bytes };
}

export { loadWorkbook, extract, refresh, writePlanJson, writeCsv, writeXlsx, readPlan, PLAN_JSON }

/** The cached plan on disk, or null if nothing has been synced yet. */
function readPlan() {
  try { return JSON.parse(fs.readFileSync(PLAN_JSON, 'utf8')) } catch { return null }
}

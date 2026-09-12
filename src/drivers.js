/**
 * The driver credentials table — a "Drivers" tab in a separate workbook:
 *
 *   Driver | Route | Area | Phone Number | (sent?) | Email | Password | (note)
 *
 * Read by header name so column order does not matter. Only ever served
 * behind the dashboard password; printed on the Drivers sheet for the driver
 * manager, who hands each driver their login in person.
 */
import fs from 'node:fs'
import path from 'node:path'
import ExcelJS from 'exceljs'
import { config, DATA_DIR } from './config.js'

const clean = (s) => String(s == null ? '' : s).replace(/\s+/g, ' ').trim()
function txt(c) {
  const v = c == null ? null : c.value
  if (v == null) return ''
  if (typeof v !== 'object') return clean(v)
  if (Array.isArray(v.richText)) return clean(v.richText.map((r) => r.text).join(''))
  if ('result' in v) { const r = v.result; return r == null || (typeof r === 'object' && !(r instanceof Date)) ? '' : txt({ value: r }) }
  if ('text' in v) return txt({ value: v.text })
  if (v.hyperlink) return clean(v.hyperlink.replace(/^mailto:/i, ''))
  return ''
}
const phone = (s) => clean(s).replace(/^'+/, '')

export const DRIVERS_SHEET_URL = config.driversSheetId
  ? `https://docs.google.com/spreadsheets/d/${config.driversSheetId}/export?format=xlsx`
  : ''

/** A workbook in data/ that has a Drivers tab — so dropping the file there is enough. */
async function findLocalDriversFile() {
  if (config.driversFile) return path.resolve(DATA_DIR, '..', config.driversFile)
  let names = []
  try { names = fs.readdirSync(DATA_DIR).filter((n) => /\.xlsx$/i.test(n) && !/^sheet-cache/i.test(n)) } catch { return '' }
  for (const n of names) {
    const wb = new ExcelJS.Workbook()
    try { await wb.xlsx.readFile(path.join(DATA_DIR, n)) } catch { continue }
    if (wb.worksheets.some((w) => /drivers?/i.test(w.name))) return path.join(DATA_DIR, n)
  }
  return ''
}

const JSON_FILE = path.join(DATA_DIR, '..', 'drivers.json')

/** Committed drivers.json, or the DRIVERS_JSON variable (base64 or plain JSON) — for hosts with no persistent disk. */
function loadDriversJson() {
  let raw = process.env.DRIVERS_JSON || ''
  let source = 'DRIVERS_JSON variable'
  if (!raw && fs.existsSync(JSON_FILE)) { raw = fs.readFileSync(JSON_FILE, 'utf8'); source = JSON_FILE }
  if (!raw) return null
  const text = raw.trim().startsWith('{') ? raw : Buffer.from(raw, 'base64').toString('utf8')
  const parsed = JSON.parse(text)
  const drivers = Array.isArray(parsed) ? parsed : parsed.drivers || []
  const noCreds = drivers.filter((d) => !d.email || !d.password).length
  const warnings = noCreds ? [`${noCreds} of ${drivers.length} drivers have no email/password (from ${path.basename(source)}).`] : []
  return { drivers, source, warnings }
}

export async function loadDrivers({ file, signal } = {}) {
  if (!file && !config.driversSheetId) file = await findLocalDriversFile()
  if (!file && !config.driversSheetId) {
    // nothing live to read: fall back to the exported JSON so Railway still has the table
    try { const j = loadDriversJson(); if (j) return j } catch (e) { return { drivers: [], source: null, warnings: ['drivers.json could not be read: ' + e.message] } }
    return { drivers: [], source: null, warnings: [] }
  }
  const wb = new ExcelJS.Workbook()
  if (file) await wb.xlsx.readFile(file)
  else {
    const res = await fetch(DRIVERS_SHEET_URL, { redirect: 'follow', signal })
    if (!res.ok) throw new Error(`Drivers sheet download failed (HTTP ${res.status}). Is it shared with "anyone with the link"?`)
    const buf = Buffer.from(await res.arrayBuffer())
    if (buf.slice(0, 2).toString() !== 'PK') throw new Error('Drivers sheet: Google returned a sign-in page, not a workbook.')
    await wb.xlsx.load(buf)
  }
  return { ...extractDrivers(wb), source: file || DRIVERS_SHEET_URL }
}

export function extractDrivers(wb) {
  const want = (config.driversTab || 'Drivers').toLowerCase()
  const ws = wb.worksheets.find((w) => w.name.toLowerCase() === want)
    || wb.worksheets.find((w) => w.name.toLowerCase().includes(want))
    || wb.worksheets.find((w) => { let hit = false; w.eachRow({ includeEmpty: false }, (r) => { if (!hit && r.values.some((v) => /^password$/i.test(txt({ value: v })))) hit = true }); return hit })
  if (!ws) throw new Error(`No "${config.driversTab || 'Drivers'}" tab (or any tab with a Password column) in the drivers workbook. Tabs: ${wb.worksheets.map((w) => w.name).join(', ')}`)

  // header row = first row containing "Driver" and "Password"
  let head = -1, H = []
  ws.eachRow({ includeEmpty: false }, (r, i) => {
    if (head > 0) return
    const cells = []
    r.eachCell({ includeEmpty: true }, (c, j) => { cells[j] = txt(c) })
    if (cells.some((x) => /^driver$/i.test(x || '')) && cells.some((x) => /^password$/i.test(x || ''))) { head = i; H = cells }
  })
  if (head < 0) throw new Error(`${ws.name}: could not find a header row with "Driver" and "Password".`)

  const col = (re) => H.findIndex((h) => re.test(h || ''))
  const ix = {
    driver: col(/^driver$/i), route: col(/^route$/i), area: col(/^area$/i),
    phone: col(/^phone/i), email: col(/^e-?mail$/i), password: col(/^password$/i),
  }
  // the unlabelled column between Phone and Email holds Yes / Sent
  let status = -1
  for (let j = 1; j < H.length; j++) if (!H[j] && j > ix.phone && (ix.email < 0 || j < ix.email)) { status = j; break }
  // the unlabelled column after Password holds a free note (e.g. "Prince")
  let note = -1
  for (let j = ix.password + 1; j < H.length + 3; j++) if (!H[j]) { note = j; break }

  const drivers = []
  const warnings = []
  ws.eachRow({ includeEmpty: false }, (r, i) => {
    if (i <= head) return
    const g = (k) => (k >= 0 ? txt(r.getCell(k)) : '')
    const name = g(ix.driver)
    if (!name) return
    drivers.push({
      driver: name,
      route: g(ix.route).toUpperCase(),
      area: g(ix.area),
      phone: phone(g(ix.phone)),
      status: status > 0 ? g(status) : '',
      email: g(ix.email).toLowerCase(),
      password: g(ix.password),
      note: note > 0 ? g(note) : '',
    })
  })
  const noCreds = drivers.filter((d) => !d.email || !d.password).length
  if (noCreds) warnings.push(`${noCreds} of ${drivers.length} drivers have no email/password yet on the Drivers tab.`)
  return { drivers, warnings }
}

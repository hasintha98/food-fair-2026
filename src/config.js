/**
 * Loads .env (no dependency), validates what the server needs, and fails
 * closed: no password hash or session secret means the server will not start.
 */
import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

export const BACKEND_DIR = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')
export const ENV_FILE = path.join(BACKEND_DIR, '.env')
export const DATA_DIR = path.join(BACKEND_DIR, 'data')
export const PLAN_FILE = path.join(DATA_DIR, 'plan.json')

/** Minimal .env parser: KEY=value, quotes optional, # comments, blank lines. */
export function readEnvFile(file = ENV_FILE) {
  const out = {}
  if (!fs.existsSync(file)) return out
  for (const raw of fs.readFileSync(file, 'utf8').split(/\r?\n/)) {
    const line = raw.trim()
    if (!line || line.startsWith('#')) continue
    const eq = line.indexOf('=')
    if (eq < 0) continue
    const key = line.slice(0, eq).trim()
    let val = line.slice(eq + 1).trim()
    if ((val.startsWith('"') && val.endsWith('"')) || (val.startsWith("'") && val.endsWith("'"))) val = val.slice(1, -1)
    out[key] = val
  }
  return out
}

/** Writes keys into .env, preserving everything else in the file. */
export function writeEnvFile(updates, file = ENV_FILE) {
  const lines = fs.existsSync(file) ? fs.readFileSync(file, 'utf8').split(/\r?\n/) : []
  const done = new Set()
  const next = lines.map((raw) => {
    const m = raw.match(/^\s*([A-Z0-9_]+)\s*=/)
    if (m && m[1] in updates) { done.add(m[1]); return `${m[1]}=${updates[m[1]]}` }
    return raw
  })
  for (const [k, v] of Object.entries(updates)) if (!done.has(k)) next.push(`${k}=${v}`)
  fs.writeFileSync(file, next.join('\n').replace(/\n*$/, '\n'), { mode: 0o600 })
}

const fileEnv = readEnvFile()
const env = (k, d) => (process.env[k] ?? fileEnv[k] ?? d)

const int = (v, d) => { const n = parseInt(v, 10); return Number.isFinite(n) && n > 0 ? n : d }

export const config = {
  port: int(env('PORT'), 5181),
  passwordHash: env('DASHBOARD_PASSWORD_HASH', ''),
  sessionSecret: env('SESSION_SECRET', ''),
  sessionHours: int(env('SESSION_HOURS'), 24),
  sheetId: env('SHEET_ID', '1Bb7PPyYHRwWUCdPSTDCTV-KPqLxOC5x2rJnfET4w7QQ'),
  staticDir: path.resolve(BACKEND_DIR, env('STATIC_DIR', '../dashboard/dist')),
  exportDir: env('EXPORT_DIR', '..') ? path.resolve(BACKEND_DIR, env('EXPORT_DIR', '..')) : null,
  cookieSecure: env('COOKIE_SECURE', 'auto'),
  // Pull the sheet on boot when there is no cached plan (hosts with ephemeral disks lose it on every deploy).
  syncOnStart: String(env('SYNC_ON_START', 'true')).toLowerCase() !== 'false',
  // Origins allowed to call the API from another site (comma-separated). Empty = same-origin only.
  corsOrigins: String(env('CORS_ORIGIN', '')).split(',').map((s) => s.trim().replace(/\/+$/, '')).filter(Boolean),
  loginMaxAttempts: int(env('LOGIN_MAX_ATTEMPTS'), 5),
  loginLockMinutes: int(env('LOGIN_LOCK_MINUTES'), 15),
}

/** Throws a readable error listing everything that is missing or malformed. */
export function assertConfig(c = config) {
  const problems = []
  if (!c.passwordHash) problems.push('DASHBOARD_PASSWORD_HASH is not set — run `npm run set-password`.')
  else if (!/^scrypt:\d+:\d+:\d+:[A-Za-z0-9+/=]+:[A-Za-z0-9+/=]+$/.test(c.passwordHash)) {
    problems.push('DASHBOARD_PASSWORD_HASH is not in the expected format — run `npm run set-password` again.')
  }
  if (!c.sessionSecret) problems.push('SESSION_SECRET is not set — run `npm run set-password` (it generates one).')
  else if (c.sessionSecret.length < 32) problems.push('SESSION_SECRET is too short — it should be at least 32 characters.')
  if (!['auto', 'true', 'false'].includes(c.cookieSecure)) problems.push('COOKIE_SECURE must be auto, true or false.')
  for (const o of c.corsOrigins) {
    if (!/^https?:\/\/[^/]+$/.test(o)) problems.push(`CORS_ORIGIN entry "${o}" must be a bare origin like https://app.example.com (no path).`)
  }
  if (problems.length) {
    const err = new Error('Configuration is incomplete:\n  - ' + problems.join('\n  - '))
    err.code = 'ECONFIG'
    throw err
  }
  return c
}

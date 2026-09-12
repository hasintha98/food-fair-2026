/**
 * The API. Everything that returns customer data sits behind `requireSession`.
 *
 *   POST /api/login      { password }  -> sets the session cookie
 *   POST /api/logout                   -> clears it
 *   GET  /api/session                  -> { ok, authenticated, expiresAt }
 *   GET  /api/plan          (auth)     -> the cached plan
 *   POST /api/refresh       (auth)     -> re-read the sheet, rewrite the cache, return the plan
 *   GET  /api/health                   -> liveness only, no data
 */
import { config } from './config.js'
import { verifyPassword, issueSession, verifySession, parseCookies, sessionCookie, clearCookie, COOKIE } from './auth.js'
import { LoginLimiter, clientKey, deviceKey } from './ratelimit.js'
import { json, readJson, isHttps, Router } from './http.js'
import { refresh as refreshPlan, readPlan, SHEET_EDIT_URL } from './plan.js'

// Two guards: a soft one per device, and a hard ceiling per address so a
// script cannot dodge the first by changing its browser string.
const limiter = new LoginLimiter({ maxAttempts: config.loginMaxAttempts, lockMinutes: config.loginLockMinutes })
const addrLimiter = new LoginLimiter({ maxAttempts: Math.max(60, config.loginMaxAttempts * 10), lockMinutes: 15 })

const cookieSecure = (req) =>
  config.cookieSecure === 'true' ? true : config.cookieSecure === 'false' ? false : isHttps(req)
const crossSite = config.corsOrigins.length > 0

/**
 * Reads and verifies the session; null when absent, forged, expired or from an
 * old password. The token travels as an HttpOnly cookie and, for browsers that
 * refuse cross-site cookies (Safari on a phone), as an Authorization header.
 */
function sessionOf(req) {
  const auth = String(req.headers.authorization || '')
  const bearer = auth.startsWith('Bearer ') ? auth.slice(7).trim() : ''
  const token = bearer || parseCookies(req.headers.cookie)[COOKIE]
  return token ? verifySession(token, { secret: config.sessionSecret, passwordHash: config.passwordHash }) : null
}

function requireSession(handler) {
  return (req, res, ctx) => {
    const s = sessionOf(req)
    if (!s) return json(res, 401, { ok: false, error: 'Unlock the dashboard first.', code: 'UNAUTHENTICATED' })
    return handler(req, res, { ...ctx, session: s })
  }
}

// A refresh already running is shared with anyone else who asks meanwhile.
let inFlight = null

export const router = new Router()

router.get('/api/health', (req, res) => {
  json(res, 200, { ok: true, uptime: Math.round(process.uptime()) })
})

router.post('/api/login', async (req, res) => {
  const key = deviceKey(req, config.trustProxy)
  const addr = clientKey(req, config.trustProxy)
  const locked = limiter.lockedFor(key) || addrLimiter.lockedFor(addr)
  if (locked) {
    return json(res, 429, {
      ok: false, code: 'LOCKED',
      error: `Too many attempts. Try again in ${Math.ceil(locked / 60)} minute${locked > 60 ? 's' : ''}.`,
      retryAfter: locked,
    }, { 'Retry-After': String(locked) })
  }

  const body = await readJson(req, 2048)
  const password = typeof body.password === 'string' ? body.password : ''

  const ok = password.length > 0 && (await verifyPassword(password, config.passwordHash))
  if (!ok) {
    // Slow the loop a little even before the lock engages.
    await new Promise((r) => setTimeout(r, 350))
    const { locked: devLocked, remaining } = limiter.fail(key)
    const { locked: addrLocked } = addrLimiter.fail(addr)
    const nowLocked = addrLocked || devLocked
    if (nowLocked) {
      return json(res, 429, {
        ok: false, code: 'LOCKED',
        error: `Too many attempts. Try again in ${Math.ceil(nowLocked / 60)} minute${nowLocked > 60 ? 's' : ''}.`,
        retryAfter: nowLocked,
      }, { 'Retry-After': String(nowLocked) })
    }
    return json(res, 401, { ok: false, code: 'BAD_PASSWORD', error: 'That password is not right.', remaining })
  }

  limiter.succeed(key)
  addrLimiter.succeed(addr)
  const { token, expiresAt } = issueSession({
    secret: config.sessionSecret, passwordHash: config.passwordHash, hours: config.sessionHours,
  })
  // the token goes back in the body too: the dashboard keeps it for browsers that drop the cookie
  json(res, 200, { ok: true, expiresAt, token }, {
    'Set-Cookie': sessionCookie(token, { maxAgeSec: config.sessionHours * 3600, secure: cookieSecure(req), crossSite }),
  })
})

router.post('/api/logout', (req, res) => {
  json(res, 200, { ok: true }, { 'Set-Cookie': clearCookie({ secure: cookieSecure(req), crossSite }) })
})

router.get('/api/session', (req, res) => {
  const s = sessionOf(req)
  json(res, 200, { ok: true, authenticated: Boolean(s), expiresAt: s ? s.exp : null })
})

router.get('/api/plan', requireSession((req, res) => {
  const plan = readPlan()
  if (!plan) {
    return json(res, 503, {
      ok: false, code: 'NO_DATA',
      error: 'The plan has not been synced yet. Press Refresh, or run `npm run sync` on the server.',
    })
  }
  json(res, 200, { ok: true, payload: plan })
}))

router.post('/api/refresh', requireSession(async (req, res) => {
  try {
    if (!inFlight) inFlight = refreshPlan().finally(() => { inFlight = null })
    const out = await inFlight
    json(res, 200, {
      ok: true,
      payload: out.payload,
      warnings: out.warnings,
      totals: out.totals,
      softErrors: out.softErrors,
      syncedAt: out.payload.plan.syncedAt,
    })
  } catch (err) {
    console.error(`[${new Date().toISOString()}] refresh failed:`, err.message)
    json(res, 502, { ok: false, code: 'SHEET', error: err.message, sheet: SHEET_EDIT_URL })
  }
}))

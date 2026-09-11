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
import { LoginLimiter, clientKey } from './ratelimit.js'
import { json, readJson, isHttps, Router } from './http.js'
import { refresh as refreshPlan, readPlan, SHEET_EDIT_URL } from './plan.js'

const limiter = new LoginLimiter({ maxAttempts: config.loginMaxAttempts, lockMinutes: config.loginLockMinutes })

const cookieSecure = (req) =>
  config.cookieSecure === 'true' ? true : config.cookieSecure === 'false' ? false : isHttps(req)
const crossSite = config.corsOrigins.length > 0

/** Reads and verifies the session cookie; null when absent, forged, expired or from an old password. */
function sessionOf(req) {
  const token = parseCookies(req.headers.cookie)[COOKIE]
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
  const key = clientKey(req)
  const locked = limiter.lockedFor(key)
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
    const { locked: nowLocked, remaining } = limiter.fail(key)
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
  const { token, expiresAt } = issueSession({
    secret: config.sessionSecret, passwordHash: config.passwordHash, hours: config.sessionHours,
  })
  json(res, 200, { ok: true, expiresAt }, {
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

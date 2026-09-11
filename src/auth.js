/**
 * Password hashing and stateless sessions, built only on node:crypto.
 *
 * - The password is stored as an scrypt hash; the plaintext never touches disk.
 * - A session is a signed token (HMAC-SHA256) carrying its own expiry, so no
 *   store is needed. It is bound to the current password hash: changing the
 *   password invalidates every session.
 * - All comparisons are constant-time.
 */
import { scrypt, randomBytes, timingSafeEqual, createHmac, createHash } from 'node:crypto'
import { promisify } from 'node:util'

const scryptAsync = promisify(scrypt)

// scrypt cost. N=2^14, r=8, p=1 is ~50ms on a laptop and well under Node's
// default 32MB maxmem (128 * N * r = 16MB).
const N = 16384, R = 8, P = 1, KEYLEN = 64

export async function hashPassword(password) {
  if (typeof password !== 'string' || password.length < 8) {
    throw new Error('Password must be at least 8 characters.')
  }
  const salt = randomBytes(16)
  const key = await scryptAsync(password.normalize('NFKC'), salt, KEYLEN, { N, r: R, p: P })
  return `scrypt:${N}:${R}:${P}:${salt.toString('base64')}:${key.toString('base64')}`
}

export async function verifyPassword(password, stored) {
  const parts = String(stored || '').split(':')
  if (parts.length !== 6 || parts[0] !== 'scrypt') return false
  const [, n, r, p, saltB64, hashB64] = parts
  const expected = Buffer.from(hashB64, 'base64')
  let actual
  try {
    actual = await scryptAsync(String(password ?? '').normalize('NFKC'), Buffer.from(saltB64, 'base64'),
      expected.length, { N: +n, r: +r, p: +p })
  } catch {
    return false
  }
  return actual.length === expected.length && timingSafeEqual(actual, expected)
}

// ---------------------------------------------------------------- sessions

const b64u = (buf) => Buffer.from(buf).toString('base64url')
const unb64u = (s) => Buffer.from(s, 'base64url')

/** Short fingerprint of the password hash; sessions carry it so a password change logs everyone out. */
export const passwordVersion = (hash) => createHash('sha256').update(String(hash)).digest('base64url').slice(0, 12)

export function issueSession({ secret, passwordHash, hours }) {
  const now = Date.now()
  const payload = { v: 1, iat: now, exp: now + hours * 3600 * 1000, pv: passwordVersion(passwordHash) }
  const body = b64u(JSON.stringify(payload))
  const sig = createHmac('sha256', secret).update(body).digest('base64url')
  return { token: `${body}.${sig}`, expiresAt: payload.exp }
}

/** Returns the payload if the token is genuine, unexpired and for the current password; otherwise null. */
export function verifySession(token, { secret, passwordHash }) {
  if (typeof token !== 'string') return null
  const dot = token.indexOf('.')
  if (dot < 0) return null
  const body = token.slice(0, dot)
  const sig = token.slice(dot + 1)
  const expect = createHmac('sha256', secret).update(body).digest('base64url')
  const a = Buffer.from(sig), b = Buffer.from(expect)
  if (a.length !== b.length || !timingSafeEqual(a, b)) return null
  let payload
  try { payload = JSON.parse(unb64u(body).toString('utf8')) } catch { return null }
  if (!payload || payload.v !== 1 || typeof payload.exp !== 'number') return null
  if (payload.exp <= Date.now()) return null
  if (payload.pv !== passwordVersion(passwordHash)) return null
  return payload
}

// ---------------------------------------------------------------- cookies

export const COOKIE = 'ff_session'

export function parseCookies(header) {
  const out = {}
  for (const part of String(header || '').split(';')) {
    const i = part.indexOf('=')
    if (i < 0) continue
    out[part.slice(0, i).trim()] = decodeURIComponent(part.slice(i + 1).trim())
  }
  return out
}

/**
 * Same-origin: SameSite=Strict. Cross-site (frontend hosted elsewhere):
 * SameSite=None, which browsers only honour together with Secure.
 */
export function sessionCookie(token, { maxAgeSec, secure, crossSite = false }) {
  const bits = [
    `${COOKIE}=${encodeURIComponent(token)}`,
    'Path=/',
    'HttpOnly',
    crossSite ? 'SameSite=None' : 'SameSite=Strict',
    `Max-Age=${maxAgeSec}`,
  ]
  if (secure || crossSite) bits.push('Secure')
  return bits.join('; ')
}

export function clearCookie({ secure, crossSite = false }) {
  return sessionCookie('', { maxAgeSec: 0, secure, crossSite })
}

export const newSecret = () => randomBytes(32).toString('hex')

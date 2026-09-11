/**
 * In-memory brute-force guard for the login endpoint. Per client address:
 * after `maxAttempts` failures the address is locked for `lockMinutes`.
 * A successful login clears the counter. Old entries are swept periodically.
 */
export class LoginLimiter {
  constructor({ maxAttempts = 5, lockMinutes = 15 } = {}) {
    this.max = maxAttempts
    this.lockMs = lockMinutes * 60 * 1000
    this.map = new Map()
    this.sweep = setInterval(() => this.#sweep(), 60 * 1000)
    if (this.sweep.unref) this.sweep.unref()
  }

  /** null if allowed, otherwise seconds until the lock lifts. */
  lockedFor(key) {
    const e = this.map.get(key)
    if (!e || !e.until) return null
    const left = e.until - Date.now()
    if (left <= 0) { this.map.delete(key); return null }
    return Math.ceil(left / 1000)
  }

  /** Records a failure; returns { locked: seconds|null, remaining: attempts left }. */
  fail(key) {
    const e = this.map.get(key) || { count: 0, until: 0, seen: 0 }
    e.count += 1
    e.seen = Date.now()
    if (e.count >= this.max) { e.until = Date.now() + this.lockMs; e.count = 0 }
    this.map.set(key, e)
    return { locked: e.until ? Math.ceil((e.until - Date.now()) / 1000) : null, remaining: Math.max(0, this.max - e.count) }
  }

  succeed(key) { this.map.delete(key) }

  #sweep() {
    const cutoff = Date.now() - Math.max(this.lockMs, 60 * 60 * 1000)
    for (const [k, e] of this.map) if ((e.until || e.seen) < cutoff) this.map.delete(k)
  }
}

/** Client address, honouring a reverse proxy only when told to. */
export function clientKey(req, trustProxy = false) {
  if (trustProxy) {
    const xf = req.headers['x-forwarded-for']
    if (xf) return String(xf).split(',')[0].trim()
  }
  return req.socket.remoteAddress || 'unknown'
}

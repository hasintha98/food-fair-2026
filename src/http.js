/**
 * The small amount of HTTP plumbing the server needs: JSON in/out, a router,
 * security headers, and static serving of the built dashboard with a
 * path-traversal guard and SPA fallback. No framework.
 */
import fs from 'node:fs'
import path from 'node:path'

export function json(res, status, body, extraHeaders = {}) {
  const data = JSON.stringify(body)
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Cache-Control': 'no-store',
    ...extraHeaders,
  })
  res.end(data)
}

/** Reads a JSON body, refusing anything over `limit` bytes. */
export function readJson(req, limit = 4096) {
  return new Promise((resolve, reject) => {
    let size = 0
    let tooBig = false
    const chunks = []
    req.on('data', (c) => {
      if (tooBig) return                      // discard the rest, keep the socket sane
      size += c.length
      if (size > limit) {
        tooBig = true
        chunks.length = 0
        reject(Object.assign(new Error('Body too large'), { status: 413 }))
      } else chunks.push(c)
    })
    req.on('end', () => {
      if (tooBig) return
      if (!chunks.length) return resolve({})
      try { resolve(JSON.parse(Buffer.concat(chunks).toString('utf8'))) }
      catch { reject(Object.assign(new Error('Invalid JSON'), { status: 400 })) }
    })
    req.on('error', reject)
  })
}

export function isHttps(req) {
  return Boolean(req.socket.encrypted) || req.headers['x-forwarded-proto'] === 'https'
}

/** Applied to every response. */
export function securityHeaders(res) {
  res.setHeader('X-Content-Type-Options', 'nosniff')
  res.setHeader('X-Frame-Options', 'DENY')
  res.setHeader('Referrer-Policy', 'no-referrer')
  res.setHeader('Permissions-Policy', 'camera=(), microphone=(), geolocation=()')
  res.setHeader('Content-Security-Policy', [
    "default-src 'self'",
    "script-src 'self'",
    "style-src 'self' 'unsafe-inline'",  // React sets inline style attributes for bar widths
    "img-src 'self' data:",
    "font-src 'self'",
    "connect-src 'self'",
    "frame-ancestors 'none'",
    "base-uri 'self'",
    "form-action 'self'",
    "object-src 'none'",
  ].join('; '))
}

// ---------------------------------------------------------------- router

export class Router {
  constructor() { this.routes = [] }
  add(method, pathname, handler) { this.routes.push({ method, pathname, handler }); return this }
  get(p, h) { return this.add('GET', p, h) }
  post(p, h) { return this.add('POST', p, h) }

  /** Returns true if a route handled the request. */
  async dispatch(req, res, ctx) {
    const url = (req.url || '/').split('?')[0]
    for (const r of this.routes) {
      if (r.pathname !== url) continue
      if (r.method !== req.method) {
        if (req.method === 'HEAD' && r.method === 'GET') { /* fine */ }
        else { json(res, 405, { ok: false, error: 'Method not allowed' }, { Allow: r.method }); return true }
      }
      try {
        await r.handler(req, res, ctx)
      } catch (err) {
        const status = err.status || 500
        if (status >= 500) console.error(`[${new Date().toISOString()}] ${req.method} ${url} ->`, err)
        if (!res.headersSent) json(res, status, { ok: false, error: status >= 500 ? 'Server error' : err.message })
      }
      return true
    }
    return false
  }
}

// ---------------------------------------------------------------- static

const TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.webp': 'image/webp',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
  '.map': 'application/json',
  '.txt': 'text/plain; charset=utf-8',
  '.webmanifest': 'application/manifest+json',
}

/**
 * Serves `dir`. Anything that resolves outside it is refused; anything not
 * found falls back to index.html so client-side routes work.
 */
export function serveStatic(dir) {
  const root = path.resolve(dir)
  const index = path.join(root, 'index.html')

  return (req, res) => {
    if (req.method !== 'GET' && req.method !== 'HEAD') {
      json(res, 405, { ok: false, error: 'Method not allowed' })
      return
    }
    let rel
    try { rel = decodeURIComponent((req.url || '/').split('?')[0]) } catch { res.writeHead(400); return res.end('Bad request') }
    if (rel.includes('\0')) { res.writeHead(400); return res.end('Bad request') }

    let file = path.resolve(root, '.' + rel)
    if (file !== root && !file.startsWith(root + path.sep)) { res.writeHead(403); return res.end('Forbidden') }

    let stat = null
    try { stat = fs.statSync(file) } catch { /* fall through */ }
    if (!stat || stat.isDirectory()) file = index

    let body
    try { body = fs.readFileSync(file) } catch { res.writeHead(404); return res.end('Not found') }

    const ext = path.extname(file).toLowerCase()
    res.setHeader('Content-Type', TYPES[ext] || 'application/octet-stream')
    // Vite hashes asset names, so they can be cached hard; index.html must not be.
    res.setHeader('Cache-Control', rel.startsWith('/assets/') ? 'public, max-age=31536000, immutable' : 'no-cache')
    res.writeHead(200)
    res.end(req.method === 'HEAD' ? undefined : body)
  }
}

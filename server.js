/**
 * Food Fair delivery dashboard — backend.
 *
 * One process serves both the built React app and the API. Customer data only
 * ever leaves this process through /api/plan and /api/refresh, both of which
 * require a valid session cookie.
 *
 *   npm run set-password   first-time setup (writes .env)
 *   npm start              serve on PORT (default 5181)
 */
import http from 'node:http'
import os from 'node:os'
import fs from 'node:fs'
import path from 'node:path'
import { config, assertConfig } from './src/config.js'
import { securityHeaders, serveStatic, json, cors } from './src/http.js'
import { router } from './src/routes.js'
import { readPlan, refresh } from './src/plan.js'

try {
  assertConfig()
} catch (err) {
  console.error('\n' + err.message + '\n')
  process.exit(1)
}

const staticHandler = serveStatic(config.staticDir)
const hasBuild = fs.existsSync(path.join(config.staticDir, 'index.html'))

const server = http.createServer(async (req, res) => {
  securityHeaders(res)

  if ((req.url || '').startsWith('/api/')) {
    if (cors(req, res, config.corsOrigins)) return
    if (!(await router.dispatch(req, res, {}))) json(res, 404, { ok: false, error: 'Not found' })
    return
  }

  if (!hasBuild) {
    res.writeHead(503, { 'Content-Type': 'text/plain; charset=utf-8' })
    return res.end('The dashboard has not been built. Run `npm run build` in ../dashboard, then restart.')
  }
  staticHandler(req, res)
})

server.on('clientError', (err, socket) => {
  if (socket.writable) socket.end('HTTP/1.1 400 Bad Request\r\n\r\n')
})

server.listen(config.port, () => {
  const ips = Object.values(os.networkInterfaces()).flat()
    .filter((n) => n && n.family === 'IPv4' && !n.internal).map((n) => n.address)
  const plan = readPlan()

  console.log('\n  Food Fair delivery dashboard — backend\n')
  console.log(`  Local:    http://localhost:${config.port}/`)
  ips.forEach((ip) => console.log(`  Network:  http://${ip}:${config.port}/`))
  console.log('')
  console.log(`  Session:  ${config.sessionHours}h after unlock`)
  console.log(`  Data:     ${plan ? `${plan.orders.length} orders, synced ${plan.plan.syncedAt}` : 'not synced yet — press Refresh or run `npm run sync`'}`)
  console.log(`  Frontend: ${hasBuild ? config.staticDir : 'NOT BUILT — run `npm run build` in ../dashboard'}`)
  if (config.corsOrigins.length) {
    console.log(`  CORS:     ${config.corsOrigins.join(', ')}`)
    console.log('            cookie is SameSite=None; Secure — serve this API over https or browsers will drop it')
  }
  console.log('')
})

// Ephemeral hosts start with an empty data dir — fetch the sheet so the first
// visitor is not met with "not synced yet". Failures are logged, not fatal.
if (config.syncOnStart && !readPlan()) {
  console.log('  No cached plan — reading the sheet now…')
  refresh()
    .then((o) => console.log(`  Synced: ${o.totals.deliveries} deliveries on ${o.totals.routes} routes${o.warnings.length ? ` (${o.warnings.length} check warnings)` : ''}`))
    .catch((e) => console.error('  Startup sync failed: ' + e.message + ' — Refresh from the dashboard will retry.'))
}

const shutdown = () => { server.close(() => process.exit(0)); setTimeout(() => process.exit(0), 2000).unref() }
process.on('SIGINT', shutdown)
process.on('SIGTERM', shutdown)

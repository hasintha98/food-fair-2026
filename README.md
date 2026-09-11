# Food Fair 2026 — backend

Node server that owns the delivery plan and serves the dashboard behind a
password. No database, no framework — `node:http` and `node:crypto`, plus
`exceljs` to read the Google Sheet.

## First run

```bash
cd backend
npm install
npm run set-password        # prompts twice, input hidden; writes .env
npm run sync                # optional: pull the sheet now instead of on first Refresh
npm start
```

Then build the frontend once (`cd ../dashboard && npm run build`) and open
http://localhost:5181. The server prints a `Network:` address so phones on the
same wifi can reach it.

`set-password` also generates `SESSION_SECRET` if there isn't one. Run it again
any time to change the password — **every device is logged out** when you do.

## What is protected, and how

| Layer | Detail |
|---|---|
| Password at rest | scrypt (N=2¹⁴, r=8, p=1, 16-byte salt). The plaintext is never written anywhere. |
| Comparison | `crypto.timingSafeEqual` on both the password hash and the session signature. |
| Session | HMAC-SHA256-signed token carrying its own expiry (`SESSION_HOURS`, default 24). Stateless — no store to leak. Bound to the current password hash, so a password change invalidates all sessions. |
| Cookie | `HttpOnly` (JS cannot read it), `SameSite=Strict` (no cross-site sends), `Secure` when served over https. |
| Brute force | Per-address: 5 failures → 15-minute lock, 350 ms delay on each failure. The lock applies even if the correct password is then supplied. |
| Data | Only `/api/plan` and `/api/refresh` return customer data, and both require a valid session. The cached JSON lives in `backend/data/`, outside the static root. The built frontend contains no data at all. |
| Headers | CSP (`default-src 'self'`, no inline scripts), `X-Frame-Options: DENY`, `nosniff`, `Referrer-Policy: no-referrer`, `Cache-Control: no-store` on every API response. |
| Static files | Path-traversal guard; anything resolving outside `dist/` is refused. |
| Input | JSON bodies capped at 2 KB on login; malformed JSON → 400; wrong method → 405. |
| Startup | Fails closed: no password hash or session secret → the process exits with instructions. |

## API

| Route | Auth | Purpose |
|---|---|---|
| `POST /api/login` `{ password }` | – | Sets the session cookie. 401 wrong, 429 locked (with `Retry-After`). |
| `POST /api/logout` | – | Clears the cookie. |
| `GET /api/session` | – | `{ authenticated, expiresAt }` — no data. |
| `GET /api/plan` | ✔ | The cached plan. 503 `NO_DATA` if never synced. |
| `POST /api/refresh` | ✔ | Re-reads the sheet, rewrites the cache and exports, returns the plan. Concurrent calls share one download. |
| `GET /api/health` | – | Liveness only. |

## Configuration (`.env`)

See `.env.example`. Everything has a default except the two secrets.
`EXPORT_DIR` controls where the CSV/XLSX copies are written (project root by
default; blank disables them). `COOKIE_SECURE=true` if you put this behind an
https reverse proxy that terminates TLS.

## Development

Two terminals:

```bash
cd backend   && npm run dev      # API on :5181, restarts on change
cd dashboard && npm run dev      # Vite on :5180, proxies /api → :5181
```

Cookies stay same-origin through the proxy, so login works in dev exactly as in
production.

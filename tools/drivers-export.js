/**
 * Writes the Drivers tab to backend/drivers.json so it can be committed and
 * deployed where backend/data/ does not persist (Railway).
 *
 *   npm run drivers:export                      finds the workbook in data/
 *   npm run drivers:export -- path/to/file.xlsx
 *
 * CAUTION: the output contains plain-text passwords. Keep the repository
 * private. To keep them out of git entirely, use the printed DRIVERS_JSON
 * value as a Railway variable instead of committing the file.
 */
import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { loadDrivers } from '../src/drivers.js'

const BACKEND = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')
const OUT = path.join(BACKEND, 'drivers.json')

;(async () => {
  const file = process.argv[2]
  let out
  try {
    out = await loadDrivers({ file })
  } catch (err) {
    console.error('Could not read the Drivers tab: ' + err.message)
    process.exit(1)
  }
  if (!out.drivers.length) {
    console.error('No drivers found. Put the workbook with the Drivers tab in backend/data/ or pass its path.')
    process.exit(1)
  }
  const payload = {
    source: path.basename(String(out.source || '')),
    exportedAt: new Date().toISOString(),
    drivers: out.drivers,
  }
  fs.writeFileSync(OUT, JSON.stringify(payload, null, 2) + '\n')

  const withCreds = out.drivers.filter((d) => d.email && d.password).length
  console.log(`Wrote ${OUT}`)
  console.log(`  ${out.drivers.length} drivers, ${withCreds} with email + password, from ${payload.source}`)
  out.warnings.forEach((w) => console.log('  ! ' + w))
  console.log('\nThis file contains passwords in plain text. Commit it only to a PRIVATE repository.')
  console.log('Alternative that keeps them out of git — set this as a Railway variable named DRIVERS_JSON:\n')
  console.log(Buffer.from(JSON.stringify(payload)).toString('base64'))
  console.log('')
})()

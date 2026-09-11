/**
 * Terminal refresh — the same work as the dashboard's Refresh button.
 *
 *   npm run sync                 pull the live sheet
 *   npm run sync -- some.xlsx    read a local workbook instead
 */
import { refresh, SHEET_EDIT_URL } from '../src/plan.js'

;(async () => {
  const file = process.argv[2]
  console.log(file ? `Reading ${file}` : 'Downloading the shared sheet…')

  let out
  try {
    out = await refresh({ file })
  } catch (err) {
    console.error('\nRefresh failed: ' + err.message)
    console.error('Sheet: ' + SHEET_EDIT_URL)
    process.exit(1)
  }

  const { payload, warnings, totals, movedNotes = [], written, softErrors, bytes } = out
  if (bytes) console.log(`  got ${(bytes / 1024).toFixed(0)} KB`)

  console.log('\n=== ' + (payload.plan.title || 'Delivery plan') + ' ===')
  console.log(`routes ${totals.routes} | areas ${totals.areas} | drivers ${totals.drivers}${totals.unassigned ? ` | UNASSIGNED ${totals.unassigned}` : ''}`)
  console.log(`deliveries ${totals.deliveries} | unique doors ${totals.doors} | packs ${totals.packs} | packing notes ${totals.packingNotes}`)
  if (movedNotes.length) {
    console.log(`
--- ${movedNotes.length} instruction(s) moved from delivery to packing ---`)
    movedNotes.forEach((m) => console.log(`  ${m.orderId.padEnd(8)} ${m.text}`))
  }

  if (warnings.length) {
    console.log('\n--- CHECKS ---')
    warnings.forEach((w) => console.log('  ! ' + w))
  } else {
    console.log('\nAll cross-tab checks passed.')
  }

  console.log('')
  written.forEach((p) => console.log('wrote ' + p))
  if (softErrors.length) {
    console.log('')
    softErrors.forEach((e) => console.error('! ' + e))
    process.exitCode = 1
  }
})()

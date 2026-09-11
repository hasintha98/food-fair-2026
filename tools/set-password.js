/**
 * First-time setup and password changes.
 *
 *   npm run set-password                 prompts (input hidden)
 *   npm run set-password -- "pass"       non-interactive
 *
 * Writes DASHBOARD_PASSWORD_HASH to .env, and SESSION_SECRET too if there
 * isn't one yet. Changing the password logs every device out.
 */
import readline from 'node:readline'
import { hashPassword, newSecret } from '../src/auth.js'
import { readEnvFile, writeEnvFile, ENV_FILE } from '../src/config.js'

function ask(question, { hidden = false } = {}) {
  return new Promise((resolve) => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout, terminal: true })
    if (hidden) {
      // Echo nothing while the password is typed.
      const write = rl._writeToOutput
      rl._writeToOutput = (s) => { if (s.includes('\n')) write.call(rl, '\n') }
      process.stdout.write(question)
    }
    rl.question(hidden ? '' : question, (a) => { rl.close(); resolve(a) })
  })
}

;(async () => {
  let password = process.argv[2]
  if (!password) {
    if (!process.stdin.isTTY) {
      console.error('No terminal to prompt on. Pass the password as an argument: npm run set-password -- "your password"')
      process.exit(1)
    }
    password = await ask('New dashboard password (min 8 chars): ', { hidden: true })
    const again = await ask('Type it again: ', { hidden: true })
    if (password !== again) { console.error('They do not match. Nothing changed.'); process.exit(1) }
  }

  let hash
  try { hash = await hashPassword(password) }
  catch (err) { console.error(err.message); process.exit(1) }

  const current = readEnvFile()
  const updates = { DASHBOARD_PASSWORD_HASH: hash }
  if (!current.SESSION_SECRET) updates.SESSION_SECRET = newSecret()
  if (!current.PORT) updates.PORT = '5181'
  if (!current.SESSION_HOURS) updates.SESSION_HOURS = '24'

  writeEnvFile(updates)
  console.log(`\nSaved to ${ENV_FILE}`)
  console.log('  DASHBOARD_PASSWORD_HASH  updated (scrypt)')
  if (updates.SESSION_SECRET) console.log('  SESSION_SECRET           generated')
  console.log('\nEvery existing session is now invalid. Restart the server if it is running.\n')
})()

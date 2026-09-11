/**
 * Prints a password hash and nothing else — for hosted deployments (Railway,
 * Render, Fly…) where configuration lives in the platform's variables, not
 * in a .env file.
 *
 *   npm run hash-password                 prompts (input hidden)
 *   npm run hash-password -- "pass"       non-interactive
 *
 * Paste the output into DASHBOARD_PASSWORD_HASH on the host. The platform
 * restarts the service; the new password is live and every device is logged out.
 */
import readline from 'node:readline'
import { hashPassword } from '../src/auth.js'

function askHidden(question) {
  return new Promise((resolve) => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout, terminal: true })
    const write = rl._writeToOutput
    rl._writeToOutput = (s) => { if (s.includes('\n')) write.call(rl, '\n') }
    process.stdout.write(question)
    rl.question('', (a) => { rl.close(); resolve(a) })
  })
}

;(async () => {
  let password = process.argv[2]
  if (!password) {
    if (!process.stdin.isTTY) {
      console.error('No terminal to prompt on. Pass the password as an argument: npm run hash-password -- "your password"')
      process.exit(1)
    }
    password = await askHidden('Password to hash (min 8 chars): ')
    const again = await askHidden('Type it again: ')
    if (password !== again) { console.error('They do not match.'); process.exit(1) }
  }

  let hash
  try { hash = await hashPassword(password) }
  catch (err) { console.error(err.message); process.exit(1) }

  console.log('\nSet this as DASHBOARD_PASSWORD_HASH on your host:\n')
  console.log(hash)
  console.log('\nThe password itself is not stored anywhere. Nothing was written locally.\n')
})()

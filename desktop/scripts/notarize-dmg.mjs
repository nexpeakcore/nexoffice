#!/usr/bin/env node
/**
 * electron-builder notarizes + staples the *app bundle* before packing the
 * .dmg — so the .dmg itself carries no signature and no ticket. Users who
 * download the .dmg from the web (quarantine flag set) get blocked by
 * Gatekeeper at the disk-image step.
 *
 * Run this after `bun run dist:mac`:
 *   1. sign the .dmg with the same Developer ID that signed the app
 *   2. notarytool submit --wait
 *   3. stapler staple
 *   4. refresh the .dmg sha512 + size in latest-mac.yml (stapling changes the file)
 *
 * Required env (credentials only ever come from the environment):
 *   APPLE_ID                    Apple ID email used for notarization
 *   APPLE_TEAM_ID               Developer team id
 *   APPLE_APP_SPECIFIC_PASSWORD app-specific password for notarytool
 *   — or, instead of the three above —
 *   APPLE_KEYCHAIN_PROFILE      notarytool keychain profile name
 * Optional:
 *   CSC_NAME                    signing identity override (otherwise read from
 *                               the freshly built app bundle)
 */
import { execFileSync, spawnSync } from 'node:child_process'
import crypto from 'node:crypto'
import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')
const RELEASE_DIR = path.join(ROOT, 'release')
const FEED = path.join(RELEASE_DIR, 'latest-mac.yml')

function fail(message) {
  console.error(`❌ ${message}`)
  process.exit(1)
}

function run(cmd, args, opts = {}) {
  return execFileSync(cmd, args, { encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'], ...opts })
}

// Developer ID used for signing — read straight from the freshly built bundle.
function signingIdentity() {
  if (process.env.CSC_NAME) return process.env.CSC_NAME
  const appDir = fs
    .readdirSync(RELEASE_DIR)
    .map((f) => path.join(RELEASE_DIR, f))
    .find((p) => fs.statSync(p).isDirectory() && p.includes('mac'))
  if (!appDir) fail('No mac-* directory in release/ — build first.')
  const app = fs.readdirSync(appDir).find((f) => f.endsWith('.app'))
  if (!app) fail(`No .app found in ${appDir}`)
  // codesign -dv writes to stderr, so read both streams.
  const res = spawnSync('codesign', ['-dv', '--verbose=2', path.join(appDir, app)], { encoding: 'utf8' })
  const match = /Authority=(Developer ID Application: [^\n]+)/.exec(`${res.stdout ?? ''}${res.stderr ?? ''}`)
  if (!match) fail('Could not read the Developer ID from the app — set CSC_NAME manually.')
  return match[1].trim()
}

function notarytoolAuth() {
  if (process.env.APPLE_KEYCHAIN_PROFILE) return ['--keychain-profile', process.env.APPLE_KEYCHAIN_PROFILE]
  const { APPLE_ID, APPLE_TEAM_ID, APPLE_APP_SPECIFIC_PASSWORD } = process.env
  if (!APPLE_ID || !APPLE_TEAM_ID || !APPLE_APP_SPECIFIC_PASSWORD) {
    fail('Missing APPLE_ID / APPLE_TEAM_ID / APPLE_APP_SPECIFIC_PASSWORD (or APPLE_KEYCHAIN_PROFILE).')
  }
  return ['--apple-id', APPLE_ID, '--team-id', APPLE_TEAM_ID, '--password', APPLE_APP_SPECIFIC_PASSWORD]
}

// Refresh sha512 + size of one artifact inside latest-mac.yml.
function updateFeed(fileName, filePath) {
  if (!fs.existsSync(FEED)) {
    console.warn(`  ⚠️  ${path.basename(FEED)} not found — skipping checksum update`)
    return
  }
  const buf = fs.readFileSync(filePath)
  const sha512 = crypto.createHash('sha512').update(buf).digest('base64')
  const size = buf.length

  const lines = fs.readFileSync(FEED, 'utf8').split('\n')
  let inEntry = false
  let patched = 0
  for (let i = 0; i < lines.length; i += 1) {
    if (/^\s*-\s*url:\s*/.test(lines[i])) inEntry = lines[i].trim().endsWith(fileName)
    if (!inEntry) continue
    if (/^\s*sha512:\s*/.test(lines[i])) {
      lines[i] = lines[i].replace(/sha512:\s*.*/, `sha512: ${sha512}`)
      patched += 1
    } else if (/^\s*size:\s*/.test(lines[i])) {
      lines[i] = lines[i].replace(/size:\s*.*/, `size: ${size}`)
      patched += 1
      inEntry = false
    }
  }
  fs.writeFileSync(FEED, lines.join('\n'))
  console.log(`  ✏️  latest-mac.yml: updated ${patched} field(s) for ${fileName}`)
}

function main() {
  if (process.platform !== 'darwin') fail('macOS only.')
  if (!fs.existsSync(RELEASE_DIR)) fail('No release/ directory — build first.')

  const dmgs = fs.readdirSync(RELEASE_DIR).filter((f) => f.endsWith('.dmg'))
  if (!dmgs.length) {
    console.log('ℹ️  No .dmg in release/ — nothing to do.')
    return
  }

  const identity = signingIdentity()
  const auth = notarytoolAuth()
  console.log(`🔏 Signing + notarizing DMG with: ${identity}`)

  for (const name of dmgs) {
    const file = path.join(RELEASE_DIR, name)
    console.log(`\n📀 ${name}`)

    console.log('  🔏 codesign…')
    run('codesign', ['--sign', identity, '--timestamp', '--force', file], { stdio: 'inherit' })

    console.log('  ☁️  notarytool submit (waiting on Apple, usually 2-5 min)…')
    run('xcrun', ['notarytool', 'submit', file, ...auth, '--wait'], { stdio: 'inherit' })

    console.log('  📌 stapler staple…')
    run('xcrun', ['stapler', 'staple', file], { stdio: 'inherit' })

    // Stapling writes the ticket into the file → sha512/size change, and the
    // feed must match again.
    updateFeed(name, file)

    const verdict = (() => {
      try {
        run('spctl', ['-a', '-t', 'open', '--context', 'context:primary-signature', file])
        return 'accepted'
      } catch (err) {
        return `NOT accepted: ${String(err.stderr || err.message).trim().split('\n')[0]}`
      }
    })()
    console.log(`  ✅ Gatekeeper: ${verdict}`)
  }

  console.log('\n✅ DMG signed + notarized + stapled. Run `bun run deploy mac` to upload.')
}

main()

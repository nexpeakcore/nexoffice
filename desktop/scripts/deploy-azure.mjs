#!/usr/bin/env node
/**
 * Upload electron-builder artifacts to Azure Blob Storage (container
 * `nexoffice`, public read) so `electron-updater`'s generic provider can serve
 * them.
 *
 *   node scripts/deploy-azure.mjs [mac|win|linux|all] [--dry-run]
 *
 * Credentials (first match wins, environment only — never hardcoded):
 *   AZURE_STORAGE_SAS_URL          https://<acct>.blob.core.windows.net/nexoffice?sv=...  (preferred for CI)
 *   AZURE_STORAGE_CONNECTION_STRING  DefaultEndpointsProtocol=...;AccountName=...;AccountKey=...
 * Both can also live in `.env.deploy` next to desktop/package.json (git-ignored).
 * Optional: AZURE_BLOB_CONTAINER (default `nexoffice`), AZURE_BLOB_PREFIX.
 *
 * Binaries are uploaded first, the `latest*.yml` feed last — clients must never
 * see a manifest pointing at a file that is still uploading.
 */
import crypto from 'node:crypto'
import fs from 'node:fs'
import fsp from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')
const RELEASE_DIR = path.join(ROOT, 'release')
const API_VERSION = '2021-08-06'
const BLOCK_SIZE = 8 * 1024 * 1024 // 8 MiB
const SINGLE_PUT_LIMIT = 64 * 1024 * 1024 // above this we switch to block upload
const BLOCK_CONCURRENCY = 4

const FEEDS = {
  mac: 'latest-mac.yml',
  win: 'latest.yml',
  linux: 'latest-linux.yml',
}

const CONTENT_TYPES = {
  '.yml': 'text/yaml; charset=utf-8',
  '.zip': 'application/zip',
  '.dmg': 'application/x-apple-diskimage',
  '.exe': 'application/vnd.microsoft.portable-executable',
  '.appimage': 'application/octet-stream',
  '.blockmap': 'application/octet-stream',
}

// ─── Config ────────────────────────────────────────────────────

function loadDotEnvDeploy() {
  const file = path.join(ROOT, '.env.deploy')
  if (!fs.existsSync(file)) return
  for (const line of fs.readFileSync(file, 'utf8').split('\n')) {
    const trimmed = line.trim()
    if (!trimmed || trimmed.startsWith('#')) continue
    const eq = trimmed.indexOf('=')
    if (eq === -1) continue
    const key = trimmed.slice(0, eq).trim()
    const value = trimmed.slice(eq + 1).trim().replace(/^['"]|['"]$/g, '')
    if (!process.env[key]) process.env[key] = value
  }
}

function resolveTarget() {
  loadDotEnvDeploy()
  const container = process.env.AZURE_BLOB_CONTAINER || 'nexoffice'
  const prefix = (process.env.AZURE_BLOB_PREFIX || '').replace(/^\/+|\/+$/g, '')

  const sasUrl = process.env.AZURE_STORAGE_SAS_URL
  if (sasUrl) {
    const u = new URL(sasUrl)
    const sasContainer = u.pathname.replace(/^\/+|\/+$/g, '') || container
    return {
      mode: 'sas',
      host: u.host,
      container: sasContainer,
      prefix,
      sas: u.search.replace(/^\?/, ''),
    }
  }

  const cs = process.env.AZURE_STORAGE_CONNECTION_STRING
  if (!cs) {
    fail(
      'Missing credentials. Set AZURE_STORAGE_SAS_URL or AZURE_STORAGE_CONNECTION_STRING\n' +
        '(env var, or a .env.deploy file at the repo root).',
    )
  }
  const parts = Object.fromEntries(
    cs
      .split(';')
      .filter(Boolean)
      .map((kv) => {
        const i = kv.indexOf('=')
        return [kv.slice(0, i).trim(), kv.slice(i + 1).trim()]
      }),
  )
  if (!parts.AccountName || !parts.AccountKey) fail('Connection string is missing AccountName/AccountKey.')
  return {
    mode: 'key',
    account: parts.AccountName,
    key: Buffer.from(parts.AccountKey, 'base64'),
    host: `${parts.AccountName}.blob.${parts.EndpointSuffix || 'core.windows.net'}`,
    container,
    prefix,
  }
}

// ─── Azure REST (Shared Key signing, zero deps) ────────────────

function signedHeaders(target, { method, blobPath, query, headers, contentLength }) {
  const base = { 'x-ms-date': new Date().toUTCString(), 'x-ms-version': API_VERSION, ...headers }
  if (target.mode === 'sas') return base

  const canonHeaders = Object.entries(base)
    .filter(([k]) => k.toLowerCase().startsWith('x-ms-'))
    .map(([k, v]) => [k.toLowerCase(), String(v).trim()])
    .sort(([a], [b]) => (a < b ? -1 : 1))
    .map(([k, v]) => `${k}:${v}\n`)
    .join('')
  const canonResource =
    `/${target.account}${blobPath}` +
    Object.keys(query)
      .sort()
      .map((k) => `\n${k.toLowerCase()}:${query[k]}`)
      .join('')
  const stringToSign = [
    method,
    '', // Content-Encoding
    '', // Content-Language
    contentLength ? String(contentLength) : '',
    '', // Content-MD5
    base['Content-Type'] || '',
    '', // Date (x-ms-date used instead)
    '', // If-Modified-Since
    '', // If-Match
    '', // If-None-Match
    '', // If-Unmodified-Since
    '', // Range
    canonHeaders + canonResource,
  ].join('\n')
  const signature = crypto.createHmac('sha256', target.key).update(stringToSign, 'utf8').digest('base64')
  return { ...base, Authorization: `SharedKey ${target.account}:${signature}` }
}

async function azure(target, { method, blobPath, query = {}, headers = {}, body }) {
  const contentLength = body ? body.byteLength : 0
  const search = new URLSearchParams(query)
  if (target.mode === 'sas') for (const [k, v] of new URLSearchParams(target.sas)) search.append(k, v)
  const qs = search.toString()
  const url = `https://${target.host}${blobPath}${qs ? `?${qs}` : ''}`
  const res = await fetch(url, {
    method,
    headers: signedHeaders(target, { method, blobPath, query, headers, contentLength }),
    body,
  })
  if (!res.ok) {
    const text = await res.text().catch(() => '')
    const code = /<Code>([^<]+)<\/Code>/.exec(text)?.[1] ?? ''
    throw new Error(`${method} ${blobPath} → ${res.status} ${res.statusText} ${code}`.trim())
  }
  await res.arrayBuffer().catch(() => undefined)
  return res
}

async function withRetry(label, fn, attempts = 4) {
  let lastError
  for (let i = 1; i <= attempts; i += 1) {
    try {
      return await fn()
    } catch (err) {
      lastError = err
      if (i === attempts) break
      const wait = 500 * 2 ** (i - 1)
      console.warn(`     ↻ ${label} failed (${err.message}); retry ${i}/${attempts - 1} in ${wait}ms`)
      await new Promise((r) => setTimeout(r, wait))
    }
  }
  throw lastError
}

// ─── Upload ────────────────────────────────────────────────────

function blobHeaders(name) {
  const ext = path.extname(name).toLowerCase()
  const isFeed = /\.ya?ml$/i.test(name)
  return {
    contentType: CONTENT_TYPES[ext] ?? 'application/octet-stream',
    // Feed manifests are mutable, so they must never be cached; versioned
    // binaries never change, so they can be cached forever.
    cacheControl: isFeed ? 'no-cache, max-age=0, must-revalidate' : 'public, max-age=31536000, immutable',
  }
}

async function uploadSmall(target, blobPath, name, buffer) {
  const { contentType, cacheControl } = blobHeaders(name)
  await withRetry(name, () =>
    azure(target, {
      method: 'PUT',
      blobPath,
      headers: {
        'Content-Type': contentType,
        'x-ms-blob-type': 'BlockBlob',
        'x-ms-blob-cache-control': cacheControl,
      },
      body: buffer,
    }),
  )
}

async function uploadBlocks(target, blobPath, name, filePath, size) {
  const { contentType, cacheControl } = blobHeaders(name)
  const count = Math.ceil(size / BLOCK_SIZE)
  const blockIds = Array.from({ length: count }, (_, i) =>
    Buffer.from(String(i).padStart(6, '0')).toString('base64'),
  )
  const handle = await fsp.open(filePath, 'r')
  let done = 0
  try {
    let next = 0
    const worker = async () => {
      for (let i = next++; i < count; i = next++) {
        const start = i * BLOCK_SIZE
        const length = Math.min(BLOCK_SIZE, size - start)
        const buffer = Buffer.allocUnsafe(length)
        await handle.read(buffer, 0, length, start)
        await withRetry(`${name} block ${i + 1}/${count}`, () =>
          azure(target, {
            method: 'PUT',
            blobPath,
            query: { comp: 'block', blockid: blockIds[i] },
            body: buffer,
          }),
        )
        done += length
        process.stdout.write(`\r     ${((done / size) * 100).toFixed(0)}% (${count} blocks)   `)
      }
    }
    await Promise.all(Array.from({ length: Math.min(BLOCK_CONCURRENCY, count) }, worker))
    process.stdout.write('\r')
  } finally {
    await handle.close()
  }

  const xml =
    `<?xml version="1.0" encoding="utf-8"?><BlockList>` +
    blockIds.map((id) => `<Latest>${id}</Latest>`).join('') +
    `</BlockList>`
  await withRetry(`${name} commit`, () =>
    azure(target, {
      method: 'PUT',
      blobPath,
      query: { comp: 'blocklist' },
      headers: {
        'Content-Type': 'application/xml',
        'x-ms-blob-content-type': contentType,
        'x-ms-blob-cache-control': cacheControl,
      },
      body: Buffer.from(xml, 'utf8'),
    }),
  )
}

async function uploadFile(target, filePath, blobName, dryRun) {
  const size = (await fsp.stat(filePath)).size
  const mb = (size / 1024 / 1024).toFixed(1)
  const blobPath = `/${target.container}/${target.prefix ? `${target.prefix}/` : ''}${blobName}`
  if (dryRun) {
    console.log(`  · would upload ${blobName} (${mb} MB) → ${blobPath}`)
    return blobPath
  }
  console.log(`  📤 ${blobName} (${mb} MB)`)
  if (size <= SINGLE_PUT_LIMIT) {
    await uploadSmall(target, blobPath, blobName, await fsp.readFile(filePath))
  } else {
    await uploadBlocks(target, blobPath, blobName, filePath, size)
  }
  return blobPath
}

// ─── Release artifacts ─────────────────────────────────────────

function parseFeed(text) {
  const version = /^version:\s*(.+)$/m.exec(text)?.[1]?.trim()
  const urls = [...text.matchAll(/^\s*-\s*url:\s*(.+)$/gm)].map((m) => m[1].trim().replace(/^['"]|['"]$/g, ''))
  return { version, urls }
}

/**
 * electron-builder writes feed URLs with spaces replaced by `-`, while the file
 * on disk keeps the productName spacing until `artifactName` is customised.
 */
function resolveLocalFile(url) {
  const direct = path.join(RELEASE_DIR, url)
  if (fs.existsSync(direct)) return direct
  const decoded = path.join(RELEASE_DIR, decodeURIComponent(url))
  if (fs.existsSync(decoded)) return decoded
  const match = fs.readdirSync(RELEASE_DIR).find((f) => f.replace(/ /g, '-') === url)
  return match ? path.join(RELEASE_DIR, match) : null
}

function collect(platform) {
  const feedName = FEEDS[platform]
  const feedPath = path.join(RELEASE_DIR, feedName)
  if (!fs.existsSync(feedPath)) return null

  const { version, urls } = parseFeed(fs.readFileSync(feedPath, 'utf8'))
  const files = []
  const missing = []
  for (const url of urls) {
    const local = resolveLocalFile(url)
    if (!local) {
      missing.push(url)
      continue
    }
    files.push({ local, blobName: url })
    const blockmap = resolveLocalFile(`${url}.blockmap`)
    if (blockmap) files.push({ local: blockmap, blobName: `${url}.blockmap` })
  }
  return { platform, version, feedName, feedPath, files, missing }
}

// ─── Main ──────────────────────────────────────────────────────

function fail(message) {
  console.error(`❌ ${message}`)
  process.exit(1)
}

async function main() {
  const args = process.argv.slice(2)
  const dryRun = args.includes('--dry-run')
  const requested = args.find((a) => !a.startsWith('--')) ?? 'all'
  const platforms = requested === 'all' ? Object.keys(FEEDS) : [requested]
  for (const p of platforms) if (!FEEDS[p]) fail(`Unknown platform "${p}" (use mac | win | linux | all).`)

  if (!fs.existsSync(RELEASE_DIR)) fail('No release/ directory — run `bun run dist:mac` (or :win/:linux) first.')

  const target = resolveTarget()
  const publicBase = `https://${target.host}/${target.container}${target.prefix ? `/${target.prefix}` : ''}`
  console.log(`🎯 ${publicBase}  (auth: ${target.mode === 'sas' ? 'SAS token' : 'account key'})`)

  const bundles = platforms.map(collect).filter(Boolean)
  if (!bundles.length) {
    fail(`No feed manifest found in release/ for: ${platforms.join(', ')} (looked for ${platforms.map((p) => FEEDS[p]).join(', ')}).`)
  }

  for (const bundle of bundles) {
    console.log(`\n📦 ${bundle.platform} v${bundle.version ?? '?'} — ${bundle.files.length} artifact(s)`)
    for (const url of bundle.missing) console.warn(`  ⚠️  listed in ${bundle.feedName} but not found locally: ${url}`)

    // Binaries first…
    for (const file of bundle.files) await uploadFile(target, file.local, file.blobName, dryRun)
    // …then the manifest that points at them.
    await uploadFile(target, bundle.feedPath, bundle.feedName, dryRun)

    if (!dryRun) {
      const feedUrl = `${publicBase}/${bundle.feedName}`
      const res = await fetch(feedUrl, { method: 'HEAD', cache: 'no-store' })
      console.log(res.ok ? `  ✅ live: ${feedUrl}` : `  ⚠️  public read check failed (${res.status}): ${feedUrl}`)
      for (const file of bundle.files) {
        const head = await fetch(`${publicBase}/${encodeURI(file.blobName)}`, { method: 'HEAD' })
        if (!head.ok) console.warn(`  ⚠️  not publicly readable (${head.status}): ${file.blobName}`)
      }
    }
  }

  console.log(dryRun ? '\n🧪 Dry run — nothing uploaded.' : '\n✅ Deploy done.')
}

main().catch((err) => fail(err.stack || err.message))

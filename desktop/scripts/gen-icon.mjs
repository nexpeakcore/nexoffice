// Generate app icons from build/icon.svg:
//   build/icon.png   1024x1024, linux + fallback
//   build/icon.icns  mac (iconutil, so macOS only — skipped elsewhere)
//   build/icon.ico   win, BMP frames at 16/24/32/48/64/128/256
//
// The ICO deliberately uses uncompressed BMP frames: some Windows shell
// components cannot read PNG-compressed frames at small sizes.
import sharp from 'sharp'
import { execFileSync } from 'node:child_process'
import { mkdir, mkdtemp, rm, writeFile } from 'node:fs/promises'
import { statSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const svg = resolve(root, 'build/icon.svg')
const png = resolve(root, 'build/icon.png')
const icns = resolve(root, 'build/icon.icns')
const ico = resolve(root, 'build/icon.ico')

const renderPng = (size) => sharp(svg).resize(size, size).png().toBuffer()

async function generateIcns() {
  if (process.platform !== 'darwin') {
    console.log('icon.icns skipped (iconutil requires macOS)')
    return
  }
  const iconset = await mkdtemp(join(tmpdir(), 'nexoffice-iconset-')) + '.iconset'
  await mkdir(iconset, { recursive: true })
  try {
    for (const base of [16, 32, 128, 256, 512]) {
      await writeFile(join(iconset, `icon_${base}x${base}.png`), await renderPng(base))
      await writeFile(join(iconset, `icon_${base}x${base}@2x.png`), await renderPng(base * 2))
    }
    execFileSync('iconutil', ['-c', 'icns', '-o', icns, iconset])
    console.log(`icon.icns generated: ${icns} (${statSync(icns).size} bytes)`)
  } finally {
    await rm(iconset, { recursive: true, force: true })
  }
}

// ICO frame: BITMAPINFOHEADER + BGRA rows bottom-up + 1bpp AND mask.
function bmpFrame(size, rgba) {
  const header = Buffer.alloc(40)
  header.writeUInt32LE(40, 0)
  header.writeInt32LE(size, 4)
  header.writeInt32LE(size * 2, 8) // XOR + AND mask heights combined
  header.writeUInt16LE(1, 12)
  header.writeUInt16LE(32, 14)

  const xor = Buffer.alloc(size * size * 4)
  for (let y = 0; y < size; y += 1) {
    const src = (size - 1 - y) * size * 4
    const dst = y * size * 4
    for (let x = 0; x < size; x += 1) {
      xor[dst + x * 4] = rgba[src + x * 4 + 2]
      xor[dst + x * 4 + 1] = rgba[src + x * 4 + 1]
      xor[dst + x * 4 + 2] = rgba[src + x * 4]
      xor[dst + x * 4 + 3] = rgba[src + x * 4 + 3]
    }
  }
  // Transparency comes from the alpha channel; the AND mask stays zeroed,
  // padded to 32-bit row boundaries.
  const and = Buffer.alloc((Math.ceil(size / 32) * 4) * size)
  return Buffer.concat([header, xor, and])
}

async function generateIco() {
  const sizes = [16, 24, 32, 48, 64, 128, 256]
  const frames = []
  for (const size of sizes) {
    const rgba = await sharp(svg)
      .resize(size, size)
      .ensureAlpha()
      .raw()
      .toBuffer()
    frames.push({ size, data: bmpFrame(size, rgba) })
  }

  const dir = Buffer.alloc(6)
  dir.writeUInt16LE(1, 2) // type: icon
  dir.writeUInt16LE(frames.length, 4)

  const entries = []
  let offset = 6 + frames.length * 16
  for (const { size, data } of frames) {
    const entry = Buffer.alloc(16)
    entry.writeUInt8(size === 256 ? 0 : size, 0)
    entry.writeUInt8(size === 256 ? 0 : size, 1)
    entry.writeUInt16LE(1, 4) // planes
    entry.writeUInt16LE(32, 6) // bit depth
    entry.writeUInt32LE(data.length, 8)
    entry.writeUInt32LE(offset, 12)
    entries.push(entry)
    offset += data.length
  }

  await writeFile(ico, Buffer.concat([dir, ...entries, ...frames.map((f) => f.data)]))
  console.log(`icon.ico generated: ${ico} (${sizes.join('/')}px, ${statSync(ico).size} bytes)`)
}

await mkdir(dirname(png), { recursive: true })
await writeFile(png, await renderPng(1024))
console.log('icon.png generated:', png)
await generateIcns()
await generateIco()

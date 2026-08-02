// Generate app icons from build/icon.svg → build/icon.png (1024) + icon.icns + icon.ico
import sharp from 'sharp'
import { writeFile, mkdir } from 'node:fs/promises'
import { resolve, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const svg = resolve(root, 'build/icon.svg')
const png = resolve(root, 'build/icon.png')

// electron-builder accepts a single 1024x1024 icon.png in buildResources
// and derives .icns (mac) / .ico (win) from it automatically.
await mkdir(dirname(png), { recursive: true })
await sharp(svg).resize(1024, 1024).png().toFile(png)
console.log('icon.png generated:', png)

import { readFileSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import sharp from 'sharp'

// Generates the add-in + store icons from Lucide's square-function ("fx") glyph.
// Lucide is ISC-licensed (free, commercial use OK): https://lucide.dev/license
// Run: node scripts/makeIcons.mjs   (from packages/addin)

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const lucide = readFileSync(
  join(root, '..', '..', 'node_modules', 'lucide-static', 'icons', 'square-function.svg'),
  'utf8'
)

// Pull the glyph shapes out of the lucide wrapper; render white on the CCG accent tile.
const shapes = lucide
  .slice(lucide.indexOf('>') + 1)
  .replace('</svg>', '')
  .replace(/<svg[\s\S]*?>/, '')

const composed = `<svg xmlns="http://www.w3.org/2000/svg" width="240" height="240" viewBox="0 0 30 30">
  <rect width="30" height="30" rx="6" fill="#1a56db"/>
  <g transform="translate(3 3)" fill="none" stroke="#ffffff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
    ${shapes}
  </g>
</svg>`

const targets = [
  { size: 16, file: 'assets/icon-16.png' },
  { size: 32, file: 'assets/icon-32.png' },
  { size: 64, file: 'assets/icon-64.png' },
  { size: 80, file: 'assets/icon-80.png' },
  { size: 300, file: 'assets/store-icon-300.png' },
]

for (const { size, file } of targets) {
  await sharp(Buffer.from(composed), { density: 300 }).resize(size, size).png().toFile(join(root, file))
  console.log(`[makeIcons] ${file} (${size}x${size})`)
}

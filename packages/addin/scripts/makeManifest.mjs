import { readFileSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

// Generates the production manifest from the dev one: swaps localhost URLs for the
// hosted base URL and gives prod its own Id so dev + prod can be installed side by side.
// Usage: node makeManifest.mjs https://aceslick911.github.io/CCG

const DEV_ID = '7c1b4a9e-8f2d-4e6b-a3c5-91d0f2b7e844'
const PROD_ID = '9d2e7b6a-4c31-4f5e-8a27-63b90d14f7c2'

const base = process.argv[2]?.replace(/\/$/, '')
if (!base?.startsWith('https://')) {
  console.error('[makeManifest] Usage: node makeManifest.mjs <https base url>')
  process.exit(1)
}

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const manifest = readFileSync(join(root, 'manifest.xml'), 'utf8')
  .replaceAll('https://localhost:3000', base)
  .replace(DEV_ID, PROD_ID)

writeFileSync(join(root, 'dist', 'manifest.xml'), manifest)
console.log(`[makeManifest] dist/manifest.xml written for ${base}`)

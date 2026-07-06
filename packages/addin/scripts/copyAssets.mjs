import { cpSync, mkdirSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const dist = join(root, 'dist')

mkdirSync(dist, { recursive: true })
cpSync(join(root, 'src', 'taskpane.html'), join(dist, 'taskpane.html'))
cpSync(join(root, 'src', 'privacy.html'), join(dist, 'privacy.html'))
cpSync(join(root, 'src', 'terms.html'), join(dist, 'terms.html'))
cpSync(join(root, 'src', 'flipbox.css'), join(dist, 'flipbox.css'))
cpSync(join(root, 'src', 'ccg.css'), join(dist, 'ccg.css'))
cpSync(join(root, 'assets'), join(dist, 'assets'), { recursive: true })
console.log('[copyAssets] dist refreshed')

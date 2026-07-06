import type { Dialect } from './expression.js'
import type { CellValue } from '../workbook.js'

export const sanitizeKeyName = (key: string): string => {
  const cleaned = key
    .replace(/[^A-Za-z0-9_]/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_|_$/g, '')
  return /^[0-9]/.test(cleaned) ? `N${cleaned}` : cleaned
}

export const literalFor = (value: CellValue | undefined, dialect: Dialect): string => {
  if (value === undefined) {
    return '0'
  }
  if (typeof value === 'number') {
    return String(value)
  }
  if (typeof value === 'boolean') {
    return dialect.boolean(value)
  }
  return dialect.string(value)
}

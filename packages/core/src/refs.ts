import type { CellRef } from './ast.js'

/** 'A' → 1, 'Z' → 26, 'AA' → 27 … (1-based, like Excel) */
export const colToIndex = (col: string): number => {
  let index = 0
  for (const ch of col.toUpperCase()) {
    index = index * 26 + (ch.charCodeAt(0) - 64)
  }
  return index
}

/** 1 → 'A', 27 → 'AA' … */
export const indexToCol = (index: number): string => {
  let n = index
  let col = ''
  while (n > 0) {
    const rem = (n - 1) % 26
    col = String.fromCharCode(65 + rem) + col
    n = Math.floor((n - 1) / 26)
  }
  return col
}

const CELL_PATTERN = /^(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)$/

/** Max column is XFD — anything "later" is an identifier, not a ref. */
const MAX_COL_INDEX = colToIndex('XFD')

export const tryParseCellRef = (text: string, sheet?: string): CellRef | undefined => {
  const match = CELL_PATTERN.exec(text)
  if (!match) {
    return undefined
  }
  const [, colAbs, col, rowAbs, rowText] = match
  if (col === undefined || rowText === undefined) {
    return undefined
  }
  if (colToIndex(col) > MAX_COL_INDEX) {
    return undefined
  }
  const row = Number.parseInt(rowText, 10)
  if (row < 1 || row > 1048576) {
    return undefined
  }
  return { sheet, col: col.toUpperCase(), row, colAbs: colAbs === '$', rowAbs: rowAbs === '$' }
}

/**
 * Canonical key for a cell: uppercase sheet + ref, no dollar signs — 'SHEET1!E7'.
 * The workbook model and classification maps are keyed by this.
 */
export const cellKey = (sheet: string, ref: CellRef | string): string => {
  if (typeof ref === 'string') {
    return `${sheet.toUpperCase()}!${ref.replace(/\$/g, '').toUpperCase()}`
  }
  const sheetName = ref.sheet ?? sheet
  return `${sheetName.toUpperCase()}!${ref.col}${ref.row}`
}

/** Every cell key covered by a range (small ranges only — guard against whole-column refs). */
export const expandRangeKeys = (start: CellRef, end: CellRef, defaultSheet: string, limit = 1000): string[] => {
  const sheet = start.sheet ?? defaultSheet
  const fromCol = colToIndex(start.col)
  const toCol = colToIndex(end.col)
  const fromRow = Math.min(start.row, end.row)
  const toRow = Math.max(start.row, end.row)
  const lo = Math.min(fromCol, toCol)
  const hi = Math.max(fromCol, toCol)
  const count = (hi - lo + 1) * (toRow - fromRow + 1)
  if (count > limit) {
    throw new Error(
      `[refs] Range ${start.col}${start.row}:${end.col}${end.row} expands to ${count} cells (limit ${limit})`
    )
  }
  const keys: string[] = []
  for (let row = fromRow; row <= toRow; row++) {
    for (let col = lo; col <= hi; col++) {
      keys.push(cellKey(sheet, `${indexToCol(col)}${row}`))
    }
  }
  return keys
}

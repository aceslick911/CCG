/**
 * The CCG formula AST. Discriminated union — the modern successor to the 2013 PNode tree.
 * Nodes are plain data: no live workbook handles, no positions into mutated strings.
 */

export type BinaryOp = '+' | '-' | '*' | '/' | '^' | '&' | '=' | '<>' | '<' | '>' | '<=' | '>='

export type CellRef = {
  sheet?: string
  col: string
  row: number
  colAbs: boolean
  rowAbs: boolean
}

export type FormulaNode =
  | { kind: 'number'; value: number; text: string }
  | { kind: 'string'; value: string }
  | { kind: 'boolean'; value: boolean }
  | { kind: 'error'; code: string }
  | { kind: 'cell'; ref: CellRef }
  | { kind: 'range'; start: CellRef; end: CellRef }
  | { kind: 'name'; name: string; sheet?: string }
  | { kind: 'call'; name: string; args: FormulaNode[] }
  | { kind: 'binary'; op: BinaryOp; left: FormulaNode; right: FormulaNode }
  | { kind: 'unary'; op: '+' | '-'; operand: FormulaNode }
  | { kind: 'percent'; operand: FormulaNode }
  | { kind: 'group'; inner: FormulaNode }
  | { kind: 'structured'; table?: string; column: string; thisRow: boolean }
  | { kind: 'array'; rows: FormulaNode[][] }
  | { kind: 'empty' }

export const walk = (node: FormulaNode, visit: (node: FormulaNode) => void): void => {
  visit(node)
  switch (node.kind) {
    case 'call': {
      for (const arg of node.args) {
        walk(arg, visit)
      }
      break
    }
    case 'binary': {
      walk(node.left, visit)
      walk(node.right, visit)
      break
    }
    case 'unary':
    case 'percent': {
      walk(node.operand, visit)
      break
    }
    case 'group': {
      walk(node.inner, visit)
      break
    }
    case 'array': {
      for (const row of node.rows) {
        for (const item of row) {
          walk(item, visit)
        }
      }
      break
    }
    default:
      break
  }
}

/** Bottom-up structural rewrite — children first, then the rebuilt node is passed to fn. */
export const transform = (node: FormulaNode, fn: (node: FormulaNode) => FormulaNode): FormulaNode => {
  let rebuilt: FormulaNode
  switch (node.kind) {
    case 'call':
      rebuilt = { ...node, args: node.args.map((arg) => transform(arg, fn)) }
      break
    case 'binary':
      rebuilt = { ...node, left: transform(node.left, fn), right: transform(node.right, fn) }
      break
    case 'unary':
    case 'percent':
      rebuilt = { ...node, operand: transform(node.operand, fn) }
      break
    case 'group':
      rebuilt = { ...node, inner: transform(node.inner, fn) }
      break
    case 'array':
      rebuilt = { ...node, rows: node.rows.map((row) => row.map((item) => transform(item, fn))) }
      break
    default:
      rebuilt = node
      break
  }
  return fn(rebuilt)
}

/** All cell and range references in a tree, in source order. */
export const collectRefs = (node: FormulaNode): Array<Extract<FormulaNode, { kind: 'cell' | 'range' }>> => {
  const refs: Array<Extract<FormulaNode, { kind: 'cell' | 'range' }>> = []
  walk(node, (n) => {
    if (n.kind === 'cell' || n.kind === 'range') {
      refs.push(n)
    }
  })
  return refs
}

/** Prefix-notation dump, ported from the 2013 Flatten.PreFixNotation — handy for debugging and the visualizer. */
export const toPrefixNotation = (node: FormulaNode): string => {
  switch (node.kind) {
    case 'number':
      return node.text
    case 'string':
      return `"${node.value}"`
    case 'boolean':
      return node.value ? 'TRUE' : 'FALSE'
    case 'error':
      return node.code
    case 'cell':
      return formatCellRefText(node.ref)
    case 'range':
      return `${formatCellRefText(node.start)}:${formatCellRefText(node.end)}`
    case 'name':
      return node.sheet ? `${node.sheet}!${node.name}` : node.name
    case 'call':
      return `${node.name}(${node.args.map(toPrefixNotation).join(', ')})`
    case 'binary':
      return `${node.op}(${toPrefixNotation(node.left)}, ${toPrefixNotation(node.right)})`
    case 'unary':
      return `${node.op}(${toPrefixNotation(node.operand)})`
    case 'percent':
      return `%(${toPrefixNotation(node.operand)})`
    case 'group':
      return toPrefixNotation(node.inner)
    case 'structured':
      return `${node.table ?? ''}[${node.thisRow ? '@' : ''}${node.column}]`
    case 'array':
      return `{${node.rows.map((row) => row.map(toPrefixNotation).join(', ')).join('; ')}}`
    case 'empty':
      return ''
    default: {
      const exhaustive: never = node
      return exhaustive
    }
  }
}

export const formatCellRefText = (ref: CellRef): string => {
  const sheet = ref.sheet ? `${quoteSheetIfNeeded(ref.sheet)}!` : ''
  return `${sheet}${ref.colAbs ? '$' : ''}${ref.col}${ref.rowAbs ? '$' : ''}${ref.row}`
}

export const quoteSheetIfNeeded = (sheet: string): string => {
  if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(sheet)) {
    return sheet
  }
  return `'${sheet.replace(/'/g, "''")}'`
}

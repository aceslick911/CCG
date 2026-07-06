import type { BinaryOp, FormulaNode } from './ast.js'
import { tryParseCellRef } from './refs.js'
import { type Token, tokenize } from './tokenizer.js'

/**
 * Pratt parser with Excel's operator precedence, including the quirks:
 * unary minus binds tighter than ^ (so -2^2 = 4), and ^ is left-associative.
 * Not supported (yet): intersection (space) and union (comma) reference operators.
 */

const BINARY_BINDING: Record<string, number> = {
  '=': 10,
  '<>': 10,
  '<': 10,
  '>': 10,
  '<=': 10,
  '>=': 10,
  '&': 20,
  '+': 30,
  '-': 30,
  '*': 40,
  '/': 40,
  '^': 50,
}

export class ParseError extends Error {}

type Cursor = {
  peek: () => Token
  peekAhead: (offset: number) => Token
  next: () => Token
}

const makeCursor = (tokens: Token[]): Cursor => {
  let pos = 0
  const end: Token = { type: 'end' }
  return {
    peek: () => tokens[pos] ?? end,
    peekAhead: (offset: number) => tokens[pos + offset] ?? end,
    next: () => {
      const token = tokens[pos] ?? end
      pos += 1
      return token
    },
  }
}

export const parseFormula = (input: string): FormulaNode => {
  const text = input.trim().startsWith('=') ? input.trim().slice(1) : input.trim()
  const cursor = makeCursor(tokenize(text))
  const node = parseExpr(cursor, 0)
  const trailing = cursor.peek()
  if (trailing.type !== 'end') {
    throw new ParseError(`[parser] Unexpected trailing input near ${describe(trailing)} in: ${input}`)
  }
  return node
}

const describe = (token: Token): string => {
  switch (token.type) {
    case 'end':
      return 'end of formula'
    case 'number':
      return `number ${token.text}`
    case 'string':
      return `string "${token.value}"`
    case 'errorLit':
      return token.code
    case 'bracket':
      return `[${token.text}]`
    case 'quotedSheet':
      return `'${token.text}'`
    default:
      return `'${token.text}'`
  }
}

const parseExpr = (cursor: Cursor, minBinding: number): FormulaNode => {
  let left = parseUnary(cursor)

  for (;;) {
    const token = cursor.peek()
    if (token.type === 'op' && token.text === '%') {
      cursor.next()
      left = { kind: 'percent', operand: left }
      continue
    }
    if (token.type !== 'op') {
      break
    }
    const binding = BINARY_BINDING[token.text]
    if (binding === undefined || binding <= minBinding) {
      break
    }
    cursor.next()
    const right = parseExpr(cursor, binding)
    left = { kind: 'binary', op: token.text as BinaryOp, left, right }
  }

  return left
}

const parseUnary = (cursor: Cursor): FormulaNode => {
  const token = cursor.peek()
  if (token.type === 'op' && (token.text === '+' || token.text === '-')) {
    cursor.next()
    const operand = parseUnary(cursor)
    return { kind: 'unary', op: token.text, operand }
  }
  return parsePostfix(cursor)
}

const parsePostfix = (cursor: Cursor): FormulaNode => {
  let node = parsePrimary(cursor)
  for (;;) {
    const token = cursor.peek()
    if (token.type === 'op' && token.text === '%') {
      cursor.next()
      node = { kind: 'percent', operand: node }
      continue
    }
    break
  }
  return node
}

const parsePrimary = (cursor: Cursor): FormulaNode => {
  const token = cursor.next()

  switch (token.type) {
    case 'number':
      return { kind: 'number', value: token.value, text: token.text }
    case 'string':
      return { kind: 'string', value: token.value }
    case 'errorLit':
      return { kind: 'error', code: token.code }
    case 'bracket':
      return parseStructuredBody(token.text, undefined)
    case 'punct': {
      if (token.text === '(') {
        const inner = parseExpr(cursor, 0)
        expectPunct(cursor, ')')
        return { kind: 'group', inner }
      }
      if (token.text === '{') {
        return parseArrayLiteral(cursor)
      }
      throw new ParseError(`[parser] Unexpected ${describe(token)}`)
    }
    case 'quotedSheet':
      return parseSheetQualified(cursor, token.text)
    case 'word':
      return parseWord(cursor, token.text)
    default:
      throw new ParseError(`[parser] Unexpected ${describe(token)}`)
  }
}

const parseArrayLiteral = (cursor: Cursor): FormulaNode => {
  const rows: FormulaNode[][] = [[]]
  for (;;) {
    const row = rows[rows.length - 1]
    if (row === undefined) {
      break
    }
    row.push(parseExpr(cursor, 0))
    const token = cursor.next()
    if (token.type === 'punct' && token.text === ',') {
      continue
    }
    if (token.type === 'punct' && token.text === ';') {
      rows.push([])
      continue
    }
    if (token.type === 'punct' && token.text === '}') {
      return { kind: 'array', rows }
    }
    throw new ParseError(`[parser] Expected ',', ';' or '}' in array literal, got ${describe(token)}`)
  }
  throw new ParseError('[parser] Invalid array literal')
}

const parseSheetQualified = (cursor: Cursor, sheet: string): FormulaNode => {
  expectPunct(cursor, '!')
  const token = cursor.next()
  if (token.type !== 'word') {
    throw new ParseError(`[parser] Expected reference after '${sheet}'!, got ${describe(token)}`)
  }
  return resolveWord(cursor, token.text, sheet)
}

const parseWord = (cursor: Cursor, text: string): FormulaNode => {
  const token = cursor.peek()
  if (token.type === 'punct' && token.text === '!') {
    cursor.next()
    const refToken = cursor.next()
    if (refToken.type !== 'word') {
      throw new ParseError(`[parser] Expected reference after '${text}!', got ${describe(refToken)}`)
    }
    return resolveWord(cursor, refToken.text, text)
  }
  return resolveWord(cursor, text, undefined)
}

const resolveWord = (cursor: Cursor, text: string, sheet: string | undefined): FormulaNode => {
  const next = cursor.peek()

  // WORD( → always a function call, even if the word looks like a cell ref (LOG10(...)).
  if (next.type === 'punct' && next.text === '(') {
    cursor.next()
    return parseCallArgs(cursor, text.toUpperCase())
  }

  // Table1[Column] structured reference.
  if (next.type === 'bracket') {
    cursor.next()
    return parseStructuredBody(next.text, text)
  }

  const upper = text.toUpperCase()
  if (upper === 'TRUE' || upper === 'FALSE') {
    return { kind: 'boolean', value: upper === 'TRUE' }
  }

  const cell = tryParseCellRef(text, sheet)
  if (cell) {
    // A1:B2 range — the ':' only means range when both sides parse as cell refs.
    if (next.type === 'punct' && next.text === ':') {
      const after = cursor.peekAhead(1)
      if (after.type === 'word') {
        const endRef = tryParseCellRef(after.text)
        if (endRef) {
          cursor.next()
          cursor.next()
          return { kind: 'range', start: cell, end: endRef }
        }
      }
    }
    return { kind: 'cell', ref: cell }
  }

  return { kind: 'name', name: text, sheet }
}

const parseCallArgs = (cursor: Cursor, name: string): FormulaNode => {
  const args: FormulaNode[] = []
  const first = cursor.peek()
  if (first.type === 'punct' && first.text === ')') {
    cursor.next()
    return { kind: 'call', name, args }
  }

  for (;;) {
    const token = cursor.peek()
    if (token.type === 'punct' && (token.text === ',' || token.text === ')')) {
      args.push({ kind: 'empty' })
    } else {
      args.push(parseExpr(cursor, 0))
    }
    const separator = cursor.next()
    if (separator.type === 'punct' && separator.text === ',') {
      continue
    }
    if (separator.type === 'punct' && separator.text === ')') {
      return { kind: 'call', name, args }
    }
    throw new ParseError(`[parser] Expected ',' or ')' in ${name}(...), got ${describe(separator)}`)
  }
}

const parseStructuredBody = (body: string, table: string | undefined): FormulaNode => {
  let column = body.trim()
  let thisRow = false
  if (column.startsWith('@')) {
    thisRow = true
    column = column.slice(1)
  }
  if (column.startsWith('[') && column.endsWith(']')) {
    column = column.slice(1, -1)
  }
  return { kind: 'structured', table, column, thisRow }
}

const expectPunct = (cursor: Cursor, text: string): void => {
  const token = cursor.next()
  if (token.type !== 'punct' || token.text !== text) {
    throw new ParseError(`[parser] Expected '${text}', got ${describe(token)}`)
  }
}

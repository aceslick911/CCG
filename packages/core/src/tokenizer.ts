/**
 * Excel formula tokenizer. Replaces the 2013 regex-splitting approach with a single
 * left-to-right scan — string literals, quoted sheet names, and error literals can no
 * longer confuse operator detection.
 */

export type Token =
  | { type: 'number'; value: number; text: string }
  | { type: 'string'; value: string }
  | { type: 'word'; text: string }
  | { type: 'quotedSheet'; text: string }
  | { type: 'errorLit'; code: string }
  | { type: 'op'; text: string }
  | { type: 'punct'; text: '(' | ')' | ',' | ';' | '{' | '}' | '!' | ':' }
  | { type: 'bracket'; text: string }
  | { type: 'end' }

const ERROR_LITERALS = [
  '#GETTING_DATA',
  '#DIV/0!',
  '#VALUE!',
  '#SPILL!',
  '#NULL!',
  '#CALC!',
  '#NAME?',
  '#NUM!',
  '#REF!',
  '#N/A',
]

const TWO_CHAR_OPS = ['<=', '>=', '<>']
const ONE_CHAR_OPS = ['+', '-', '*', '/', '^', '&', '=', '<', '>', '%']

const NUMBER_PATTERN = /^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/
const WORD_PATTERN = /^[$A-Za-z_][$A-Za-z0-9_.]*/

export const tokenize = (input: string): Token[] => {
  const tokens: Token[] = []
  let rest = input

  const take = (count: number): void => {
    rest = rest.slice(count)
  }

  while (rest.length > 0) {
    const ch = rest[0]
    if (ch === undefined) {
      break
    }

    if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r') {
      take(1)
      continue
    }

    if (ch === '"') {
      let value = ''
      let i = 1
      while (i < rest.length) {
        if (rest[i] === '"') {
          if (rest[i + 1] === '"') {
            value += '"'
            i += 2
            continue
          }
          break
        }
        value += rest[i]
        i += 1
      }
      if (i >= rest.length) {
        throw new Error(`[tokenizer] Unterminated string literal in: ${input}`)
      }
      tokens.push({ type: 'string', value })
      take(i + 1)
      continue
    }

    if (ch === "'") {
      let value = ''
      let i = 1
      while (i < rest.length) {
        if (rest[i] === "'") {
          if (rest[i + 1] === "'") {
            value += "'"
            i += 2
            continue
          }
          break
        }
        value += rest[i]
        i += 1
      }
      if (i >= rest.length) {
        throw new Error(`[tokenizer] Unterminated quoted sheet name in: ${input}`)
      }
      tokens.push({ type: 'quotedSheet', text: value })
      take(i + 1)
      continue
    }

    if (ch === '#') {
      const upper = rest.toUpperCase()
      const errorLit = ERROR_LITERALS.find((code) => upper.startsWith(code))
      if (errorLit) {
        tokens.push({ type: 'errorLit', code: errorLit })
        take(errorLit.length)
        continue
      }
      throw new Error(`[tokenizer] Unknown error literal at: ${rest.slice(0, 12)}`)
    }

    if (ch === '[') {
      // Structured-reference body: capture raw text to the matching bracket, nesting-aware
      // (handles [@[Column With Spaces]]). The parser interprets the contents.
      let depth = 0
      let i = 0
      while (i < rest.length) {
        if (rest[i] === '[') {
          depth += 1
        } else if (rest[i] === ']') {
          depth -= 1
          if (depth === 0) {
            break
          }
        }
        i += 1
      }
      if (depth !== 0) {
        throw new Error(`[tokenizer] Unbalanced structured reference in: ${input}`)
      }
      tokens.push({ type: 'bracket', text: rest.slice(1, i) })
      take(i + 1)
      continue
    }

    const numberMatch = NUMBER_PATTERN.exec(rest)
    if (numberMatch) {
      const text = numberMatch[0]
      tokens.push({ type: 'number', value: Number.parseFloat(text), text })
      take(text.length)
      continue
    }

    const twoOp = TWO_CHAR_OPS.find((op) => rest.startsWith(op))
    if (twoOp) {
      tokens.push({ type: 'op', text: twoOp })
      take(2)
      continue
    }

    if (ch === '(' || ch === ')' || ch === ',' || ch === ';' || ch === '{' || ch === '}' || ch === '!' || ch === ':') {
      tokens.push({ type: 'punct', text: ch })
      take(1)
      continue
    }

    if (ONE_CHAR_OPS.includes(ch)) {
      tokens.push({ type: 'op', text: ch })
      take(1)
      continue
    }

    const wordMatch = WORD_PATTERN.exec(rest)
    if (wordMatch) {
      tokens.push({ type: 'word', text: wordMatch[0] })
      take(wordMatch[0].length)
      continue
    }

    throw new Error(`[tokenizer] Unexpected character '${ch}' in: ${input}`)
  }

  tokens.push({ type: 'end' })
  return tokens
}

import assert from 'node:assert/strict'
import { test } from 'node:test'
import { parseFormula, toPrefixNotation } from '../src/index.js'

const prefix = (formula: string): string => toPrefixNotation(parseFormula(formula))

test('numbers, strings, booleans, errors', () => {
  assert.equal(prefix('=42'), '42')
  assert.equal(prefix('=1.5e3'), '1.5e3')
  assert.equal(prefix('="he said ""hi"""'), '"he said "hi""')
  assert.equal(prefix('=TRUE'), 'TRUE')
  assert.equal(prefix('=#REF!'), '#REF!')
})

test('operator precedence matches Excel', () => {
  assert.equal(prefix('=1+2*3'), '+(1, *(2, 3))')
  assert.equal(prefix('=1*2+3'), '+(*(1, 2), 3)')
  assert.equal(prefix('=2^3^2'), '^(^(2, 3), 2)')
  assert.equal(prefix('=1&2=3'), '=(&(1, 2), 3)')
  assert.equal(prefix('=1<2'), '<(1, 2)')
})

test('unary minus binds tighter than power (Excel quirk: -2^2 = 4)', () => {
  assert.equal(prefix('=-2^2'), '^(-(2), 2)')
  assert.equal(prefix('=2^-3'), '^(2, -(3))')
})

test('percent postfix', () => {
  assert.equal(prefix('=50%'), '%(50)')
  assert.equal(prefix('=A1%*2'), '*(%(A1), 2)')
})

test('cell refs, absolute refs, sheets', () => {
  assert.equal(prefix('=E7'), 'E7')
  assert.equal(prefix('=$E$7'), '$E$7')
  assert.equal(prefix('=Sheet2!$F$9'), 'Sheet2!$F$9')
  assert.equal(prefix("='My Sheet'!A1"), "'My Sheet'!A1")
})

test('ranges', () => {
  assert.equal(prefix('=SUM(A1:A3)'), 'SUM(A1:A3)')
  assert.equal(prefix('=SUM(Sheet2!A1:B2)'), 'SUM(Sheet2!A1:B2)')
})

test('function calls, including ref-like names (LOG10)', () => {
  assert.equal(prefix('=POWER(1+E8,E9)'), 'POWER(+(1, E8), E9)')
  assert.equal(prefix('=LOG10(100)'), 'LOG10(100)')
  assert.equal(prefix('=IF(A1>0,"pos","neg")'), 'IF(>(A1, 0), "pos", "neg")')
  assert.equal(prefix('=PI()'), 'PI()')
  assert.equal(prefix('=IF(A1,,2)'), 'IF(A1, , 2)')
})

test('LOG10 without parens is a cell reference (column LOG, row 10)', () => {
  assert.equal(prefix('=LOG10'), 'LOG10')
  const node = parseFormula('=LOG10')
  assert.equal(node.kind, 'cell')
})

test('structured references', () => {
  assert.equal(prefix('=[@Rate]*2'), '*([@Rate], 2)')
  assert.equal(prefix('=Table1[Amount]'), 'Table1[Amount]')
})

test('the golden formula: Future_Value', () => {
  assert.equal(prefix('=E7*POWER(1+E8,E9)'), '*(E7, POWER(+(1, E8), E9))')
})

test('groups are preserved', () => {
  assert.equal(prefix('=(1+2)*3'), '*(+(1, 2), 3)')
})

test('trailing garbage rejected', () => {
  assert.throws(() => parseFormula('=1 2'))
})

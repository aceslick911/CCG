import assert from 'node:assert/strict'
import { test } from 'node:test'
import {
  CircularDependencyError,
  type Classification,
  emitCSharpModule,
  extractFunctions,
  makeSnapshot,
} from '../src/index.js'

test('chained functions come out in dependency order with pass-through params', () => {
  const snapshot = makeSnapshot([
    { sheet: 'Sheet1', ref: 'A1', value: 10 },
    { sheet: 'Sheet1', ref: 'A2', formula: '=A1*2', value: 20 },
    { sheet: 'Sheet1', ref: 'A3', formula: '=A2+5', value: 25 },
  ])
  const classification: Classification = new Map([
    ['SHEET1!A1', { kind: 'variable' as const, name: 'Base' }],
    ['SHEET1!A2', { kind: 'function' as const, name: 'Doubled' }],
    ['SHEET1!A3', { kind: 'function' as const, name: 'Final' }],
  ])
  const result = extractFunctions(snapshot, 'SHEET1!A3', classification)

  assert.deepEqual(
    result.functions.map((fn) => fn.name),
    ['Doubled', 'Final']
  )
  assert.deepEqual(
    result.target.params.map((p) => p.name),
    ['Base']
  )

  const cs = emitCSharpModule(result)
  assert.match(cs, /public static double Doubled\(double Base\)/)
  assert.match(cs, /public static double Final\(double Base\)/)
  assert.match(cs, /return Doubled\(Base\) \+ 5;/)
})

test('circular dependencies are a hard error with the cycle path', () => {
  const snapshot = makeSnapshot([
    { sheet: 'Sheet1', ref: 'A1', formula: '=A2+1' },
    { sheet: 'Sheet1', ref: 'A2', formula: '=A1+1' },
  ])
  const classification: Classification = new Map([
    ['SHEET1!A1', { kind: 'function' as const, name: 'First' }],
    ['SHEET1!A2', { kind: 'function' as const, name: 'Second' }],
  ])
  assert.throws(() => extractFunctions(snapshot, 'SHEET1!A1', classification), CircularDependencyError)
})

test('unclassified referenced cells are auto-promoted to parameters', () => {
  const snapshot = makeSnapshot([
    { sheet: 'Sheet1', ref: 'D1', value: 7 },
    { sheet: 'Sheet1', ref: 'D2', formula: '=D1*3', value: 21 },
  ])
  const classification: Classification = new Map([['SHEET1!D2', { kind: 'function' as const, name: 'Tripled' }]])
  const result = extractFunctions(snapshot, 'SHEET1!D2', classification)
  assert.deepEqual(
    result.target.params.map((p) => p.name),
    ['Sheet1_D1']
  )
})

test('cross-sheet references resolve', () => {
  const snapshot = makeSnapshot([
    { sheet: 'Inputs', ref: 'B1', value: 3 },
    { sheet: 'Calc', ref: 'A1', formula: '=Inputs!B1^2', value: 9 },
  ])
  const classification: Classification = new Map([
    ['INPUTS!B1', { kind: 'variable' as const, name: 'Side' }],
    ['CALC!A1', { kind: 'function' as const, name: 'Area' }],
  ])
  const result = extractFunctions(snapshot, 'CALC!A1', classification)
  assert.deepEqual(
    result.target.params.map((p) => p.name),
    ['Side']
  )
  const cs = emitCSharpModule(result)
  assert.match(cs, /return Math\.Pow\(Side, 2\);/)
})

test('duplicate refs collapse to one parameter', () => {
  const snapshot = makeSnapshot([
    { sheet: 'Sheet1', ref: 'E1', value: 4 },
    { sheet: 'Sheet1', ref: 'E2', formula: '=E1*E1', value: 16 },
  ])
  const classification: Classification = new Map([
    ['SHEET1!E1', { kind: 'variable' as const, name: 'X' }],
    ['SHEET1!E2', { kind: 'function' as const, name: 'Square' }],
  ])
  const result = extractFunctions(snapshot, 'SHEET1!E2', classification)
  assert.equal(result.target.params.length, 1)
})

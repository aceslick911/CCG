import assert from 'node:assert/strict'
import { test } from 'node:test'
import {
  type Classification,
  emitCSharpModule,
  emitLambdaNames,
  emitVbaModule,
  emitDelphiModule,
  extractFunctions,
  makeSnapshot,
} from '../src/index.js'

/**
 * The golden scenario, straight from fixtures/2013/InterestSpreadsheet.xlsm:
 * E7 Present_Value (1000), E8 Interest_Rate (0.05), E9 Years (20),
 * E10 Future_Value = E7*POWER(1+E8,E9) → 2653.297705.
 */
const interestSnapshot = makeSnapshot([
  { sheet: 'Sheet1', ref: 'E7', value: 1000 },
  { sheet: 'Sheet1', ref: 'E8', value: 0.05 },
  { sheet: 'Sheet1', ref: 'E9', value: 20 },
  { sheet: 'Sheet1', ref: 'E10', formula: '=E7*POWER(1+E8,E9)', value: 2653.297705 },
])

const interestClassification: Classification = new Map([
  ['SHEET1!E7', { kind: 'variable' as const, name: 'Present_Value' }],
  ['SHEET1!E8', { kind: 'variable' as const, name: 'Interest_Rate' }],
  ['SHEET1!E9', { kind: 'variable' as const, name: 'Years' }],
  ['SHEET1!E10', { kind: 'function' as const, name: 'Future_Value' }],
])

const extractInterest = () => extractFunctions(interestSnapshot, 'SHEET1!E10', interestClassification)

test('VBA output matches the 2013 golden shape', () => {
  const vba = emitVbaModule(extractInterest())
  assert.equal(
    vba,
    [
      'Function Future_Value(Present_Value, Interest_Rate, Years)',
      '  Future_Value = Present_Value * (1 + Interest_Rate) ^ Years',
      'End Function',
    ].join('\n')
  )
})

test('C# output is typed, modern, and correct', () => {
  const cs = emitCSharpModule(extractInterest(), { withTests: true })
  assert.match(cs, /public static double Future_Value\(double Present_Value, double Interest_Rate, double Years\)/)
  assert.match(cs, /return Present_Value \* Math\.Pow\(1 \+ Interest_Rate, Years\);/)
  assert.match(cs, /public static bool Test_Future_Value\(\)/)
  assert.match(cs, /Math\.Abs\(Future_Value\(1000, 0\.05, 20\) - 2653\.297705\) < 0\.000001/)
})

test('Delphi output for old times sake', () => {
  const pas = emitDelphiModule(extractInterest())
  assert.match(pas, /function Future_Value\(Present_Value: Double; Interest_Rate: Double; Years: Double\): Double;/)
  assert.match(pas, /Result := Present_Value \* Power\(1 \+ Interest_Rate, Years\);/)
})

test('LAMBDA output — the modern VBA-injection replacement', () => {
  const [lambda] = emitLambdaNames(extractInterest())
  assert.ok(lambda)
  assert.equal(lambda.name, 'Future_Value')
  assert.equal(
    lambda.formula,
    '=LAMBDA(Present_Value, Interest_Rate, Years, Present_Value * POWER(1 + Interest_Rate, Years))'
  )
})

test('IF becomes ternary in C#, inlineIf shim in VBA', () => {
  const snapshot = makeSnapshot([
    { sheet: 'Sheet1', ref: 'A1', value: 5 },
    { sheet: 'Sheet1', ref: 'A2', formula: '=IF(A1>0,A1,0)', value: 5 },
  ])
  const classification: Classification = new Map([
    ['SHEET1!A1', { kind: 'variable' as const, name: 'Amount' }],
    ['SHEET1!A2', { kind: 'function' as const, name: 'Clamped' }],
  ])
  const result = extractFunctions(snapshot, 'SHEET1!A2', classification)

  const cs = emitCSharpModule(result)
  assert.match(cs, /return \(Amount > 0 \? Amount : 0\);/)

  const vba = emitVbaModule(result)
  assert.match(vba, /Function inlineIf\(/)
  assert.match(vba, /Clamped = inlineIf\(Amount > 0, Amount, 0\)/)
})

test('constants are emitted as consts, not parameters', () => {
  const snapshot = makeSnapshot([
    { sheet: 'Sheet1', ref: 'B1', value: 0.1 },
    { sheet: 'Sheet1', ref: 'B2', value: 200 },
    { sheet: 'Sheet1', ref: 'B3', formula: '=B2*(1-B1)', value: 180 },
  ])
  const classification: Classification = new Map([
    ['SHEET1!B1', { kind: 'constant' as const, name: 'Discount' }],
    ['SHEET1!B2', { kind: 'variable' as const, name: 'Price' }],
    ['SHEET1!B3', { kind: 'function' as const, name: 'Net_Price' }],
  ])
  const cs = emitCSharpModule(extractFunctions(snapshot, 'SHEET1!B3', classification))
  assert.match(cs, /public static double Net_Price\(double Price\)/)
  assert.match(cs, /const double Discount = 0\.1;/)
})

test('SUM over a range expands to spread parameters', () => {
  const snapshot = makeSnapshot([
    { sheet: 'Sheet1', ref: 'C1', value: 1 },
    { sheet: 'Sheet1', ref: 'C2', value: 2 },
    { sheet: 'Sheet1', ref: 'C3', value: 3 },
    { sheet: 'Sheet1', ref: 'C4', formula: '=SUM(C1:C3)', value: 6 },
  ])
  const classification: Classification = new Map([['SHEET1!C4', { kind: 'function' as const, name: 'Total' }]])
  const cs = emitCSharpModule(extractFunctions(snapshot, 'SHEET1!C4', classification))
  assert.match(cs, /public static double Total\(double Sheet1_C1, double Sheet1_C2, double Sheet1_C3\)/)
  assert.match(cs, /return \(Sheet1_C1 \+ Sheet1_C2 \+ Sheet1_C3\);/)
})

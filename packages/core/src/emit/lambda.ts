import type { Dialect, EmitContext } from './expression.js'
import { emitExpression, makeContext } from './expression.js'
import { quoteSheetIfNeeded } from '../ast.js'
import type { ExtractResult, FunctionSpec } from '../workbook.js'
import { literalFor } from './shared.js'

/**
 * LAMBDA emitter — the modern replacement for the 2013 VBA-injection feature.
 * The output is an Excel formula to register as a workbook Name: no macros, no
 * "Trust access to the VBA project" setting, works on Mac/Windows/web.
 */

export const lambdaDialect: Dialect = {
  name: 'lambda',
  string: (value) => `"${value.replace(/"/g, '""')}"`,
  boolean: (value) => (value ? 'TRUE' : 'FALSE'),
  percent: (inner) => `${inner}%`,
  opSymbol: (op) => op,
  call: (name, args) => `${name}(${args.join(', ')})`,
  fallbackRef: (key) => {
    const [sheet, ref] = key.split('!')
    if (sheet === undefined || ref === undefined) {
      return key
    }
    return `${quoteSheetIfNeeded(sheet)}!${ref}`
  },
}

/** `=LAMBDA(p1, p2, body)` — consts become LET bindings so the name stays readable. */
export const emitLambdaName = (spec: FunctionSpec, ctx: EmitContext): { name: string; formula: string } => {
  const body = emitExpression(spec.body, lambdaDialect, ctx)
  const params = spec.params.map((p) => p.name)
  const wrapped =
    spec.consts.length > 0
      ? `LET(${spec.consts.map((c) => `${c.name}, ${literalFor(c.value, lambdaDialect)}`).join(', ')}, ${body})`
      : body
  return { name: spec.name, formula: `=LAMBDA(${[...params, wrapped].join(', ')})` }
}

/** All functions in dependency order — register each as a workbook Name, earliest first. */
export const emitLambdaNames = (result: ExtractResult): Array<{ name: string; formula: string }> => {
  const ctx = makeContext(result.target.sheet, result.nameByKey, result.functions)
  return result.functions.map((spec) => emitLambdaName(spec, ctx))
}

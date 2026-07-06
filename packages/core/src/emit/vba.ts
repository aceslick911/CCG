import type { Dialect, EmitContext } from './expression.js'
import { emitExpression, makeContext } from './expression.js'
import type { FormulaNode } from '../ast.js'
import { transform } from '../ast.js'
import type { ExtractResult, FunctionSpec } from '../workbook.js'
import { literalFor, sanitizeKeyName } from './shared.js'

/**
 * VBA emitter. Direct descendant of the 2013 Gen_VBA — same shim strategy
 * (VBA lacks ternaries and several Excel functions), same Function/End Function shape.
 */

export const VBA_SHIMS: Record<string, string> = {
  inlineIf: [
    'Function inlineIf(statement, iftrue, Optional iffalse)',
    '  If statement Then',
    '    inlineIf = iftrue',
    '  ElseIf IsMissing(iffalse) Then',
    '    inlineIf = False',
    '  Else',
    '    inlineIf = iffalse',
    '  End If',
    'End Function',
  ].join('\n'),
  PI: ['Function PI()', '  PI = 4 * Atn(1)', 'End Function'].join('\n'),
  MIN: [
    'Function MIN(ParamArray values())',
    '  Dim v As Variant',
    '  MIN = values(LBound(values))',
    '  For Each v In values',
    '    If v < MIN Then MIN = v',
    '  Next v',
    'End Function',
  ].join('\n'),
  MAX: [
    'Function MAX(ParamArray values())',
    '  Dim v As Variant',
    '  MAX = values(LBound(values))',
    '  For Each v In values',
    '    If v > MAX Then MAX = v',
    '  Next v',
    'End Function',
  ].join('\n'),
}

const joinInfix = (args: string[], op: string): string => `(${args.join(` ${op} `)})`

export const vbaDialect: Dialect = {
  name: 'vba',
  string: (value) => `"${value.replace(/"/g, '""')}"`,
  boolean: (value) => (value ? 'True' : 'False'),
  percent: (inner) => `(${inner} / 100)`,
  opSymbol: (op) => op,
  call: (name, args, ctx) => {
    switch (name) {
      case 'IF': {
        ctx.usedShims.add('inlineIf')
        return `inlineIf(${args.join(', ')})`
      }
      case 'AND':
        return joinInfix(args, 'And')
      case 'OR':
        return joinInfix(args, 'Or')
      case 'NOT':
        return `Not (${args.join(', ')})`
      case 'SUM':
        return joinInfix(args, '+')
      case 'POWER':
        // Fallback for direct emitExpression calls; emitVbaFunction rewrites POWER to the
        // ^ operator at AST level so precedence-aware parenthesization applies.
        return `((${args[0]}) ^ (${args[1]}))`
      case 'SQRT':
        return `Sqr(${args.join(', ')})`
      case 'ABS':
        return `Abs(${args.join(', ')})`
      case 'PI': {
        ctx.usedShims.add('PI')
        return 'PI()'
      }
      case 'MIN':
      case 'MAX': {
        ctx.usedShims.add(name)
        return `${name}(${args.join(', ')})`
      }
      case 'ROUND':
        return `Round(${args.join(', ')})`
      default: {
        ctx.unknownFunctions.add(name)
        return `${name}(${args.join(', ')})`
      }
    }
  },
  fallbackRef: sanitizeKeyName,
}

const powerToOperator = (node: FormulaNode): FormulaNode => {
  if (node.kind === 'call' && node.name === 'POWER' && node.args.length === 2) {
    const [left, right] = node.args
    if (left && right && left.kind !== 'empty' && right.kind !== 'empty') {
      return { kind: 'binary', op: '^', left, right }
    }
  }
  return node
}

export const emitVbaFunction = (spec: FunctionSpec, ctx: EmitContext): string => {
  const body = emitExpression(transform(spec.body, powerToOperator), vbaDialect, ctx)
  const lines: string[] = [`Function ${spec.name}(${spec.params.map((p) => p.name).join(', ')})`]
  for (const constant of spec.consts) {
    lines.push(`  Const ${constant.name} = ${literalFor(constant.value, vbaDialect)}`)
  }
  lines.push(`  ${spec.name} = ${body}`)
  lines.push('End Function')
  return lines.join('\n')
}

/** Whole extraction → module text: shims first (only the ones used), then functions in dependency order. */
export const emitVbaModule = (result: ExtractResult): string => {
  const ctx = makeContext(result.target.sheet, result.nameByKey, result.functions)
  const functions = result.functions.map((spec) => emitVbaFunction(spec, ctx))
  const shims = [...ctx.usedShims]
    .sort()
    .map((shim) => VBA_SHIMS[shim])
    .filter((s): s is string => s !== undefined)
  return [...shims, ...functions].join('\n\n')
}

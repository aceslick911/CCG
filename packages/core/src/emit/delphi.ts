import type { Dialect, EmitContext } from './expression.js'
import { emitExpression, makeContext } from './expression.js'
import type { ExtractResult, FunctionSpec } from '../workbook.js'
import { literalFor, sanitizeKeyName } from './shared.js'

/**
 * Delphi emitter — kept for old times' sake (and AW6). Requires `uses Math;` for
 * Power/Min/Max/IfThen. Note IfThen evaluates both branches (eager), matching the 2013
 * inlineIf semantics.
 */

export const delphiDialect: Dialect = {
  name: 'delphi',
  string: (value) => `'${value.replace(/'/g, "''")}'`,
  boolean: (value) => (value ? 'True' : 'False'),
  percent: (inner) => `(${inner} / 100)`,
  binary: (op, left, right) => {
    if (op === '^') {
      return `Power(${left}, ${right})`
    }
    return undefined
  },
  opSymbol: (op) => {
    switch (op) {
      case '&':
        return '+'
      default:
        return op
    }
  },
  call: (name, args, ctx) => {
    switch (name) {
      case 'IF':
        return `IfThen(${args.join(', ')})`
      case 'AND':
        return `(${args.join(' and ')})`
      case 'OR':
        return `(${args.join(' or ')})`
      case 'NOT':
        return `not (${args.join(', ')})`
      case 'SUM':
        return `(${args.join(' + ')})`
      case 'POWER':
        return `Power(${args.join(', ')})`
      case 'SQRT':
        return `Sqrt(${args.join(', ')})`
      case 'ABS':
        return `Abs(${args.join(', ')})`
      case 'PI':
        return 'Pi'
      case 'MIN':
        return `Min(${args.join(', ')})`
      case 'MAX':
        return `Max(${args.join(', ')})`
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

export const emitDelphiFunction = (spec: FunctionSpec, ctx: EmitContext): string => {
  const body = emitExpression(spec.body, delphiDialect, ctx)
  const params = spec.params.map((p) => `${p.name}: Double`).join('; ')
  const lines: string[] = [`function ${spec.name}(${params}): Double;`]
  if (spec.consts.length > 0) {
    lines.push('const')
    for (const constant of spec.consts) {
      lines.push(`  ${constant.name} = ${literalFor(constant.value, delphiDialect)};`)
    }
  }
  lines.push('begin')
  lines.push(`  Result := ${body};`)
  lines.push('end;')
  return lines.join('\n')
}

export const emitDelphiModule = (
  result: ExtractResult,
  options: { onContext?: (ctx: EmitContext) => void } = {}
): string => {
  const ctx = makeContext(result.target.sheet, result.nameByKey, result.functions)
  const parts = result.functions.map((spec) => emitDelphiFunction(spec, ctx))
  options.onContext?.(ctx)
  return parts.join('\n\n')
}

import type { Dialect, EmitContext } from './expression.js'
import { emitExpression, makeContext } from './expression.js'
import type { CellValue, ExtractResult, FunctionSpec, ParamSpec } from '../workbook.js'
import { literalFor, sanitizeKeyName } from './shared.js'

/**
 * C# emitter. Modernized from the 2013 Gen_CSharp: IF becomes a ternary instead of an
 * inlineIf shim, POWER/^ become Math.Pow, and the generated test compares with a tolerance
 * instead of ==. Types are inferred from cached values (double unless proven otherwise) —
 * the same pragmatic call the 2013 version made, minus its bool-return bug.
 */

const nestPairwise = (fn: string, args: string[]): string => {
  if (args.length === 1) {
    return args[0] ?? ''
  }
  const [first, ...rest] = args
  return `${fn}(${first}, ${nestPairwise(fn, rest)})`
}

export const csharpDialect: Dialect = {
  name: 'csharp',
  string: (value) => `"${value.replace(/\\/g, '\\\\').replace(/"/g, '\\"')}"`,
  boolean: (value) => (value ? 'true' : 'false'),
  percent: (inner) => `(${inner} / 100.0)`,
  binary: (op, left, right) => {
    if (op === '^') {
      return `Math.Pow(${left}, ${right})`
    }
    return undefined
  },
  opSymbol: (op) => {
    switch (op) {
      case '=':
        return '=='
      case '<>':
        return '!='
      case '&':
        return '+'
      default:
        return op
    }
  },
  call: (name, args, ctx) => {
    switch (name) {
      case 'IF': {
        const [cond, whenTrue, whenFalse] = args
        return `(${cond} ? ${whenTrue} : ${whenFalse ?? 'false'})`
      }
      case 'AND':
        return `(${args.join(' && ')})`
      case 'OR':
        return `(${args.join(' || ')})`
      case 'NOT':
        return `!(${args.join(', ')})`
      case 'SUM':
        return `(${args.join(' + ')})`
      case 'POWER':
        return `Math.Pow(${args.join(', ')})`
      case 'SQRT':
        return `Math.Sqrt(${args.join(', ')})`
      case 'ABS':
        return `Math.Abs(${args.join(', ')})`
      case 'PI':
        return 'Math.PI'
      case 'SIN':
      case 'COS':
      case 'TAN':
        return `Math.${name.charAt(0)}${name.slice(1).toLowerCase()}(${args.join(', ')})`
      case 'ATAN':
        return `Math.Atan(${args.join(', ')})`
      case 'ROUND':
        return `Math.Round(${args.join(', ')})`
      case 'MIN':
        return nestPairwise('Math.Min', args)
      case 'MAX':
        return nestPairwise('Math.Max', args)
      default: {
        ctx.unknownFunctions.add(name)
        return `${name}(${args.join(', ')})`
      }
    }
  },
  fallbackRef: sanitizeKeyName,
}

const typeOf = (value: CellValue | undefined): string => {
  if (typeof value === 'string') {
    return 'string'
  }
  if (typeof value === 'boolean') {
    return 'bool'
  }
  return 'double'
}

const paramList = (params: ParamSpec[]): string => params.map((p) => `${typeOf(p.value)} ${p.name}`).join(', ')

export const emitCSharpFunction = (spec: FunctionSpec, ctx: EmitContext): string => {
  const body = emitExpression(spec.body, csharpDialect, ctx)
  const lines: string[] = [`public static ${typeOf(spec.value)} ${spec.name}(${paramList(spec.params)})`, '{']
  for (const constant of spec.consts) {
    lines.push(`  const ${typeOf(constant.value)} ${constant.name} = ${literalFor(constant.value, csharpDialect)};`)
  }
  lines.push(`  return ${body};`)
  lines.push('}')
  return lines.join('\n')
}

/**
 * The 2013 test_Method trick, kept: the workbook's cached values become the test oracle.
 * Only emitted when the target and every parameter have numeric cached values.
 */
export const emitCSharpTest = (spec: FunctionSpec): string | undefined => {
  if (typeof spec.value !== 'number' || spec.params.some((p) => typeof p.value !== 'number')) {
    return undefined
  }
  const args = spec.params.map((p) => String(p.value)).join(', ')
  return [
    `public static bool Test_${spec.name}()`,
    '{',
    `  return Math.Abs(${spec.name}(${args}) - ${spec.value}) < 0.000001;`,
    '}',
  ].join('\n')
}

export const emitCSharpModule = (
  result: ExtractResult,
  options: { withTests?: boolean; onContext?: (ctx: EmitContext) => void } = {}
): string => {
  const ctx = makeContext(result.target.sheet, result.nameByKey, result.functions)
  const parts = result.functions.map((spec) => emitCSharpFunction(spec, ctx))
  if (options.withTests) {
    for (const spec of result.functions) {
      const test = emitCSharpTest(spec)
      if (test) {
        parts.push(test)
      }
    }
  }
  options.onContext?.(ctx)
  return parts.join('\n\n')
}

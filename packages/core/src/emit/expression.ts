import type { BinaryOp, FormulaNode } from '../ast.js'
import { formatCellRefText } from '../ast.js'
import { cellKey, expandRangeKeys } from '../refs.js'
import type { FunctionSpec } from '../workbook.js'

/**
 * Shared expression walker. The 2013 codebase routed VB, C# and Delphi through one
 * flattener configured by global mutable flags; this is the same idea with the dialect
 * passed explicitly and no state leakage between calls.
 */

export type EmitContext = {
  defaultSheet: string
  nameForKey: (key: string) => string | undefined
  functionByKey: (key: string) => FunctionSpec | undefined
  usedShims: Set<string>
  unknownFunctions: Set<string>
}

export type Dialect = {
  name: string
  string: (value: string) => string
  boolean: (value: boolean) => string
  percent: (inner: string) => string
  /** Return undefined to keep the operator infix with the same symbol. */
  binary?: (op: BinaryOp, left: string, right: string) => string | undefined
  opSymbol: (op: BinaryOp) => string
  call: (name: string, args: string[], ctx: EmitContext) => string
  /** Fallback identifier when a referenced cell has no classification-provided name. */
  fallbackRef: (key: string) => string
  structured?: (table: string | undefined, column: string, thisRow: boolean) => string
}

const BINDING: Record<BinaryOp, number> = {
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

export const emitExpression = (node: FormulaNode, dialect: Dialect, ctx: EmitContext): string =>
  emitNode(node, dialect, ctx, 0)

const emitNode = (node: FormulaNode, dialect: Dialect, ctx: EmitContext, parentBinding: number): string => {
  switch (node.kind) {
    case 'number':
      return node.text
    case 'string':
      return dialect.string(node.value)
    case 'boolean':
      return dialect.boolean(node.value)
    case 'error':
      throw new Error(`[emit:${dialect.name}] Cannot generate code from error literal ${node.code}`)
    case 'cell':
      return refName(cellKey(node.ref.sheet ?? ctx.defaultSheet, node.ref), dialect, ctx)
    case 'range':
      throw new Error(
        `[emit:${dialect.name}] Range ${formatCellRefText(node.start)}:${formatCellRefText(node.end)} used outside a function call`
      )
    case 'name':
      return node.name
    case 'call': {
      const args: string[] = []
      for (const arg of node.args) {
        if (arg.kind === 'range') {
          for (const key of expandRangeKeys(arg.start, arg.end, ctx.defaultSheet)) {
            args.push(refName(key, dialect, ctx))
          }
        } else if (arg.kind === 'empty') {
          args.push('')
        } else {
          args.push(emitNode(arg, dialect, ctx, 0))
        }
      }
      return dialect.call(node.name, args, ctx)
    }
    case 'binary': {
      const binding = BINDING[node.op]
      const left = emitNode(node.left, dialect, ctx, binding - 1)
      const right = emitNode(node.right, dialect, ctx, binding)
      const transformed = dialect.binary?.(node.op, left, right)
      if (transformed !== undefined) {
        return transformed
      }
      const text = `${left} ${dialect.opSymbol(node.op)} ${right}`
      return binding <= parentBinding ? `(${text})` : text
    }
    case 'unary': {
      const operand = emitNode(node.operand, dialect, ctx, 60)
      return `${node.op}${operand}`
    }
    case 'percent':
      return dialect.percent(emitNode(node.operand, dialect, ctx, 60))
    case 'group':
      return `(${emitNode(node.inner, dialect, ctx, 0)})`
    case 'structured': {
      if (dialect.structured) {
        return dialect.structured(node.table, node.column, node.thisRow)
      }
      throw new Error(`[emit:${dialect.name}] Structured references (tables) are not supported yet`)
    }
    case 'array':
      throw new Error(`[emit:${dialect.name}] Array literals are not supported yet`)
    case 'empty':
      return ''
    default: {
      const exhaustive: never = node
      return exhaustive
    }
  }
}

const refName = (key: string, dialect: Dialect, ctx: EmitContext): string => {
  const fn = ctx.functionByKey(key)
  if (fn) {
    // A reference to a function cell becomes a call to the generated function —
    // built directly so it never collides with the dialect's Excel-function mapping.
    return `${fn.name}(${fn.params.map((param) => param.name).join(', ')})`
  }
  return ctx.nameForKey(key) ?? dialect.fallbackRef(key)
}

export const makeContext = (
  defaultSheet: string,
  nameByKey: Map<string, string>,
  functions: FunctionSpec[] = []
): EmitContext => {
  const fnMap = new Map(functions.map((fn) => [fn.key, fn]))
  return {
    defaultSheet,
    nameForKey: (key) => nameByKey.get(key),
    functionByKey: (key) => fnMap.get(key),
    usedShims: new Set<string>(),
    unknownFunctions: new Set<string>(),
  }
}

import { collectRefs, type FormulaNode } from './ast.js'
import { parseFormula } from './parser.js'
import { cellKey, expandRangeKeys } from './refs.js'

/**
 * The workbook model. Data crosses into core as plain snapshots — never live
 * Office.js/interop objects. This is the 2013 lesson (Cell held a live Excel.Range,
 * and codegen re-read formulas from Excel mid-generation) applied in reverse.
 */

export type CellValue = number | string | boolean

export type CellSnapshot = {
  sheet: string
  ref: string
  formula?: string
  value?: CellValue
}

export type WorkbookSnapshot = {
  cells: Map<string, CellSnapshot>
}

export const makeSnapshot = (cells: CellSnapshot[]): WorkbookSnapshot => {
  const map = new Map<string, CellSnapshot>()
  for (const cell of cells) {
    map.set(cellKey(cell.sheet, cell.ref), cell)
  }
  return { cells: map }
}

export type CellKind = 'function' | 'variable' | 'constant'

export type CellMeta = {
  kind: CellKind
  name?: string
  description?: string
  unit?: string
}

/** Keyed by canonical cell key ('SHEET1!E7'). The successor to the 2013 CustomXMLParts payload. */
export type Classification = Map<string, CellMeta>

export type ParamSpec = {
  name: string
  key: string
  sheet: string
  ref: string
  value?: CellValue
  description?: string
  unit?: string
}

export type FunctionSpec = {
  name: string
  key: string
  sheet: string
  ref: string
  formula: string
  body: FormulaNode
  params: ParamSpec[]
  consts: ParamSpec[]
  calls: string[]
  value?: CellValue
}

export type ExtractResult = {
  /** Dependencies first, target last — the 2013 TraceTree ordering, minus the retry loop. */
  functions: FunctionSpec[]
  target: FunctionSpec
  nameByKey: Map<string, string>
}

export type ExtractOptions = {
  /** Identifier for unclassified referenced cells. Default: Sheet_Ref. */
  autoName?: (sheet: string, ref: string) => string
}

export class CircularDependencyError extends Error {
  constructor(public readonly path: string[]) {
    super(`[workbook] Circular dependency: ${path.join(' -> ')}`)
  }
}

const sanitizeIdentifier = (raw: string): string => {
  const cleaned = raw
    .replace(/[^A-Za-z0-9_]/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_|_$/g, '')
  return /^[0-9]/.test(cleaned) ? `N${cleaned}` : cleaned || 'Value'
}

const defaultAutoName = (sheet: string, ref: string): string => sanitizeIdentifier(`${sheet}_${ref}`)

/**
 * Feature 1/2: turn a formula cell (and, transitively, every function cell it references)
 * into ordered FunctionSpecs ready for any emitter. Referenced cells become parameters,
 * constants, or calls to other generated functions, per the classification map.
 */
export const extractFunctions = (
  snapshot: WorkbookSnapshot,
  targetKey: string,
  classification: Classification,
  options: ExtractOptions = {}
): ExtractResult => {
  const { autoName = defaultAutoName } = options
  const nameByKey = new Map<string, string>()
  const usedNames = new Set<string>()

  const nameFor = (key: string, meta: CellMeta | undefined, cell: CellSnapshot | undefined): string => {
    const existing = nameByKey.get(key)
    if (existing !== undefined) {
      return existing
    }
    const [sheetPart, refPart] = key.split('!')
    const base = meta?.name
      ? sanitizeIdentifier(meta.name)
      : autoName(cell?.sheet ?? sheetPart ?? '', cell?.ref ?? refPart ?? '')
    let candidate = base
    let suffix = 2
    while (usedNames.has(candidate)) {
      candidate = `${base}_${suffix}`
      suffix += 1
    }
    usedNames.add(candidate)
    nameByKey.set(key, candidate)
    return candidate
  }

  const done = new Map<string, FunctionSpec>()
  const order: FunctionSpec[] = []
  const visiting = new Set<string>()

  const build = (key: string, path: string[]): FunctionSpec => {
    const cached = done.get(key)
    if (cached) {
      return cached
    }
    if (visiting.has(key)) {
      throw new CircularDependencyError([...path, key])
    }
    visiting.add(key)

    const cell = snapshot.cells.get(key)
    if (!cell) {
      throw new Error(`[workbook] No snapshot for cell ${key}`)
    }
    if (!cell.formula) {
      throw new Error(`[workbook] Cell ${key} has no formula — cannot extract a function from it`)
    }

    const body = parseFormula(cell.formula)
    const params: ParamSpec[] = []
    const consts: ParamSpec[] = []
    const calls: string[] = []
    const seen = new Set<string>()

    const referencedKeys: string[] = []
    for (const refNode of collectRefs(body)) {
      if (refNode.kind === 'cell') {
        referencedKeys.push(cellKey(refNode.ref.sheet ?? cell.sheet, refNode.ref))
      } else {
        referencedKeys.push(...expandRangeKeys(refNode.start, refNode.end, cell.sheet))
      }
    }

    for (const refKey of referencedKeys) {
      if (seen.has(refKey) || refKey === key) {
        continue
      }
      seen.add(refKey)

      const meta = classification.get(refKey)
      const refCell = snapshot.cells.get(refKey)
      const spec: ParamSpec = {
        name: nameFor(refKey, meta, refCell),
        key: refKey,
        sheet: refCell?.sheet ?? refKey.split('!')[0] ?? '',
        ref: refCell?.ref ?? refKey.split('!')[1] ?? '',
        value: refCell?.value,
        description: meta?.description,
        unit: meta?.unit,
      }

      if (meta?.kind === 'function') {
        const child = build(refKey, [...path, key])
        calls.push(refKey)
        // Pass-through: the caller needs the callee's inputs in its own signature.
        for (const childParam of child.params) {
          if (!seen.has(childParam.key)) {
            seen.add(childParam.key)
            params.push(childParam)
          }
        }
      } else if (meta?.kind === 'constant') {
        consts.push(spec)
      } else {
        // Variables and unclassified cells become parameters — the 2013 auto-promotion rule.
        params.push(spec)
      }
    }

    const meta = classification.get(key)
    const spec: FunctionSpec = {
      name: nameFor(key, meta, cell),
      key,
      sheet: cell.sheet,
      ref: cell.ref,
      formula: cell.formula,
      body,
      params,
      consts,
      calls,
      value: cell.value,
    }

    visiting.delete(key)
    done.set(key, spec)
    order.push(spec)
    return spec
  }

  const target = build(targetKey.toUpperCase(), [])
  return { functions: order, target, nameByKey }
}

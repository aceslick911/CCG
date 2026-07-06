import {
  type CellMeta,
  type CellSnapshot,
  type CellValue,
  type Classification,
  cellKey,
  collectRefs,
  emitCSharpModule,
  emitDelphiModule,
  emitLambdaNames,
  emitVbaModule,
  expandRangeKeys,
  extractFunctions,
  makeSnapshot,
  parseFormula,
  toPrefixNotation,
} from '@ccg/core'

const byId = <T extends HTMLElement>(id: string): T => {
  const node = document.getElementById(id)
  if (!node) {
    throw new Error(`[taskpane] Missing element #${id}`)
  }
  return node as T
}

const ui = {
  address: () => byId<HTMLSpanElement>('cell-address'),
  formula: () => byId<HTMLPreElement>('cell-formula'),
  tree: () => byId<HTMLPreElement>('parse-tree'),
  fnName: () => byId<HTMLInputElement>('fn-name'),
  lang: () => byId<HTMLSelectElement>('lang'),
  chain: () => byId<HTMLInputElement>('chain'),
  generate: () => byId<HTMLButtonElement>('generate'),
  createLambda: () => byId<HTMLButtonElement>('create-lambda'),
  copy: () => byId<HTMLButtonElement>('copy'),
  output: () => byId<HTMLPreElement>('output'),
  error: () => byId<HTMLDivElement>('error'),
}

type Selected = { sheet: string; ref: string; formula?: string }

let selected: Selected | undefined

const showError = (err: unknown): void => {
  ui.error().textContent = err instanceof Error ? err.message : String(err)
}

const clearError = (): void => {
  ui.error().textContent = ''
}

/** 'Sheet1!E10:F12' / "'My Sheet'!E10" → { sheet, ref of first cell } */
const parseAddress = (address: string): { sheet: string; ref: string } => {
  const bang = address.lastIndexOf('!')
  const sheetRaw = bang === -1 ? '' : address.slice(0, bang)
  const sheet = sheetRaw.startsWith("'") ? sheetRaw.slice(1, -1).replace(/''/g, "'") : sheetRaw
  const refPart = address.slice(bang + 1)
  const firstCell = refPart.split(':')[0] ?? refPart
  return { sheet, ref: firstCell.replace(/\$/g, '') }
}

const asFormula = (raw: unknown): string | undefined =>
  typeof raw === 'string' && raw.startsWith('=') ? raw : undefined

const asValue = (raw: unknown): CellValue | undefined =>
  typeof raw === 'number' || typeof raw === 'string' || typeof raw === 'boolean' ? raw : undefined

const refreshSelection = async (): Promise<void> => {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange()
    range.load(['address', 'formulas'])
    await context.sync()

    const { sheet, ref } = parseAddress(range.address)
    const formula = asFormula(range.formulas[0]?.[0])
    selected = { sheet, ref, formula }

    ui.address().textContent = `${sheet}!${ref}`
    ui.formula().textContent = formula ?? 'Select a cell that contains a formula to get started.'
    ui.generate().disabled = !formula
    ui.createLambda().disabled = !formula
    if (formula) {
      try {
        ui.tree().textContent = toPrefixNotation(parseFormula(formula))
        clearError()
      } catch (err) {
        ui.tree().textContent = '—'
        showError(err)
      }
    } else {
      ui.tree().textContent = '—'
    }
  })
}

/**
 * Breadth-first snapshot builder: batched reads, one sync per depth level —
 * never a per-cell round trip. Data leaves Excel as plain CellSnapshots.
 */
const buildSnapshot = async (
  target: Selected,
  chain: boolean
): Promise<{ cells: CellSnapshot[]; classification: Classification; targetKey: string }> => {
  const targetKey = cellKey(target.sheet, target.ref)
  const cells: CellSnapshot[] = []
  const classification: Classification = new Map<string, CellMeta>()
  const fetched = new Set<string>()
  let pending = new Map<string, { sheet: string; ref: string }>([[targetKey, { sheet: target.sheet, ref: target.ref }]])

  await Excel.run(async (context) => {
    while (pending.size > 0) {
      const batch = [...pending.entries()].filter(([key]) => !fetched.has(key))
      pending = new Map()
      if (batch.length === 0) {
        break
      }

      const proxies = batch.map(([key, loc]) => {
        fetched.add(key)
        const range = context.workbook.worksheets.getItem(loc.sheet).getRange(loc.ref)
        range.load(['formulas', 'values'])
        return { key, loc, range }
      })
      await context.sync()

      for (const { key, loc, range } of proxies) {
        const formula = asFormula(range.formulas[0]?.[0])
        const value = asValue(range.values[0]?.[0])
        cells.push({ sheet: loc.sheet, ref: loc.ref, formula, value })

        if (key === targetKey) {
          classification.set(key, { kind: 'function', name: ui.fnName().value.trim() || undefined })
        } else if (formula && chain) {
          classification.set(key, { kind: 'function' })
        } else {
          classification.set(key, { kind: 'variable' })
        }

        const followRefs = formula !== undefined && (chain || key === targetKey)
        if (followRefs && formula) {
          for (const refNode of collectRefs(parseFormula(formula))) {
            const targets =
              refNode.kind === 'cell'
                ? [
                    {
                      key: cellKey(refNode.ref.sheet ?? loc.sheet, refNode.ref),
                      sheet: refNode.ref.sheet ?? loc.sheet,
                      ref: `${refNode.ref.col}${refNode.ref.row}`,
                    },
                  ]
                : expandRangeKeys(refNode.start, refNode.end, loc.sheet).map((k) => {
                    const [, refPart] = k.split('!')
                    return { key: k, sheet: refNode.start.sheet ?? loc.sheet, ref: refPart ?? '' }
                  })
            for (const next of targets) {
              if (!fetched.has(next.key) && !pending.has(next.key)) {
                pending.set(next.key, { sheet: next.sheet, ref: next.ref })
              }
            }
          }
        }
      }
    }
  })

  return { cells, classification, targetKey }
}

const unknownFunctionWarning = (unknown: Set<string>): string =>
  unknown.size > 0
    ? `\n\n' NOTE: no native mapping for: ${[...unknown].sort().join(', ')} — emitted verbatim; supply your own implementation.`
    : ''

const generate = async (): Promise<void> => {
  clearError()
  if (!selected?.formula) {
    showError('Select a cell that contains a formula first.')
    return
  }
  try {
    ui.generate().disabled = true
    const { cells, classification, targetKey } = await buildSnapshot(selected, ui.chain().checked)
    const result = extractFunctions(makeSnapshot(cells), targetKey, classification)

    const lang = ui.lang().value
    let unknown = new Set<string>()
    const onContext = (ctx: { unknownFunctions: Set<string> }): void => {
      unknown = ctx.unknownFunctions
    }
    let output: string
    if (lang === 'vba') {
      output = emitVbaModule(result, { onContext })
    } else if (lang === 'csharp') {
      output = emitCSharpModule(result, { withTests: true, onContext })
    } else if (lang === 'delphi') {
      output = emitDelphiModule(result, { onContext })
    } else {
      output = emitLambdaNames(result)
        .map((entry) => `${entry.name}\n${entry.formula}`)
        .join('\n\n')
    }
    ui.output().textContent = output + unknownFunctionWarning(unknown)
  } catch (err) {
    showError(err)
  } finally {
    ui.generate().disabled = false
  }
}

/**
 * Verify the calc engine actually has LAMBDA before creating names that would #NAME?.
 * Hard-won: a current build can run an old feature set (stale licence/flight cache — see
 * docs/troubleshooting.md). Probes on a temporary hidden sheet, which is removed after.
 */
const probeLambdaSupport = async (): Promise<boolean> => {
  try {
    return await Excel.run(async (context) => {
      const leftover = context.workbook.worksheets.getItemOrNullObject('__ccg_probe__')
      await context.sync()
      if (!leftover.isNullObject) {
        leftover.delete()
      }
      const sheet = context.workbook.worksheets.add('__ccg_probe__')
      sheet.visibility = Excel.SheetVisibility.hidden
      const cell = sheet.getRange('A1')
      cell.formulas = [['=LAMBDA(x, x*2)(21)']]
      await context.sync()
      cell.load('values')
      await context.sync()
      const supported = cell.values[0]?.[0] === 42
      sheet.delete()
      await context.sync()
      return supported
    })
  } catch {
    // Probe itself failed (protected workbook etc.) — don't block the user on it.
    return true
  }
}

/** The modern VBA-injection replacement: register generated LAMBDAs as workbook names. */
const createLambdaNames = async (): Promise<void> => {
  clearError()
  if (!selected?.formula) {
    showError('Select a cell that contains a formula first.')
    return
  }
  try {
    ui.createLambda().disabled = true
    if (!(await probeLambdaSupport())) {
      showError(
        "This Excel's calculation engine doesn't support LAMBDA yet, so created names would return #NAME?. " +
          'Use the Visual Basic output instead. If this is a current Microsoft 365 build, see the troubleshooting ' +
          'guide (github.com/aceslick911/CCG/blob/main/docs/troubleshooting.md) — a licence/policy refresh and a ' +
          'full double restart of Excel usually restores it.'
      )
      return
    }
    const { cells, classification, targetKey } = await buildSnapshot(selected, ui.chain().checked)
    const result = extractFunctions(makeSnapshot(cells), targetKey, classification)
    const lambdas = emitLambdaNames(result)

    await Excel.run(async (context) => {
      const existing = lambdas.map((entry) => context.workbook.names.getItemOrNullObject(entry.name))
      await context.sync()
      existing.forEach((item, index) => {
        if (!item.isNullObject) {
          item.delete()
        }
        const entry = lambdas[index]
        if (entry) {
          context.workbook.names.add(entry.name, entry.formula)
        }
      })
      await context.sync()
    })

    const usage = result.target.params.map((p) => p.value ?? p.name).join(', ')
    ui.output().textContent = [
      `Created workbook name${lambdas.length > 1 ? 's' : ''}: ${lambdas.map((l) => l.name).join(', ')}`,
      '',
      `Try it in any cell:`,
      `=${result.target.name}(${usage})`,
    ].join('\n')
  } catch (err) {
    showError(err)
  } finally {
    ui.createLambda().disabled = false
  }
}

const copyOutput = async (): Promise<void> => {
  const text = ui.output().textContent ?? ''
  try {
    await navigator.clipboard.writeText(text)
    ui.copy().textContent = 'Copied!'
  } catch {
    const area = document.createElement('textarea')
    area.value = text
    document.body.appendChild(area)
    area.select()
    document.execCommand('copy')
    area.remove()
    ui.copy().textContent = 'Copied!'
  }
  setTimeout(() => {
    ui.copy().textContent = 'Copy'
  }, 1500)
}

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Excel) {
    showError('CCG only runs in Excel.')
    return
  }
  ui.generate().addEventListener('click', () => void generate())
  ui.createLambda().addEventListener('click', () => void createLambdaNames())
  ui.copy().addEventListener('click', () => void copyOutput())

  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => void refreshSelection())
  await refreshSelection()
})

# `@ccg/cli` — Design Document (ON ICE)

> **Status: not implemented.** This document is the full design so the CLI can be built later
> without re-deriving decisions. Nothing in here is speculative about the *core* — everything
> the CLI needs already exists (or is planned) in `@ccg/core`; the CLI is a thin I/O shell.

## Why it exists

The 2013 add-in had a killer workflow beyond interactive use: engineering calculations lived in
an Excel workbook, and the add-in regenerated the engineering module's source files (for the
AW6 window/wall system software) whenever the workbook changed. Engineers owned the maths in a
tool they understood; the application consumed generated, type-safe code.

The CLI is that workflow, modernised. Because `@ccg/core` parses formulas itself (it never
needed Excel's object model to compute the dependency graph), **codegen does not require Excel
at all** — a `.xlsx` file is sufficient input. That unlocks what 2013 never could: headless
runs in CI, pre-commit hooks, and watch mode on any OS.

## Architecture

```
.xlsx file ──▶ reader adapter (exceljs/SheetJS) ──▶ WorkbookSnapshot ──▶ @ccg/core ──▶ files
                                                    { sheet, ref,
                                                      formula, value }
```

- One dependency boundary: the reader produces the same `WorkbookSnapshot` structure the add-in
  produces from Office.js. Core neither knows nor cares which side fed it.
- Values read from the file are the **cached values** Excel stored at last save. That is fine:
  values are only used for const inlining and generated test expectations. The CLI never
  recalculates — the formulas are the source of truth, which is the entire premise of CCG.
- Cell classification metadata (which cells are functions/variables/constants, their names,
  units, descriptions) is read from the workbook's **custom XML part** — the same
  `ccg` JSON payload the add-in writes (successor to the 2013 `Angelo.CellAnalysis`
  XMS part). A workbook annotated in the add-in on Mac is immediately buildable by the CLI in
  CI with zero extra config.
- Overrides / annotation-free operation via `ccg.config.json` (see below) for workbooks that
  cannot be annotated (e.g. third-party files).

## Installation (future)

```sh
npm install -D @ccg/cli        # per-project (preferred, pins the codegen version)
npm install -g @ccg/cli        # global, for exploration
```

Node ≥ 20. No Excel, no macOS/Windows requirement — runs anywhere Node runs.

## Commands

### `ccg inspect <file.xlsx>`

Prints the workbook map: sheets, annotated cells (functions/variables/constants) with names and
formulas, tables, and any CCG metadata found. Sanity-check what the generator will see.

```
$ ccg inspect InterestSpreadsheet.xlsm
Sheet1
  E7   Present_Value   variable   1000
  E8   Interest_Rate   variable   0.05
  E9   Years           variable   20
  E10  Future_Value    function   =E7*POWER(1+E8,E9)
```

### `ccg analyze <file.xlsx> <cell>`

Analyze one cell: parsed AST (pretty tree), transitive precedents, and the inferred
parameter/const classification. `<cell>` is `Sheet1!E10` form.

### `ccg gen <file.xlsx> --cell <ref> --lang <lang> [--out <path>]`

Feature 1/2: single cell + dependents → one function (or a chain of functions when precedent
cells are themselves classified as functions).

- `--lang`: `vba` | `csharp` | `delphi` | `lambda` | `ts`
- `--style flat|struct` — flat parameter list vs generated input/output structs (ports the 2013
  `Gen_CSharp` vs `Gen_CSharp_Struct` distinction)
- `--with-tests` — also emit the test function seeded with the workbook's cached values (the
  2013 `test_Method*` trick, still a great idea)
- `--shims inline|import` — prepend the language shim library (`POWER`, `inlineIf`, `MIN`…) to
  the output file, or emit an import against a shared shims file
- No `--out` → stdout (pipeable)

```sh
$ ccg gen InterestSpreadsheet.xlsm --cell Sheet1!E10 --lang vba
Function Future_Value(Present_Value, Interest_Rate, Years)
Future_Value = Present_Value * POWER(1 + Interest_Rate, Years)
End Function
```

### `ccg gen <file.xlsx> --module <name> --lang <lang> --out <dir>`

Feature 3: whole-workbook → importable code module. Convention over configuration:

- A sheet named `Inputs` (or marked in metadata) defines the module's input surface: each named
  cell/row becomes a parameter of the module facade.
- A sheet named `Outputs` defines the return surface.
- Every annotated function cell in between becomes a private function; the dependency graph is
  topologically ordered (the 2013 `TraceTree` algorithm) so functions are emitted
  before their callers, and cycles are a hard error listing the cycle path.
- Emits one file per language module + the shims file + optionally a test file asserting every
  output cell against cached values.

### `ccg gen <file.xlsx> --table <Sheet!TableName> --lang <lang>`

Feature 4: table → loop. Columns become parameters (header row = names), each computed column's
formula becomes the loop body, rows become iterations. Generates a function taking arrays (or
an iterable of row-structs with `--style struct`) and returning computed columns.

### `ccg watch <file.xlsx> --out <dir> [--lang ...] [--module ...]`

The AW6 workflow. Watches the file (with debounce — Excel saves are multi-event), regenerates
on change, writes only when output actually differs (keeps mtimes stable for build tools),
prints a one-line diff summary per regeneration.

### `ccg check <file.xlsx> --out <dir>`

CI drift guard: regenerates in memory and exits non-zero if committed generated files differ.

```yaml
# .github/workflows/codegen.yml
- run: npm run build
- run: ./node_modules/.bin/ccg check calcs/engineering.xlsx --out src/generated/
```

## `ccg.config.json`

Optional; CLI flags win over config, config wins over workbook metadata.

```json
{
  "file": "calcs/engineering.xlsx",
  "out": "src/generated",
  "languages": ["csharp", "ts"],
  "module": { "name": "EngineeringCalcs", "inputs": "Inputs", "outputs": "Outputs" },
  "cells": {
    "Sheet1!E7": { "name": "Present_Value", "kind": "variable" },
    "Sheet1!E10": { "name": "Future_Value", "kind": "function" }
  },
  "onUnsupportedFunction": "shim-stub"
}
```

`onUnsupportedFunction`: `error` (default) | `shim-stub` (emit a `NotImplemented` shim and keep
going, listing them in the summary) — the policy for Excel functions the emitter has no mapping
or shim for.

## Capabilities & limits

- **Can**: parse any formula `@ccg/core` supports; resolve cross-sheet references; follow the
  dependency graph transitively; emit all core languages; run without Excel installed.
- **Cannot**: recalculate (values are last-saved cache); read `.xlsb`; evaluate volatile
  functions (`NOW()` etc. — these emit as runtime calls in the target language, flagged in the
  summary); understand VBA already in the workbook (out of scope — CCG generates code *from
  formulas*, it does not parse VBA).
- **Won't** (non-goals): a formula evaluator/spreadsheet engine; a general xlsx diff tool;
  writing anything back into the workbook (the CLI is strictly read-from-xlsx).

## Implementation notes (for whenever this thaws)

- Reader: `exceljs` for streaming + style access, or SheetJS CE if licensing/size wins. Decide
  at build time; the adapter is ~100 lines either way.
- House rules apply: TypeScript 5, compiled with `tsc`, `node --test`, no `tsx`/`npx`, Biome
  formatting (see `coding-style.md`).
- The watcher uses `fs.watch` with a 500 ms debounce and a content hash to suppress Excel's
  double-save events.
- Package layout mirrors `@ccg/core`: `src/`, `tests/`, golden fixtures reused from
  `packages/core/fixtures/`.

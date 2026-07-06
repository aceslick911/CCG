# CCG Coding Style

Conventions for all TypeScript in this repo. Derived from the house style in `~/Developer/web`
(Flipped accounts Astro project) on 2026-07-05, adapted for what CCG is: an Office.js Excel
add-in plus pure-logic packages — **not** a web app. React/Astro/a11y-specific rules were
dropped; everything else carries over.

## Formatting — Biome, not Prettier/ESLint

Biome 2.x is the single formatter + linter. Run `npm run format` before committing; CI treats
diagnostics as failures.

- **2-space indentation**, spaces never tabs
- **Line width 120**
- **Single quotes**
- **No semicolons** (`semicolons: "asNeeded"`)
- **Trailing commas: es5**
- JSON files may contain comments (`.jsonc` parser mode is on)

## TypeScript

TypeScript 5, `strict: true` everywhere.

- `noImplicitAny` is **never** disabled. Take 5 seconds to fix errors caused by it — disabling
  it causes dangerous type assumptions. (Verbatim policy from the web repo.)
- Prefer `type` aliases over `interface`:
  `export type EmitOptions = { indent?: string }`
- Public API functions are `export const` arrow functions **with explicit return types**:
  `export const parseFormula = (text: string): FormulaNode => { ... }`
  Plain `export function` is fine for long/recursive bodies — consistency within a file beats
  dogma.
- Options-object parameters with destructured defaults:
  `const { timeoutMs = 120000, context = 'operation' } = options`
- Model data as **discriminated unions** (`kind` field) — the AST is the canonical example.
  Exhaustiveness via `never` checks in `switch` defaults.
- No classes unless there is real mutable state to encapsulate. The 2013 codebase's global
  mutable static flags (`trans_*`) are the cautionary tale: emitter configuration is passed as
  an explicit context argument, never module-level state.

## Lint rules we enforce (Biome)

Carried over from the web repo where they are `error`:

- `useBlockStatements` — always braces, even single-statement `if`s
- `useTemplate` — template literals over string concatenation
- `noDoubleEquals` — `===`/`!==` only
- `useOptionalChain`, `useLiteralKeys`, `noUselessTernary`, `useSingleVarDeclarator`
- `noConfusingVoidType`, `noGlobalIsNan`, `noSwitchDeclarations`
- `noExplicitAny` is **error** here (the web repo relaxes it for app code; CCG is library code,
  so we hold the line)

## Naming

- `camelCase` functions/variables, `PascalCase` types, `UPPER_SNAKE_CASE` for true constants
- Log lines get a bracketed context prefix: `console.log('[parser] ...')`
- File names are `camelCase.ts` (`tokenizer.ts`, `workbookModel.ts`); one concern per file

## Build & run rules (house rules, non-negotiable)

- **No `tsx`, no `ts-node`**: compile with `tsc`, run the emitted JS
- **No `npx`**: tools run via npm scripts (npm resolves `node_modules/.bin`) or an explicit
  `./node_modules/.bin/<tool>` path
- Add dependencies with `npm install <pkg>@<major>` and keep the range npm writes
- Tests use **`node --test`** against compiled output (`npm test` = build + run). No test
  framework dependency.

## Office add-in specifics (where this diverges from the web repo)

CCG targets the current JavaScript Office Add-ins platform — the only add-in model Excel on
macOS supports.

- **Manifest**: the add-in-only **XML manifest**, which is what Excel on Mac requires today.
  Revisit the unified JSON manifest once Mac support is GA; keep manifest-derived constants in
  one module so the swap is contained.
- **Office.js** is loaded from the Microsoft CDN via `<script>` in the task pane HTML — it is
  never bundled or imported. Types come from `@types/office-js` (globals `Office`, `Excel`).
- **Requirement sets**: code against **ExcelApi 1.12** as the floor, feature-detect anything
  newer with `Office.context.requirements.isSetSupported('ExcelApi', '1.14')`. Never call an
  API above the floor without a guard.
- **Batch pattern discipline**: inside `Excel.run`, queue reads, `await context.sync()`, then
  compute, then queue writes and `sync()` again. Never interleave per-cell syncs in a loop —
  that is the Office.js equivalent of the 2013 add-in's per-cell COM chatter.
- **No React** (or any UI framework). The task pane is vanilla TS + DOM. Keep DOM code in the
  add-in package only — `@ccg/core` must never touch `document`, `Office`, or `Excel` globals.
- **FlipBox styles**: layout uses the FlipBox CSS vendored from the web repo
  (`packages/addin/src/flipbox.css`) — Figma-auto-layout-style classes:
  `class="flipbox y NW x-fill"` (direction `x`/`y`, compass alignment `N/S/E/W/NW/...`,
  sizing `x-hug|x-fill|x-fixed` / `y-hug|y-fill|y-fixed`). Respect
  `prefers-color-scheme: dark` and the `--color-text` custom property; Office task panes
  follow the OS theme.
- The add-in bundle is produced by esbuild (via npm script); `tsc --noEmit` typechecks first.
  esbuild is a bundler here, not a runner — the no-`tsx` rule stands.

## Architecture boundaries

- `@ccg/core` — parser, AST, workbook model, emitters. Pure functions, zero I/O, zero globals,
  runs identically in Node and the task-pane webview.
- `@ccg/addin` — Office.js glue + task pane UI. All `Excel.run` calls live here. Data crosses
  into core as plain snapshots (`{ sheet, ref, formula, value }`), never as live proxy objects
  — the 2013 lesson (live `Excel.Range` handles woven through the model) in reverse.
- `@ccg/cli` — on ice; see `docs/cli.md`.

## Commits

One line, no body, no trailers: `dev: <≤10-word description>` (or `FL-XXXX <description>` for
ticketed work — real ticket numbers only).

# Roadmap & cleanup list

Working list of what's rough, what's missing, and what's next. Ordered roughly by
priority within each group.

## Cleanup / correctness (current code)

- [ ] **LAMBDA-support probe at startup** — evaluate a throwaway LAMBDA when the pane loads;
      if the engine lacks it, disable "Create LAMBDA name" with a hint (VBA path + link to
      docs/troubleshooting.md). Hard-won requirement: modern builds can run old feature sets.
- [ ] **C# `&` concatenation semantics** — Excel's `&` coerces both sides to text; our C#
      emitter maps it to `+`, which *adds* numbers. Needs a `ToString`/`Convert` wrapper (or
      an explicit shim) when operand types aren't provably strings. Same review for Delphi.
- [ ] **Percent-literal check** — `50%` in a formula is a *value* (0.5); our `(x / 100)`
      emit is correct, but verify interaction with precedence and LAMBDA passthrough.
- [ ] **Unknown-function surfacing** — `ctx.unknownFunctions` is collected but never shown.
      Render a warning strip in the pane ("no mapping for FOO(); emitted verbatim").
- [ ] **Range→parameter explosion limits** — ranges expand to individual scalar params
      (limit 1000). Large ranges should become array parameters + a loop in the target
      language instead. (Feeds into the tables→loops feature.)
- [ ] **Boolean test values** — the 2013 generator coerced TRUE/FALSE→1/0 in generated
      tests; ours skips non-numeric oracles entirely. Decide and implement.
- [ ] **Delphi `IfThen` eager evaluation** — matches 2013 semantics but differs from Excel's
      lazy IF; document or emit a real if/else via an out-of-line function.
- [ ] Real icons (current ones are 1×1 placeholder PNGs), manifest polish (localized
      strings, better description).
- [ ] `flipbox.css` is vendored from the private web repo — note the sync provenance in the
      file header and re-sync deliberately, not automatically.

## Features (ports from 2013 + the two that were never built)

- [ ] **Classification UI** — the EditVariable/EditFunction workflow reborn: name/describe
      referenced cells, param vs const, reorder arguments. Persist as JSON in
      `workbook.customXmlParts` (successor to the 2013 `Angelo.CellAnalysis` XMS part) so
      annotations travel with the file and the CLI can read them.
- [ ] **Formula visualizer** — math-notation rendering in the pane (KaTeX), drill-down into
      precedents with breadcrumbs. The nostalgic centrepiece.
- [ ] **Mark Wrong/Right/Changed** — border colouring via `range.format.borders`, plus
      Clear Marks.
- [ ] **Workbook → module** (feature 3, never actually implemented in 2013): Inputs/Outputs
      sheets convention, whole-workbook topological emit, one file per language.
- [ ] **Tables → loops** (feature 4, also never implemented): ListObject columns as
      parameters, rows as iterations; structured-reference formulas become loop bodies.
- [ ] **TypeScript emitter** — the language we now live in; near-clone of the C# emitter.
- [ ] **`@ccg/cli`** — thaw per docs/cli.md once core annotations (customXmlParts) exist.

## Distribution (Excel on the web, other machines, the store)

Same add-in runs on Excel for web/Windows/iPad with zero code changes — it's all Office.js.
What's actually required:

- [ ] **Static hosting** — the pane is pure static files; GitHub Pages is enough (free,
      HTTPS). Build → publish `packages/addin/dist/` → swap `localhost:3000` URLs in the
      manifest for the public URL (parameterize the manifest build).
- [ ] **Excel on the web sideload** — Insert → Add-ins → Upload My Add-in with the manifest
      (needs the hosted URL; web can't reach your localhost dev server from Microsoft's
      side, though same-machine browser testing works).
- [ ] **Org rollout** — Centralized Deployment via admin center → Integrated apps (works for
      your whole tenant, no store involvement).
- [ ] **AppSource (the public store)** — needs a Microsoft Partner Center account, validation
      pass (working support URL, privacy policy, terms, proper icons, no localhost anywhere),
      and review turnaround. Do this only once the UX is solid.

## Testing

- [ ] Golden-file test that byte-compares generated VBA against a checked-in expected module
      for the InterestSpreadsheet fixture (currently semantic assertions only).
- [ ] Parser fuzz/property tests: parse→re-emit→re-parse stability.
- [ ] An e2e smoke against Excel web via Playwright (once hosted): open workbook, run
      generate, assert output.

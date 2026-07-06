# AppSource submission checklist

Everything needed to list CCG on Microsoft AppSource (free app, free to publish).
Validation policies: https://learn.microsoft.com/en-us/legal/marketplace/certification-policies
(section 1120 covers Office add-ins).

## Account (human steps)

- [ ] Enrol at https://partner.microsoft.com/dashboard/registration — **Commercial
      Marketplace** program, no fee. Publisher: Flipped Energy (company, ABN verification)
      or individual.
- [ ] Wait for identity verification (days; watch email for document requests).
- [ ] Partner Center → Marketplace offers → **+ New offer → Office add-in**.

## Technical prerequisites (repo — status)

- [x] Assets hosted on public HTTPS (GitHub Pages, auto-deployed from `main`)
- [x] Production manifest with no localhost, own GUID — https://aceslick911.github.io/CCG/manifest.xml
- [x] Real icons: 16/32/64/80 px in manifest + 300×300 store icon
      (`packages/addin/assets/store-icon-300.png`) — Lucide `square-function`, ISC licence
- [x] Privacy policy — https://aceslick911.github.io/CCG/privacy.html
- [x] Terms of use — https://aceslick911.github.io/CCG/terms.html
- [x] Support URL — https://github.com/aceslick911/CCG
- [x] Buttons disabled with guidance when no formula selected (validators poke empty states)
- [x] LAMBDA-support probe: creating names on an engine without LAMBDA gives a clear
      explanation instead of silent #NAME?
- [x] Unknown Excel functions surfaced as a note in generated output
- [ ] Manifest passes production validation:
      `./node_modules/.bin/office-addin-manifest validate -p packages/addin/manifest.xml`
      (run against the *prod* manifest: `packages/addin/dist/manifest.xml`)
- [ ] Screenshots for the listing: at least 1, ideally 3–5, **1366×768** PNG. Suggested:
      (1) task pane with parse tree next to the InterestSpreadsheet, (2) generated C# code,
      (3) =FutVal(...) evaluating after Create LAMBDA name.
- [ ] Test the prod manifest end-to-end once more before submitting (sideload the hosted
      manifest, not the localhost one).

## Listing content (drafts — paste into Partner Center)

**Name:** CCG — exCel Code Generator

**Summary (100 chars):**
Turn Excel formulas into real code: VBA, C#, Delphi and reusable LAMBDA functions.

**Description:**
CCG analyses the formula in any cell — and all the cells it depends on — and generates
clean, ready-to-use code. Referenced cells become named parameters or constants; chains of
formula cells become chains of functions.

- Generate Visual Basic (VBA), C#, or Delphi from any formula
- Create reusable spreadsheet functions instantly: one click registers your formula as a
  native Excel LAMBDA name — no macros, no VBA trust settings
- View a parse tree of any formula as you navigate your workbook
- 100% local: your workbook data never leaves Excel (no accounts, no telemetry)
- Free and open source (Apache 2.0) — github.com/aceslick911/CCG

CCG is the modern revival of the 2013 "Cell Analysis Add-in for Microsoft Excel".

**Categories:** Productivity; Developer tools
**Industries:** (leave generic)

## Validator testing notes (paste into "Notes for certification")

1. Open any workbook (or the sample at
   https://github.com/aceslick911/CCG/raw/main/packages/core/fixtures/2013/InterestSpreadsheet.xlsm
   — save as .xlsx first if prompted; legacy .xls files do not support add-ins).
2. Enter values: E7=1000, E8=0.05, E9=20, and in E11: `=E7*POWER(1+E8,E9)`
3. Select E11 → the task pane shows the formula and its parse tree.
4. Enter a function name (e.g. FutVal), pick a language, click **Generate code** →
   generated source appears in Output.
5. Pick "Excel LAMBDA" → **Create LAMBDA name** → in any empty cell type
   `=FutVal(1000, 0.05, 20)` → returns 2653.297705.
6. No sign-in, no external services; all processing is client-side.

## Common rejection reasons to pre-empt

- Icons unclear at 16px / missing sizes → done (four sizes generated).
- Add-in errors on empty/edge states → buttons disabled without a formula; errors surfaced
  in-pane with guidance.
- Privacy policy URL broken or generic → dedicated page, accurate (no data collection).
- Description overpromises → keep the LAMBDA caveat honest (requires a LAMBDA-capable
  Excel build; the add-in explains this itself when unsupported).
- Manifest/version mismatch → bump `<Version>` in manifest.xml on every resubmission.

## After approval

- Updates to task-pane code/UI ship via GitHub Pages on push — **no resubmission**.
- Only manifest or listing changes require a new Partner Center submission round.
- Add the AppSource badge/link to README once live.

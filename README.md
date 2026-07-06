# CCG — exCel Code Generator

**CCG** is a modern, open-source rewrite of the original **Cell Analysis Add-in for Microsoft Excel** that I released as free software in 2013.

It helps users analyze, visualize, and work with complex Excel formulas by rendering them in clean mathematical notation (similar to LaTeX), navigating through cell dependencies, applying validation color coding, and instantly converting formulas into C#, Visual Basic, or Delphi code.

## History

This project is a complete rewrite and open-source revival of my earlier free tool:

**Original software (2013):**\
**Cell Analysis Add-in for Microsoft Excel**

- CNET download page: https://download.cnet.com/cell-analysis-add-in-for-microsoft-excel/3000-2077_4-76037251.html
- Softpedia page: https://www.softpedia.com/get/Office-tools/Other-Office-Tools/Cell-Analysis-Add-in.shtml

The original add-in added an "Analyse Cell" button to Excel that opened a Visualizer window showing formulas in proper mathematical notation, allowed drilling into dependent cells, supported color marking (Wrong/Right/Changed), and included one-click conversion of formulas into programming code.

CCG brings that same core idea into the modern era as a fully open-source project under the Apache 2.0 license.

## Features (Current / Planned)

- **Working now:** parse any cell's formula, walk its precedents, and generate a function
  from it in **VBA**, **C#**, **Delphi**, or as an **Excel LAMBDA** — chained precedent
  formulas become chained functions
- **Working now:** register generated LAMBDAs as workbook names (the modern, macro-free
  successor to the 2013 add-in's VBA injection) — works on Mac, Windows, and web
- Formula visualizer with clean mathematical notation *(planned)*
- Color marking system (Wrong/Right/Changed) *(planned)*
- Whole-workbook → importable code module with input/output sheets *(planned)*
- Tables → programming loops (columns as parameters, rows as iterations) *(planned)*
- Headless CLI for CI codegen from `.xlsx` files *(designed, on ice — see [docs/cli.md](docs/cli.md))*

## Project structure

TypeScript monorepo targeting the current JavaScript Office Add-ins platform — the only
add-in model Excel on macOS supports. See [coding-style.md](coding-style.md).

- **`packages/core`** (`@ccg/core`) — formula tokenizer + parser (a real Pratt parser with
  Excel's precedence quirks), plain-data AST, workbook snapshot model with dependency
  extraction and cycle detection, and one emitter per target language. Pure logic, zero
  Office.js — tested with `node --test` against golden fixtures from the 2013 add-in
  (`packages/core/fixtures/2013/`).
- **`packages/addin`** (`@ccg/addin`) — the Excel task pane (vanilla TypeScript, no
  framework; layout via vendored FlipBox CSS). Reads cells as plain snapshots, hands them
  to core, writes generated code and LAMBDA names back.

## License & Attribution

This project is licensed under the **Apache License 2.0**.

If you redistribute this software or a derivative of it, the license requires you to keep the [NOTICE](NOTICE) file's attribution to the original author (Angelo Perera) intact.

See [LICENSE](LICENSE) for full details.

## Getting Started

Requires Node 20+ and Excel for Microsoft 365 (Mac or Windows).

```sh
npm install
npm run build          # compile @ccg/core, typecheck + bundle @ccg/addin
npm test               # core test suite (node --test)
```

To run the add-in in Excel on macOS:

```sh
npm run certs -w @ccg/addin      # one-time: install localhost dev certificates
npm run dev -w @ccg/addin        # serve the task pane on https://localhost:3000
npm run sideload -w @ccg/addin   # copy the manifest into Excel's wef folder
```

Then restart Excel — the **CCG** group with **Analyse Cell** appears on the Home tab
(if not, check Insert → Add-ins → My Add-ins → Developer Add-ins).

The same add-in runs unchanged in Excel on the web, Windows, and iPad once hosted on a
public HTTPS URL — see [docs/roadmap.md](docs/roadmap.md) for the distribution plan.

## Troubleshooting

Corporate Microsoft 365 tenants can block web add-ins in several independent, slowly-
propagating ways (store policy, connected-experiences Cloud Policy, feature-flight caches).
If the add-in won't appear, the Add-ins pane says "disabled by your IT administrator", or
LAMBDA returns #NAME? on a current build, see **[docs/troubleshooting.md](docs/troubleshooting.md)**
— it was earned the hard way.

## Contributing

Contributions are very welcome! Please read [CONTRIBUTING.md](CONTRIBUTING.md) and the [Code of Conduct](CODE_OF_CONDUCT.md).

## Author

**Angelo Perera**\
GitHub: [@aceslick911](https://github.com/aceslick911)

---

*CCG is the spiritual successor to the 2013 Cell Analysis Add-in — rebuilt from the ground up as open source.*

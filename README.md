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

- Formula visualizer with clean mathematical notation
- Navigate between cells and drill into precedents/dependents
- Color marking system with option to apply marks back to the worksheet
- One-click conversion of formulas to C#, Visual Basic, or Delphi code
- Define variables and functions for cleaner code generation
- Modern .NET implementation (targeting current Excel / Office Add-in model)

## License & Attribution

This project is licensed under the **Apache License 2.0**.

**Attribution Requirement**\
If you use this software or substantial portions of it in your own Excel add-in, commercial product, or derivative work, you **must** provide clear and prominent attribution to the original author:

> Includes CCG (exCel Code Generator) functionality originally developed by Angelo Perera\
> https://github.com/aceslick911/CCG

This requirement applies even if you modify the code.

See the [NOTICE](NOTICE) file and [LICENSE](LICENSE) for full details.

## Getting Started

*(Installation and build instructions will be added once the first working version is ready)*

## Contributing

Contributions are very welcome! Please read [CONTRIBUTING.md](CONTRIBUTING.md) and the [Code of Conduct](CODE_OF_CONDUCT.md).

## Author

**Angelo Perera**\
GitHub: [@aceslick911](https://github.com/aceslick911)

---

*CCG is the spiritual successor to the 2013 Cell Analysis Add-in — rebuilt from the ground up as open source.*

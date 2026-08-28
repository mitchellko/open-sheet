# open-sheet

<img src=".github/assets/banner-light.png" alt="open-sheet — a dependency graph drawn across a spreadsheet grid: two cells feeding one" width="100%">

[![npm](https://img.shields.io/npm/v/@open-sheet/core?style=flat&label=%40open-sheet%2Fcore)](https://www.npmjs.com/package/@open-sheet/core)
[![CI](https://github.com/lianghsun/open-sheet/actions/workflows/ci.yml/badge.svg)](https://github.com/lianghsun/open-sheet/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg?style=flat)](https://opensource.org/licenses/MIT)

**English** · [繁體中文](README.zh-TW.md)

**The spreadsheet framework built for agents.** Describe the model you need in natural language — your coding agent writes the React. open-sheet handles the cell addressing, the formula references, the recalculation, and the export.

If [open-slide](https://github.com/1weiho/open-slide) is Google Slides for agents and [open-doc](https://github.com/simonliu-ai-product/open-doc) is Google Docs, open-sheet is **Google Sheets**. Same idea, different medium — but the medium changes the problem.

> A deck is pixels. A document is pixels on paper. **A spreadsheet is a dependency graph.**
> A workbook baked down to static numbers is just a picture of a table. The value of a spreadsheet is that the person who receives it can change an assumption and watch it recalculate.

```bash
npx @open-sheet/cli init my-sheets
```

> Practice exercise note: this change is only for local Git workflow practice.

## Why

Agents write excellent analysis and terrible spreadsheets, and the reason is specific: **`=SUM(B2:B13)`**.

Cell addresses are the one thing an agent cannot hold onto. It miscounts the header row. It forgets the total row shifted when the data grew. It writes a reference to `B7` and then inserts a quarter above it, and now every formula in the file is quietly wrong — not crashed, just *wrong*, which is worse.

So open-sheet takes the addresses away. You never write one:

```tsx
col('grossProfit', {
  header: 'Gross profit',
  formula: (r) => sub(r.cell('revenue'), r.cell('cogs')),
})
```

The framework owns every coordinate. Add a row to the data array and every reference re-resolves. **There is not a single A1 address in an open-sheet source file.**

## The lineage

Each of the three frameworks absorbs the one chore that its medium makes hard for an agent:

| | open-slide | open-doc | **open-sheet** |
| --- | --- | --- | --- |
| Medium | 1920 × 1080 canvas | A4 sheets | the cell grid |
| The chore it absorbs | scaling, navigation, present mode | pagination, contents, page numbers | **cell addressing, formula references, recalculation** |
| You get | a deck you can present | a PDF that survives a printer | **a model a CFO can change** |

## Highlights

### 🔢 References, not addresses

`r.cell('revenue')` · `r.prev().cell('revenue')` · `ref('pl').column('revenue')` · `ref('pl').total('revenue')` · `ref('assumptions').get('growth')`

These resolve to A1 only at the very end, after layout has decided where everything sits. Scalar assumptions are additionally emitted as **Excel defined names**, so the exported workbook reads `=B4*growth` — legible to the human who opens it, not just to the machine.

### 📐 Auto-placement

`<Stack>` and `<Row>` pack blocks onto the grid without collisions. You describe the *order*; the framework decides the *coordinates*. This is open-sheet's answer to open-doc's `flow()`.

### 🧮 One formula tree, two backends

Every formula is an expression AST with two consumers: `serialize()` writes an Excel formula string into the `.xlsx`, and `evaluate()` computes the number the viewer shows. They are checked against each other in CI by **recalculating the exported workbook in LibreOffice and diffing** — which proves both that the export contains live formulas and that our evaluator agrees with a real spreadsheet engine.

Formulas we can't evaluate are shown as `#NOT_EVALUATED`, never as an invented number.

### 📤 Live `.xlsx`, plus `.csv` / `.html` / `.pdf`

The `.xlsx` contains **formulas, not baked values** — with number formats, defined names, frozen panes, and conditional formatting. One style model feeds both the Excel writer and the HTML/PDF renderer, so the printed report and the workbook match.

### 🤖 Agent-native

Skills ship with the scaffolder: `/create-sheet`, `/sheet-authoring`, `/current-sheet`, `/apply-comments`. An MCP server (`open-sheet dev --mcp`) lets any agent framework drive it. Inspect mode lets you click a cell to see its source line, its resolved formula, and its computed value — or leave a note for your agent.

## Status

**Early development.** Nothing is published to npm yet. Follow the [milestones](https://github.com/lianghsun/open-sheet/milestones).

Everything described above is built and tested — the compiler, references and the
formula engine, the viewer and dev server, the skills, the MCP server, inspect
mode, themes, the design panel, native charts, and all four export formats.
144 tests, including two that drive a real spreadsheet application.

Published as `0.1.0`:

```bash
npx @open-sheet/cli init my-sheets
cd my-sheets && npm install && npm run dev
```

The proof-of-life, which now holds end to end: `apps/demo` exports a workbook
whose net-income column is `=F6*(1-taxRate)`, and a test changes `taxRate` and
asserts the whole column moves — with LibreOffice, not open-sheet, doing the
arithmetic.


Trying it before it ships: [TESTING.md](TESTING.md).

## Repo layout

pnpm + Turbo monorepo.

| Path | Description |
| --- | --- |
| `packages/core` | `@open-sheet/core` — compiler, placement, references, formula engine, viewer, Vite plugin, and the `open-sheet` CLI. |
| `packages/cli` | `@open-sheet/cli` — `npx @open-sheet/cli init` scaffolder + template. |
| `packages/mcp` | `@open-sheet/mcp` — MCP server over Streamable HTTP. |
| `apps/demo` | Example workspace consuming `@open-sheet/core` via `workspace:*`. Dogfood target. |

## Development

```bash
pnpm install
pnpm dev        # runs the demo against the local @open-sheet/core
pnpm build      # builds all packages
pnpm typecheck  # tsc across the graph
pnpm check      # biome (format + lint + organize imports)
pnpm test       # vitest
```

## Credits

open-sheet follows [open-slide](https://github.com/1weiho/open-slide) by [@1weiho](https://github.com/1weiho) and [open-doc](https://github.com/simonliu-ai-product/open-doc) by [@simonliu-ai-product](https://github.com/simonliu-ai-product) — the virtual-module content discovery, the scaffolder, and the skills-as-documentation approach are theirs. This is the third medium.

## License

MIT

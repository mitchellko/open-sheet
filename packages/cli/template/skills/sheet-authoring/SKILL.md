---
name: sheet-authoring
description: Technical reference for writing open-sheet workbooks — the file contract, the component surface, references and formulas, number formats, and the rules that keep a workbook a live model. Use this whenever writing or editing a file under `sheets/<id>/index.tsx`. The workflow for drafting a new workbook from scratch lives in the `create-sheet` skill.
---

# Authoring an open-sheet workbook

## The one rule

**Never write a cell address.**

Not `B5`. Not `SUM(B2:B13)`. Not `$A$1`. The framework assigns every coordinate
*after* it has laid the blocks out, and references resolve against that layout.
Writing an address by hand means writing a number that is correct only until
someone adds a row — and then it is silently wrong, which in a financial model is
worse than broken.

If you catch yourself counting rows to work out where something landed, stop.
The thing you want is a reference.

The single exception is `raw('...')`, the deliberate escape hatch — see
[references/formulas.md](references/formulas.md).

## The file contract

```tsx
// sheets/<id>/index.tsx
import { Sheet, Table, Workbook, col, sub, type SheetMeta } from '@open-sheet/core'

export const meta: SheetMeta = {
  title: 'FY26 Budget',        // shown in the viewer and used as the export name
  description: 'One line.',    // optional
  theme: 'corporate-neutral',  // optional; links back to themes/<id>.md
  createdAt: '2026-08-18T00:00:00.000Z',
}

export const design: DesignSystem = { … }   // optional; see the Design panel note

export default (
  <Workbook>
    <Sheet name="…">…</Sheet>
  </Workbook>
)
```

The default export must be a `<Workbook>` containing `<Sheet>` children. `meta`
and `design` are plain object literals — the dev UI parses and rewrites them, and
a spread or an imported value makes the workbook untweakable.

## Structure in JSX, data in TypeScript

JSX describes the *report*. Rows are plain arrays.

```tsx
const quarters = [
  { quarter: 'Q1', revenue: 12_400_000, cogs: 5_100_000 },
  { quarter: 'Q2', revenue: 13_900_000, cogs: 5_600_000 },
]

<Table
  name="pl"
  data={quarters}
  columns={[
    col('quarter', { header: 'Quarter', width: 12 }),
    col('revenue', { header: 'Revenue', format: 'currency' }),
    col('grossProfit', {
      header: 'Gross profit',
      format: 'currency',
      formula: (r) => sub(r.cell('revenue'), r.cell('cogs')),
    }),
  ]}
  total={{ revenue: 'sum', grossProfit: 'sum' }}
/>
```

Do not write a thousand rows of JSX. If the data is generated, generate the
array and hand it to `data`.

## The component surface

| Component | What it is |
| --- | --- |
| `<Workbook>` | The root. Children must be `<Sheet>`. |
| `<Sheet name freeze origin>` | One tab. `freeze="B2"` freezes the rows above and columns left of that cell. |
| `<Stack gap>` | Stacks blocks downward. `gap` is in rows (default 1). |
| `<Row gap>` | Places blocks side by side. `gap` is in columns. |
| `<Table>` | A named data block. See below. |
| `<KpiBand items>` | A label row above a value row — a headline strip. |
| `<Cell value formula format style span>` | One cell. |
| `<Note cols>` | A line of prose spanning `cols` columns. |
| `<Spacer rows cols>` | Deliberate empty space. |

`<Table>` has two shapes:

- **grid** (default) — `name`, `data`, `columns`, optional `title`, `showHeader`,
  `total`
- **key-value** — `kind="keyValue"` with `data={[{ key, label, value, format }]}`.
  Every `key` becomes an **Excel defined name**, so formulas referencing it read
  `=B5*growth` in the exported file. This is how assumptions should be written.

`name` is workbook-global, because `ref()` looks blocks up by name. A duplicate
throws at compile time.

`total` applies to **grid** tables only — a key-value block is a list of named
scalars, not a column to aggregate. To total a key-value block, reference the
entries you want and add them.

## Separate assumptions from calculations

Put every number a reader might want to change on its own sheet, in a
`kind="keyValue"` table, and reference it. That is what makes the export a model
rather than a picture of one.

```tsx
<Sheet name="Assumptions">
  <Table name="assumptions" kind="keyValue" data={[
    { key: 'growth',  label: 'QoQ growth', value: 0.08, format: 'percent' },
    { key: 'taxRate', label: 'Tax rate',   value: 0.2,  format: 'percent' },
  ]} />
</Sheet>
```

A number hard-coded inside a formula is a number the recipient cannot change.

## Further reference

- [references/placement.md](references/placement.md) — how blocks are sized and placed
- [references/formulas.md](references/formulas.md) — references, the builder API, the function whitelist
- [references/formats.md](references/formats.md) — number formats, styles, data bars, themes
- [references/charts.md](references/charts.md) — what is live and what is not
- [references/printing.md](references/printing.md) — headers, footers, page breaks, print areas
- [references/recipients.md](references/recipients.md) — validation and the other affordances the person opening the file gets

## What the framework cannot check for you

The framework guarantees **referential integrity**: addresses are right, formulas
point where you meant, nothing breaks when a row is inserted. It cannot guarantee
that the things being compared are **comparable**.

Every failure of this kind looks identical to a correct result. There is no error,
no `#NOT_EVALUATED`, no trace in any cell — just a number that is arithmetically
right and analytically wrong. Three shapes seen in real workbooks:

| Shape | Example |
| --- | --- |
| Ordering decided in `data` | a sorted array: right this month, silently stale next month |
| Periods of different length | two days of data beside thirty → a growth rate of +790% |
| Sources with different coverage | 30 days of cost ÷ 20 days of requests → wrong by 50% |

So whenever two numbers go into one expression — a division, a difference, a
ranking — ask once: **do they cover the same ground?**

### Normalise before you combine

When two figures come from different sources, divide each by its own coverage
first. The spans cancel, and what is left is comparable:

```tsx
// Wrong: two totals carrying different numbers of days
col('costPerRequest', { formula: (r) => div(r.cell('cost'), r.cell('requests')) })

// Right: each normalised first
col('dailyCost',      { formula: (r) => div(r.cell('cost'), r.cell('costDays')) })
col('dailyRequests',  { formula: (r) => div(r.cell('requests'), r.cell('requestDays')) })
col('costPerRequest', { formula: (r) => div(r.cell('dailyCost'), r.cell('dailyRequests')) })
```

Two extra columns, and both earn their place as diagnostics: a daily cost of 810,
1,225, 1,068 across three months shows which one is off. Collapsed into a single
`costPerRequest`, it does not.

Make the coverage gap itself a column when the sources may disagree — a visible
number with a threshold flag beats a footnote nobody reads.

## Self-review before finishing

- [ ] No A1 address anywhere in the file
- [ ] Every number a reader might change lives in an assumptions block, not inside a formula
- [ ] Every `col` that computes uses `formula`, not a pre-computed value in `data`
- [ ] **Nothing that depends on the values is decided in the `data` array** — no
      sorting, filtering, grouping, or top-N. Those must be formulas. Easier to
      miss than a pre-computed value, because a sorted array leaves no trace in
      any cell: the workbook is right this month and quietly wrong next month,
      while still looking like a sorted table.
- [ ] `r.prev()` / `r.next()` are guarded with `r.isFirst` / `r.isLast`
- [ ] Formats are set on the columns that need them — a bare `0.0829` reads as noise
- [ ] Nothing was invented: every figure came from the user or a named source
- [ ] **Everything compared covers the same ground** — see "What the framework
      cannot check for you". Provenance is not comparability, and this class of
      error leaves no trace in any cell. Prefer **not generating** a comparison
      over generating one with a caveat: a caveat gets skipped, a column that
      does not exist cannot be misread.
- [ ] The viewer shows no unexpected `#NOT_EVALUATED`

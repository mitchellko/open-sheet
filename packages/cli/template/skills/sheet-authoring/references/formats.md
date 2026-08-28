# Formats, styles, and themes

## Number formats

Set `format` on a column, a `<Cell>`, or a KPI item. Named formats resolve to
Excel codes and are interpreted identically by the HTML renderer:

| Name | Code | Renders |
| --- | --- | --- |
| `number` | `#,##0` | `12,400,000` |
| `decimal` | `#,##0.00` | `1,234.50` |
| `currency` | `#,##0` | `12,400,000` |
| `currency2` | `#,##0.00` | |
| `accounting` | `_(* #,##0_);…` | negatives in parentheses |
| `percent` | `0.0%` | `60.3%` |
| `percent0` / `percent2` | `0%` / `0.00%` | |
| `thousands` | `#,##0,"K"` | `12,400K` |
| `millions` | `#,##0,,"M"` | `12M` |
| `date` | `yyyy-mm-dd` | |
| `text` | `@` | forces text |

Any other string is passed through as a literal Excel format code.

**A ratio without a format reads as noise.** `0.6029159519725558` in a report is
a defect. Put `format: 'percent'` on it.

## What each output does with a format

| Output | Formats applied? |
| --- | --- |
| viewer, `.html`, `.pdf` | yes — rendered as Excel would |
| `.xlsx` | yes — the code is written and Excel applies it |
| `.csv` | **no** — raw values at full precision |

CSV is data, not presentation. A `percent` column exports as
`0.5352386237513873`, not `53.5%`, because something downstream is going to do
arithmetic on it and a rounded string would be the wrong thing to hand over.

## Styles

Cells carry a style key resolved against the active theme. The compiler assigns
sensible ones (`tableHeader`, `tableTotal`, `kpiLabel`, `kpiValue`, `kvLabel`,
`kvValue`, `note`, `tableTitle`); override with `style` on a column or `<Cell>`.

Every style is one `CellStyle` — font, fill, alignment, borders, number format —
translated to CSS for the viewer and HTML, and to Excel formatting for the xlsx.
Define it once; you never write it twice.

## Data bars

```tsx
col('revenue', { header: 'Revenue', format: 'currency', bar: true })
col('variance', { bar: { color: '#16a34a' } })
```

Emits a **native** Excel `dataBar` rule over the column's data range, so it
rescales when the numbers change, and the same rule renders as a gradient in
HTML. Use it where a reader is comparing magnitudes down a column.

## Long text

Excel does not wrap. A description column set narrower than its content spills
into the neighbouring cell, or is clipped when printed — and no test that reads
values will notice:

```tsx
col('spec', { header: '說明', width: 30, wrap: true })
```

Set it on the columns that hold sentences, not on the whole table: wrapping a
figures column just makes the rows taller.

## Printing

A workbook of grids and a workbook of forms want opposite defaults, so the sheet
says which it is:

```tsx
<Sheet name="請款單" print={{ orientation: 'portrait', size: 'A4', fitToWidth: true, repeatHeader: true }}>
```

| Option | What it does | When |
| --- | --- | --- |
| `orientation` | `'portrait'` or `'landscape'` (default) | **forms are portrait** — a landscape invoice is unusable |
| `size` | `'A4'` (default), `'A3'`, `'Letter'`, `'Legal'` | regional |
| `fitToWidth` | scale to one page wide | any sheet meant to be printed |
| `repeatHeader` | repeat the table header on every page | any table longer than a page |
| `margin` | inches, all four sides | tight forms |

Without this, Excel prints landscape at 100%, so a form comes out sideways and a
long table's second page arrives with no header. Neither is recoverable by the
person holding the paper.

The HTML and PDF exports follow the same declaration: if any sheet asks for
portrait, the document prints portrait, since a form printed sideways is unusable
while a grid merely wraps.

## Themes

`themes/<id>.md` is a house style — palette, type scale, and paste-ready
components — with an optional `<id>.demo.tsx` the gallery previews. Link a
workbook to one with `meta.theme`.

A theme that omits a style key falls back to the default for that key, so a
half-written theme degrades to plain rather than to nothing.

## Writing for the Design panel

Keep `design` a plain object literal. The panel parses and rewrites it through an
AST edit; a spread, a computed key, or a value imported from elsewhere makes the
workbook untweakable.

## Conditional formatting

Four kinds, all **live** — the rule travels into the file, so it stays true when
the numbers change. A rule evaluated at export time and written as a static fill
is the same mistake as a chart pasted in as an image.

```tsx
col('margin',   { scale: ['#fee2e2', '#ffffff', '#dcfce7'] })
col('trend',    { icons: 'arrows' })
col('variance', { highlight: { above: 0, fill: '#dcfce7' } })
col('amount',   { bar: true })
```

`highlight` takes one rule or an array applied in order, so a later rule wins
where two overlap. The test is one of:

| | |
| --- | --- |
| `above` · `below` · `atLeast` · `atMost` | numeric comparison |
| `between: [a, b]` | inclusive |
| `equals` | number, string or boolean |
| `contains` | case-insensitive substring |
| `duplicates: true` | appears more than once in the column |
| `top: n` · `bottom: n` | rank within the column |

and the look is any of `fill`, `color`, `bold`.

`scale` takes two or three stops, lowest value to highest. `icons` is `'arrows'`
or `'trafficLights'`, split into thirds of the column's range.

**Both renderers draw the same thing**, and that is a constraint rather than a
nicety — the viewer and Excel disagreeing about one cell is a bug that has
happened here three times. Two consequences worth knowing:

- The colour scale's midpoint is the **linear** midpoint of the range, not the
  median. Excel's own default is the median, which the HTML export could not
  reproduce without reimplementing Excel's percentile; the two would then drift
  apart on any column that is not evenly distributed.
- A column where every value is identical gets the **first** stop, not the
  middle one, because that is what Excel paints.

A bar and a highlight can share a column: one is a background image and the
other a background colour, so neither is lost.

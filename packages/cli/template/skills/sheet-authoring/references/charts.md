# Charts

## The rule

A chart embedded as an image stops being true the moment someone changes a
number — and the whole point of an open-sheet export is that the recipient *can*
change a number. So:

**Anything in the .xlsx must recalculate. If it cannot, it does not go in.**

`<Chart>` writes real OOXML chart parts bound to real cell ranges. It is not a
picture. Change a cell in Excel and the chart moves.

## Using it

```tsx
<Chart
  kind="line"                                   // 'bar' | 'line' | 'pie'
  title="Trials by month"
  categories={ref('funnel').column('month')}
  series={[
    { name: 'Trials', values: ref('funnel').column('trials') },
    { name: 'Paid',   values: ref('funnel').column('paid') },
  ]}
  rows={16}
  cols={7}
/>
```

`categories` and `series[].values` are ordinary references, so they resolve after
layout like everything else — add a row to the data and the chart's range grows
with it.

`rows` and `cols` are the chart's footprint on the grid. The placement engine
treats it as any other block, so it stacks and gaps normally.

## What renders where

| Output | What you get |
| --- | --- |
| `.xlsx` | a native chart bound to ranges — live |
| viewer, `.html`, `.pdf` | an SVG drawn from the same evaluated values |
| `.csv` | nothing; a chart is not data |

The SVG twin reads the values the grid shows, so the two cannot disagree.

## Also live, and often better

For comparing magnitudes down a single column, `col(key, { bar: true })` is
usually the clearer choice — a native `dataBar` sits in the cells themselves, so
there is no chart to position and nothing to fall out of date.

## Structuring data for charts

- Keep a series in one contiguous column with a header
- Put category labels in their own column
- No blank rows inside a table — `<Spacer>` goes between blocks, not inside one

## The chart kinds

| | |
| --- | --- |
| `bar` · `stackedBar` | comparison; stacked for composition over time |
| `line` | a trend |
| `area` · `stackedArea` | the same, filled — stacked is most cost reporting |
| `pie` | a share of one total |
| `scatter` | two measures against each other; `categories` becomes the x values |
| `combo` | bars with a line over them — actual against target |

## Making a chart readable rather than decorative

```tsx
<Chart
  kind="bar"
  categories={ref('pl').column('quarter')}
  series={[{ name: 'Revenue', values: ref('pl').column('revenue') }]}
  axes={{ category: 'Quarter', value: 'NT$', valueFormat: 'currency', min: 0 }}
  dataLabels
/>
```

An axis with no title and unformatted numbers is decoration: the reader cannot
tell thousands from millions, or margin from revenue. `valueFormat` takes the
same codes cells take. `min`/`max` pin the axis — Excel's automatic zero-based
scale flattens a series that lives in a narrow band.

`dataLabels` is off by default: on a dense series the numbers collide into an
unreadable smear.

## Combo, and the second axis

```tsx
series={[
  { name: 'Actual', values: ref('pl').column('revenue'), as: 'bar' },
  { name: 'Target', values: ref('pl').column('attainment'), as: 'line', axis: 'secondary' },
]}
```

Each series says how it is drawn. `axis: 'secondary'` gives it the right-hand
axis, titled with `axes.secondary` — a percentage next to revenue in dollars is
otherwise a flat line along the bottom of the plot. Bars are drawn before the
line, so the line sits over them.

## Sparklines

```tsx
col('trend', { header: 'Trend', sparkline: { of: ['q1', 'q2', 'q3', 'q4'] } })
```

An in-cell chart of that row's own numbers, one per data row. `kind` is `'line'`
or `'column'`.

The columns it reads **must be next to each other**. The file format stores one
range per sparkline, so a gap would silently pull in whatever sits between them
— possibly a text column. A non-contiguous set is refused at compile time with
an error naming the columns.

Not every reader keeps them: sparklines are a 2010 extension, and an older
application will show an empty cell rather than a broken one. The number they
summarise is in the row beside them either way.

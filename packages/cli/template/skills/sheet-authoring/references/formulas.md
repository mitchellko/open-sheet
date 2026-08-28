# References and formulas

## References

None of these carry an address. They are descriptions, resolved after layout.

| Reference | Points at |
| --- | --- |
| `r.cell('revenue')` | the named column, same row |
| `r.prev().cell('revenue')` | the row above — **guard with `r.isFirst`** |
| `r.next().cell('revenue')` | the row below — **guard with `r.isLast`** |
| `r.index`, `r.isFirst`, `r.isLast`, `r.data` | position and the row's own object |
| `ref('pl').column('revenue')` | that column's whole data range |
| `ref('pl').total('revenue')` | that column's total cell (needs `total={{ … }}`) |
| `ref('pl').cell('revenue', 0)` | a specific data row |
| `ref('pl').body()` | the whole data area |
| `ref('assumptions').get('growth')` | a key-value entry, by its defined name |

`r` is the argument to a column's `formula`. `ref(name)` works anywhere,
including across sheets — the qualifier is added for you.

## Letting the data decide which rows compute

`r.data` is the row's own object, so a flag on the data can decide whether a cell
computes at all — which keeps "who does this apply to" a property of the data
rather than a condition buried in a formula:

```tsx
const MONTHS = [
  { key: 'm202605', label: '2026-05*', partial: true },   // export started mid-month
  { key: 'm202606', label: '2026-06', partial: false },
]

col('mom', {
  formula: (r) => (r.isFirst || r.data.prevPartial ? null : sub(div(…), 1)),
})
```

Returning `null` leaves the cell empty, which is the honest answer when the
comparison is not meaningful. Better still, do not generate the column at all —
see below.

## Columns are an ordinary array

`columns` is computed like any other array, which is how period-as-column layouts
stay maintainable:

```tsx
columns={[
  col('service', { header: 'Service' }),
  ...MONTHS.map((m) => col(m.key, { header: m.label, format: 'currency' })),
]}
```

You can filter as well as map. A comparison whose baseline is a partial period is
better **not produced** than produced with a caveat — a caveat gets skipped, a
column that does not exist cannot be misread:

```tsx
...MONTHS.slice(1)
  .filter((_, i) => !MONTHS[i]?.partial)
  .map((m) => col(`mom_${m.key}`, { … }))
```

**Watch the index after a filter.** `.filter().map()` gives you the position in
the *filtered* array, not the original. Use `MONTHS.indexOf(m)` when you need the
original position — pairing a column with the wrong month is easy to write and
usually only shows up in the header label.

## `r.cell(key)` vs `r.data.key`

`r.cell('revenue')` points at a **cell**, so the exported formula reads `B5` and
the recipient can change it. `r.data.revenue` reads the **raw value**, which is
baked into the formula as a literal:

```tsx
formula: (r) => div(r.cell('mar'), r.cell('prev'))   // → =D6/E6
formula: (r) => div(r.cell('mar'), r.data.prev)      // → =D6/128400
```

Both compute the same number here. Only the first stays a model.

Use `r.data` when the value genuinely is not part of the model — a flag deciding
*which* formula to build, a label, a lookup key. If it is a number the reader
might want to change, give it a column so it lands on the grid.

A field that exists in your `data` array but has no `col()` has no cell to point
at, and `r.cell()` will say so.

Using `r.prev()` on the first row is an error at resolve time naming the guard
you forgot, not a silent `#REF!`.

## Building expressions

```tsx
sub(r.cell('revenue'), r.cell('cogs'))
div(sum(ref('pl').column('grossProfit')), sum(ref('pl').column('revenue')))
mul(r.cell('operatingIncome'), sub(1, ref('assumptions').get('taxRate')))
if_(gt(r.cell('revenue'), 0), div(r.cell('cogs'), r.cell('revenue')), 0)
```

**Operators** — arithmetic `add` `sub` `mul` `div` `pow` `neg`, text `concat`
(that is `&`), comparison `eq` `neq` `lt` `gt` `lte` `gte`.

**Functions**, by what you reach for them for:

| | |
| --- | --- |
| Aggregate | `sum` `avg` `count` `counta` `min` `max` `product` `subtotal` `aggregate` |
| Conditional | `if_` `ifs` `switch_` `and` `or` `not` `xor` `iferror` `ifna` |
| Conditional aggregate | `sumif` `sumifs` `countif` `countifs` `averageif` `averageifs` `maxifs` `minifs` |
| Rounding | `round` `roundup` `rounddown` `ceiling` `floor` `int` `trunc` `mod` `sign` `abs` |
| Math | `sqrt` `power` `exp` `ln` `log` `log10` `sumproduct` |
| Lookup | `lookup` `index` `match` `choose` `large` `small` |
| Statistics | `median` `mode` `rank` `percentile` `quartile` `stdev` `stdeva` `var_` `varp` `correl` `slope` `intercept` `forecast` `trend` |
| Finance | `npv` `xnpv` `irr` `xirr` `pmt` `ipmt` `ppmt` `fv` `pv` `rate` `nper` `sln` `db` `ddb` `syd` |
| Text | `len` `left` `right` `mid` `find` `search` `substitute` `replace` `trim` `upper` `lower` `proper` `rept` `text` `value` `textjoin` `join` `concatenate` |
| Dates | `today` `now` `date` `year` `month` `day` `hour` `minute` `weekday` `weeknum` `edate` `eomonth` `days` `datedif` `yearfrac` `networkdays` `workday` |
| Tests | `isblank` `isnumber` `istext` `iserror` `isna` `iseven` `isodd` |

Every one of these is checked against LibreOffice on each build — a function is
on this list only because our value and a real spreadsheet's agree.

**Naming.** A builder is Excel's name lowercased, all one word — `sumproduct`,
`sortby`, `networkdays`, not `sumProduct`. Three deviate, because the obvious
name was taken: `var_` and `switch_` and `if_` (JavaScript keywords), and `join`
for Excel's `CONCAT`, since `concat` is already the `&` operator.

**`text(value, code)` takes the same format codes a cell does**, date codes
included — `text(edate(start, n), 'yyyy-mm')`. One caveat: month and day
*names* (`mmm`, `dddd`) render in English here, while Excel renders them in the
reader's own language. Numeric codes agree everywhere, so prefer `yyyy-mm-dd`
over `d mmm yyyy` when the reader's locale is not yours.

Bare numbers, strings, and references all lift automatically — a reference can go
anywhere an expression can, including a KPI value or a whole column formula:

```tsx
{ label: 'Total cost', value: ref('costs').total('total') }
col('mirror', { formula: (r) => r.cell('amount') })
```

## Periods as columns

The reference examples above put periods in rows. Cost and budget analysis
usually does the opposite — one row per account, one column per month — and that
layout needs two things the row-wise examples never show.

**`sum` is variadic**, so a row total is a spread, not a range:

```tsx
const MONTHS = [
  { key: 'jan', header: 'Jan' },
  { key: 'feb', header: 'Feb' },
  { key: 'mar', header: 'Mar' },
]

const services = [
  { service: 'Compute', jan: 128_400, feb: 141_200, mar: 155_900 },
  { service: 'Storage', jan: 22_100, feb: 21_800, mar: 24_600 },
]

<Table
  name="costs"
  data={services}
  columns={[
    col('service', { header: 'Service', width: 20 }),
    ...MONTHS.map((m) => col(m.key, { header: m.header, format: 'currency' })),
    col('total', {
      header: 'Q1',
      format: 'currency',
      formula: (r) => sum(...MONTHS.map((m) => r.cell(m.key))),
    }),
    col('mom', {
      header: 'MoM',
      format: 'percent',
      // Month-on-month reads sideways: this column against the one before it.
      formula: (r) => sub(div(r.cell('mar'), r.cell('feb')), 1),
    }),
  ]}
  total={Object.fromEntries(MONTHS.map((m) => [m.key, 'sum' as const]))}
/>
```

**Two tables fed the same array line up row for row**, so a second sheet can
reference across without any alignment work:

```tsx
// On another sheet, same `services` array, same order:
col('share', {
  formula: (r) => div(ref('costs').cell('total', r.index), ref('costs').total('total')),
})
```

`r.index` is the row's position in the data, which is what makes this safe —
insert a service and both tables move together.

## Looking a value up in another table

`lookup()` names the columns instead of counting them:

```tsx
col('price', {
  formula: (r) =>
    lookup({ value: r.cell('sku'), from: 'products', match: 'sku', get: 'price', ifMissing: 0 }),
})
```

compiles to `INDEX(…, MATCH(…, 0))`. It is not `VLOOKUP` on purpose: `VLOOKUP`
takes a **positional** column index, so inserting a column in the lookup table
silently repoints it — the exact failure this framework exists to remove. It also
requires the matched column to be leftmost, which is not the author's choice to
make.

Without `ifMissing` an unmatched row reads `#N/A`, which is Excel's answer and
often the right one: a missing match in a reconciliation should be loud. Supply
`ifMissing` when a blank is genuinely the answer.

## Conditional aggregation

`sumif` `countif` `averageif` take Excel's criteria syntax as a string:

```tsx
formula: () => sumif(ref('costs').column('amount'), '>1000')
formula: () => countif(ref('costs').column('service'), 'Cloud Run')
```

The criteria language (`">100"`, `"<>done"`) is passed through as written. It is
a syntax Excel already defines and every spreadsheet user already knows, so
inventing a builder for it would add a dialect rather than remove one.

## Guarding a division

Two ways, and both are honest. Use whichever reads better:

```tsx
// return null: the cell is simply empty
formula: (r) => (r.isFirst ? null : sub(div(r.cell('cur'), r.prev().cell('cur')), 1))

// iferror: keep the formula, name what happens when it cannot compute
formula: (r) => iferror(sub(div(r.cell('cur'), r.cell('prev')), 1), '')
```

**Do not "fix" a `#DIV/0!` by padding the denominator.** `max(prev, 1)` turns an
honest blank into a number that looks real and is not. `iferror` exists so you
never have to.

Return `null` from a `formula` to leave the cell empty — that is how a
first-row growth figure should be written:

```tsx
formula: (r) => (r.isFirst ? null : sub(div(r.cell('revenue'), r.prev().cell('revenue')), 1))
```

## A formula that fills more than one cell

`SORT`, `FILTER`, `UNIQUE`, `SORTBY`, `SEQUENCE`, and `TRANSPOSE` return a
rectangle, not a value. They need `<Spill>`, which reserves the room:

```tsx
<Spill formula={sort(ref('reps').column('revenue'), 1, -1)} rows={3} cols={1} />
```

`rows` and `cols` are required and not inferred. The size of a `FILTER` result
is not knowable until the file is recalculated, and a layout engine that guessed
would be guessing about collisions too — so you declare the footprint and the
placement engine reserves exactly that.

The rectangle is a contract, not a suggestion. Excel would *spill* such a result
and grow to fit; we emit an array formula over the declared range, so the result
fills that range and stops. Cells the result does not reach show `#N/A` — the
same thing a spreadsheet shows, and visibly not a number. If you ask for three
rows and the filter matches five, you see three: size the footprint for the
largest case you expect.

Two of them, `FILTER` and `SEQUENCE`, are not in our function library. They
export correctly and Excel computes them; the viewer shows `#NOT_EVALUATED`
across the footprint until then.

**Who can compute the result.** These are recent additions to the file format.
Excel has them, and so does a current LibreOffice — but an older LibreOffice
shows `#NAME?` across the whole footprint. `TRANSPOSE` is the exception; it has
been there since the beginning. If the reader's spreadsheet application is not
something you control, sort in the `data` array instead and let `<Spill>` be a
convenience rather than the thing the workbook depends on.

## The whitelist, and `raw()`

The builders above are the functions open-sheet can both write *and* evaluate.
Anything else goes through the escape hatch:

```tsx
formula: () => raw('=BESSELJ(A1,1)')
```

`raw()` exports verbatim and works in Excel. It is **not evaluated here**, so the
viewer shows `#NOT_EVALUATED` and the CSV cell is empty. That is deliberate: a
plausible-looking wrong number in a financial model is the worst failure this
project can have, so it never guesses.

Prefer a whitelisted expression. Reach for `raw()` when the function genuinely
has no equivalent, and say so in a `<Note>` if the reader will wonder.

Check the table above before reaching for it — `xirr`, `pmt`, and `sumifs` are
all supported, and a `raw()` cell that could have been an expression shows
`#NOT_EVALUATED` to your reader for no reason.

## Formula strings

A string like `"=A1+B2"` is parsed where possible so it still evaluates, but it
is not the recommended path — it is exactly the thing that breaks when a row is
inserted. Use references.

## What happens after you hand the file over

By default a table exports as ordinary A1 ranges. That decides what a recipient
can safely do:

| They do this in Excel | Ranges follow? |
| --- | --- |
| Insert a row **inside** the data | **Yes** — Excel rewrites `B2:B13` to `B2:B14` |
| Change a value | Yes, everything recalculates |
| **Append** a row below the last | **No** — the range still ends where it did |
| Delete a row inside the data | Yes |

So "insert a row above the total" is safe advice; "add new rows at the bottom" is
not — nothing breaks visibly, the total just stops including the new row.

### Unless you say `appendable`

```tsx
<Table name="costs" appendable data={costs} total={{ amount: 'sum' }} … />
```

This exports an Excel Table, and every whole-column reference is written as
`costs[Amount]` instead of `B2:B13`. Excel takes a new row into the table and
every structured reference follows it, derived columns included — a row the
reader adds computes itself.

**How they add it depends on whether the table has a total row**, and the
difference is not guessable:

| | No total row | With a total row |
| --- | --- | --- |
| Type below the last row | **works** | does nothing — the table does not grow |
| Tab from the last cell | works | does nothing — Tab moves to the total |
| **Insert** a row above the total | works | **works** |

So a table with a total needs "insert a row above the total", not "add rows at
the bottom". Nobody reads documentation before they start typing, so put it in a
`<Note>` next to the table — `open-sheet build` prints a reminder for every
appendable table that has one.

All of this is measured in Excel, not inferred from the file format.

Reach for it on anything the reader is meant to extend: a register, a cost
table, a request form. Leave it off for a fixed statement — a P&L with four
quarters has no rows to append, and an Excel Table brings banded styling and
filter arrows that a printed document does not want.

Four things it requires, each refused at compile time rather than at open time:

- **A header row.** A structured reference names the column by its header text.
- **Distinct headers.** Two columns headed "Amount" make `costs[Amount]`
  ambiguous, and Excel repairs the file by renaming one — silently changing what
  the formula means.
- **A name Excel accepts** — a letter or underscore first, no spaces, and not
  something that reads as a cell address like `AB12`.
- **Not `filter` as well.** An Excel Table brings its own filter arrows.

The total row becomes the table's own totals row, written as `SUBTOTAL`. Left
outside it, Excel would grow the table into the total the first time someone
appended a row, and the total would become a data row of itself.

A derived column carries its formula into an appended row, and its per-row
formulas are written `costs[[#This Row],[Q1]]` rather than `B4` — the same-row
form, which is what lets one stored formula serve every row. Excel displays that
in the formula bar as the shorthand `[@Q1]`; the long form is what the file has
to contain, and both Excel and LibreOffice compute it.

**Except a column that reads another row.** `r.prev()` cannot be expressed as a
single row-independent formula, so that column comes out blank on an appended
row and the reader has to fill it down. Export warns, naming the column. If it
matters, say so in a `<Note>` beside the table.

Adding rows in the *source* is always safe — that is what the framework is for,
and every reference re-resolves on the next build.

## Cycles

A circular reference is reported with every participating cell. Break it in the
source; there is no iterative-calculation mode.

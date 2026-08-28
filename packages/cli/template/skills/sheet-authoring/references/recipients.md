# What the recipient can do with the file

Everything here is about the person who opens the `.xlsx` — not the author, and
not the viewer. These are the affordances that make a workbook a tool rather
than a printout.

## Validation: a dropdown, not a free-text field

A form exported without validation collects typos, and a typo in a status column
breaks every formula that reads it.

```tsx
col('status', {
  header: '狀態',
  validate: {
    list: ['待審', '核准', '退回'],
    prompt: '從清單挑一個',
    error: '請從下拉選單選擇',
  },
})
```

**Write both messages.** `prompt` shows when the cell is selected; `error` shows
when the entry is refused. A cell that rejects input without saying why is worse
than one with no validation at all — the recipient learns only that the file will
not take their answer.

Attach it to a column, and every data cell in it gets the rule. Attach it to one
cell with `<Cell validate={…} />`.

### The rules

| | |
| --- | --- |
| `{ list: [...] }` | a fixed set of answers |
| `{ list: ref('statuses').column('name') }` | the answers live in a range |
| `{ whole: { min, max } }` · `{ decimal: { … } }` | numeric bounds; either end is optional |
| `{ date: { from: '2026-01-01', to: '2026-12-31' } }` | ISO strings or Excel serials |
| `{ textLength: { max: 40 } }` | length bounds |
| `{ custom: gt(r.cell('end'), r.cell('start')) }` | anything else — must evaluate true to pass |

Point a list at a `ref()` when the options belong to the workbook: the lookup
sheet stays the single source of truth for every dropdown that reads it.

**Make the lookup table `appendable` and the list keeps up on its own.**

```tsx
<Table name="statuses" appendable data={statuses} columns={[col('name', { header: '狀態' })]} />
```

The validation then reads `INDIRECT("statuses[狀態]")`, which resolves to
whatever the table has grown to — so appending an option at the bottom of the
lookup sheet reaches every dropdown that reads it.

Without `appendable` the validation stores a fixed range (`'Lists'!$A$2:$A$4`),
and then:

| Editing the lookup sheet in Excel | Dropdown follows? |
| --- | --- |
| **Insert** a row inside the list | yes |
| Change an option | yes |
| **Append** below the last option | **no** — and nothing says so |

That last one fails silently: no error, the new option simply never appears, and
whoever maintains the list has no way to notice. Either make the table
`appendable`, or tell them to insert rather than append.

Two things about the INDIRECT form. It must be `INDIRECT("statuses[狀態]")` and
not the bare `statuses[狀態]` — a structured reference written straight into a
validation makes Excel refuse to open the workbook at all, not ignore the rule.
And INDIRECT is volatile, so it re-evaluates on every recalculation; lookup
lists are small, but if yours is not, leave the table plain and take the fixed
range.

An inline list is stored as one comma-separated string, so an option containing
a comma would silently become two — that is refused at export with an error
telling you to use a range.

`allowBlank` defaults to true. A blank is not a typo, and refusing one makes a
half-filled form impossible to save.

`style: 'warning'` or `'info'` lets a non-matching entry through after a prompt;
the default `'error'` refuses it.

## Filters, and the total that has to agree with them

```tsx
<Table name="costs" filter data={costs} total={{ amount: 'sum' }} />
```

Puts the sort/filter arrows on the header row. **Off by default** — a register
wants them, a printed invoice does not, and they show up in print.

The part that matters is the total. On a filtered table the total is written as
`SUBTOTAL(109, …)` instead of `SUM(…)`, so hiding rows changes it. A plain `SUM`
would keep totalling rows the reader can no longer see, and the number at the
bottom would silently disagree with the rows above it with nothing on screen
saying why.

Each aggregate maps to the same function it always was — `sum` → 109, `avg` →
101, `count` → 102, `min` → 105, `max` → 104 — so turning the filter on changes
*which rows count*, never *what is counted*.

The filter covers the header and the data. It never covers the total row: inside
it, Excel would treat the total as data and hide it the first time anyone
filtered. Excel allows one filter per sheet, so a second filtered table on the
same sheet is refused at export rather than silently losing its arrows.

## Locking the formulas, leaving the inputs open

```tsx
<Sheet name="P&L" protect={{ allow: ['assumptions'] }}>
```

Every cell in Excel is *already* locked; locking only takes effect once the sheet
is protected. So `protect` names what to **un**lock and turns protection on.

Inside an allowed block, only the inputs open up — a cell holding a literal. A
cell holding a formula stays locked even there, because a derived cell is not an
input and typing over it is exactly what the protection exists to prevent.

Sorting and filtering stay available on a protected sheet: they read the model,
they do not change it. **Hiding rows by hand does not** — that is a formatting
change, and it stays blocked. A reader who expects "if I can filter, I can hide"
will find otherwise, so say which one you meant them to use.

**There is no password, deliberately.** Sheet protection is an accident-prevention
affordance, not access control — anyone with a zip tool removes it in a minute.
A password field would imply a guarantee the file format cannot keep.

Protection is off unless you ask for it. Protection that gets in the way is worse
than none: people pass the password around and it becomes noise. Pair it with a
visual convention — give the input cells a distinct style — so an input *looks*
like an input before anyone discovers what is locked.

## Notes: where a number came from

```tsx
col('revenue', {
  header: 'Revenue',
  note: (row) =>
    row.month === thisMonth ? 'Only 19 days of data — the export ran mid-month' : undefined,
})
```

Shown on hover in Excel, and as a tooltip with a corner marker in the HTML
export. Return `undefined` for rows that need no note.

This matters more here than in a hand-made spreadsheet. These workbooks are
generated: the reader cannot ask the author, and the source file may not be
something they can open. A caveat like "this month has 19 days of data" is a
property of the cell, not a footnote — put it on the cell.

`<Cell note="From the billing export" />` for a single cell.

**Not the same thing as `<Note>`.** `<Note>` is a block that occupies cells and
is part of the layout; a note is metadata attached to a cell that already exists.
Use `<Note>` for a caption the reader should see without hunting, and `note` for
provenance they only need when they wonder.

We write the **legacy** note form rather than a threaded comment. Excel,
LibreOffice, Google Sheets and Numbers all read it; a threaded comment shows as
nothing at all in the ones that do not.

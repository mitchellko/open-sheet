# Printing

A workbook that gets handed round a meeting is a printed document, and the
defaults for one are not the defaults for the other.

```tsx
<Sheet
  name="Report"
  print={{
    orientation: 'portrait',
    fitToWidth: true,
    repeatHeader: true,
    header: { center: { bold: 'FY26 Budget' } },
    footer: { right: ['Page ', pageNumber, ' of ', pageCount] },
    breakBefore: ['detail'],
  }}
>
```

## Headers and footers

You name the field; the framework produces Excel's `&`-code. `pageNumber`,
`pageCount`, `printDate`, `printTime`, `sheetName`, `fileName`, plus plain
strings, `{ bold: … }` and `{ italic: … }`. Mix them in an array.

Nobody remembers that `&P` is the page and `&N` is the count, and a literal `&`
in your title has to be doubled or it starts a code — that is the sort of
encoding this framework exists to own.

**What survives where.** The xlsx keeps everything. The PDF export goes through
Chromium, which has no CSS margin boxes but does have its own header template,
so page numbers, the date and the title survive and `printTime` does not — you
get a warning naming what was dropped rather than silence. The standalone HTML
export emits CSS margin boxes, which Firefox and print engines honour and Chrome
ignores entirely; only the page counters have a CSS equivalent at all.

If the header matters, ship the PDF or the xlsx, not the HTML.

## Where the pages break

```tsx
print={{ breakBefore: ['detail'], printArea: ['summary'] }}
```

Both name **blocks**, never rows. A break at "row 47" is exactly the coordinate
this framework exists to stop people writing, and it goes stale the moment a row
is inserted above it. Naming the block means the break moves when the content
does.

`printArea` limits printing to the named blocks, for a workbook whose working
sheets are not meant to come out of a printer at all. `breakBefore` starts a new
page above each named block — a break before the first row is dropped, since
there is no page to end.

Also available: `blackAndWhite` for something that will be photocopied or faxed,
and `center: { horizontal: true }` for a short table that would otherwise sit in
the top-left corner of a sheet of A4.

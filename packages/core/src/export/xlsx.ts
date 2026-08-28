import ExcelJS from 'exceljs'
import type { CompiledSheet, CompiledWorkbook, ConditionalFormat } from '../compile/emit.js'
import type { TableAnchor } from '../compile/registry.js'
import { serialize } from '../formula/serialize.js'
import { type Computed, isExcelError, isNotEvaluated } from '../formula/value.js'
import { columnName, rangeToA1, toA1 } from '../model/a1.js'
import { type Cell, cellKey, parseCellKey } from '../model/cell.js'
import { EXCEL_ICON_SETS, type HighlightStyle, testOf } from '../model/highlight.js'
import { isListRule, ruleOf, type Validation } from '../model/validation.js'
import type { Ref } from '../refs/ref.js'
import { type ResolveContext, resolveRef } from '../refs/resolve.js'
import { themeFor } from '../style/design.js'
import { toArgb, toExcelStyle } from '../style/excel.js'
import { DEFAULT_THEME, resolveStyle } from '../style/theme.js'
import type { Theme } from '../style/types.js'
import { numberFormat } from './formats.js'
import { encodeMargin } from './margin.js'
import { injectCharts } from './ooxml-chart.js'
import type { WorkbookWriter, WriteOptions } from './writer.js'

export class XlsxWriter implements WorkbookWriter {
  readonly extension = 'xlsx'

  async write(book: CompiledWorkbook, options: WriteOptions = {}): Promise<Buffer> {
    const theme = themeFor(book.design, options.theme ?? DEFAULT_THEME)
    const workbook = new ExcelJS.Workbook()
    workbook.creator = options.creator ?? 'open-sheet'
    // Excel honours this and recalculates on open. LibreOffice does not — its
    // default for xlsx is "never recalculate on load", so it shows our cached
    // results until the user edits a cell. That is why the CI check exports
    // without cached results rather than relying on this flag.
    workbook.calcProperties.fullCalcOnLoad = true

    for (const sheet of book.sheets) {
      const worksheet = workbook.addWorksheet(sheet.name)
      const context: ResolveContext = {
        registry: book.registry,
        definedNames: book.definedNames,
        sheet: sheet.name,
      }

      // Before the cells, not after: ExcelJS's addTable writes its own row
      // values, so a table declared afterwards drops the styles, formats and
      // notes our writer had already put on those cells.
      writeTables(worksheet, book, sheet)

      const cached = options.cacheValues === false ? undefined : options.values
      for (const [key, cell] of sheet.cells) {
        const { r, c } = parseCellKey(key)
        writeCell(worksheet, r, c, cell, context, cached, theme)
      }

      // Excel takes one per sheet. A second table's arrows would replace the
      // first's silently, so the extra ones are refused rather than dropped.
      if (sheet.autoFilters.length > 1) {
        throw new Error(
          `sheet "${sheet.name}" has ${sheet.autoFilters.length} tables with filter — Excel allows one filter per sheet. ` +
            'Put the others on their own sheets, or drop filter from all but one.',
        )
      }
      const filter = sheet.autoFilters[0]
      if (filter) worksheet.autoFilter = rangeToA1(filter)

      writeValidations(worksheet, sheet, context)

      if (sheet.protect) {
        for (const [key, cell] of sheet.cells) {
          if (!cell.unlocked) continue
          const { r, c } = parseCellKey(key)
          worksheet.getCell(r + 1, c + 1).protection = { locked: false }
        }
        // Sorting and filtering read the sheet, they do not change the model —
        // protecting the formulas should not take away the affordance we just
        // put on the header row.
        await worksheet.protect(undefined as never, {
          selectLockedCells: true,
          selectUnlockedCells: true,
          autoFilter: sheet.autoFilters.length > 0,
          sort: sheet.autoFilters.length > 0,
        })
      }

      for (let c = 0; c < sheet.bounds.cols; c += 1) {
        worksheet.getColumn(c + 1).width = sheet.columnWidths.get(c) ?? theme.defaultColumnWidth
      }

      sheet.conditionalFormats.forEach((format, index) => {
        worksheet.addConditionalFormatting({
          ref: rangeToA1(format.rect),
          rules: [conditionalRule(format, index + 1, context)],
        })
      })

      // A form printed landscape, or a table whose header does not repeat on
      // page two, is a form nobody can use. ExcelJS defaults neither.
      if (sheet.print) {
        const { orientation, size, fitToWidth, margin, blackAndWhite, center } = sheet.print
        worksheet.pageSetup = {
          ...worksheet.pageSetup,
          orientation: orientation ?? 'portrait',
          paperSize: PAPER[size ?? 'A4'],
          ...(blackAndWhite ? { blackAndWhite: true } : {}),
          ...(center?.horizontal ? { horizontalCentered: true } : {}),
          ...(center?.vertical ? { verticalCentered: true } : {}),
          ...(sheet.printArea.length > 0
            ? {
                // Row absolute, column not: ExcelJS prefixes its own `$` to
                // each side of whatever it is given, so `$A$1` becomes `$$A$1`
                // and Excel reads the defined name as corrupt. `A$1` is what
                // comes out the far end as `$A$1`.
                printArea: sheet.printArea
                  .map((rect) =>
                    [
                      toA1({ r: rect.r, c: rect.c }, { absoluteRow: true }),
                      toA1(
                        { r: rect.r + rect.rows - 1, c: rect.c + rect.cols - 1 },
                        {
                          absoluteRow: true,
                        },
                      ),
                    ].join(':'),
                  )
                  .join(','),
              }
            : {}),
          ...(fitToWidth ? { fitToPage: true, fitToWidth: 1, fitToHeight: 0 } : {}),
          ...(margin === undefined
            ? {}
            : {
                margins: {
                  left: margin,
                  right: margin,
                  top: margin,
                  bottom: margin,
                  header: margin / 2,
                  footer: margin / 2,
                },
              }),
        }
      }

      if (sheet.print?.header) worksheet.headerFooter.oddHeader = encodeMargin(sheet.print.header)
      if (sheet.print?.footer) worksheet.headerFooter.oddFooter = encodeMargin(sheet.print.footer)

      // ExcelJS takes 1-based rows; ours are 0-based, and the break belongs
      // above the named block's first row.
      if (sheet.pageBreaks.length > 0) {
        for (const row of sheet.pageBreaks) worksheet.getRow(row).addPageBreak()
      }

      if (sheet.repeatRows) {
        worksheet.pageSetup = {
          ...worksheet.pageSetup,
          printTitlesRow: `${sheet.repeatRows.from + 1}:${sheet.repeatRows.to + 1}`,
        }
      }

      if (sheet.freeze && (sheet.freeze.r > 0 || sheet.freeze.c > 0)) {
        worksheet.views = [
          {
            state: 'frozen',
            xSplit: sheet.freeze.c,
            ySplit: sheet.freeze.r,
            topLeftCell: undefined,
          },
        ]
      }
    }

    for (const [name, target] of book.definedNames) {
      const address = `${quote(target.sheet)}!${toA1(target.addr, { absoluteRow: true, absoluteCol: true })}`
      workbook.definedNames.add(address, name)
    }

    const buffer = await workbook.xlsx.writeBuffer()
    return injectCharts(Buffer.from(buffer as ArrayBuffer), book)
  }
}

/**
 * Cells sharing a validation are emitted as one element over a range, not one
 * per cell: a validated column of 500 rows would otherwise be 500 elements in
 * the file, and Excel's own UI shows it as one rule.
 */
function writeValidations(
  worksheet: ExcelJS.Worksheet,
  sheet: CompiledSheet,
  context: ResolveContext,
): void {
  const groups = new Map<string, { validation: Validation; cells: { r: number; c: number }[] }>()
  for (const [key, cell] of sheet.cells) {
    if (!cell.validate) continue
    const id = JSON.stringify(cell.validate)
    const group = groups.get(id) ?? { validation: cell.validate, cells: [] }
    group.cells.push(parseCellKey(key))
    groups.set(id, group)
  }

  for (const { validation, cells } of groups.values()) {
    const spec = toExcelValidation(validation, context)
    for (const range of contiguousRanges(cells)) {
      ;(
        worksheet as unknown as { dataValidations: { add(r: string, v: unknown): void } }
      ).dataValidations.add(range, spec)
    }
  }
}

/** Runs of consecutive rows in one column collapse to `C2:C40`; anything else stays a cell. */
function contiguousRanges(cells: readonly { r: number; c: number }[]): string[] {
  const byColumn = new Map<number, number[]>()
  for (const { r, c } of cells) byColumn.set(c, [...(byColumn.get(c) ?? []), r])

  const out: string[] = []
  for (const [c, rows] of byColumn) {
    rows.sort((a, b) => a - b)
    let start = rows[0] as number
    let previous = start
    for (const r of rows.slice(1)) {
      if (r === previous + 1) {
        previous = r
        continue
      }
      out.push(rangeOf(start, previous, c))
      start = r
      previous = r
    }
    out.push(rangeOf(start, previous, c))
  }
  return out
}

function rangeOf(first: number, last: number, c: number): string {
  const from = toA1({ r: first, c })
  return first === last ? from : `${from}:${toA1({ r: last, c })}`
}

function boundsFormulae(bounds: { min?: number; max?: number }): {
  operator: string
  formulae: number[]
} {
  if (bounds.min !== undefined && bounds.max !== undefined) {
    return { operator: 'between', formulae: [bounds.min, bounds.max] }
  }
  if (bounds.min !== undefined) return { operator: 'greaterThanOrEqual', formulae: [bounds.min] }
  if (bounds.max !== undefined) return { operator: 'lessThanOrEqual', formulae: [bounds.max] }
  throw new Error('a numeric validation needs a min, a max, or both')
}

function toExcelValidation(validation: Validation, context: ResolveContext): unknown {
  const messages = {
    // Excel's own default. A blank cell is not a typo, and refusing one makes a
    // half-filled form impossible to save.
    allowBlank: validation.allowBlank ?? true,
    showInputMessage: validation.prompt !== undefined || validation.promptTitle !== undefined,
    ...(validation.promptTitle === undefined ? {} : { promptTitle: validation.promptTitle }),
    ...(validation.prompt === undefined ? {} : { prompt: validation.prompt }),
    showErrorMessage: validation.error !== undefined || validation.errorTitle !== undefined,
    // The file format's enum is stop | warning | information, not the words the
    // authoring surface uses. Writing `error` or `info` produces a document that
    // Excel tolerates and that openpyxl refuses to open at all — so every
    // workbook using `validate` was unreadable by the Python side of a pipeline,
    // which is the case this project is most for.
    errorStyle: ERROR_STYLES[validation.style ?? 'error'],
    ...(validation.errorTitle === undefined ? {} : { errorTitle: validation.errorTitle }),
    ...(validation.error === undefined ? {} : { error: validation.error }),
  }
  const rule = ruleOf(validation)

  if (isListRule(rule)) {
    const formula = Array.isArray(rule.list)
      ? // Excel reads an inline list as one quoted, comma-separated string, so a
        // value containing a comma would silently become two options.
        `"${rule.list
          .map((item: string) => {
            if (String(item).includes(',')) {
              throw new Error(
                `validation list item "${item}" contains a comma, which Excel reads as a separator. ` +
                  'Put the options in a range and pass a ref() instead.',
              )
            }
            return String(item)
          })
          .join(',')}"`
      : listSource(rule.list as Ref, context)
    return { type: 'list', ...messages, formulae: [formula] }
  }
  if ('whole' in rule) return { type: 'whole', ...messages, ...boundsFormulae(rule.whole) }
  if ('decimal' in rule) return { type: 'decimal', ...messages, ...boundsFormulae(rule.decimal) }
  if ('textLength' in rule) {
    return { type: 'textLength', ...messages, ...boundsFormulae(rule.textLength) }
  }
  if ('date' in rule) {
    // ExcelJS converts Date objects to serials itself; handing it a serial makes
    // it convert *that* as if it were a date, and 2026-01-01 became 1970-01-01.
    const bounds = boundsFormulae({
      min: rule.date.from === undefined ? undefined : toDateSerial(rule.date.from),
      max: rule.date.to === undefined ? undefined : toDateSerial(rule.date.to),
    })
    return {
      type: 'date',
      ...messages,
      operator: bounds.operator,
      formulae: bounds.formulae.map(fromSerial),
    }
  }
  return {
    type: 'custom',
    ...messages,
    formulae: [`=${serialize(rule.custom, context)}`.replace(/^==/, '=')],
  }
}

function fromSerial(serial: number): Date {
  return new Date(Date.UTC(1899, 11, 30) + serial * 86_400_000)
}

function toDateSerial(value: string | number): number {
  if (typeof value === 'number') return value
  const parsed = Date.parse(`${value}T00:00:00Z`)
  if (Number.isNaN(parsed)) throw new Error(`validation date "${value}" is not an ISO date`)
  return Math.round((parsed - Date.UTC(1899, 11, 30)) / 86_400_000)
}

/**
 * Where the dropdown reads its options from.
 *
 * An absolute range is fixed: append an option to the bottom of the lookup
 * sheet and the dropdown never sees it — silently, with no error and nothing to
 * notice. When the options live in an `appendable` table, `INDIRECT` over the
 * table column resolves to whatever the table has grown to, so appending works.
 *
 * It has to be `INDIRECT("statuses[Status]")` and not the bare
 * `statuses[Status]`: a structured reference written straight into a
 * validation's formula makes Excel refuse to open the workbook at all — not
 * ignore the rule, refuse the file. Measured in Excel, and LibreOffice resolves
 * the INDIRECT form too.
 *
 * The cost is that INDIRECT is volatile, so it re-evaluates on every
 * recalculation. Lookup lists are small; if yours is not, leave the table plain
 * and take the fixed range.
 */
function listSource(reference: Ref, context: ResolveContext): string {
  if (reference.kind === 'range' && reference.part === 'column' && reference.column !== undefined) {
    const anchor = context.registry.get(reference.block)
    const header =
      anchor?.kind === 'table' ? anchor.table?.headers.get(reference.column) : undefined
    if (anchor?.kind === 'table' && header !== undefined) {
      const escaped = header.replace(/([[\]#'@])/g, "'$1")
      // The name is inside a formula *string*, so a double quote in a header
      // would close it early.
      return `INDIRECT("${anchor.name}[${escaped.replace(/"/g, '""')}]")`
    }
  }
  return absoluteRange(reference, context)
}

function absoluteRange(reference: Ref, context: ResolveContext): string {
  const resolved = resolveRef(reference, context)
  const { r, c, rows, cols } = resolved.rect
  const first = toA1({ r, c }, { absoluteRow: true, absoluteCol: true })
  const last = toA1({ r: r + rows - 1, c: c + cols - 1 }, { absoluteRow: true, absoluteCol: true })
  return `${quote(resolved.sheet)}!${first}:${last}`
}

const TOTALS_FUNCTIONS: Record<string, string> = {
  sum: 'sum',
  avg: 'average',
  count: 'count',
  min: 'min',
  max: 'max',
}

/**
 * An Excel Table over the header and data rows, so a row typed below the last
 * one is taken in and every structured reference follows it.
 */
function writeTables(
  worksheet: ExcelJS.Worksheet,
  book: CompiledWorkbook,
  sheet: CompiledSheet,
): void {
  for (const anchor of book.registry.values()) {
    if (anchor.kind !== 'table' || !anchor.table || anchor.sheet !== sheet.name) continue

    const headers = [...anchor.table.headers]
    const first = anchor.headerRow ?? anchor.firstDataRow
    // Our own total row becomes the table's totals row. Left outside, Excel
    // would grow the table into it the first time someone appended a row.
    const hasTotals = anchor.totalRow !== undefined

    worksheet.addTable({
      name: anchor.name,
      ref: toA1({ r: first, c: anchor.rect.c }),
      headerRow: anchor.headerRow !== undefined,
      totalsRow: hasTotals,
      columns: headers.map(([key, label], index) => {
        const totalsFunction = hasTotals ? anchor.table?.totals.get(key) : undefined
        return {
          name: label,
          ...(hasTotals && totalsFunction
            ? { totalsRowFunction: TOTALS_FUNCTIONS[totalsFunction] ?? 'none' }
            : {}),
          // Read from the cell we already wrote rather than restated, so the
          // table part and the visible total row cannot disagree.
          ...(hasTotals && !totalsFunction && index === 0
            ? { totalsRowLabel: totalLabel(sheet, anchor) }
            : {}),
        }
      }) as never,
      // Placeholders: every one is overwritten by our own cell writing, which
      // runs after this. ExcelJS requires the shape, not the content.
      rows: Array.from({ length: anchor.rowCount }, () => headers.map(() => null)),
    })
  }
}

function totalLabel(sheet: CompiledSheet, anchor: TableAnchor): string {
  if (anchor.totalRow === undefined) return 'Total'
  const cell = sheet.cells.get(cellKey(anchor.totalRow, anchor.rect.c))
  return typeof cell?.value === 'string' ? cell.value : 'Total'
}

const ERROR_STYLES: Record<string, string> = {
  error: 'stop',
  warning: 'warning',
  info: 'information',
}

function dxf(style: HighlightStyle): Record<string, unknown> {
  return {
    ...(style.fill
      ? { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: toArgb(style.fill) } } }
      : {}),
    ...(style.color || style.bold
      ? {
          font: {
            ...(style.bold ? { bold: true } : {}),
            ...(style.color ? { color: { argb: toArgb(style.color) } } : {}),
          },
        }
      : {}),
  }
}

/** Excel quotes a text comparand; a number must not be quoted or it compares as text. */
function comparand(value: number | string | boolean): string {
  return typeof value === 'string' ? `"${value.replace(/"/g, '""')}"` : String(value)
}

function highlightRule(
  format: Extract<ConditionalFormat, { kind: 'highlight' }>,
  priority: number,
): Record<string, unknown> {
  const test = testOf(format.rule)
  const base = { priority, style: dxf(format.rule) }
  const cellIs = (operator: string, formulae: string[]) => ({
    ...base,
    type: 'cellIs',
    operator,
    formulae,
  })

  if ('above' in test) return cellIs('greaterThan', [String(test.above)])
  if ('below' in test) return cellIs('lessThan', [String(test.below)])
  if ('atLeast' in test) return cellIs('greaterThanOrEqual', [String(test.atLeast)])
  if ('atMost' in test) return cellIs('lessThanOrEqual', [String(test.atMost)])
  if ('equals' in test) return cellIs('equal', [comparand(test.equals)])
  if ('between' in test) {
    return cellIs('between', [String(test.between[0]), String(test.between[1])])
  }
  if ('contains' in test) {
    return { ...base, type: 'containsText', operator: 'containsText', text: test.contains }
  }
  if ('top' in test) return { ...base, type: 'top10', rank: test.top }
  if ('bottom' in test) return { ...base, type: 'top10', rank: test.bottom, bottom: true }

  // ExcelJS drops a `duplicateValues` rule silently — it writes no element at
  // all. COUNTIF over the range says the same thing, goes through our own
  // serializer, and is what Excel's own UI produces for this rule anyway.
  const { r, c, rows } = format.rect
  const first = toA1({ r, c }, { absoluteRow: true, absoluteCol: true })
  const last = toA1({ r: r + rows - 1, c }, { absoluteRow: true, absoluteCol: true })
  return {
    ...base,
    type: 'expression',
    formulae: [`COUNTIF(${first}:${last},${toA1({ r, c })})>1`],
  }
}

function conditionalRule(
  format: ConditionalFormat,
  priority: number,
  context: ResolveContext,
): never {
  void context
  if (format.kind === 'dataBar') {
    return {
      type: 'dataBar',
      priority,
      minLength: 0,
      maxLength: 100,
      gradient: false,
      showValue: true,
      cfvo: [{ type: 'min' }, { type: 'max' }],
      color: { argb: toArgb(format.color) },
    } as never
  }
  if (format.kind === 'colorScale') {
    // `percent`, not Excel's default `percentile`: percent is the linear
    // midpoint of min..max, which the HTML export can reproduce exactly.
    // Percentile is the median, and the two renderers would drift apart on any
    // column that is not evenly distributed.
    const stops =
      format.scale.length === 2
        ? [{ type: 'min' }, { type: 'max' }]
        : [{ type: 'min' }, { type: 'percent', value: 50 }, { type: 'max' }]
    return {
      type: 'colorScale',
      priority,
      cfvo: stops,
      color: format.scale.map((color) => ({ argb: toArgb(color) })),
    } as never
  }
  if (format.kind === 'iconSet') {
    return {
      type: 'iconSet',
      priority,
      iconSet: EXCEL_ICON_SETS[format.icons],
      cfvo: [
        { type: 'percent', value: 0 },
        { type: 'percent', value: 33 },
        { type: 'percent', value: 67 },
      ],
    } as never
  }
  return highlightRule(format, priority) as never
}

/** Excel's numeric paper sizes; the names mean nothing to it. */
const PAPER: Record<string, number> = { Letter: 1, Legal: 5, A3: 8, A4: 9 }

function quote(sheet: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet) ? sheet : `'${sheet.replace(/'/g, "''")}'`
}

function writeCell(
  worksheet: ExcelJS.Worksheet,
  r: number,
  c: number,
  cell: Cell,
  context: ResolveContext,
  values: Map<string, Computed> | undefined,
  theme: Theme,
): void {
  const target = worksheet.getCell(r + 1, c + 1)

  if (cell.expr) {
    // The row is part of the context: inside an Excel Table a same-row
    // reference is written `[@Amount]`, which is what makes one stored formula
    // able to fill every appended row.
    const formula = serialize(cell.expr, { ...context, row: r })
    const cached = values?.get(`${context.sheet}!${r},${c}`)
    const result =
      cached === undefined || isNotEvaluated(cached) ? undefined : toExcelResult(cached)
    const base =
      result === undefined ? { formula, date1904: false } : { formula, result, date1904: false }
    // A legacy array formula, not a dynamic-array spill: it fills exactly the
    // rectangle named in `ref` and no more, which is what the placement engine
    // reserved. Excel and LibreOffice have both honoured this since 2007.
    if (cell.spill) {
      const last = { r: r + cell.spill.rows - 1, c: c + cell.spill.cols - 1 }
      // ExcelJS writes array formulas correctly but does not describe them in
      // its types, which have not moved since 2023. Verified by reading the
      // file back in LibreOffice, not by trusting the cast.
      target.value = {
        ...base,
        shareType: 'array',
        ref: `${toA1({ r, c })}:${toA1(last)}`,
      } as unknown as typeof target.value
    } else {
      target.value = base
    }
  } else if (cell.spillFrom) {
    // Written by the origin's array formula; a value here would fight it.
  } else if (cell.value !== null && cell.value !== undefined) {
    target.value = cell.value
  }

  // The legacy note form, not a threaded comment: Excel, LibreOffice, Google
  // Sheets and Numbers all read it, and a threaded comment shows as nothing at
  // all in the ones that do not.
  if (cell.note) target.note = cell.note

  const style = resolveStyle(theme, cell.style)
  if (style) Object.assign(target, toExcelStyle(style))
  if (cell.wrap) {
    target.alignment = { ...target.alignment, wrapText: true, vertical: 'top' }
  }

  const format = numberFormat(cell.format ?? style?.format)
  if (format) target.numFmt = format

  if (cell.span && (cell.span.rows > 1 || cell.span.cols > 1)) {
    worksheet.mergeCells(r + 1, c + 1, r + cell.span.rows, c + cell.span.cols)
  }
}

type FormulaResult = string | number | boolean | Date | ExcelJS.CellErrorValue | undefined

function toExcelResult(value: Computed): FormulaResult {
  if (isExcelError(value)) return { error: value.code } as ExcelJS.CellErrorValue
  if (value === null) return undefined
  return value as FormulaResult
}

export function a1(r: number, c: number): string {
  return `${columnName(c)}${r + 1}`
}

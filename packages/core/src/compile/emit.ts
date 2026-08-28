import {
  avg,
  count,
  type Expr,
  type ExprInput,
  isExpr,
  lift,
  max,
  min,
  subtotal,
  sum,
} from '../formula/expr.js'
import { parseFormula } from '../formula/parse.js'
import { type Placement, placeSheet } from '../layout/place.js'
import { type Cell, type CellKey, type CellValue, cellKey } from '../model/cell.js'
import type { Addr, Rect, Size } from '../model/geometry.js'
import type { ColorScale, Highlight, IconSet } from '../model/highlight.js'
import { isRef, type Ref } from '../refs/ref.js'
import { type ResolveContext, resolveRef } from '../refs/resolve.js'
import type { DesignSystem } from '../style/design.js'
import type { KeyValueEntry } from './components.js'
import type {
  Aggregate,
  Block,
  ChartAxes,
  ChartKind,
  ChartSeries,
  PrintSetup,
  SheetNode,
  SheetProtection,
  TableNode,
  WorkbookNode,
} from './nodes.js'
import { type Registry, requireAnchor, type TableAnchor } from './registry.js'
import { makeRowContext } from './row-context.js'

export type ConditionalFormat =
  | { kind: 'dataBar'; rect: Rect; color: string; negativeColor?: string }
  | { kind: 'colorScale'; rect: Rect; scale: ColorScale }
  | { kind: 'iconSet'; rect: Rect; icons: IconSet }
  | { kind: 'highlight'; rect: Rect; rule: Highlight }

export interface PlacedSparkline {
  kind: 'line' | 'column'
  color: string
  /** Where the sparkline is drawn. */
  cell: Addr
  /** The cells it reads, on the same row. */
  source: Rect
}

export interface PlacedChart {
  chart: ChartKind
  title?: string
  axes?: ChartAxes
  dataLabels?: boolean
  rect: Rect
  categories: Ref
  // The node's own type, not a restatement of it: as an inline duplicate it
  // fell behind the moment a series gained `as` and `axis`.
  series: ChartSeries[]
}

export interface CompiledSheet {
  name: string
  print?: PrintSetup
  cells: Map<CellKey, Cell>
  columnWidths: Map<number, number>
  conditionalFormats: ConditionalFormat[]
  charts: PlacedChart[]
  /** Header-plus-data rectangles that get filter arrows. Never includes a total row. */
  autoFilters: Rect[]
  /** When set, every cell is locked except those marked `unlocked`. */
  protect?: boolean
  /** One entry per sparkline cell, already resolved to the row it reads. */
  sparklines: PlacedSparkline[]
  /** Rectangles to print, resolved from the named blocks. Empty means the whole sheet. */
  printArea: Rect[]
  /** Zero-based rows to start a new page on, resolved from the named blocks. */
  pageBreaks: number[]
  /** Rows to repeat at the top of every printed page, zero-based inclusive. */
  repeatRows?: { from: number; to: number }
  freeze?: Addr
  bounds: Size
}

export interface DefinedName {
  sheet: string
  addr: Addr
  /** The block that claimed it, so a collision can name both sides. */
  owner: string
}

export interface CompiledWorkbook {
  sheets: CompiledSheet[]
  registry: Registry
  definedNames: Map<string, DefinedName>
  /** From the module's `design` const; drives the theme for every renderer. */
  design?: DesignSystem
}

const AGGREGATES: Record<Aggregate, (expr: Expr) => Expr> = {
  sum: (expr) => sum(expr),
  avg: (expr) => avg(expr),
  count: (expr) => count(expr),
  min: (expr) => min(expr),
  max: (expr) => max(expr),
}

/**
 * The 10x codes ignore rows a filter has hidden. On a filtered table a plain
 * SUM keeps totalling rows the reader can no longer see, so the total silently
 * disagrees with the rows above it — the exact failure this project cares most
 * about. Each code is the same function as its AGGREGATES entry, so turning the
 * filter on changes which rows count and never what is counted.
 */
const FILTERED_AGGREGATES: Record<Aggregate, number> = {
  sum: 109,
  avg: 101,
  count: 102,
  min: 105,
  max: 104,
}

/**
 * Unlocks the inputs of the named blocks — the cells holding a literal. A cell
 * with a formula is derived, not an input, so it stays locked even inside a
 * block the author opened up: that is the whole point of protecting the sheet.
 */
function unlockInputs(
  protection: SheetProtection,
  sheetName: string,
  registry: Registry,
  cells: Map<CellKey, Cell>,
): void {
  for (const name of protection.allow) {
    const anchor = requireAnchor(registry, name)
    if (anchor.sheet !== sheetName) {
      throw new Error(
        `<Sheet name="${sheetName}" protect> allows "${name}", which is on sheet "${anchor.sheet}". ` +
          'Protection is per sheet, so name it on the sheet it lives on.',
      )
    }
    // Only the cells someone is meant to fill in: a table's data rows, or a
    // key-value block's value cells. A title, a header and a total are not
    // inputs, and unlocking a header invites a reader to rename the column
    // every formula in the sheet refers to.
    const inputs =
      anchor.kind === 'table'
        ? rowsOf(anchor.firstDataRow, anchor.lastDataRow, anchor.rect)
        : [...anchor.keys.values()]

    for (const addr of inputs) {
      const cell = cells.get(cellKey(addr.r, addr.c))
      if (cell && cell.expr === undefined) cell.unlocked = true
    }
  }
}

const DEFAULT_SPARKLINE_COLOR = '#1d4ed8'

/**
 * One sparkline per data row, each reading its own row. The file format stores
 * a single range per sparkline, so the columns it reads have to be adjacent —
 * a non-contiguous set is refused here rather than written as a range that
 * quietly includes whatever sits between them.
 */
function placeSparklines(table: TableNode, rect: Rect): PlacedSparkline[] {
  const out: PlacedSparkline[] = []
  const firstDataRow = rect.r + (table.title ? 1 : 0) + (table.showHeader ? 1 : 0)

  table.columns.forEach((column, index) => {
    const spec = column.sparkline
    if (!spec) return
    if (spec.of.length < 2) {
      throw new Error(
        `col('${column.key}') has a sparkline of ${spec.of.length} column(s) — a trend needs at least two`,
      )
    }

    const positions = spec.of.map((key) => {
      const at = table.columns.findIndex((candidate) => candidate.key === key)
      if (at < 0) {
        throw new Error(
          `col('${column.key}') has a sparkline over "${key}", which is not a column of table "${table.name}" ` +
            `(columns: ${table.columns.map((candidate) => candidate.key).join(', ')})`,
        )
      }
      return at
    })

    const from = Math.min(...positions)
    const to = Math.max(...positions)
    if (to - from + 1 !== positions.length) {
      throw new Error(
        `col('${column.key}') has a sparkline over ${spec.of.join(', ')}, which are not next to each other. ` +
          'A sparkline stores one range, so the columns it reads must be adjacent — reorder them, or read a contiguous run.',
      )
    }

    for (let row = firstDataRow; row < firstDataRow + table.data.length; row += 1) {
      out.push({
        kind: spec.kind ?? 'line',
        color: spec.color ?? DEFAULT_SPARKLINE_COLOR,
        cell: { r: row, c: rect.c + index },
        source: { r: row, c: rect.c + from, rows: 1, cols: to - from + 1 },
      })
    }
  })
  return out
}

function headerLabel(column: { key: string; header?: string }): string {
  return column.header ?? column.key
}

/**
 * Excel Table names share the rules defined names have — and one more: the name
 * lives in the same space as the defined names we already emit, so a collision
 * is a file Excel offers to repair.
 */
function assertTableName(name: string): void {
  if (!/^[A-Za-z_\\][A-Za-z0-9._\\]*$/.test(name)) {
    throw new Error(
      `<Table name="${name}" appendable> — Excel Table names must start with a letter or underscore and hold no spaces. ` +
        'Rename the block, or drop appendable.',
    )
  }
  if (/^[A-Za-z]{1,3}[0-9]+$/.test(name)) {
    throw new Error(
      `<Table name="${name}" appendable> — that reads as the cell address ${name.toUpperCase()}, which Excel refuses as a table name.`,
    )
  }
}

function onThisSheet(registry: Registry, name: string, sheetName: string, field: string) {
  const anchor = requireAnchor(registry, name)
  if (anchor.sheet !== sheetName) {
    throw new Error(
      `<Sheet name="${sheetName}" print={{ ${field} }}> names "${name}", which is on sheet "${anchor.sheet}". ` +
        'Printing is per sheet, so name it on the sheet it lives on.',
    )
  }
  return anchor
}

function rowsOf(first: number, last: number, rect: Rect): Addr[] {
  const out: Addr[] = []
  for (let r = first; r <= last; r += 1) {
    for (let c = rect.c; c < rect.c + rect.cols; c += 1) out.push({ r, c })
  }
  return out
}

export function emitWorkbook(workbook: WorkbookNode): CompiledWorkbook {
  assertUniqueNames(workbook)
  const registry: Registry = new Map()
  const definedNames = new Map<string, DefinedName>()
  const sheets = workbook.children.map((sheet) => emitSheet(sheet, registry, definedNames))
  return { sheets, registry, definedNames }
}

function emitSheet(
  sheet: SheetNode,
  registry: Registry,
  definedNames: Map<string, DefinedName>,
): CompiledSheet {
  const cells = new Map<CellKey, Cell>()
  const columnWidths = new Map<number, number>()
  const conditionalFormats: ConditionalFormat[] = []
  const charts: PlacedChart[] = []
  const sparklines: PlacedSparkline[] = []
  const autoFilters: Rect[] = []
  const placements = placeSheet(sheet)

  for (const placement of placements) {
    if (placement.block.kind === 'chart') {
      const node = placement.block
      const placed: PlacedChart = {
        chart: node.chart,
        rect: placement.rect,
        categories: node.categories,
        series: node.series,
      }
      if (node.title !== undefined) placed.title = node.title
      if (node.axes !== undefined) placed.axes = node.axes
      if (node.dataLabels !== undefined) placed.dataLabels = node.dataLabels
      charts.push(placed)
      continue
    }
    if (placement.block.kind === 'table' && placement.block.filter) {
      const table = placement.block
      // Header through last data row. Including the total row would make Excel
      // filter it like data and hide the total the moment anyone uses a filter.
      autoFilters.push({
        r: placement.rect.r + (table.title ? 1 : 0),
        c: placement.rect.c,
        rows: table.data.length + 1,
        cols: table.columns.length,
      })
    }
    if (placement.block.kind === 'table') {
      sparklines.push(...placeSparklines(placement.block, placement.rect))
    }
    emitPlacement(
      placement,
      sheet.name,
      cells,
      columnWidths,
      conditionalFormats,
      registry,
      definedNames,
    )
  }

  if (sheet.protect) {
    unlockInputs(sheet.protect, sheet.name, registry, cells)
  }

  const printArea = (sheet.print?.printArea ?? []).map(
    (name) => onThisSheet(registry, name, sheet.name, 'printArea').rect,
  )
  // The break goes *before* the block, so the row it starts on is the first row
  // of the new page. Sorted and deduplicated: two blocks starting on the same
  // row is one break, and Excel reads an unsorted list as corrupt.
  const pageBreaks = [
    ...new Set(
      (sheet.print?.breakBefore ?? []).map(
        (name) => onThisSheet(registry, name, sheet.name, 'breakBefore').rect.r,
      ),
    ),
  ]
    .filter((row) => row > 0)
    // A break on the first row of the print area starts a page that has already
    // started. Excel ignores it, so keeping it would only leave a break in the
    // file that does nothing and reads as though it should.
    .filter((row) => !printArea.some((rect) => rect.r === row))
    .sort((a, b) => a - b)

  const compiled: CompiledSheet = {
    name: sheet.name,
    ...(sheet.protect ? { protect: true } : {}),
    cells,
    columnWidths,
    conditionalFormats,
    charts,
    sparklines,
    autoFilters,
    printArea,
    pageBreaks,
    bounds: boundsOf(placements),
  }
  if (sheet.freeze) compiled.freeze = parseFreeze(sheet.freeze)
  if (sheet.print) compiled.print = sheet.print

  // The header row of the first table is what a reader needs on page two.
  if (sheet.print?.repeatHeader) {
    for (const anchor of registry.values()) {
      if (
        anchor.kind === 'table' &&
        anchor.sheet === sheet.name &&
        anchor.headerRow !== undefined
      ) {
        compiled.repeatRows = { from: anchor.headerRow, to: anchor.headerRow }
        break
      }
    }
  }
  return compiled
}

function emitPlacement(
  placement: Placement,
  sheetName: string,
  cells: Map<CellKey, Cell>,
  columnWidths: Map<number, number>,
  conditionalFormats: ConditionalFormat[],
  registry: Registry,
  definedNames: Map<string, DefinedName>,
): void {
  const { block, rect } = placement
  switch (block.kind) {
    case 'cell': {
      const cell: Cell = {}
      if (block.value !== undefined) cell.value = block.value
      if (block.expr !== undefined) cell.expr = block.expr
      if (block.validate !== undefined) cell.validate = block.validate
      if (block.note !== undefined) cell.note = block.note
      if (block.format !== undefined) cell.format = block.format
      if (block.style !== undefined) cell.style = block.style
      if (block.span !== undefined) cell.span = block.span
      cells.set(cellKey(rect.r, rect.c), cell)
      return
    }
    case 'note': {
      const text = block.runs.map((run) => run.text).join('')
      cells.set(cellKey(rect.r, rect.c), {
        value: text,
        style: block.style ?? 'note',
        span: { rows: 1, cols: block.cols },
      })
      return
    }
    case 'spill': {
      const origin = cellKey(rect.r, rect.c)
      const cell: Cell = { expr: block.expr, spill: { rows: block.rows, cols: block.cols } }
      if (block.format !== undefined) cell.format = block.format
      if (block.style !== undefined) cell.style = block.style
      cells.set(origin, cell)
      // The rest of the footprint is written as belonging to the origin, so
      // placement sees it as occupied and the evaluator knows where to put the
      // values the array produces.
      for (let r = 0; r < block.rows; r += 1) {
        for (let c = 0; c < block.cols; c += 1) {
          if (r === 0 && c === 0) continue
          const member: Cell = { spillFrom: origin }
          if (block.format !== undefined) member.format = block.format
          if (block.style !== undefined) member.style = block.style
          cells.set(cellKey(rect.r + r, rect.c + c), member)
        }
      }
      return
    }
    case 'spacer':
    case 'chart':
      return
    case 'kpiBand': {
      block.items.forEach((item, i) => {
        cells.set(cellKey(rect.r, rect.c + i), { value: item.label, style: 'kpiLabel' })
        const valueCell: Cell = { style: 'kpiValue' }
        if (isExpr(item.value) || isRef(item.value)) valueCell.expr = lift(item.value)
        else valueCell.value = item.value as CellValue
        if (item.format) valueCell.format = item.format
        cells.set(cellKey(rect.r + 1, rect.c + i), valueCell)
      })
      return
    }
    case 'table':
      emitTable(
        block,
        rect,
        sheetName,
        cells,
        columnWidths,
        conditionalFormats,
        registry,
        definedNames,
      )
      return
  }
}

const DEFAULT_BAR_COLOR = '#93c5fd'

function emitTable(
  table: TableNode,
  rect: Rect,
  sheetName: string,
  cells: Map<CellKey, Cell>,
  columnWidths: Map<number, number>,
  conditionalFormats: ConditionalFormat[],
  registry: Registry,
  definedNames: Map<string, DefinedName>,
): void {
  let row = rect.r

  if (table.title) {
    cells.set(cellKey(row, rect.c), {
      value: table.title,
      style: 'tableTitle',
      span: { rows: 1, cols: rect.cols },
    })
    row += 1
  }
  const titleRow = table.title ? rect.r : undefined

  if (table.variant === 'keyValue') {
    emitKeyValue(table, rect, row, sheetName, cells, registry, definedNames)
    return
  }

  let headerRow: number | undefined
  if (table.showHeader) {
    headerRow = row
    table.columns.forEach((column, i) => {
      cells.set(cellKey(row, rect.c + i), {
        value: column.header ?? column.key,
        style: column.style ? `${column.style}Header` : 'tableHeader',
      })
    })
    row += 1
  }

  const firstDataRow = row
  const columns = new Map<string, number>()
  table.columns.forEach((column, i) => {
    columns.set(column.key, rect.c + i)
    if (column.width !== undefined) columnWidths.set(rect.c + i, column.width)
  })

  table.data.forEach((dataRow, index) => {
    table.columns.forEach((column, i) => {
      const target = cellKey(firstDataRow + index, rect.c + i)
      const cell: Cell = {}
      // Set before the formula branch, which has its own early return.
      if (column.validate) cell.validate = column.validate
      if (column.formula) {
        // A formula that throws is a formula somebody wrote; without the block,
        // column and row, a 450-line workbook means bisecting by hand to find it.
        const produced =
          typeof column.formula === 'string'
            ? column.formula
            : withContext(table.name, column.key, index, () =>
                (column.formula as (r: never) => ExprInput | null | undefined)(
                  makeRowContext(table.name, table.data, index) as never,
                ),
              )
        if (produced === null || produced === undefined) {
          cells.set(target, cell)
          return
        }
        cell.expr = typeof produced === 'string' ? parseFormula(produced).expr : lift(produced)
      } else if (column.value) {
        cell.value = column.value(dataRow, index)
      } else {
        cell.value = readField(dataRow, column.key)
      }
      if (column.format) cell.format = column.format
      if (column.style) cell.style = column.style
      if (column.wrap) cell.wrap = true
      if (column.note) {
        const note = column.note(dataRow, index)
        if (note !== undefined && note !== '') cell.note = note
      }
      cells.set(target, cell)
    })
  })

  if (table.data.length > 0) {
    table.columns.forEach((column, i) => {
      const area: Rect = {
        r: firstDataRow,
        c: rect.c + i,
        rows: table.data.length,
        cols: 1,
      }
      if (column.bar) {
        const bar = column.bar === true ? {} : column.bar
        const format: ConditionalFormat = {
          kind: 'dataBar',
          rect: area,
          color: bar.color ?? DEFAULT_BAR_COLOR,
        }
        if (bar.negativeColor) format.negativeColor = bar.negativeColor
        conditionalFormats.push(format)
      }
      if (column.scale)
        conditionalFormats.push({ kind: 'colorScale', rect: area, scale: column.scale })
      if (column.icons)
        conditionalFormats.push({ kind: 'iconSet', rect: area, icons: column.icons })
      for (const rule of column.highlight === undefined
        ? []
        : Array.isArray(column.highlight)
          ? column.highlight
          : [column.highlight]) {
        conditionalFormats.push({ kind: 'highlight', rect: area, rule })
      }
    })
  }

  const lastDataRow = firstDataRow + Math.max(table.data.length - 1, 0)
  let totalRow: number | undefined

  if (table.total) {
    totalRow = firstDataRow + table.data.length
    table.columns.forEach((column, i) => {
      const aggregate = table.total?.[column.key]
      const target = cellKey(totalRow as number, rect.c + i)
      if (!aggregate) {
        cells.set(
          target,
          i === 0 ? { value: 'Total', style: 'tableTotal' } : { style: 'tableTotal' },
        )
        return
      }
      const range: Expr = {
        k: 'ref',
        target: { kind: 'range', block: table.name, part: 'column', column: column.key },
      }
      // An Excel Table's totals row must hold SUBTOTAL, not the plain
      // aggregate. Given `totalsRowFunction="sum"` and a cell holding
      // `SUM(register[Q1])`, Excel decides the row is not a totals row at all,
      // drops `totalsRowCount`, and swallows it as data — at which point
      // `register[Q1]` includes the total itself and the cell is a circular
      // reference. It still *shows* the cached number, so nothing looks wrong.
      const cell: Cell = {
        expr:
          table.filter || table.appendable
            ? subtotal(FILTERED_AGGREGATES[aggregate], range)
            : AGGREGATES[aggregate](range),
        style: 'tableTotal',
      }
      if (column.format) cell.format = column.format
      cells.set(target, cell)
    })
  }

  const anchor: TableAnchor = {
    kind: 'table',
    name: table.name,
    sheet: sheetName,
    rect,
    firstDataRow,
    lastDataRow,
    rowCount: table.data.length,
    columns,
  }
  if (titleRow !== undefined) anchor.titleRow = titleRow
  if (headerRow !== undefined) anchor.headerRow = headerRow
  if (totalRow !== undefined) anchor.totalRow = totalRow
  if (table.appendable) {
    assertTableName(table.name)
    anchor.table = {
      headers: new Map(table.columns.map((c) => [c.key, headerLabel(c)])),
      totals: new Map(
        Object.entries(table.total ?? {}).filter(([, v]) => v !== undefined) as [string, string][],
      ),
      noFillDown: [],
    }
    const seen = new Set<string>()
    for (const [key, label] of anchor.table.headers) {
      // Excel keys a structured reference by the header text, so two columns
      // sharing one header make `costs[Amount]` ambiguous and Excel repairs the
      // file by renaming one of them — silently changing what a formula means.
      if (seen.has(label)) {
        throw new Error(
          `<Table name="${table.name}" appendable> has two columns headed "${label}" (one of them is col('${key}')). ` +
            'An Excel Table names its columns by their headers, so they must be distinct.',
        )
      }
      seen.add(label)
    }
  }
  registry.set(table.name, anchor)

  const emitted = anchor.table
  if (emitted) {
    const context: ResolveContext = { registry, definedNames, sheet: sheetName }
    for (const [key, header] of emitted.headers) {
      const column = anchor.columns.get(key)
      if (column === undefined) continue
      const cell = cells.get(cellKey(anchor.lastDataRow, column))
      if (!cell?.expr) continue
      if (!rowIndependent(cell.expr, anchor, anchor.lastDataRow, context)) {
        emitted.noFillDown.push(header)
      }
    }
  }
}

/**
 * True when this formula means the same thing in every row — every reference it
 * makes into its own table lands on its own row, so one stored formula can fill
 * an appended row unchanged. A formula reaching another row (`r.prev()`), a raw
 * address, or a `raw()` escape cannot be reasoned about and is not claimed.
 */
export function rowIndependent(
  expr: Expr,
  anchor: TableAnchor,
  row: number,
  context: ResolveContext,
): boolean {
  let ok = true
  const visit = (node: Expr): void => {
    if (!ok) return
    switch (node.k) {
      case 'ref': {
        const resolved = resolveRef(node.target, context)
        const insideTable =
          resolved.sheet === anchor.sheet &&
          resolved.rect.r >= anchor.firstDataRow &&
          resolved.rect.r <= anchor.lastDataRow
        if (insideTable && resolved.rect.rows === 1 && resolved.rect.r !== row) ok = false
        return
      }
      case 'addr':
      case 'raw':
      case 'rawTemplate':
        ok = false
        return
      case 'op':
        visit(node.l)
        visit(node.r)
        return
      case 'neg':
        visit(node.e)
        return
      case 'fn':
        for (const arg of node.args) visit(arg)
        return
      default:
    }
  }
  visit(expr)
  return ok
}

function emitKeyValue(
  table: TableNode,
  rect: Rect,
  startRow: number,
  sheetName: string,
  cells: Map<CellKey, Cell>,
  registry: Registry,
  definedNames: Map<string, DefinedName>,
): void {
  const entries = table.data as readonly KeyValueEntry[]
  const keys = new Map<string, Addr>()

  entries.forEach((entry, index) => {
    const r = startRow + index
    cells.set(cellKey(r, rect.c), { value: entry.label, style: 'kvLabel' })
    const cell: Cell = { style: 'kvValue' }
    if (isExpr(entry.value) || isRef(entry.value)) cell.expr = lift(entry.value)
    else cell.value = entry.value as CellValue
    if (entry.format) cell.format = entry.format
    cells.set(cellKey(r, rect.c + 1), cell)
    keys.set(entry.key, { r, c: rect.c + 1 })

    // Excel's defined names are workbook-global and case-insensitive. Two blocks
    // claiming one name means the exported formula points at whichever was
    // written last, while the evaluator resolves through the block the author
    // named — the viewer and Excel then disagree about the same cell.
    const existing = findDefinedName(definedNames, entry.key)
    if (existing && existing.owner !== table.name) {
      throw new Error(
        `duplicate defined name "${entry.key}" — claimed by "${existing.owner}" and ` +
          `"${table.name}". Excel's defined names are workbook-global and ` +
          'case-insensitive, so one would silently win and formulas referencing the ' +
          'other would read the wrong cell. Rename one of the keys.',
      )
    }
    assertUsableName(entry.key, table.name)
    definedNames.set(entry.key, {
      sheet: sheetName,
      addr: { r, c: rect.c + 1 },
      owner: table.name,
    })
  })

  registry.set(table.name, {
    kind: 'keyValue',
    name: table.name,
    sheet: sheetName,
    rect,
    keys,
  })
}

/** Names the construct a thrown error came from, which the stack alone does not. */
function withContext<T>(block: string, column: string, row: number, fn: () => T): T {
  try {
    return fn()
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error)
    throw new Error(`table "${block}", column "${column}", row ${row + 1}: ${message}`, {
      cause: error,
    })
  }
}

function findDefinedName(names: Map<string, DefinedName>, key: string): DefinedName | undefined {
  const lower = key.toLowerCase()
  for (const [name, value] of names) if (name.toLowerCase() === lower) return value
  return undefined
}

/** Excel rejects some names outright; better to say so than to write a broken file. */
const NAME_START = /^[A-Za-z_\\]/
const NAME_BODY = /^[A-Za-z0-9_.\\]+$/
const LOOKS_LIKE_ADDRESS = /^\$?[A-Za-z]{1,3}\$?\d{1,7}$/
const RESERVED = new Set(['r', 'c'])

function assertUsableName(key: string, block: string): void {
  const problem =
    key.length === 0
      ? 'is empty'
      : key.length > 255
        ? 'is longer than 255 characters'
        : !NAME_START.test(key)
          ? 'must start with a letter or underscore'
          : !NAME_BODY.test(key)
            ? 'may only contain letters, digits, underscore and full stop'
            : LOOKS_LIKE_ADDRESS.test(key)
              ? 'looks like a cell address'
              : RESERVED.has(key.toLowerCase())
                ? 'is reserved by Excel'
                : undefined

  if (problem) {
    throw new Error(
      `key "${key}" in block "${block}" cannot be an Excel defined name: it ${problem}. ` +
        'Key-value keys become defined names so exported formulas read `=B5*growth`, ' +
        'so they have to satisfy Excel’s rules for one.',
    )
  }
}

function readField(row: unknown, key: string): CellValue {
  if (typeof row !== 'object' || row === null) return null
  const value = (row as Record<string, unknown>)[key]
  if (value === undefined || value === null) return null
  if (typeof value === 'number' || typeof value === 'string' || typeof value === 'boolean') {
    return value
  }
  if (value instanceof Date) return value.toISOString().slice(0, 10)
  return String(value)
}

function assertUniqueNames(workbook: WorkbookNode): void {
  const seen = new Map<string, string>()
  for (const sheet of workbook.children) {
    walkBlocks(sheet.children, (block) => {
      if (block.kind !== 'table') return
      const previous = seen.get(block.name)
      if (previous) {
        const where =
          previous === sheet.name
            ? `twice on sheet "${sheet.name}"`
            : `on sheets "${previous}" and "${sheet.name}"`
        throw new Error(
          `duplicate block name "${block.name}" — used ${where}. ` +
            'Block names are workbook-global because ref() looks them up by name.',
        )
      }
      seen.set(block.name, sheet.name)
    })
  }
}

function walkBlocks(blocks: readonly Block[], visit: (block: Block) => void): void {
  for (const block of blocks) {
    visit(block)
    if (block.kind === 'stack' || block.kind === 'row') walkBlocks(block.children, visit)
  }
}

function boundsOf(placements: readonly Placement[]): Size {
  let rows = 0
  let cols = 0
  for (const { rect } of placements) {
    rows = Math.max(rows, rect.r + rect.rows)
    cols = Math.max(cols, rect.c + rect.cols)
  }
  return { rows, cols }
}

function parseFreeze(freeze: string): Addr {
  const match = /^(\$?)([A-Za-z]+)(\$?)(\d+)$/.exec(freeze.trim())
  if (!match) throw new SyntaxError(`<Sheet freeze> expects a cell like "B2", got "${freeze}"`)
  let c = 0
  for (const ch of (match[2] as string).toUpperCase()) c = c * 26 + (ch.charCodeAt(0) - 64)
  return { r: Number(match[4]) - 1, c: c - 1 }
}

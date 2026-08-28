import type { Expr, ExprInput } from '../formula/expr.js'
import { lift } from '../formula/expr.js'
import { parseFormula } from '../formula/parse.js'
import type { CellValue } from '../model/cell.js'
import type { Addr, Size } from '../model/geometry.js'
import type { Validation } from '../model/validation.js'
import type { Ref } from '../refs/ref.js'
import { isRef } from '../refs/ref.js'
import { asBlocks, asRuns, asSheets } from './children.js'
import type {
  Aggregate,
  Block,
  CellNode,
  ChartAxes,
  ChartKind,
  ChartNode,
  ChartSeries,
  ColumnSpec,
  KpiBandNode,
  KpiItem,
  NoteNode,
  PrintSetup,
  RowNode,
  SheetNode,
  SheetProtection,
  SpacerNode,
  SpillNode,
  StackNode,
  TableNode,
  WorkbookNode,
} from './nodes.js'

/**
 * Formula strings are a compatibility shim. They are parsed so the cell still
 * evaluates, and a dev-mode warning points at the structural equivalent — an
 * address written by hand survives exactly until someone inserts a row.
 */
function asExpr(formula: ExprInput): Expr {
  if (isRef(formula)) return { k: 'ref', target: formula }
  if (typeof formula === 'number' || typeof formula === 'boolean') return { k: 'lit', v: formula }
  if (typeof formula !== 'string') return formula
  const parsed = parseFormula(formula)
  if (typeof process !== 'undefined' && process.env?.NODE_ENV !== 'production') {
    const detail = parsed.degraded ? ` (${parsed.reason}; it will show as #NOT_EVALUATED)` : ''
    process.emitWarning?.(
      `open-sheet: formula string "${formula}" will not survive a row insert${detail}. ` +
        'Use r.cell(...) / ref(...) instead — see the sheet-authoring skill.',
    )
  }
  return parsed.expr
}

export function Workbook(props: { children?: unknown }): WorkbookNode {
  return { kind: 'workbook', children: asSheets(props.children) }
}

export function Sheet(props: {
  name: string
  freeze?: string
  origin?: Addr
  /** How this sheet prints. Forms want `{ orientation: 'portrait', fitToWidth: true }`. */
  print?: PrintSetup
  /** Lock everything but the named blocks' inputs, so a reader cannot type over a formula. */
  protect?: SheetProtection
  children?: unknown
}): SheetNode {
  if (!props.name) throw new TypeError('<Sheet> requires a name')
  const node: SheetNode = {
    kind: 'sheet',
    name: props.name,
    children: asBlocks(props.children, 'Sheet'),
  }
  if (props.freeze !== undefined) node.freeze = props.freeze
  if (props.origin !== undefined) node.origin = props.origin
  if (props.print !== undefined) node.print = props.print
  if (props.protect !== undefined) node.protect = props.protect
  return node
}

export function Stack(props: { gap?: number; children?: unknown }): StackNode {
  return { kind: 'stack', gap: props.gap ?? 1, children: asBlocks(props.children, 'Stack') }
}

export function Row(props: { gap?: number; children?: unknown }): RowNode {
  return { kind: 'row', gap: props.gap ?? 1, children: asBlocks(props.children, 'Row') }
}

/**
 * Derived rather than restated. As a hand-written duplicate it silently fell
 * behind ColumnSpec — `validate` was added to the spec, accepted by the
 * compiler, and rejected by `col()`, which is the only way anyone writes one.
 */
export type ColumnOptions<T> = Omit<ColumnSpec<T>, 'key'>

export function col<T = any>(key: string, options: ColumnOptions<T> = {}): ColumnSpec<T> {
  return { key, ...options }
}

export interface KeyValueEntry {
  key: string
  label: string
  value: CellValue | ExprInput
  format?: string
}

interface GridTableProps<T> {
  name: string
  /**
   * Off by default. A register wants the arrows; a printed invoice does not, and
   * they show up in print.
   */
  filter?: boolean
  /** Let the recipient append rows below the table and have the ranges follow. */
  appendable?: boolean
  data: readonly T[]
  columns: ColumnSpec<T>[]
  kind?: 'grid'
  title?: string
  showHeader?: boolean
  total?: Partial<Record<string, Aggregate>>
  style?: string
}

interface KeyValueTableProps {
  name: string
  data: readonly KeyValueEntry[]
  kind: 'keyValue'
  title?: string
  style?: string
}

export type TableProps<T> = GridTableProps<T> | KeyValueTableProps

const KEY_VALUE_COLUMNS: ColumnSpec<KeyValueEntry>[] = [
  { key: 'label', header: 'Name', value: (row) => row.label },
  { key: 'value', header: 'Value' },
]

export function Table<T = any>(props: TableProps<T>): TableNode<T> {
  if (!props.name) throw new TypeError('<Table> requires a name — it is what ref() points at')
  if (props.kind === 'keyValue') {
    const node: TableNode<any> = {
      kind: 'table',
      name: props.name,
      variant: 'keyValue',
      showHeader: false,
      data: props.data,
      columns: KEY_VALUE_COLUMNS,
    }
    if (props.title !== undefined) node.title = props.title
    if (props.style !== undefined) node.style = props.style
    return node as TableNode<T>
  }
  if (!props.columns?.length) throw new TypeError(`<Table name="${props.name}"> requires columns`)
  const node: TableNode<T> = {
    kind: 'table',
    name: props.name,
    variant: 'grid',
    showHeader: props.showHeader ?? true,
    data: props.data,
    columns: props.columns,
  }
  if (props.title !== undefined) node.title = props.title
  if (props.total !== undefined) node.total = props.total
  if (props.style !== undefined) node.style = props.style
  if (props.filter !== undefined) node.filter = props.filter
  if (props.appendable !== undefined) node.appendable = props.appendable
  if (props.appendable && props.showHeader === false) {
    throw new TypeError(
      `<Table name="${props.name}" appendable> has showHeader={false} — an Excel Table names its columns by their headers, so it needs them`,
    )
  }
  if (props.appendable && props.filter) {
    throw new TypeError(
      `<Table name="${props.name}"> has both appendable and filter — an Excel Table brings its own filter arrows, so filter is redundant and Excel rejects the second one. Drop filter.`,
    )
  }
  if (props.filter && props.showHeader === false) {
    throw new TypeError(
      `<Table name="${props.name}" filter> has showHeader={false} — the filter arrows live on the header row, so there is nowhere to put them`,
    )
  }
  return node
}

export function KpiBand(props: { items: KpiItem[]; style?: string }): KpiBandNode {
  if (!props.items?.length) throw new TypeError('<KpiBand> requires items')
  const node: KpiBandNode = { kind: 'kpiBand', items: props.items }
  if (props.style !== undefined) node.style = props.style
  return node
}

export function Cell(props: {
  value?: CellValue
  formula?: ExprInput
  format?: string
  style?: string
  span?: Size
  validate?: Validation
  note?: string
}): CellNode {
  const node: CellNode = { kind: 'cell' }
  if (props.value !== undefined) node.value = props.value
  if (props.formula !== undefined) node.expr = asExpr(props.formula)
  if (props.format !== undefined) node.format = props.format
  if (props.style !== undefined) node.style = props.style
  if (props.span !== undefined) node.span = props.span
  if (props.validate !== undefined) node.validate = props.validate
  if (props.note !== undefined) node.note = props.note
  return node
}

export function Note(props: { cols?: number; style?: string; children?: unknown }): NoteNode {
  const node: NoteNode = { kind: 'note', runs: asRuns(props.children), cols: props.cols ?? 4 }
  if (props.style !== undefined) node.style = props.style
  return node
}

/**
 * A native chart. In the .xlsx it is real chart XML bound to real ranges, so it
 * moves when the numbers do; an embedded picture would go stale the moment
 * someone changed a cell, which is the one thing this export must never do.
 */
export function Chart(props: {
  kind?: ChartKind
  title?: string
  categories: Ref
  series: ChartSeries[]
  axes?: ChartAxes
  dataLabels?: boolean
  rows?: number
  cols?: number
}): ChartNode {
  if (!props.series?.length) throw new TypeError('<Chart> requires at least one series')
  const node: ChartNode = {
    kind: 'chart',
    chart: props.kind ?? 'bar',
    categories: props.categories,
    series: props.series,
    rows: props.rows ?? 15,
    cols: props.cols ?? 6,
  }
  if (props.title !== undefined) node.title = props.title
  if (props.axes !== undefined) node.axes = props.axes
  if (props.dataLabels !== undefined) node.dataLabels = props.dataLabels
  return node
}

/**
 * `rows` and `cols` are required rather than inferred: the size of a FILTER or
 * SORT result is not knowable until recalculation, and a placement engine that
 * guessed would be guessing about collisions too. Declaring it makes the
 * footprint the author's decision and the layout deterministic.
 */
export function Spill(props: {
  formula: Expr | Ref
  rows: number
  cols: number
  format?: string
  style?: string
}): SpillNode {
  if (!(props.rows >= 1 && props.cols >= 1)) {
    throw new TypeError('<Spill> needs rows and cols of at least 1 — how much room to reserve')
  }
  const node: SpillNode = {
    kind: 'spill',
    expr: lift(props.formula),
    rows: props.rows,
    cols: props.cols,
  }
  if (props.format !== undefined) node.format = props.format
  if (props.style !== undefined) node.style = props.style
  return node
}

export function Spacer(props: { rows?: number; cols?: number }): SpacerNode {
  return { kind: 'spacer', rows: props.rows ?? 1, cols: props.cols ?? 1 }
}

export type { Block }

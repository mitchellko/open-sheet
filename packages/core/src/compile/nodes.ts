import type { Expr, ExprInput } from '../formula/expr.js'
import type { CellValue } from '../model/cell.js'
import type { Addr, Size } from '../model/geometry.js'
import type { ColorScale, Highlight, IconSet } from '../model/highlight.js'
import type { Validation } from '../model/validation.js'
import type { Ref, RowContext } from '../refs/ref.js'

/**
 * A chart the width of one cell, showing this row's own numbers. `of` names the
 * columns it reads; they must be next to each other, because the file format
 * stores one range per sparkline.
 */
export interface Sparkline {
  of: readonly string[]
  kind?: 'line' | 'column'
  color?: string
}

export interface DataBar {
  color?: string
  /** Negative values grow left from a zero baseline. */
  negativeColor?: string
}

export interface ColumnSpec<T = any> {
  key: string
  header?: string
  format?: string
  width?: number
  style?: string
  /** Draw an in-cell bar across this column's data range. Live in Excel, drawn in HTML. */
  bar?: boolean | DataBar
  /** A heatmap across the column's data range: two or three stops, lowest to highest. */
  scale?: ColorScale
  /** Arrows or traffic lights, split into thirds of the column's range. */
  icons?: IconSet
  /** One rule, or several applied in order. Live in Excel, drawn the same way in HTML. */
  highlight?: Highlight | readonly Highlight[]
  /** An in-cell trend across this row's other columns. */
  sparkline?: Sparkline
  /**
   * Wrap long text instead of letting it overflow. Excel does not wrap by
   * default, so a description column set narrower than its content spills into
   * the neighbouring cell or is clipped when printed.
   */
  wrap?: boolean
  /** Constrains what the recipient may type in this column's data cells. */
  validate?: Validation
  /**
   * Provenance for each cell, shown on hover: which export, which date range,
   * which caveat. Returning undefined leaves that row's cell without one.
   */
  note?: (row: T, index: number) => string | undefined
  value?: (row: T, index: number) => CellValue
  /**
   * A builder expression, or a formula string for compatibility. A string is
   * parsed where possible but is not the recommended path — it is exactly what
   * breaks when a row is inserted.
   */
  formula?: ((row: RowContext<T>) => ExprInput | null | undefined) | string
}

export type Aggregate = 'sum' | 'avg' | 'count' | 'min' | 'max'

export interface InlineRun {
  text: string
  emphasis?: 'bold' | 'italic' | 'code'
}

export interface StackNode {
  kind: 'stack'
  gap: number
  children: Block[]
}

export interface RowNode {
  kind: 'row'
  gap: number
  children: Block[]
}

export interface TableNode<T = any> {
  kind: 'table'
  /** Dropdown arrows on the header row, so the recipient can sort and filter. */
  filter?: boolean
  /**
   * Emit an Excel Table, so a row typed below the last one is taken into the
   * ranges. Without it a recipient can insert a row inside the data and the
   * formulas follow, but appending below silently does not — the total just
   * stops including the new row.
   */
  appendable?: boolean
  name: string
  variant: 'grid' | 'keyValue'
  title?: string
  showHeader: boolean
  data: readonly T[]
  columns: ColumnSpec<T>[]
  total?: Partial<Record<string, Aggregate>>
  style?: string
}

export interface KpiItem {
  label: string
  value: ExprInput | CellValue
  format?: string
}

export interface KpiBandNode {
  kind: 'kpiBand'
  items: KpiItem[]
  style?: string
}

export interface CellNode {
  kind: 'cell'
  value?: CellValue
  expr?: Expr
  validate?: Validation
  note?: string
  format?: string
  style?: string
  span?: Size
}

export interface NoteNode {
  kind: 'note'
  runs: InlineRun[]
  cols: number
  style?: string
}

export type ChartKind =
  | 'bar'
  | 'stackedBar'
  | 'line'
  | 'area'
  | 'stackedArea'
  | 'pie'
  /** Two measures against each other; `categories` becomes the x values. */
  | 'scatter'
  /** Bars with a line over them — actual against target. Each series says which it is. */
  | 'combo'

/**
 * An unlabelled axis with unformatted numbers is decoration, not information —
 * the reader cannot tell thousands from millions, or margin from revenue.
 */
export interface ChartAxes {
  /** Axis titles. */
  category?: string
  value?: string
  /** A number format for the value axis, in the same codes cells use. */
  valueFormat?: string
  /** Pin the value axis. Excel's automatic zero-based scale flattens a narrow series. */
  min?: number
  max?: number
  /** Title for the right-hand axis of a combo chart. */
  secondary?: string
}

export interface ChartSeries {
  name: string
  /** A column reference; resolved to a range after layout, like any other. */
  values: Ref
  /** Only read on a `combo` chart: how this one series is drawn. */
  as?: 'bar' | 'line'
  /**
   * Only read on a `combo` chart. A target expressed as a percentage next to
   * revenue in dollars needs its own axis, or one of the two is a flat line at
   * the bottom of the plot.
   */
  axis?: 'primary' | 'secondary'
}

export interface ChartNode {
  kind: 'chart'
  chart: ChartKind
  title?: string
  axes?: ChartAxes
  /** Print each point's value on the chart. */
  dataLabels?: boolean
  categories: Ref
  series: ChartSeries[]
  rows: number
  cols: number
}

/**
 * A formula that fills a declared rectangle. Excel would "spill" such a result,
 * deciding at recalculation time how many cells it occupies — which fights a
 * placement engine that owns every coordinate. We emit a legacy array formula
 * over the declared range instead: the footprint is a contract the file format
 * enforces, so the author reserves the space and nothing can collide with it.
 */
export interface SpillNode {
  kind: 'spill'
  expr: Expr
  rows: number
  cols: number
  format?: string
  style?: string
}

/**
 * Excel's protection is inverted from what people expect: every cell is locked
 * already, and locking only takes effect once the sheet is protected. So this
 * names what to *un*lock, and turns protection on.
 *
 * There is deliberately no password. Sheet protection is an accident-prevention
 * affordance, not access control — anyone with a zip tool removes it in a
 * minute — and offering a password field would imply a guarantee the file
 * format cannot keep.
 */
export interface SheetProtection {
  /** Blocks whose input cells the recipient may edit. Formulas inside them stay locked. */
  allow: readonly string[]
}

export interface SpacerNode {
  kind: 'spacer'
  rows: number
  cols: number
}

export type Block =
  | StackNode
  | RowNode
  | TableNode
  | KpiBandNode
  | CellNode
  | NoteNode
  | SpacerNode
  | SpillNode
  | ChartNode

export type PageSize = 'A4' | 'A3' | 'Letter' | 'Legal'
export type Orientation = 'portrait' | 'landscape'

/**
 * What goes in the margin of every printed page. Excel encodes these as
 * `&`-codes; you name the field and the framework produces the encoding, as
 * everywhere else. Mixing text and fields is an array:
 * `center: ['FY26 Budget']`, `right: ['Page ', pageNumber, ' of ', pageCount]`.
 */
export type MarginField =
  | { field: 'pageNumber' | 'pageCount' | 'date' | 'time' | 'sheetName' | 'fileName' }
  | { bold: MarginContent }
  | { italic: MarginContent }

export type MarginContent = string | MarginField | readonly (string | MarginField)[]

export interface PageMargin {
  left?: MarginContent
  center?: MarginContent
  right?: MarginContent
}

export interface PrintSetup {
  /** Portrait for forms, landscape for wide grids. Default: landscape. */
  orientation?: Orientation
  size?: PageSize
  /** Scale the sheet to one page wide. Forms almost always want this. */
  fitToWidth?: boolean
  /** Repeat the table header on every printed page. */
  repeatHeader?: boolean
  margin?: number
  header?: PageMargin
  footer?: PageMargin
  /**
   * Print only these blocks. Named, not a range: a working sheet nobody is meant
   * to print stays out of the printed output without anyone tracking rows.
   */
  printArea?: readonly string[]
  /**
   * Start a new page before each named block. Named rather than numbered — a
   * break at "row 47" is exactly the coordinate this framework exists to stop
   * people writing, and it goes stale the moment a row is inserted.
   */
  breakBefore?: readonly string[]
  /** Ignore colours when printing. Useful for a form that is faxed or photocopied. */
  blackAndWhite?: boolean
  /** Centre the printed content on the page. */
  center?: { horizontal?: boolean; vertical?: boolean }
}

export interface SheetNode {
  kind: 'sheet'
  protect?: SheetProtection
  name: string
  freeze?: string
  origin?: Addr
  print?: PrintSetup
  children: Block[]
}

export interface WorkbookNode {
  kind: 'workbook'
  children: SheetNode[]
}

/**
 * Typed as a total record of the union rather than a Set of strings, so adding a
 * node type without listing it here fails to compile. As a Set it silently did
 * not, and a new block kind was accepted by the type checker and then rejected
 * at runtime by `<Stack>` with "cannot contain".
 */
const BLOCK_KINDS: Record<Block['kind'], true> = {
  stack: true,
  row: true,
  table: true,
  kpiBand: true,
  cell: true,
  note: true,
  spacer: true,
  chart: true,
  spill: true,
}

export function isBlock(value: unknown): value is Block {
  if (typeof value !== 'object' || value === null) return false
  const kind = (value as { kind?: string }).kind
  return kind !== undefined && Object.hasOwn(BLOCK_KINDS, kind)
}

export function isSheet(value: unknown): value is SheetNode {
  return typeof value === 'object' && value !== null && (value as SheetNode).kind === 'sheet'
}

export function isWorkbook(value: unknown): value is WorkbookNode {
  return typeof value === 'object' && value !== null && (value as WorkbookNode).kind === 'workbook'
}

export function isContainer(block: Block): block is StackNode | RowNode {
  return block.kind === 'stack' || block.kind === 'row'
}

import type { CompiledWorkbook } from '../compile/emit.js'
import { originOf } from '../compile/origin.js'
import { fromA1, toA1 } from '../model/a1.js'
import { type Cell, cellKey, parseCellKey } from '../model/cell.js'
import type { Size } from '../model/geometry.js'
import { type ResolveContext, resolveRef } from '../refs/resolve.js'
import type { BinaryOp, Expr } from './expr.js'
import { lookup } from './functions.js'
import {
  type Computed,
  DIV0,
  errorFrom,
  isExcelError,
  isNotEvaluated,
  NA,
  NOT_EVALUATED,
  NUM,
  VALUE,
} from './value.js'

export type ValueMap = Map<string, Computed>

/**
 * A range operand evaluates to an array, and Excel's array semantics let one
 * flow through comparisons and arithmetic — `(B2:B4>B2)*1` is an array of ones
 * and zeros. Without this, SUMPRODUCT is on the whitelist but the only reason
 * anyone reaches for it does not work: summing a single plain range is just SUM.
 */
type Value = Computed | Computed[]

function isArray(value: Value): value is Computed[] {
  return Array.isArray(value)
}

/**
 * A range is read row-major into a flat array, which loses whether it was a
 * column or a row. Most functions do not care — SUM of nine cells is the same
 * either way. The array-returning ones do: formulajs reads a flat array as one
 * row, so SORT on a column of numbers sorts a single row by its first column
 * and hands it back untouched, looking for all the world like it worked.
 */
const SHAPE_SENSITIVE: ReadonlySet<string> = new Set([
  'SORT',
  'SORTBY',
  'UNIQUE',
  'FILTER',
  'TRANSPOSE',
])

const shapes = new WeakMap<Computed[], Size>()

function shaped(items: Computed[], rows: number, cols: number): Computed[] {
  shapes.set(items, { rows, cols })
  return items
}

/** Carries a shape onto a derived array, so `(range>4)` stays a column. */
function reshape(out: Computed[], from: Value): Computed[] {
  const shape = isArray(from) ? shapes.get(from) : undefined
  if (shape && shape.rows * shape.cols === out.length) shapes.set(out, shape)
  return out
}

function toGrid(value: Value): Computed[][] {
  if (!isArray(value)) return [[value]]
  const shape = shapes.get(value) ?? { rows: 1, cols: value.length }
  const grid: Computed[][] = []
  for (let r = 0; r < shape.rows; r += 1) {
    grid.push(value.slice(r * shape.cols, (r + 1) * shape.cols))
  }
  return grid
}

function key(sheet: string, r: number, c: number): string {
  return `${sheet}!${cellKey(r, c)}`
}

interface Node {
  sheet: string
  r: number
  c: number
  cell: Cell
}

export class CycleError extends Error {
  readonly cells: string[]
  constructor(cells: string[], described?: string[]) {
    // Named by the construct that produced each cell, not by a coordinate. The
    // author never wrote a coordinate, so `S!1,2` tells them nothing about which
    // line to change.
    super(
      `circular reference: ${(described ?? cells).join(' → ')}. ` +
        'Break the cycle in the source that produced these cells.',
    )
    this.name = 'CycleError'
    this.cells = cells
  }
}

function describeCell(id: string, book: CompiledWorkbook): string {
  const bang = id.lastIndexOf('!')
  const sheet = id.slice(0, bang)
  const { r, c } = parseCellKey(id.slice(bang + 1))
  const origin = originOf(book.registry, sheet, { r, c })
  if (!origin) return `${sheet}!${toA1({ r, c })}`

  const parts = [`"${origin.block}"`]
  if (origin.column) parts.push(`column "${origin.column}"`)
  if (origin.row !== undefined) parts.push(`row ${origin.row + 1}`)
  else if (origin.part !== 'data') parts.push(origin.part)
  return parts.join(' ')
}

export function evaluateWorkbook(book: CompiledWorkbook): ValueMap {
  const nodes = new Map<string, Node>()
  for (const sheet of book.sheets) {
    for (const [cellRef, cell] of sheet.cells) {
      const comma = cellRef.indexOf(',')
      const r = Number(cellRef.slice(0, comma))
      const c = Number(cellRef.slice(comma + 1))
      nodes.set(key(sheet.name, r, c), { sheet: sheet.name, r, c, cell })
    }
  }

  const values: ValueMap = new Map()
  const state = new Map<string, 'visiting' | 'done'>()
  const stack: string[] = []

  const contextFor = (sheet: string): ResolveContext => ({
    registry: book.registry,
    definedNames: book.definedNames,
    sheet,
  })

  function visit(id: string): Computed {
    const existing = values.get(id)
    if (state.get(id) === 'done') return existing ?? null

    if (state.get(id) === 'visiting') {
      const from = stack.indexOf(id)
      const cells = [...stack.slice(from), id]
      throw new CycleError(
        cells,
        cells.map((cell) => describeCell(cell, book)),
      )
    }

    const node = nodes.get(id)
    if (!node) return null

    if (node.cell.spillFrom) {
      // Visiting the origin fills this cell in passing. Marked visiting first,
      // so a formula that spills over a cell it reads is caught as the cycle
      // it is rather than reading a half-written value.
      state.set(id, 'visiting')
      stack.push(id)
      try {
        const from = parseCellKey(node.cell.spillFrom)
        visit(key(node.sheet, from.r, from.c))
      } finally {
        stack.pop()
      }
      const filled = values.get(id) ?? null
      values.set(id, filled)
      state.set(id, 'done')
      return filled
    }

    if (!node.cell.expr) {
      const literal = node.cell.value ?? null
      values.set(id, literal)
      state.set(id, 'done')
      return literal
    }

    state.set(id, 'visiting')
    stack.push(id)
    let result: Computed
    try {
      const raw = evaluateValue(node.cell.expr, contextFor(node.sheet), read)
      if (node.cell.spill) {
        fill(raw, node, node.cell.spill)
        result = values.get(id) ?? null
      } else {
        result = collapse(raw)
      }
    } catch (error) {
      if (error instanceof CycleError) throw error
      // A resolution failure surfaces here, far from the formula that caused it.
      // Without the construct, finding it in a long workbook means bisecting by
      // hand — which is what one tester actually had to do.
      const message = error instanceof Error ? error.message : String(error)
      throw new Error(`${describeCell(id, book)}: ${message}`, { cause: error })
    } finally {
      stack.pop()
    }
    values.set(id, result)
    state.set(id, 'done')
    return result
  }

  /**
   * Writes one array across the rectangle its formula reserved. Cells the array
   * does not reach get #N/A — which is what a real spreadsheet shows for an
   * array formula wider than its result, and is visibly not a number.
   */
  function fill(value: Value, node: Node, size: Size): void {
    // Not computed means not computed everywhere in the rectangle. Filling the
    // first cell with #NOT_EVALUATED and the rest with #N/A would claim a
    // spreadsheet condition we have no evidence for.
    const items = isNotEvaluated(value) ? null : Array.isArray(value) ? value : [value]
    let i = 0
    for (let r = 0; r < size.rows; r += 1) {
      for (let c = 0; c < size.cols; c += 1) {
        const item = items === null ? NOT_EVALUATED : (items[i] ?? NA)
        i += 1
        values.set(key(node.sheet, node.r + r, node.c + c), item)
      }
    }
  }

  function read(sheet: string, r: number, c: number): Computed {
    return visit(key(sheet, r, c))
  }

  for (const id of nodes.keys()) visit(id)
  return values
}

type Reader = (sheet: string, r: number, c: number) => Computed

export function evaluateExpr(expr: Expr, context: ResolveContext, read: Reader): Computed {
  return collapse(evaluateValue(expr, context, read))
}

/** A cell holds one value. An array that is not a single element cannot be shown. */
function collapse(value: Value): Computed {
  if (!isArray(value)) return value
  if (value.length === 1) return value[0] as Computed
  for (const item of value) if (isNotEvaluated(item)) return NOT_EVALUATED
  // Excel would spill this; we have no representation for that, and claiming
  // #VALUE! would assert Excel errors here when it does not.
  return NOT_EVALUATED
}

function evaluateValue(expr: Expr, context: ResolveContext, read: Reader): Value {
  switch (expr.k) {
    case 'lit':
      return expr.v
    case 'raw':
    case 'rawTemplate':
      return NOT_EVALUATED
    case 'addr': {
      const values = readAddr(expr.ref, context, read)
      return values.length === 1 ? (values[0] as Computed) : values
    }
    case 'ref': {
      const values = readRef(expr, context, read)
      return values.length === 1 ? (values[0] as Computed) : values
    }
    case 'neg':
      return mapValue(evaluateValue(expr.e, context, read), (inner) => {
        if (isNotEvaluated(inner) || isExcelError(inner)) return inner
        const n = toNumber(inner)
        return typeof n === 'number' ? -n : n
      })
    case 'op':
      return broadcast(
        expr.op,
        evaluateValue(expr.l, context, read),
        evaluateValue(expr.r, context, read),
      )
    case 'fn':
      return applyFn(expr, context, read)
  }
}

/** Literal addresses from a parsed formula string, read against the current sheet. */
function readAddr(reference: string, context: ResolveContext, read: Reader): Computed[] {
  const [start, end] = reference.split(':')
  const a = fromA1(start as string)
  const b = end ? fromA1(end) : a
  const out: Computed[] = []
  for (let r = Math.min(a.r, b.r); r <= Math.max(a.r, b.r); r += 1) {
    for (let c = Math.min(a.c, b.c); c <= Math.max(a.c, b.c); c += 1) {
      out.push(read(context.sheet, r, c))
    }
  }
  return shaped(out, Math.abs(b.r - a.r) + 1, Math.abs(b.c - a.c) + 1)
}

function readRef(expr: Expr & { k: 'ref' }, context: ResolveContext, read: Reader): Computed[] {
  const resolved = resolveRef(expr.target, context)
  const out: Computed[] = []
  for (let r = resolved.rect.r; r < resolved.rect.r + resolved.rect.rows; r += 1) {
    for (let c = resolved.rect.c; c < resolved.rect.c + resolved.rect.cols; c += 1) {
      out.push(read(resolved.sheet, r, c))
    }
  }
  shaped(out, resolved.rect.rows, resolved.rect.cols)
  return out
}

function toNumber(value: Computed): number | Computed {
  if (value === null || value === undefined || value === '') return 0
  if (typeof value === 'number') return value
  if (typeof value === 'boolean') return value ? 1 : 0
  if (typeof value === 'string') {
    const n = Number(value)
    return Number.isNaN(n) ? VALUE : n
  }
  return value
}

function toText(value: Computed): string | Computed {
  if (value === null || value === undefined) return ''
  if (typeof value === 'string') return value
  if (typeof value === 'number') return String(value)
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE'
  return value
}

function mapValue(value: Value, fn: (item: Computed) => Computed): Value {
  return isArray(value) ? reshape(value.map(fn), value) : fn(value)
}

/** Elementwise where either side is an array, as Excel does. */
function broadcast(op: BinaryOp, left: Value, right: Value): Value {
  if (!isArray(left) && !isArray(right)) return applyOp(op, left, right)

  const length = Math.max(isArray(left) ? left.length : 1, isArray(right) ? right.length : 1)
  const at = (value: Value, i: number): Computed =>
    isArray(value) ? ((value[i] ?? null) as Computed) : value

  const out: Computed[] = []
  for (let i = 0; i < length; i += 1) out.push(applyOp(op, at(left, i), at(right, i)))
  return reshape(out, isArray(left) && shapes.has(left) ? left : right)
}

function applyOp(op: BinaryOp, left: Computed, right: Computed): Computed {
  if (isNotEvaluated(left) || isNotEvaluated(right)) return NOT_EVALUATED
  if (isExcelError(left)) return left
  if (isExcelError(right)) return right

  if (op === '&') {
    const l = toText(left)
    const r = toText(right)
    if (typeof l !== 'string') return l
    if (typeof r !== 'string') return r
    return l + r
  }

  if (op === '=' || op === '<>') {
    const equal = equals(left, right)
    return op === '=' ? equal : !equal
  }

  const l = toNumber(left)
  const r = toNumber(right)
  if (typeof l !== 'number') return l
  if (typeof r !== 'number') return r

  switch (op) {
    case '+':
      return l + r
    case '-':
      return l - r
    case '*':
      return l * r
    case '/':
      return r === 0 ? DIV0 : l / r
    case '^': {
      const result = l ** r
      return Number.isFinite(result) ? result : NUM
    }
    case '<':
      return l < r
    case '>':
      return l > r
    case '<=':
      return l <= r
    case '>=':
      return l >= r
  }
}

function isBlank(value: Computed): boolean {
  return value === null || value === undefined
}

/**
 * A blank cell takes the empty value of whatever it is compared against, which
 * is why `blank = 0` and `blank = ""` are both TRUE while `0 = ""` is FALSE —
 * the relation is not transitive, so a blank cannot simply be normalised to one
 * or the other. Verified against LibreOffice rather than assumed.
 */
function equals(left: Computed, right: Computed): boolean {
  const leftBlank = isBlank(left)
  const rightBlank = isBlank(right)
  if (leftBlank && rightBlank) return true
  if (leftBlank) return isEmptyFor(right)
  if (rightBlank) return isEmptyFor(left)

  if (typeof left === 'string' && typeof right === 'string') {
    // Excel's text comparison ignores case.
    return left.toLowerCase() === right.toLowerCase()
  }
  return left === right
}

/** The value a blank equals when compared with something of this type. */
function isEmptyFor(value: Computed): boolean {
  if (typeof value === 'number') return value === 0
  if (typeof value === 'string') return value === ''
  if (typeof value === 'boolean') return value === false
  return false
}

/**
 * Functions whose whole purpose is to see an error, or to not evaluate a branch.
 * Evaluating their arguments eagerly — the ordinary path below — means IFERROR
 * can never catch anything and IF fails on exactly the rows its guard exists
 * for. Excel evaluates these lazily; so must we.
 *
 * This matters beyond correctness: when a guarded division reports #DIV/0!
 * anyway, the tempting fix is to fudge the denominator, and a fabricated number
 * is the one failure this project cannot afford. IFERROR working is what lets an
 * author stay honest.
 */
function applyLazyFn(
  expr: Expr & { k: 'fn' },
  context: ResolveContext,
  read: Reader,
): Computed | undefined {
  const name = expr.name.toUpperCase()

  if (name === 'IFERROR' || name === 'IFNA') {
    const [value, fallback] = expr.args
    if (!value || !fallback) return undefined
    const result = evaluateExpr(value, context, read)
    if (isNotEvaluated(result)) return NOT_EVALUATED
    const caught = isExcelError(result) && (name === 'IFERROR' || result.code === '#N/A')
    return caught ? evaluateExpr(fallback, context, read) : result
  }

  // These exist to *inspect* a value, so the argument loop's "an error argument
  // makes the call an error" rule is exactly wrong for them: ISERROR(1/0) was
  // returning #DIV/0! where every spreadsheet returns TRUE. Implemented here
  // rather than passed to the library, which has no way to read our errors.
  if (PREDICATES.has(name)) {
    const [only] = expr.args
    if (!only) return undefined
    const value = evaluateExpr(only, context, read)
    if (isNotEvaluated(value)) return NOT_EVALUATED
    switch (name) {
      case 'ISERROR':
        return isExcelError(value)
      case 'ISNA':
        return isExcelError(value) && value.code === '#N/A'
      case 'ISBLANK':
        return value === null
      case 'ISNUMBER':
        return typeof value === 'number'
      case 'ISTEXT':
        return typeof value === 'string'
      default:
        return undefined
    }
  }

  if (name === 'IF') {
    const [test, then, otherwise] = expr.args
    if (!test || !then) return undefined
    const condition = evaluateExpr(test, context, read)
    if (isNotEvaluated(condition)) return NOT_EVALUATED
    if (isExcelError(condition)) return condition
    const taken = truthy(condition) ? then : otherwise
    return taken ? evaluateExpr(taken, context, read) : false
  }

  return undefined
}

function truthy(value: Computed): boolean {
  if (typeof value === 'boolean') return value
  if (typeof value === 'number') return value !== 0
  if (typeof value === 'string') return value.toUpperCase() === 'TRUE'
  return false
}

/**
 * Functions that act on one value at a time. Given a range, Excel maps them over
 * it — `ABS(A1:A9)` inside SUMPRODUCT is an array of magnitudes, and that is the
 * ordinary way to rank by size of change.
 *
 * The aggregates are deliberately absent: MIN over a range is a minimum, not
 * nine minimums. Getting that backwards would be worse than not mapping at all.
 */
const ELEMENTWISE: ReadonlySet<string> = new Set(['ABS', 'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'NOT'])

const PREDICATES: ReadonlySet<string> = new Set([
  'ISERROR',
  'ISNA',
  'ISBLANK',
  'ISNUMBER',
  'ISTEXT',
])

function applyFn(expr: Expr & { k: 'fn' }, context: ResolveContext, read: Reader): Value {
  const lazy = applyLazyFn(expr, context, read)
  if (lazy !== undefined) return lazy

  const implementation = lookup(expr.name)
  if (!implementation) return NOT_EVALUATED

  const args: Value[] = []
  for (const arg of expr.args) {
    const value = evaluateValue(arg, context, read)
    const items = isArray(value) ? value : [value]
    for (const item of items) {
      if (isNotEvaluated(item)) return NOT_EVALUATED
      if (isExcelError(item)) return item
    }
    args.push(isArray(value) && value.length === 1 ? (value[0] as Computed) : value)
  }

  if (ELEMENTWISE.has(expr.name.toUpperCase())) {
    const spread = args.find(isArray)
    if (spread) {
      return spread.map((_, i) =>
        fromLibrary(implementation(...args.map((arg) => (isArray(arg) ? arg[i] : arg)))),
      )
    }
  }

  // These read their arguments as rectangles, so they must be handed rectangles.
  const result = SHAPE_SENSITIVE.has(expr.name.toUpperCase())
    ? implementation(...args.map(toGrid))
    : implementation(...args)
  return valueFromLibrary(result)
}

/**
 * `#VALUE!` from the function library usually means *we* handed it something it
 * did not understand, not that Excel would error. Reporting it as an Excel error
 * puts a fabricated error into the exported cache and lets `iferror` swallow a
 * gap in this evaluator as though it were a real spreadsheet condition.
 * #NOT_EVALUATED says the true thing: we did not compute this. It is not
 * catchable, and it counts in the "not evaluated" badge.
 */
/**
 * The library returns array results as rows of columns. Flattened row-major
 * here and given its shape by the footprint the author declared — the array
 * supplies the order, `<Spill rows cols>` supplies the rectangle.
 */
function valueFromLibrary(result: unknown): Value {
  if (!Array.isArray(result)) return fromLibrary(result)
  const flat: Computed[] = []
  for (const row of result) {
    if (Array.isArray(row)) for (const item of row) flat.push(fromLibrary(item))
    else flat.push(fromLibrary(row))
  }
  return flat
}

function fromLibrary(result: unknown): Computed {
  if (result === null || result === undefined) return null
  if (typeof result === 'number') return Number.isFinite(result) ? result : NUM
  if (typeof result === 'string') {
    if (!result.startsWith('#')) return result
    return result === '#VALUE!' ? NOT_EVALUATED : errorFrom(result)
  }
  if (typeof result === 'boolean') return result
  // The library reports Excel's own errors as Error objects whose message is the
  // error code. Treating those as "we could not compute" was wrong twice over:
  // it hid a real spreadsheet condition, and it made IFNA unable to catch the
  // #N/A that a failed MATCH exists to produce.
  if (result instanceof Error) {
    const code = result.message.trim()
    return /^#[A-Z0-9/?!]+$/.test(code) ? errorFrom(code) : NOT_EVALUATED
  }
  if (Array.isArray(result)) return NOT_EVALUATED
  // A date in a workbook *is* a number — the serial — with a format on top. The
  // library hands back a Date, and treating that as "cannot compute" broke every
  // date chain at its first call.
  if (result instanceof Date) return toSerial(result)
  return NOT_EVALUATED
}

/**
 * Excel counts days from 1899-12-30, an epoch that exists because 1900 is
 * treated as a leap year for compatibility with Lotus 1-2-3. Dates before
 * 1900-03-01 are off by one in every spreadsheet; matching that is the point.
 */
const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30)
const DAY_MS = 86_400_000

export function toSerial(date: Date): number {
  const utc = Date.UTC(
    date.getFullYear(),
    date.getMonth(),
    date.getDate(),
    date.getHours(),
    date.getMinutes(),
    date.getSeconds(),
  )
  return (utc - EXCEL_EPOCH_UTC) / DAY_MS
}

export function fromSerial(serial: number): Date {
  return new Date(EXCEL_EPOCH_UTC + Math.round(serial * DAY_MS))
}

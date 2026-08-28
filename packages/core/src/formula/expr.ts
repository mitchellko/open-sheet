import { isRef, type Ref } from '../refs/ref.js'

export type Scalar = number | string | boolean

export type BinaryOp = '+' | '-' | '*' | '/' | '^' | '&' | '=' | '<' | '>' | '<=' | '>=' | '<>'

export interface LitExpr {
  k: 'lit'
  v: Scalar
}
export interface RefExpr {
  k: 'ref'
  target: Ref
}
export interface OpExpr {
  k: 'op'
  op: BinaryOp
  l: Expr
  r: Expr
}
export interface NegExpr {
  k: 'neg'
  e: Expr
}
export interface FnExpr {
  k: 'fn'
  name: string
  args: Expr[]
}
export interface RawExpr {
  k: 'raw'
  src: string
}

/**
 * A literal A1 address or range, from a hand-written formula string. Kept as a
 * distinct node rather than as `raw` so it can still be evaluated — but kept as
 * an *address* rather than resolved into a Ref, because it is precisely the
 * thing that breaks when a row is inserted, and that should stay visible.
 */
export interface AddrExpr {
  k: 'addr'
  ref: string
}

export type Expr =
  | LitExpr
  | RefExpr
  | OpExpr
  | NegExpr
  | FnExpr
  | RawExpr
  | AddrExpr
  | RawTemplateExpr

export type ExprInput = Expr | Ref | Scalar

const EXPR_KINDS = new Set(['lit', 'ref', 'op', 'neg', 'fn', 'raw', 'addr', 'rawTemplate'])

export function isExpr(value: unknown): value is Expr {
  if (typeof value !== 'object' || value === null) return false
  const k = (value as { k?: unknown }).k
  return typeof k === 'string' && EXPR_KINDS.has(k)
}

export function lift(input: ExprInput): Expr {
  if (isExpr(input)) return input
  if (isRef(input)) return { k: 'ref', target: input }
  return { k: 'lit', v: input }
}

/**
 * The v0 whitelist. Dispatch is never dynamic on an arbitrary name: a function
 * outside this set has no evaluation path and must go through `raw()`, which
 * renders as #NOT_EVALUATED rather than as an invented number.
 */
export const FUNCTIONS = [
  'SUM',
  'AVERAGE',
  'COUNT',
  'COUNTA',
  'MIN',
  'MAX',
  'ROUND',
  'ROUNDUP',
  'ROUNDDOWN',
  'ABS',
  'IF',
  'IFERROR',
  'IFNA',
  'AND',
  'OR',
  'NOT',
  'CONCATENATE',
  'NPV',
  'IRR',
  'PMT',
  'SORT',
  'SORTBY',
  'UNIQUE',
  'FILTER',
  'SEQUENCE',
  'TRANSPOSE',
  'SUMPRODUCT',
  // tier 1 — lookup and conditional aggregation
  'INDEX',
  'MATCH',
  'LARGE',
  'SMALL',
  'SUMIF',
  'SUMIFS',
  'COUNTIF',
  'COUNTIFS',
  'AVERAGEIF',
  'AVERAGEIFS',
  'MAXIFS',
  'MINIFS',
  // tier 2 — text
  'LEN',
  'LEFT',
  'RIGHT',
  'MID',
  'TRIM',
  'UPPER',
  'LOWER',
  'PROPER',
  'SUBSTITUTE',
  'REPLACE',
  'FIND',
  'SEARCH',
  'TEXT',
  'VALUE',
  'REPT',
  'TEXTJOIN',
  'CONCAT',
  // tier 2 — dates
  'DATE',
  'TODAY',
  'NOW',
  'YEAR',
  'MONTH',
  'DAY',
  'HOUR',
  'MINUTE',
  'WEEKDAY',
  'WEEKNUM',
  'EOMONTH',
  'EDATE',
  'DATEDIF',
  'DAYS',
  'NETWORKDAYS',
  'WORKDAY',
  'YEARFRAC',
  // tier 3 — finance and statistics
  'FV',
  'PV',
  'RATE',
  'NPER',
  'IPMT',
  'PPMT',
  'SLN',
  'DB',
  'DDB',
  'SYD',
  'XIRR',
  'XNPV',
  'MEDIAN',
  'RANK',
  'RANK.EQ',
  'RANK.AVG',
  'PERCENTILE',
  'QUARTILE',
  'STDEV',
  'MODE',
  'VAR',
  'VARP',
  'STDEVA',
  'CORREL',
  'SLOPE',
  'INTERCEPT',
  'FORECAST',
  'TREND',
  // tier 4 — logic, predicates and arithmetic
  'IFS',
  'SWITCH',
  'CHOOSE',
  'XOR',
  'ISBLANK',
  'ISNUMBER',
  'ISTEXT',
  'ISERROR',
  'ISNA',
  'ISEVEN',
  'ISODD',
  'MOD',
  'INT',
  'SIGN',
  'SQRT',
  'POWER',
  'EXP',
  'LN',
  'LOG',
  'LOG10',
  'CEILING',
  'FLOOR',
  'TRUNC',
  'PRODUCT',
  'SUBTOTAL',
  'AGGREGATE',
] as const

export type FunctionName = (typeof FUNCTIONS)[number]

const FUNCTION_SET: ReadonlySet<string> = new Set(FUNCTIONS)

export function isWhitelisted(name: string): name is FunctionName {
  return FUNCTION_SET.has(name.toUpperCase())
}

function fn(name: FunctionName, args: readonly ExprInput[]): FnExpr {
  return { k: 'fn', name, args: args.map(lift) }
}

function binary(op: BinaryOp, l: ExprInput, r: ExprInput): OpExpr {
  return { k: 'op', op, l: lift(l), r: lift(r) }
}

export const add = (l: ExprInput, r: ExprInput): OpExpr => binary('+', l, r)
export const sub = (l: ExprInput, r: ExprInput): OpExpr => binary('-', l, r)
export const mul = (l: ExprInput, r: ExprInput): OpExpr => binary('*', l, r)
export const div = (l: ExprInput, r: ExprInput): OpExpr => binary('/', l, r)
export const pow = (l: ExprInput, r: ExprInput): OpExpr => binary('^', l, r)
export const concat = (l: ExprInput, r: ExprInput): OpExpr => binary('&', l, r)
export const eq = (l: ExprInput, r: ExprInput): OpExpr => binary('=', l, r)
export const lt = (l: ExprInput, r: ExprInput): OpExpr => binary('<', l, r)
export const gt = (l: ExprInput, r: ExprInput): OpExpr => binary('>', l, r)
export const lte = (l: ExprInput, r: ExprInput): OpExpr => binary('<=', l, r)
export const gte = (l: ExprInput, r: ExprInput): OpExpr => binary('>=', l, r)
export const neq = (l: ExprInput, r: ExprInput): OpExpr => binary('<>', l, r)

export const neg = (e: ExprInput): NegExpr => ({ k: 'neg', e: lift(e) })

export const sum = (...args: ExprInput[]): FnExpr => fn('SUM', args)
export const avg = (...args: ExprInput[]): FnExpr => fn('AVERAGE', args)
export const count = (...args: ExprInput[]): FnExpr => fn('COUNT', args)
export const min = (...args: ExprInput[]): FnExpr => fn('MIN', args)
export const max = (...args: ExprInput[]): FnExpr => fn('MAX', args)
export const round = (value: ExprInput, digits: ExprInput = 0): FnExpr =>
  fn('ROUND', [value, digits])
export const abs = (value: ExprInput): FnExpr => fn('ABS', [value])
export const if_ = (test: ExprInput, then: ExprInput, otherwise: ExprInput): FnExpr =>
  fn('IF', [test, then, otherwise])
export const iferror = (value: ExprInput, fallback: ExprInput): FnExpr =>
  fn('IFERROR', [value, fallback])
export const ifna = (value: ExprInput, fallback: ExprInput): FnExpr => fn('IFNA', [value, fallback])
export const npv = (rate: ExprInput, ...values: ExprInput[]): FnExpr => fn('NPV', [rate, ...values])
export const irr = (values: ExprInput, guess?: ExprInput): FnExpr =>
  fn('IRR', guess === undefined ? [values] : [values, guess])
export const sumproduct = (...args: ExprInput[]): FnExpr => fn('SUMPRODUCT', args)

export const large = (range: ExprInput, k: ExprInput): FnExpr => fn('LARGE', [range, k])
export const small = (range: ExprInput, k: ExprInput): FnExpr => fn('SMALL', [range, k])
export const index = (range: ExprInput, position: ExprInput): FnExpr =>
  fn('INDEX', [range, position])
export const match = (value: ExprInput, range: ExprInput, kind: ExprInput = 0): FnExpr =>
  fn('MATCH', [value, range, kind])
/**
 * The criteria argument is a small expression language of its own — `">100"`,
 * `"<>done"`, `"apple"`. It is passed through as written, since inventing a
 * builder for it would mean re-implementing a syntax Excel already defines and
 * every spreadsheet user already knows.
 */
export const sumif = (range: ExprInput, criteria: ExprInput, sumRange?: ExprInput): FnExpr =>
  fn('SUMIF', sumRange === undefined ? [range, criteria] : [range, criteria, sumRange])
export const countif = (range: ExprInput, criteria: ExprInput): FnExpr =>
  fn('COUNTIF', [range, criteria])
export const averageif = (range: ExprInput, criteria: ExprInput, avgRange?: ExprInput): FnExpr =>
  fn('AVERAGEIF', avgRange === undefined ? [range, criteria] : [range, criteria, avgRange])

// --- text -------------------------------------------------------------------
export const len = (text: ExprInput): FnExpr => fn('LEN', [text])
export const left = (text: ExprInput, count: ExprInput = 1): FnExpr => fn('LEFT', [text, count])
export const right = (text: ExprInput, count: ExprInput = 1): FnExpr => fn('RIGHT', [text, count])
export const mid = (text: ExprInput, start: ExprInput, count: ExprInput): FnExpr =>
  fn('MID', [text, start, count])
export const trim = (text: ExprInput): FnExpr => fn('TRIM', [text])
export const upper = (text: ExprInput): FnExpr => fn('UPPER', [text])
export const lower = (text: ExprInput): FnExpr => fn('LOWER', [text])
export const proper = (text: ExprInput): FnExpr => fn('PROPER', [text])
export const substitute = (text: ExprInput, find: ExprInput, replace: ExprInput): FnExpr =>
  fn('SUBSTITUTE', [text, find, replace])
export const find = (needle: ExprInput, haystack: ExprInput, start?: ExprInput): FnExpr =>
  fn('FIND', start === undefined ? [needle, haystack] : [needle, haystack, start])
export const search = (needle: ExprInput, haystack: ExprInput, start?: ExprInput): FnExpr =>
  fn('SEARCH', start === undefined ? [needle, haystack] : [needle, haystack, start])
/**
 * Formats a number *inside* a formula, which is not the same as a cell's number
 * format: the result is text. Use a column `format` when the cell should stay a
 * number the reader can compute with.
 */
export const text = (value: ExprInput, format: ExprInput): FnExpr => fn('TEXT', [value, format])
export const value = (text: ExprInput): FnExpr => fn('VALUE', [text])
export const rept = (t: ExprInput, times: ExprInput): FnExpr => fn('REPT', [t, times])
export const textjoin = (
  delimiter: ExprInput,
  ignoreEmpty: ExprInput,
  ...parts: ExprInput[]
): FnExpr => fn('TEXTJOIN', [delimiter, ignoreEmpty, ...parts])

// --- dates ------------------------------------------------------------------
export const date = (year: ExprInput, month: ExprInput, day: ExprInput): FnExpr =>
  fn('DATE', [year, month, day])
export const today = (): FnExpr => fn('TODAY', [])
export const year = (serial: ExprInput): FnExpr => fn('YEAR', [serial])
export const month = (serial: ExprInput): FnExpr => fn('MONTH', [serial])
export const day = (serial: ExprInput): FnExpr => fn('DAY', [serial])
export const weekday = (serial: ExprInput, kind: ExprInput = 1): FnExpr =>
  fn('WEEKDAY', [serial, kind])
export const eomonth = (start: ExprInput, months: ExprInput = 0): FnExpr =>
  fn('EOMONTH', [start, months])
export const edate = (start: ExprInput, months: ExprInput): FnExpr => fn('EDATE', [start, months])
export const days = (end: ExprInput, start: ExprInput): FnExpr => fn('DAYS', [end, start])
export const networkdays = (start: ExprInput, end: ExprInput, holidays?: ExprInput): FnExpr =>
  fn('NETWORKDAYS', holidays === undefined ? [start, end] : [start, end, holidays])
export const workday = (start: ExprInput, days: ExprInput, holidays?: ExprInput): FnExpr =>
  fn('WORKDAY', holidays === undefined ? [start, days] : [start, days, holidays])
export const yearfrac = (start: ExprInput, end: ExprInput, basis: ExprInput = 0): FnExpr =>
  fn('YEARFRAC', [start, end, basis])

// --- finance ----------------------------------------------------------------
export const fv = (
  rate: ExprInput,
  periods: ExprInput,
  payment: ExprInput,
  present?: ExprInput,
): FnExpr =>
  fn('FV', present === undefined ? [rate, periods, payment] : [rate, periods, payment, present])
export const pv = (
  rate: ExprInput,
  periods: ExprInput,
  payment: ExprInput,
  future?: ExprInput,
): FnExpr =>
  fn('PV', future === undefined ? [rate, periods, payment] : [rate, periods, payment, future])
export const rate = (periods: ExprInput, payment: ExprInput, present: ExprInput): FnExpr =>
  fn('RATE', [periods, payment, present])
export const nper = (r: ExprInput, payment: ExprInput, present: ExprInput): FnExpr =>
  fn('NPER', [r, payment, present])
export const sln = (cost: ExprInput, salvage: ExprInput, life: ExprInput): FnExpr =>
  fn('SLN', [cost, salvage, life])
/** Cash flows with their own dates — the standard irregular-interval return. */
export const xirr = (values: ExprInput, dates: ExprInput, guess?: ExprInput): FnExpr =>
  fn('XIRR', guess === undefined ? [values, dates] : [values, dates, guess])
export const xnpv = (r: ExprInput, values: ExprInput, dates: ExprInput): FnExpr =>
  fn('XNPV', [r, values, dates])

// --- statistics -------------------------------------------------------------
export const median = (...args: ExprInput[]): FnExpr => fn('MEDIAN', args)
/**
 * Excel's own ranking. The `SUMPRODUCT((range>cell)*1)+1` idiom exists because
 * this was unavailable; it still works, but this reads as what it is.
 * `order` 0 ranks largest first, which is what people mean by "rank".
 */
export const rank = (value: ExprInput, range: ExprInput, order: ExprInput = 0): FnExpr =>
  fn('RANK', [value, range, order])
export const percentile = (range: ExprInput, k: ExprInput): FnExpr => fn('PERCENTILE', [range, k])
export const stdev = (...args: ExprInput[]): FnExpr => fn('STDEV', args)
export const correl = (a: ExprInput, b: ExprInput): FnExpr => fn('CORREL', [a, b])
export const forecast = (x: ExprInput, ys: ExprInput, xs: ExprInput): FnExpr =>
  fn('FORECAST', [x, ys, xs])

// --- logic and predicates ---------------------------------------------------
/**
 * Pairs of condition and result, ending with an optional fallback. Reads better
 * than nesting `if_` three deep, which is the usual alternative.
 */
export const ifs = (...args: ExprInput[]): FnExpr => fn('IFS', args)
export const switch_ = (value: ExprInput, ...args: ExprInput[]): FnExpr =>
  fn('SWITCH', [value, ...args])
export const choose = (index: ExprInput, ...options: ExprInput[]): FnExpr =>
  fn('CHOOSE', [index, ...options])
export const isblank = (v: ExprInput): FnExpr => fn('ISBLANK', [v])
export const isnumber = (v: ExprInput): FnExpr => fn('ISNUMBER', [v])
export const istext = (v: ExprInput): FnExpr => fn('ISTEXT', [v])
export const iserror = (v: ExprInput): FnExpr => fn('ISERROR', [v])
export const isna = (v: ExprInput): FnExpr => fn('ISNA', [v])

// --- arithmetic -------------------------------------------------------------
export const mod = (a: ExprInput, b: ExprInput): FnExpr => fn('MOD', [a, b])
export const int = (v: ExprInput): FnExpr => fn('INT', [v])
export const sign = (v: ExprInput): FnExpr => fn('SIGN', [v])
export const sqrt = (v: ExprInput): FnExpr => fn('SQRT', [v])
export const power = (base: ExprInput, exponent: ExprInput): FnExpr => fn('POWER', [base, exponent])
export const ln = (v: ExprInput): FnExpr => fn('LN', [v])
export const log = (v: ExprInput, base?: ExprInput): FnExpr =>
  fn('LOG', base === undefined ? [v] : [v, base])
export const ceiling = (v: ExprInput, significance: ExprInput = 1): FnExpr =>
  fn('CEILING', [v, significance])
export const floor = (v: ExprInput, significance: ExprInput = 1): FnExpr =>
  fn('FLOOR', [v, significance])
export const trunc = (v: ExprInput, digits: ExprInput = 0): FnExpr => fn('TRUNC', [v, digits])
export const product = (...args: ExprInput[]): FnExpr => fn('PRODUCT', args)
/**
 * The filter-aware aggregate. `SUBTOTAL(109, range)` sums only visible rows,
 * which is what a reader expects from a total under a filtered table — a plain
 * SUM would keep counting rows they have hidden.
 */
export const subtotal = (code: ExprInput, ...ranges: ExprInput[]): FnExpr =>
  fn('SUBTOTAL', [code, ...ranges])

export interface RawTemplateExpr {
  k: 'rawTemplate'
  strings: readonly string[]
  /**
   * Lifted to expressions, not stored as bare Refs. `r.cell('x')` already
   * returns a RefExpr, so the natural interpolation is an expression — an
   * earlier version typed it as `Ref | Expr` and implemented only `Ref`, which
   * made the commonest form crash inside the writer.
   */
  parts: Expr[]
}

/**
 * Escape hatch for formulas outside the whitelist. Exports verbatim; the viewer
 * shows #NOT_EVALUATED for it rather than guessing a value.
 *
 * Also usable as a tagged template, which is the form to prefer. A plain string
 * can only contain hand-written addresses — and this framework's one rule is
 * that you never write one, so the escape hatch should not be the thing that
 * forces you to. Interpolated references resolve after layout like any other,
 * so the formula survives an inserted row:
 *
 * ```ts
 * raw`=LARGE(${ref('costs').column('delta')}, 1)`
 * ```
 */
export function raw(src: string): RawExpr
export function raw(strings: TemplateStringsArray, ...values: ExprInput[]): RawTemplateExpr
export function raw(
  src: string | TemplateStringsArray,
  ...values: ExprInput[]
): RawExpr | RawTemplateExpr {
  if (typeof src === 'string') return { k: 'raw', src: src.replace(/^=/, '') }

  const strings = [...src]
  if (strings.length > 0) strings[0] = (strings[0] as string).replace(/^\s*=/, '')

  return {
    k: 'rawTemplate',
    strings,
    parts: values.map((value, i) => {
      if (value === undefined || value === null) {
        throw new TypeError(
          `raw\`…\` interpolation ${i + 1} is ${String(value)}. Interpolate a reference ` +
            "(ref('block').column('key')), a row cell (r.cell('key')), an expression, or a literal.",
        )
      }
      return lift(value)
    }),
  }
}

export const counta = (...args: ExprInput[]): FnExpr => fn('COUNTA', args)
export const roundup = (value: ExprInput, digits: ExprInput = 0): FnExpr =>
  fn('ROUNDUP', [value, digits])
export const rounddown = (value: ExprInput, digits: ExprInput = 0): FnExpr =>
  fn('ROUNDDOWN', [value, digits])
export const and = (...args: ExprInput[]): FnExpr => fn('AND', args)
export const or = (...args: ExprInput[]): FnExpr => fn('OR', args)
export const not = (value: ExprInput): FnExpr => fn('NOT', [value])
export const xor = (...args: ExprInput[]): FnExpr => fn('XOR', args)
/** `concat` is the `&` operator; this is Excel's legacy variadic function. */
export const concatenate = (...args: ExprInput[]): FnExpr => fn('CONCATENATE', args)
/**
 * Excel's CONCAT. Named `join` because `concat` is the `&` operator — and
 * unlike either of those, this one takes whole ranges.
 */
export const join = (...args: ExprInput[]): FnExpr => fn('CONCAT', args)
export const replace = (
  text: ExprInput,
  start: ExprInput,
  count: ExprInput,
  replacement: ExprInput,
): FnExpr => fn('REPLACE', [text, start, count, replacement])

/** Excel returns a payment as a negative number; negate it if you want it positive. */
export const pmt = (
  rate: ExprInput,
  periods: ExprInput,
  present: ExprInput,
  future: ExprInput = 0,
  type: ExprInput = 0,
): FnExpr => fn('PMT', [rate, periods, present, future, type])
export const ipmt = (
  rate: ExprInput,
  period: ExprInput,
  periods: ExprInput,
  present: ExprInput,
): FnExpr => fn('IPMT', [rate, period, periods, present])
export const ppmt = (
  rate: ExprInput,
  period: ExprInput,
  periods: ExprInput,
  present: ExprInput,
): FnExpr => fn('PPMT', [rate, period, periods, present])
export const db = (
  cost: ExprInput,
  salvage: ExprInput,
  life: ExprInput,
  period: ExprInput,
): FnExpr => fn('DB', [cost, salvage, life, period])
export const ddb = (
  cost: ExprInput,
  salvage: ExprInput,
  life: ExprInput,
  period: ExprInput,
): FnExpr => fn('DDB', [cost, salvage, life, period])
export const syd = (
  cost: ExprInput,
  salvage: ExprInput,
  life: ExprInput,
  period: ExprInput,
): FnExpr => fn('SYD', [cost, salvage, life, period])

/**
 * The plural forms take the summed/averaged range first, then criteria in
 * `range, test` pairs — the opposite order from `sumif`. That is Excel's
 * inconsistency, not ours; keeping it means a reader of the exported file sees
 * what they expect.
 */
export const sumifs = (sumRange: ExprInput, ...criteria: ExprInput[]): FnExpr =>
  fn('SUMIFS', [sumRange, ...criteria])
export const countifs = (...criteria: ExprInput[]): FnExpr => fn('COUNTIFS', criteria)
export const averageifs = (avgRange: ExprInput, ...criteria: ExprInput[]): FnExpr =>
  fn('AVERAGEIFS', [avgRange, ...criteria])
export const maxifs = (maxRange: ExprInput, ...criteria: ExprInput[]): FnExpr =>
  fn('MAXIFS', [maxRange, ...criteria])
export const minifs = (minRange: ExprInput, ...criteria: ExprInput[]): FnExpr =>
  fn('MINIFS', [minRange, ...criteria])

export const now = (): FnExpr => fn('NOW', [])
export const hour = (serial: ExprInput): FnExpr => fn('HOUR', [serial])
export const minute = (serial: ExprInput): FnExpr => fn('MINUTE', [serial])
export const weeknum = (serial: ExprInput, kind: ExprInput = 1): FnExpr =>
  fn('WEEKNUM', [serial, kind])
export const datedif = (start: ExprInput, end: ExprInput, unit: ExprInput): FnExpr =>
  fn('DATEDIF', [start, end, unit])

export const quartile = (range: ExprInput, q: ExprInput): FnExpr => fn('QUARTILE', [range, q])
export const mode = (...args: ExprInput[]): FnExpr => fn('MODE', args)
/** `var` is a reserved word; the trailing underscore follows `if_` and `switch_`. */
export const var_ = (...args: ExprInput[]): FnExpr => fn('VAR', args)
export const varp = (...args: ExprInput[]): FnExpr => fn('VARP', args)
export const stdeva = (...args: ExprInput[]): FnExpr => fn('STDEVA', args)
export const slope = (ys: ExprInput, xs: ExprInput): FnExpr => fn('SLOPE', [ys, xs])
export const intercept = (ys: ExprInput, xs: ExprInput): FnExpr => fn('INTERCEPT', [ys, xs])
export const trend = (ys: ExprInput, xs: ExprInput, newXs?: ExprInput): FnExpr =>
  fn('TREND', newXs === undefined ? [ys, xs] : [ys, xs, newXs])

export const iseven = (value: ExprInput): FnExpr => fn('ISEVEN', [value])
export const isodd = (value: ExprInput): FnExpr => fn('ISODD', [value])
export const exp = (value: ExprInput): FnExpr => fn('EXP', [value])
export const log10 = (value: ExprInput): FnExpr => fn('LOG10', [value])
export const aggregate = (code: ExprInput, options: ExprInput, ...ranges: ExprInput[]): FnExpr =>
  fn('AGGREGATE', [code, options, ...ranges])

/**
 * These return a rectangle rather than one value, so they belong in a `<Spill>`
 * whose declared footprint says how much room to reserve. Used in an ordinary
 * cell the result has nowhere to go, and the cell shows #NOT_EVALUATED.
 */
export const sort = (range: ExprInput, byColumn: ExprInput = 1, order: ExprInput = 1): FnExpr =>
  fn('SORT', [range, byColumn, order])
export const sortby = (range: ExprInput, by: ExprInput, order: ExprInput = 1): FnExpr =>
  fn('SORTBY', [range, by, order])
export const unique = (range: ExprInput): FnExpr => fn('UNIQUE', [range])
export const filter = (range: ExprInput, keep: ExprInput, ifEmpty?: ExprInput): FnExpr =>
  fn('FILTER', ifEmpty === undefined ? [range, keep] : [range, keep, ifEmpty])
export const sequence = (rows: ExprInput, cols: ExprInput = 1, start: ExprInput = 1): FnExpr =>
  fn('SEQUENCE', [rows, cols, start])
export const transpose = (range: ExprInput): FnExpr => fn('TRANSPOSE', [range])

import type { Expr, ExprInput } from './expr.js'
import { lift } from './expr.js'

/**
 * Builds a call to any function, whitelisted or not. The harness exists to
 * decide what belongs on the whitelist, so it must be able to ask about
 * functions that are not on it yet.
 */
export function call(name: string, ...args: ExprInput[]): Expr {
  return { k: 'fn', name, args: args.map(lift) }
}

export interface FunctionCase {
  /** Groups the report; several cases per function is normal. */
  fn: string
  /** What this case is checking, when the arguments do not say it. */
  note?: string
  /** Laid out in one column; `cell(i)` and `range(a, b)` point into it. */
  data: (number | string)[]
  build: (cell: (index: number) => Expr, range: (from: number, to: number) => Expr) => Expr
  /** What a real spreadsheet produces. Floats compare with a tolerance. */
  expect: number | string | boolean
}

const N = [3, 9, 5, 1, 7]

/**
 * One case per behaviour worth trusting, not one per function. A function joins
 * the whitelist only when its cases agree with a real engine — anything else is
 * a promise the whitelist cannot keep.
 */
export const CASES: FunctionCase[] = [
  // --- already whitelisted: these guard against regressions -----------------
  { fn: 'SUM', data: N, build: (_, r) => call('SUM', r(0, 4)), expect: 25 },
  { fn: 'AVERAGE', data: N, build: (_, r) => call('AVERAGE', r(0, 4)), expect: 5 },
  { fn: 'MIN', data: N, build: (_, r) => call('MIN', r(0, 4)), expect: 1 },
  { fn: 'MAX', data: N, build: (_, r) => call('MAX', r(0, 4)), expect: 9 },
  { fn: 'COUNT', data: N, build: (_, r) => call('COUNT', r(0, 4)), expect: 5 },
  {
    fn: 'ROUND',
    data: [1234.567],
    build: (c) => call('ROUND', c(0), 2),
    expect: 1234.57,
  },
  {
    fn: 'ROUND',
    note: 'negative digits round to the left of the point',
    data: [1234.567],
    build: (c) => call('ROUND', c(0), -2),
    expect: 1200,
  },
  { fn: 'ABS', data: [-42.5], build: (c) => call('ABS', c(0)), expect: 42.5 },
  {
    fn: 'SUMPRODUCT',
    note: 'array comparison — the only reason to use it',
    data: N,
    build: (c, r) =>
      call('SUMPRODUCT', {
        k: 'op',
        op: '*',
        l: { k: 'op', op: '>', l: r(0, 4), r: c(0) },
        r: { k: 'lit', v: 1 },
      }),
    expect: 3,
  },

  // --- operators and coercion ------------------------------------------------
  // Suggested by @ericweichun (#53): conformance is not only about functions.
  // A blank cell takes the empty value of whatever it is compared against, so
  // `blank = 0` and `blank = ""` are both true while `0 = ""` is false. The
  // relation is not transitive, which is why blanks cannot just be normalised.
  {
    fn: '= (blank vs zero)',
    data: [],
    build: (c) => ({ k: 'op', op: '=', l: c(1), r: { k: 'lit', v: 0 } }),
    expect: true,
  },
  {
    fn: '= (blank vs empty string)',
    data: [],
    build: (c) => ({ k: 'op', op: '=', l: c(1), r: { k: 'lit', v: '' } }),
    expect: true,
  },
  {
    fn: '<> (blank vs zero)',
    data: [],
    build: (c) => ({ k: 'op', op: '<>', l: c(1), r: { k: 'lit', v: 0 } }),
    expect: false,
  },
  {
    fn: '<> (blank vs empty string)',
    data: [],
    build: (c) => ({ k: 'op', op: '<>', l: c(1), r: { k: 'lit', v: '' } }),
    expect: false,
  },
  {
    fn: '+ (blank as zero)',
    data: [],
    build: (c) => ({ k: 'op', op: '+', l: c(1), r: { k: 'lit', v: 1 } }),
    expect: 1,
  },
  {
    fn: '& (blank as empty text)',
    data: [],
    build: (c) => ({ k: 'op', op: '&', l: c(1), r: { k: 'lit', v: 'x' } }),
    expect: 'x',
  },
  {
    fn: '= (text is case-insensitive)',
    data: ['A'],
    build: (c) => ({ k: 'op', op: '=', l: c(0), r: { k: 'lit', v: 'a' } }),
    expect: true,
  },
  {
    fn: '^ (left associative, unlike most languages)',
    note: '=2^3^2 is 64 in a spreadsheet, not 512',
    data: [],
    build: () => ({
      k: 'op',
      op: '^',
      l: { k: 'op', op: '^', l: { k: 'lit', v: 2 }, r: { k: 'lit', v: 3 } },
      r: { k: 'lit', v: 2 },
    }),
    expect: 64,
  },

  // --- tier 1: lookup and conditional aggregation ---------------------------
  { fn: 'LARGE', data: N, build: (_, r) => call('LARGE', r(0, 4), 2), expect: 7 },
  { fn: 'SMALL', data: N, build: (_, r) => call('SMALL', r(0, 4), 2), expect: 3 },
  {
    fn: 'MATCH',
    note: 'exact match returns a 1-based position',
    data: N,
    build: (_, r) => call('MATCH', 5, r(0, 4), 0),
    expect: 3,
  },
  {
    fn: 'INDEX',
    data: N,
    build: (_, r) => call('INDEX', r(0, 4), 2),
    expect: 9,
  },
  {
    fn: 'INDEX+MATCH',
    note: 'the pair that replaces VLOOKUP',
    data: N,
    build: (_, r) => call('INDEX', r(0, 4), call('MATCH', 7, r(0, 4), 0)),
    expect: 7,
  },
  {
    fn: 'SUMIF',
    note: 'criteria as a comparison string',
    data: N,
    build: (_, r) => call('SUMIF', r(0, 4), '>4'),
    expect: 21,
  },
  {
    fn: 'COUNTIF',
    data: N,
    build: (_, r) => call('COUNTIF', r(0, 4), '>4'),
    expect: 3,
  },
  {
    fn: 'AVERAGEIF',
    data: N,
    build: (_, r) => call('AVERAGEIF', r(0, 4), '>4'),
    expect: 7,
  },
  {
    fn: 'RANK',
    note: 'the library exposes it as RANK.EQ, so it resolves through an alias',
    data: N,
    build: (c, r) => call('RANK', c(1), r(0, 4), 0),
    expect: 1,
  },

  // --- tier 2: text and dates ----------------------------------------------
  { fn: 'LEN', data: ['hello'], build: (c) => call('LEN', c(0)), expect: 5 },
  { fn: 'LEFT', data: ['hello'], build: (c) => call('LEFT', c(0), 2), expect: 'he' },
  { fn: 'RIGHT', data: ['hello'], build: (c) => call('RIGHT', c(0), 2), expect: 'lo' },
  { fn: 'MID', data: ['hello'], build: (c) => call('MID', c(0), 2, 3), expect: 'ell' },
  { fn: 'TRIM', data: ['  a b  '], build: (c) => call('TRIM', c(0)), expect: 'a b' },
  { fn: 'UPPER', data: ['abc'], build: (c) => call('UPPER', c(0)), expect: 'ABC' },
  {
    fn: 'SUBSTITUTE',
    data: ['a-b-c'],
    build: (c) => call('SUBSTITUTE', c(0), '-', '+'),
    expect: 'a+b+c',
  },
  {
    fn: 'TEXT',
    note: 'number formatting inside a formula',
    data: [0.1234],
    build: (c) => call('TEXT', c(0), '0.0%'),
    expect: '12.3%',
  },
  {
    fn: 'YEAR',
    note: 'a date literal, not a serial number',
    data: [],
    build: () => call('YEAR', call('DATE', 2026, 8, 22)),
    expect: 2026,
  },
  {
    fn: 'EOMONTH',
    note: 'end of month, as a serial — the comparison is on the number',
    data: [],
    build: () => call('DAY', call('EOMONTH', call('DATE', 2026, 2, 10), 0)),
    expect: 28,
  },
  {
    fn: 'EDATE',
    data: [],
    build: () => call('MONTH', call('EDATE', call('DATE', 2026, 1, 31), 1)),
    expect: 2,
  },

  {
    fn: 'DATEDIF',
    note: 'whole months between two dates',
    data: [],
    build: () => call('DATEDIF', call('DATE', 2026, 1, 15), call('DATE', 2026, 8, 20), 'M'),
    expect: 7,
  },
  {
    fn: 'NETWORKDAYS',
    note: 'working days, excluding weekends',
    data: [],
    build: () => call('NETWORKDAYS', call('DATE', 2026, 8, 3), call('DATE', 2026, 8, 14)),
    expect: 10,
  },
  {
    fn: 'WEEKDAY',
    note: '2026-08-24 is a Monday; default numbering makes that 2',
    data: [],
    build: () => call('WEEKDAY', call('DATE', 2026, 8, 24)),
    expect: 2,
  },
  {
    fn: 'DAYS',
    data: [],
    build: () => call('DAYS', call('DATE', 2026, 3, 1), call('DATE', 2026, 2, 1)),
    expect: 28,
  },
  {
    fn: 'YEARFRAC',
    data: [],
    build: () =>
      call('ROUND', call('YEARFRAC', call('DATE', 2026, 1, 1), call('DATE', 2026, 7, 1)), 4),
    expect: 0.5,
  },
  {
    fn: 'TEXTJOIN',
    note: 'post-2007, so it needs the _xlfn prefix to work anywhere',
    data: [],
    build: () => call('TEXTJOIN', '-', true, 'a', 'b', 'c'),
    expect: 'a-b-c',
  },
  {
    fn: 'FIND',
    note: 'case sensitive, 1-based',
    data: ['hello'],
    build: (c) => call('FIND', 'll', c(0)),
    expect: 3,
  },
  {
    fn: 'SEARCH',
    note: 'case insensitive, unlike FIND',
    data: ['Hello'],
    build: (c) => call('SEARCH', 'H', c(0)),
    expect: 1,
  },
  {
    fn: 'PROPER',
    data: ['hello world'],
    build: (c) => call('PROPER', c(0)),
    expect: 'Hello World',
  },
  {
    fn: 'REPT',
    data: [],
    build: () => call('REPT', 'ab', 3),
    expect: 'ababab',
  },
  {
    fn: 'VALUE',
    data: ['42'],
    build: (c) => call('VALUE', c(0)),
    expect: 42,
  },

  // --- tier 3: finance and statistics --------------------------------------
  { fn: 'MEDIAN', data: N, build: (_, r) => call('MEDIAN', r(0, 4)), expect: 5 },
  {
    fn: 'PMT',
    note: 'negative by convention: it is a payment out',
    data: [],
    build: () => call('ROUND', call('PMT', 0.05 / 12, 60, 100000), 2),
    expect: -1887.12,
  },
  {
    fn: 'FV',
    data: [],
    build: () => call('ROUND', call('FV', 0.05, 10, 0, -1000), 2),
    expect: 1628.89,
  },
  {
    fn: 'PV',
    data: [],
    build: () => call('ROUND', call('PV', 0.05, 10, 0, -1000), 2),
    expect: 613.91,
  },
  {
    fn: 'SLN',
    data: [],
    build: () => call('SLN', 10000, 1000, 5),
    expect: 1800,
  },

  // --- tier 4: logic and predicates ----------------------------------------
  { fn: 'MOD', data: [], build: () => call('MOD', 7, 3), expect: 1 },
  { fn: 'INT', data: [], build: () => call('INT', 7.8), expect: 7 },
  { fn: 'SIGN', data: [], build: () => call('SIGN', -3), expect: -1 },
  { fn: 'SQRT', data: [], build: () => call('SQRT', 16), expect: 4 },
  { fn: 'POWER', data: [], build: () => call('POWER', 2, 10), expect: 1024 },
  { fn: 'CEILING', data: [], build: () => call('CEILING', 4.1, 1), expect: 5 },
  { fn: 'FLOOR', data: [], build: () => call('FLOOR', 4.9, 1), expect: 4 },
  {
    fn: 'CHOOSE',
    data: [],
    build: () => call('CHOOSE', 2, 'a', 'b', 'c'),
    expect: 'b',
  },
  {
    fn: 'SWITCH',
    data: [],
    build: () => call('SWITCH', 2, 1, 'one', 2, 'two', 'other'),
    expect: 'two',
  },
  {
    fn: 'ISNUMBER',
    data: N,
    build: (c) => call('ISNUMBER', c(0)),
    expect: true,
  },
  {
    fn: 'IFS',
    data: N,
    build: (c) =>
      call('IFS', { k: 'op', op: '>', l: c(1), r: { k: 'lit', v: 5 } }, 'big', true, 'small'),
    expect: 'big',
  },
  { fn: 'TRUNC', data: [], build: () => call('TRUNC', 7.89, 1), expect: 7.8 },
  { fn: 'PRODUCT', data: N, build: (_, r) => call('PRODUCT', r(0, 4)), expect: 945 },
  { fn: 'LN', data: [], build: () => call('ROUND', call('LN', 100), 4), expect: 4.6052 },
  { fn: 'LOG', data: [], build: () => call('LOG', 1000, 10), expect: 3 },
  // The engine's rounded answer is the point of the case; Math.E is a different test.
  // biome-ignore lint/suspicious/noApproximativeNumericConstant: see above
  { fn: 'EXP', data: [], build: () => call('ROUND', call('EXP', 1), 4), expect: 2.7183 },
  { fn: 'ISBLANK', data: [], build: (c) => call('ISBLANK', c(1)), expect: true },
  { fn: 'ISTEXT', data: ['x'], build: (c) => call('ISTEXT', c(0)), expect: true },
  { fn: 'ISEVEN', data: [], build: () => call('ISEVEN', 4), expect: true },
  {
    fn: 'SUBTOTAL',
    note: '109 sums visible rows only — what a filtered total should do',
    data: N,
    build: (_, r) => call('SUBTOTAL', 109, r(0, 4)),
    expect: 25,
  },
  {
    fn: 'XOR',
    data: [],
    build: () => call('XOR', true, false),
    expect: true,
  },
  {
    fn: 'RATE',
    note: 'solved iteratively; agreement here is worth checking',
    data: [],
    build: () => call('ROUND', call('RATE', 10, -1000, 8000), 4),
    expect: 0.0428,
  },
  {
    fn: 'NPER',
    data: [],
    build: () => call('ROUND', call('NPER', 0.05, -1000, 8000), 2),
    expect: 10.47,
  },
  {
    fn: 'CORREL',
    note: 'perfectly correlated series, so the answer is exactly 1',
    data: [1, 2, 3],
    build: (_, r) => call('ROUND', call('CORREL', r(0, 2), r(0, 2)), 4),
    expect: 1,
  },
  {
    fn: 'XNPV',
    note: 'the gap M7 left open — dates are a parallel range of serials',
    data: [-1000, 600, 700, 45658, 45839, 46023],
    build: (_, r) => call('ROUND', call('XNPV', 0.1, r(0, 2), r(3, 5)), 2),
    expect: 208.67,
  },

  // --- M7 builders reachable for the first time -----------------------------
  {
    fn: 'RANK',
    note: 'order 1 ranks smallest first',
    data: N,
    build: (c, r) => call('RANK', c(1), r(0, 4), 1),
    expect: 5,
  },
  { fn: 'AND', data: N, build: (c) => call('AND', call('ISNUMBER', c(0)), 1 < 2), expect: true },
  { fn: 'OR', data: N, build: (c) => call('OR', call('ISTEXT', c(0)), 1 < 2), expect: true },
  { fn: 'NOT', data: N, build: (c) => call('NOT', call('ISTEXT', c(0))), expect: true },
  { fn: 'XOR', data: N, build: () => call('XOR', true, true), expect: false },
  { fn: 'COUNTA', data: ['a', '', 'c'], build: (_, r) => call('COUNTA', r(0, 2)), expect: 2 },
  { fn: 'ROUNDUP', data: [1.234], build: (c) => call('ROUNDUP', c(0), 1), expect: 1.3 },
  { fn: 'ROUNDDOWN', data: [1.789], build: (c) => call('ROUNDDOWN', c(0), 1), expect: 1.7 },
  {
    fn: 'PMT',
    note: 'Excel returns a payment as negative — the sign is the contract',
    data: [0.05 / 12, 360, 500000],
    build: (c) => call('ROUND', call('PMT', c(0), c(1), c(2)), 2),
    expect: -2684.11,
  },
  {
    fn: 'IPMT',
    data: [0.05 / 12, 1, 360, 500000],
    build: (c) => call('ROUND', call('IPMT', c(0), c(1), c(2), c(3)), 2),
    expect: -2083.33,
  },
  {
    fn: 'PPMT',
    data: [0.05 / 12, 1, 360, 500000],
    build: (c) => call('ROUND', call('PPMT', c(0), c(1), c(2), c(3)), 2),
    expect: -600.77,
  },
  {
    fn: 'SUMIFS',
    note: 'range first, then criteria pairs — the opposite of SUMIF',
    data: N,
    build: (_, r) => call('SUMIFS', r(0, 4), r(0, 4), '>4'),
    expect: 21,
  },
  { fn: 'COUNTIFS', data: N, build: (_, r) => call('COUNTIFS', r(0, 4), '>4'), expect: 3 },
  {
    fn: 'AVERAGEIFS',
    data: N,
    build: (_, r) => call('AVERAGEIFS', r(0, 4), r(0, 4), '>4'),
    expect: 7,
  },
  { fn: 'MAXIFS', data: N, build: (_, r) => call('MAXIFS', r(0, 4), r(0, 4), '<8'), expect: 7 },
  { fn: 'MINIFS', data: N, build: (_, r) => call('MINIFS', r(0, 4), r(0, 4), '>2'), expect: 3 },
  { fn: 'LOG10', data: [1000], build: (c) => call('LOG10', c(0)), expect: 3 },
  { fn: 'ISODD', data: [4], build: (c) => call('ISODD', c(0)), expect: false },
  { fn: 'MODE', data: [3, 9, 3, 1], build: (_, r) => call('MODE', r(0, 3)), expect: 3 },
  {
    fn: 'STDEV',
    note: 'the bare name is the sample form, STDEV.S',
    data: N,
    build: (_, r) => call('ROUND', call('STDEV', r(0, 4)), 4),
    expect: 3.1623,
  },
  { fn: 'VAR', data: N, build: (_, r) => call('VAR', r(0, 4)), expect: 10 },
  { fn: 'VARP', data: N, build: (_, r) => call('VARP', r(0, 4)), expect: 8 },
  { fn: 'QUARTILE', data: N, build: (_, r) => call('QUARTILE', r(0, 4), 1), expect: 3 },
  { fn: 'PERCENTILE', data: N, build: (_, r) => call('PERCENTILE', r(0, 4), 0.5), expect: 5 },
  {
    fn: 'SLOPE',
    data: [1, 2, 3, 2, 4, 6],
    build: (_, r) => call('SLOPE', r(3, 5), r(0, 2)),
    expect: 2,
  },
  {
    fn: 'INTERCEPT',
    data: [1, 2, 3, 2, 4, 6],
    build: (_, r) => call('INTERCEPT', r(3, 5), r(0, 2)),
    expect: 0,
  },
  { fn: 'CONCAT', data: ['a', 'b', 'c'], build: (_, r) => call('CONCAT', r(0, 2)), expect: 'abc' },
  {
    fn: 'CONCATENATE',
    data: ['a', 'b'],
    build: (c) => call('CONCATENATE', c(0), c(1)),
    expect: 'ab',
  },
  {
    fn: 'REPLACE',
    data: ['2026-08'],
    build: (c) => call('REPLACE', c(0), 1, 4, '2027'),
    expect: '2027-08',
  },
  {
    fn: 'WEEKNUM',
    data: [],
    build: () => call('WEEKNUM', call('DATE', 2026, 8, 24), 2),
    expect: 35,
  },
  {
    fn: 'DATEDIF',
    data: [],
    build: () => call('DATEDIF', call('DATE', 2020, 1, 1), call('DATE', 2026, 8, 24), 'Y'),
    expect: 6,
  },
  {
    fn: 'DB',
    data: [100000, 10000, 5, 1],
    build: (c) => call('ROUND', call('DB', c(0), c(1), c(2), c(3)), 2),
    expect: 36900,
  },
  {
    fn: 'SYD',
    data: [100000, 10000, 5, 1],
    build: (c) => call('SYD', c(0), c(1), c(2), c(3)),
    expect: 30000,
  },
  {
    fn: 'DDB',
    data: [100000, 10000, 5, 1],
    build: (c) => call('DDB', c(0), c(1), c(2), c(3)),
    expect: 40000,
  },

  // --- whitelisted but never verified until now -----------------------------
  {
    fn: 'IF',
    data: N,
    build: (c) => call('IF', { k: 'op', op: '>', l: c(0), r: { k: 'lit', v: 2 } }, 'big', 'small'),
    expect: 'big',
  },
  {
    fn: 'IFERROR',
    data: [],
    build: () =>
      call('IFERROR', { k: 'op', op: '/', l: { k: 'lit', v: 1 }, r: { k: 'lit', v: 0 } }, 'caught'),
    expect: 'caught',
  },
  {
    fn: 'IFNA',
    note: 'catches #N/A and nothing else',
    data: ['x'],
    build: (c) => call('IFNA', call('MATCH', 'nope', c(0), 0), 'missing'),
    expect: 'missing',
  },
  {
    fn: 'NPV',
    data: [-1000, 600, 700],
    build: (_, r) => call('ROUND', call('NPV', 0.1, r(0, 2)), 2),
    expect: 112.7,
  },
  {
    fn: 'IRR',
    data: [-1000, 600, 700],
    build: (_, r) => call('ROUND', call('IRR', r(0, 2)), 4),
    expect: 0.1888,
  },
  {
    fn: 'XIRR',
    data: [-1000, 600, 700, 45658, 45839, 46023],
    build: (_, r) => call('ROUND', call('XIRR', r(0, 2), r(3, 5)), 4),
    expect: 0.4147,
  },
  { fn: 'LOWER', data: ['MiXeD'], build: (c) => call('LOWER', c(0)), expect: 'mixed' },
  { fn: 'DATE', data: [], build: () => call('YEAR', call('DATE', 2026, 8, 24)), expect: 2026 },
  { fn: 'MONTH', data: [], build: () => call('MONTH', call('DATE', 2026, 8, 24)), expect: 8 },
  { fn: 'DAY', data: [], build: () => call('DAY', call('DATE', 2026, 8, 24)), expect: 24 },
  {
    fn: 'TODAY',
    note: 'the value moves, so the case asks something that does not',
    data: [],
    build: () => call('ISNUMBER', call('TODAY')),
    expect: true,
  },
  {
    fn: 'NOW',
    note: 'likewise — comparing the instant would compare two clocks',
    data: [],
    build: () => call('INT', { k: 'op', op: '-', l: call('NOW'), r: call('TODAY') }),
    expect: 0,
  },
  { fn: 'HOUR', data: [], build: () => call('HOUR', 0.5), expect: 12 },
  { fn: 'MINUTE', data: [], build: () => call('MINUTE', 0.51), expect: 14 },
  {
    fn: 'WORKDAY',
    data: [],
    build: () => call('DAY', call('WORKDAY', call('DATE', 2026, 8, 21), 1)),
    expect: 24,
  },
  {
    fn: 'ISERROR',
    data: [],
    build: () =>
      call('ISERROR', { k: 'op', op: '/', l: { k: 'lit', v: 1 }, r: { k: 'lit', v: 0 } }),
    expect: true,
  },
  {
    fn: 'ISNA',
    data: ['x'],
    build: (c) => call('ISNA', call('MATCH', 'nope', c(0), 0)),
    expect: true,
  },
  {
    fn: 'STDEVA',
    data: N,
    build: (_, r) => call('ROUND', call('STDEVA', r(0, 4)), 4),
    expect: 3.1623,
  },
  {
    fn: 'RANK.EQ',
    note: 'the explicit name for what bare RANK means',
    data: N,
    build: (c, r) => call('RANK.EQ', c(1), r(0, 4), 0),
    expect: 1,
  },
  {
    fn: 'RANK.AVG',
    data: [3, 9, 3, 1],
    build: (c, r) => call('RANK.AVG', c(0), r(0, 3), 0),
    expect: 2.5,
  },
  {
    fn: 'FORECAST',
    data: [1, 2, 3, 2, 4, 6],
    build: (_, r) => call('FORECAST', 4, r(3, 5), r(0, 2)),
    expect: 8,
  },
  {
    fn: 'TREND',
    note: 'one new x, so the result fits a single cell',
    data: [1, 2, 3, 2, 4, 6],
    build: (_, r) => call('TREND', r(3, 5), r(0, 2), 4),
    expect: 8,
  },
  {
    fn: 'AGGREGATE',
    note: 'code 9 is SUM, option 6 ignores errors',
    data: N,
    build: (_, r) => call('AGGREGATE', 9, 6, r(0, 4)),
    expect: 25,
  },

  // --- TEXT with a date code, which the harness had never asked about --------
  {
    fn: 'TEXT',
    note: 'a date code, not a number one — the library returns the serial untouched',
    data: [],
    build: () => call('TEXT', call('DATE', 2026, 8, 1), 'yyyy-mm'),
    expect: '2026-08',
  },
  {
    fn: 'TEXT',
    note: 'literals in a code pass through; deliberately no month *name*, which follows the locale',
    data: [],
    build: () => call('TEXT', call('DATE', 2026, 8, 1), 'yyyy/mm/dd'),
    expect: '2026/08/01',
  },
  {
    fn: 'TEXT',
    note: 'and a number code still formats a number',
    data: [1234.5],
    build: (c) => call('TEXT', c(0), '#,##0.00'),
    expect: '1,234.50',
  },
]

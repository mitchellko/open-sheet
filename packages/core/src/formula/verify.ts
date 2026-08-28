import { compile } from '../compile/compile.js'
import type { CompiledWorkbook } from '../compile/emit.js'
import { toA1 } from '../model/a1.js'
import type { FunctionCase } from './cases.js'
import { evaluateWorkbook } from './evaluate.js'
import type { Expr } from './expr.js'
import { isWhitelisted } from './expr.js'
import { type Computed, isNotEvaluated } from './value.js'

/**
 * One case per row: its data across the left, its formula on the right. The
 * width follows the widest case — it was once a constant that counted the
 * index column too, so a case with six values wrote its last one into the
 * formula column and the whole run died on a circular reference.
 */
function dataWidth(cases: readonly FunctionCase[]): number {
  return Math.max(1, ...cases.map((testCase) => testCase.data.length))
}

export interface Layout {
  book: CompiledWorkbook
  /** A1 address of each case's formula cell, in case order. */
  addresses: string[]
  /** Zero-based column the results land in — for reading the engine's CSV back. */
  resultColumn: number
}

/**
 * Compiles every case into a single sheet, so verifying forty functions costs
 * one LibreOffice invocation rather than forty.
 */
export function layout(cases: readonly FunctionCase[]): Layout {
  const rows = cases.map((testCase, index) => {
    const row: Record<string, number | string | null> = { case: index }
    testCase.data.forEach((value, i) => {
      row[`d${i}`] = value
    })
    return row
  })

  const width = dataWidth(cases)
  const columns: unknown[] = [{ key: 'case', header: 'case' }]
  for (let i = 0; i < width; i += 1) {
    columns.push({ key: `d${i}`, header: `d${i}` })
  }

  columns.push({
    key: 'result',
    header: 'result',
    formula: (r: { index: number }) => {
      const testCase = cases[r.index] as FunctionCase
      // Both helpers point into this row's own data cells, so a case never has
      // to know where the harness placed it.
      const cell = (i: number): Expr => ({
        k: 'addr',
        ref: toA1({ r: r.index + 1, c: i + 1 }),
      })
      const range = (from: number, to: number): Expr => ({
        k: 'addr',
        ref: `${toA1({ r: r.index + 1, c: from + 1 })}:${toA1({ r: r.index + 1, c: to + 1 })}`,
      })
      return testCase.build(cell, range)
    },
  })

  const book = compile({
    kind: 'workbook',
    children: [
      {
        kind: 'sheet',
        name: 'Cases',
        children: [
          { kind: 'table', name: 'cases', variant: 'grid', showHeader: true, data: rows, columns },
        ],
      },
    ],
  })

  return {
    book,
    addresses: cases.map((_, index) => toA1({ r: index + 1, c: width + 1 })),
    resultColumn: width + 1,
  }
}

export type Outcome = 'agrees' | 'disagrees' | 'not-evaluated' | 'engine-failed'

export interface Result {
  fn: string
  note?: string
  outcome: Outcome
  whitelisted: boolean
  expected: number | string | boolean
  ours?: string
  theirs?: string
}

const TOLERANCE = 1e-6

function same(a: unknown, b: unknown): boolean {
  const x = Number(a)
  const y = Number(b)
  if (!Number.isNaN(x) && !Number.isNaN(y)) return Math.abs(x - y) < TOLERANCE
  return String(a).trim().toUpperCase() === String(b).trim().toUpperCase()
}

/**
 * Three outcomes, not two. "We cannot evaluate it" is the state that matters:
 * a function can sit on the whitelist and be useless, which is how SUMPRODUCT
 * shipped for weeks unable to do the one thing anyone wanted from it.
 */
export function compare(
  cases: readonly FunctionCase[],
  ours: readonly (Computed | undefined)[],
  theirs: readonly (string | undefined)[],
): Result[] {
  return cases.map((testCase, i) => {
    const mine = ours[i]
    const engine = theirs[i]

    const result: Result = {
      fn: testCase.fn,
      outcome: 'agrees',
      whitelisted: testCase.fn.split('+').every((name) => isWhitelisted(name)),
      expected: testCase.expect,
    }
    if (testCase.note) result.note = testCase.note

    if (engine === undefined || engine === '' || String(engine).startsWith('#')) {
      result.outcome = 'engine-failed'
      result.theirs = engine ?? '(missing)'
      return result
    }
    result.theirs = String(engine)

    if (mine === undefined || isNotEvaluated(mine)) {
      result.outcome = 'not-evaluated'
      return result
    }
    result.ours = String(mine)

    // The engine is the authority; `expect` guards the case itself from drifting.
    if (!same(mine, engine) || !same(engine, testCase.expect)) result.outcome = 'disagrees'
    return result
  })
}

export function evaluateCases(book: CompiledWorkbook, cases: readonly FunctionCase[]): Computed[] {
  const values = evaluateWorkbook(book)
  const column = dataWidth(cases) + 1
  return cases.map((_, index) => values.get(`Cases!${index + 1},${column}`) as Computed)
}

export function summarise(results: readonly Result[]): string {
  const by: Record<Outcome, Result[]> = {
    agrees: [],
    disagrees: [],
    'not-evaluated': [],
    'engine-failed': [],
  }
  for (const result of results) by[result.outcome].push(result)

  const lines = [
    `${results.length} cases: ${by.agrees.length} agree, ${by.disagrees.length} disagree, ` +
      `${by['not-evaluated'].length} not evaluated, ${by['engine-failed'].length} engine failed`,
  ]

  const ready = by['not-evaluated'].filter((r) => !r.whitelisted)
  if (ready.length > 0) {
    lines.push(
      '',
      'The engine computes these and we do not — candidates for the whitelist:',
      ...unique(ready.map((r) => `  ${r.fn} → ${r.theirs}`)),
    )
  }

  const broken = by['not-evaluated'].filter((r) => r.whitelisted)
  if (broken.length > 0) {
    lines.push(
      '',
      'Whitelisted but not evaluated — the whitelist is promising what it cannot do:',
      ...unique(broken.map((r) => `  ${r.fn}${r.note ? ` (${r.note})` : ''}`)),
    )
  }

  if (by.disagrees.length > 0) {
    lines.push(
      '',
      'We disagree with the engine:',
      ...by.disagrees.map(
        (r) => `  ${r.fn}: ours ${r.ours}, engine ${r.theirs}, case expects ${r.expected}`,
      ),
    )
  }

  if (by['engine-failed'].length > 0) {
    lines.push(
      '',
      'The engine could not compute these either — the case may be wrong:',
      ...unique(by['engine-failed'].map((r) => `  ${r.fn} → ${r.theirs}`)),
    )
  }

  return lines.join('\n')
}

function unique(values: string[]): string[] {
  return [...new Set(values)]
}

import { describe, expect, it } from 'vitest'
import { SORT, SORTBY, UNIQUE } from './arrays.js'

const column = (values: unknown[]) => values.map((v) => [v])
const flat = (grid: unknown[][]) => grid.map((row) => row[0])

/**
 * These three are ours because the function library's are wrong: its `SORT`
 * compares as strings, and its `UNIQUE` returns the input untouched. Both
 * failures are invisible on the small, same-width test data anyone reaches for
 * first — the cross-engine test that was supposed to cover SORT used
 * [300, 900, 500], three values of three digits each, and passed.
 */
describe('SORT', () => {
  it('orders numbers numerically, not as strings', () => {
    // The library gives 960000, 95000, 81000, 210000, 102000 — lexicographic,
    // and indistinguishable from correct while every value has six digits.
    const money = [960000, 102000, 210000, 95000, 81000]
    expect(flat(SORT(column(money), 1, -1))).toEqual([960000, 210000, 102000, 95000, 81000])
    expect(flat(SORT(column(money), 1, 1))).toEqual([81000, 95000, 102000, 210000, 960000])
  })

  it('sorts text case-insensitively', () => {
    expect(flat(SORT(column(['banana', 'Apple', 'cherry']), 1, 1))).toEqual([
      'Apple',
      'banana',
      'cherry',
    ])
  })

  it('puts numbers before text, as Excel does', () => {
    expect(flat(SORT(column(['b', 2, 'a', 1]), 1, 1))).toEqual([1, 2, 'a', 'b'])
  })

  it('sorts a table by the column named, keeping rows together', () => {
    const rows = [
      ['ana', 300],
      ['ben', 900],
      ['cai', 500],
    ]
    expect(SORT(rows, 2, -1)).toEqual([
      ['ben', 900],
      ['cai', 500],
      ['ana', 300],
    ])
  })

  it('is stable, so equal keys keep their order', () => {
    const rows = [
      ['first', 1],
      ['second', 1],
    ]
    expect(SORT(rows, 2, 1)).toEqual(rows)
  })
})

describe('UNIQUE', () => {
  it('actually deduplicates', () => {
    // The library returns the input unchanged — a whitelisted function that
    // answers wrongly is worse than one that is absent.
    expect(flat(UNIQUE(column([3, 9, 3, 1])))).toEqual([3, 9, 1])
  })

  it('compares a row by its whole contents', () => {
    expect(
      UNIQUE([
        ['a', 1],
        ['a', 2],
        ['a', 1],
      ]),
    ).toEqual([
      ['a', 1],
      ['a', 2],
    ])
  })

  it('keeps only what appears once when asked', () => {
    expect(flat(UNIQUE(column([3, 9, 3, 1]), false, true))).toEqual([9, 1])
  })
})

describe('SORTBY', () => {
  it('orders one range by another', () => {
    const names = column(['ana', 'ben', 'cai'])
    const scores = column([2, 3, 1])
    expect(flat(SORTBY(names, scores, 1))).toEqual(['cai', 'ana', 'ben'])
    expect(flat(SORTBY(names, scores, -1))).toEqual(['ben', 'ana', 'cai'])
  })

  it('falls through to the next key on a tie', () => {
    const names = column(['ana', 'ben', 'cai'])
    expect(flat(SORTBY(names, column([1, 1, 0]), 1, column([2, 1, 9]), 1))).toEqual([
      'cai',
      'ben',
      'ana',
    ])
  })
})

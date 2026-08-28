import { describe, expect, it } from 'vitest'
import * as entry from '../index.js'
import * as expr from './expr.js'
import { FUNCTIONS, isExpr } from './expr.js'

/**
 * M7 added 57 builders that lived in expr.ts for a whole release without ever
 * reaching the package entry. The verification harness imported them directly,
 * so every case passed while none of them was callable by a user. A builder
 * nobody can import is a builder that does not exist.
 */
describe('the public surface', () => {
  it('re-exports every formula builder', () => {
    expect(Object.keys(expr).filter((name) => !(name in entry))).toEqual([])
  })

  /**
   * By what each builder emits, not by its identifier: `avg` produces AVERAGE
   * and `if_` produces IF, so matching on the exported name would report both
   * as missing and hide the ones genuinely absent.
   */
  it('names a builder for every whitelisted function', () => {
    const built = new Set<string>()
    for (const value of Object.values(expr)) {
      if (typeof value !== 'function') continue
      try {
        const result = (value as (...args: unknown[]) => unknown)(1, 1, 1, 1)
        if (isExpr(result) && result.k === 'fn') built.add(result.name)
      } catch {
        // a builder that rejects placeholder arguments still counts as absent
      }
    }
    // Reached through lookup() or as a dotted form of a bare name.
    const byOtherRoute = new Set(['RANK.EQ', 'RANK.AVG'])
    expect(FUNCTIONS.filter((n) => !built.has(n) && !byOtherRoute.has(n))).toEqual([])
  })

  it('whitelists nothing whose value changes between our engine and Excel', () => {
    const nondeterministic = ['RAND', 'RANDBETWEEN', 'RANDARRAY']
    expect(FUNCTIONS.filter((n) => nondeterministic.includes(n))).toEqual([])
  })
})

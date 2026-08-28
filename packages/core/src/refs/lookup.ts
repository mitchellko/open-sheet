import { type Expr, type ExprInput, lift } from '../formula/expr.js'
import { ref } from './ref.js'

export interface LookupSpec {
  /** The value to find — usually a cell on the current row. */
  value: ExprInput
  /** The block to search. */
  from: string
  /** Column in `from` to match against. */
  match: string
  /** Column in `from` whose value to return. */
  get: string
  /**
   * What to show when nothing matches. Without it the cell reads `#N/A`, which
   * is Excel's answer and sometimes the right one — a missing match in a
   * reconciliation should be loud.
   */
  ifMissing?: ExprInput
}

/**
 * A lookup that names its columns instead of counting them.
 *
 * Compiles to `INDEX(…, MATCH(…, 0))` rather than `VLOOKUP`, deliberately.
 * `VLOOKUP(A2, products, 3, FALSE)` carries a positional column index, and
 * inserting a column in the lookup table silently makes the 3 point somewhere
 * else — the exact failure this framework exists to remove. It also forces the
 * matched column to be leftmost, which the author did not choose.
 *
 * ```tsx
 * formula: (r) => lookup({
 *   value: r.cell('sku'),
 *   from: 'products',
 *   match: 'sku',
 *   get: 'price',
 *   ifMissing: 0,
 * })
 * ```
 *
 * Both ranges resolve after layout, so adding a product moves them together.
 */
export function lookup(spec: LookupSpec): Expr {
  const block = ref(spec.from)
  const found: Expr = {
    k: 'fn',
    name: 'INDEX',
    args: [
      { k: 'ref', target: block.column(spec.get) },
      {
        k: 'fn',
        name: 'MATCH',
        args: [
          lift(spec.value),
          { k: 'ref', target: block.column(spec.match) },
          { k: 'lit', v: 0 },
        ],
      },
    ],
  }

  if (spec.ifMissing === undefined) return found
  return { k: 'fn', name: 'IFNA', args: [found, lift(spec.ifMissing)] }
}

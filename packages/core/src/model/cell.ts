import type { Expr } from '../formula/expr.js'
import type { Size } from './geometry.js'
import type { Validation } from './validation.js'

export type CellValue = number | string | boolean | null

export interface SourceLoc {
  file: string
  line: number
  column: number
}

export interface Cell {
  value?: CellValue
  expr?: Expr
  style?: string
  format?: string
  span?: Size
  /** Overrides the style's wrap setting for this cell. */
  wrap?: boolean
  /** This cell's formula fills a rectangle of this size, itself included. */
  spill?: Size
  /** This cell is inside another's spill footprint; that cell's key. */
  spillFrom?: CellKey
  /** What the recipient is allowed to type here. */
  validate?: Validation
  /** Only meaningful on a protected sheet, where everything else is locked. */
  unlocked?: boolean
  /** Where this number came from, shown on hover. Not a <Note> — that is a cell of its own. */
  note?: string
  loc?: SourceLoc
}

export type CellKey = string

export function cellKey(r: number, c: number): CellKey {
  return `${r},${c}`
}

export function parseCellKey(key: CellKey): { r: number; c: number } {
  const comma = key.indexOf(',')
  return { r: Number(key.slice(0, comma)), c: Number(key.slice(comma + 1)) }
}

/**
 * What a renderer should show for this cell: the computed value when the cell
 * carries a formula, and the literal otherwise.
 *
 * A cell inside a `<Spill>` footprint carries neither — the origin's formula
 * fills it, and the result lives only in the value map. Six call sites each
 * decided this for themselves and every one of them got it wrong the same way,
 * showing the top-left corner of a spill and blank everywhere else.
 */
export function readsFromValues(cell: Cell | undefined): boolean {
  return cell?.expr !== undefined || cell?.spillFrom !== undefined
}

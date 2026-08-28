import type { CellValue } from '../model/cell.js'

export interface NotEvaluated {
  readonly kind: 'not-evaluated'
}

export interface ExcelError {
  readonly kind: 'error'
  readonly code: string
}

export type Computed = CellValue | NotEvaluated | ExcelError

/**
 * The honest answer for a formula we did not compute — a `raw()` escape hatch,
 * or a dependency that was itself not evaluated. It exports fine; the viewer
 * shows it greyed. We never substitute a plausible number for one we don't have.
 */
export const NOT_EVALUATED: NotEvaluated = Object.freeze({ kind: 'not-evaluated' })

export const DIV0: ExcelError = Object.freeze({ kind: 'error', code: '#DIV/0!' })
export const VALUE: ExcelError = Object.freeze({ kind: 'error', code: '#VALUE!' })
export const REF: ExcelError = Object.freeze({ kind: 'error', code: '#REF!' })
export const NAME: ExcelError = Object.freeze({ kind: 'error', code: '#NAME?' })
export const NUM: ExcelError = Object.freeze({ kind: 'error', code: '#NUM!' })
export const NA: ExcelError = Object.freeze({ kind: 'error', code: '#N/A' })

export function isNotEvaluated(value: unknown): value is NotEvaluated {
  return (
    typeof value === 'object' && value !== null && (value as NotEvaluated).kind === 'not-evaluated'
  )
}

export function isExcelError(value: unknown): value is ExcelError {
  return typeof value === 'object' && value !== null && (value as ExcelError).kind === 'error'
}

export function errorFrom(code: string): ExcelError {
  return Object.freeze({ kind: 'error', code })
}

export function display(value: Computed): string {
  if (value === null || value === undefined) return ''
  if (isNotEvaluated(value)) return '#NOT_EVALUATED'
  if (isExcelError(value)) return value.code
  return String(value)
}

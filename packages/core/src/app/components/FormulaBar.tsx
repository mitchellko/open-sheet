/** @jsxImportSource react */
import type { CompiledWorkbook } from '../../compile/emit.js'
import type { ValueMap } from '../../formula/evaluate.js'
import { toFormula } from '../../formula/serialize.js'
import { display, isNotEvaluated } from '../../formula/value.js'
import { toA1 } from '../../model/a1.js'
import { readsFromValues } from '../../model/cell.js'
import type { Selection } from './Grid.js'

interface Props {
  book: CompiledWorkbook
  values: ValueMap
  sheetIndex: number
  selection: Selection
}

/**
 * Shows the *resolved* formula. The author wrote references, so this is the
 * first place they see what those became — and the only place to check that the
 * framework picked the range they meant.
 */
export function FormulaBar({ book, values, sheetIndex, selection }: Props) {
  const sheet = book.sheets[sheetIndex]
  const cell = sheet?.cells.get(`${selection.r},${selection.c}`)

  let formula = ''
  let error: string | undefined
  if (cell?.expr && sheet) {
    try {
      formula = toFormula(cell.expr, {
        registry: book.registry,
        definedNames: book.definedNames,
        sheet: sheet.name,
      })
    } catch (caught) {
      error = caught instanceof Error ? caught.message : String(caught)
    }
  } else if (cell?.value !== null && cell?.value !== undefined) {
    formula = String(cell.value)
  }

  const computed =
    sheet && readsFromValues(cell)
      ? values.get(`${sheet.name}!${selection.r},${selection.c}`)
      : undefined

  return (
    <div className="os-formula-bar">
      <span className="os-addr">{toA1(selection)}</span>
      <span className="os-fx">fx</span>
      {error ? (
        <span className="os-formula os-formula-error">{error}</span>
      ) : (
        <span className="os-formula">{formula}</span>
      )}
      {computed !== undefined ? (
        <span className={`os-computed${isNotEvaluated(computed) ? ' is-skipped' : ''}`}>
          {display(computed)}
        </span>
      ) : null}
    </div>
  )
}

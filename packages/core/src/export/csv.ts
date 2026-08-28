import type { CompiledSheet } from '../compile/emit.js'
import { type Computed, display, isExcelError, isNotEvaluated } from '../formula/value.js'
import { parseCellKey, readsFromValues } from '../model/cell.js'

export interface CsvOptions {
  delimiter?: string
  bom?: boolean
  /** Cells we could not compute export empty; their addresses are reported instead. */
  onSkipped?: (addresses: string[]) => void
}

export function toCsv(
  sheet: CompiledSheet,
  values: Map<string, Computed>,
  options: CsvOptions = {},
): string {
  const delimiter = options.delimiter ?? ','
  const grid: string[][] = []
  const skipped: string[] = []

  for (const [key, cell] of sheet.cells) {
    const { r, c } = parseCellKey(key)
    const computed = readsFromValues(cell)
      ? values.get(`${sheet.name}!${key}`)
      : (cell.value ?? null)

    let text: string
    if (isNotEvaluated(computed)) {
      skipped.push(`${sheet.name}!r${r}c${c}`)
      text = ''
    } else if (isExcelError(computed)) {
      text = computed.code
    } else {
      text = display(computed ?? null)
    }

    grid[r] ??= []
    ;(grid[r] as string[])[c] = text
  }

  if (skipped.length) options.onSkipped?.(skipped)

  const width = sheet.bounds.cols
  const lines: string[] = []
  for (let r = 0; r < sheet.bounds.rows; r += 1) {
    const row = grid[r] ?? []
    const fields: string[] = []
    for (let c = 0; c < width; c += 1) fields.push(quote(row[c] ?? '', delimiter))
    lines.push(fields.join(delimiter))
  }

  return (options.bom ? '﻿' : '') + lines.join('\r\n')
}

function quote(field: string, delimiter: string): string {
  if (!field.includes(delimiter) && !field.includes('"') && !/[\r\n]/.test(field)) return field
  return `"${field.replace(/"/g, '""')}"`
}

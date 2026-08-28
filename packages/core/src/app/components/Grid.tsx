/** @jsxImportSource react */
import { useCallback, useMemo, useRef, useState } from 'react'
import type { CompiledSheet } from '../../compile/emit.js'
import type { ValueMap } from '../../formula/evaluate.js'
import { isExcelError, isNotEvaluated } from '../../formula/value.js'
import { columnName } from '../../model/a1.js'
import { readsFromValues } from '../../model/cell.js'
import { formatValue, toStyleObject } from '../../style/css.js'
import { DEFAULT_THEME, resolveStyle } from '../../style/theme.js'
import { mergeStyle } from '../../style/types.js'

const ROW_HEIGHT = 26
const HEADER_HEIGHT = 26
const GUTTER_WIDTH = 52
const OVERSCAN = 6
const MIN_COLS = 12
const MIN_ROWS = 40

export interface Selection {
  r: number
  c: number
}

interface GridProps {
  sheet: CompiledSheet
  values: ValueMap
  selection: Selection
  onSelect: (selection: Selection) => void
}

function widthOf(sheet: CompiledSheet, c: number): number {
  return Math.round((sheet.columnWidths.get(c) ?? DEFAULT_THEME.defaultColumnWidth) * 8)
}

export function Grid({ sheet, values, selection, onSelect }: GridProps) {
  const viewport = useRef<HTMLDivElement>(null)
  const [scroll, setScroll] = useState({ top: 0, left: 0 })
  const [size, setSize] = useState({ width: 1200, height: 600 })

  const cols = Math.max(sheet.bounds.cols + 2, MIN_COLS)
  const rows = Math.max(sheet.bounds.rows + 2, MIN_ROWS)

  const offsets = useMemo(() => {
    const out = [0]
    for (let c = 0; c < cols; c += 1) out.push((out[c] as number) + widthOf(sheet, c))
    return out
  }, [sheet, cols])

  const totalWidth = offsets[cols] as number
  const totalHeight = rows * ROW_HEIGHT

  // A zero measurement is not information — it means the container has not been
  // laid out yet (or is hidden). Accepting it would window the grid down to
  // nothing and render an empty sheet.
  const measure = useCallback((node: HTMLDivElement | null) => {
    if (!node) return
    viewport.current = node
    const apply = () => {
      const width = node.clientWidth
      const height = node.clientHeight
      if (width > 0 && height > 0) setSize({ width, height })
    }
    apply()
    if (typeof ResizeObserver === 'undefined') return
    const observer = new ResizeObserver(apply)
    observer.observe(node)
    return () => observer.disconnect()
  }, [])

  const firstRow = Math.max(0, Math.floor(scroll.top / ROW_HEIGHT) - OVERSCAN)
  const lastRow = Math.min(rows, Math.ceil((scroll.top + size.height) / ROW_HEIGHT) + OVERSCAN)

  let firstCol = 0
  while (firstCol < cols && (offsets[firstCol + 1] as number) < scroll.left) firstCol += 1
  firstCol = Math.max(0, firstCol - 1)
  let lastCol = firstCol
  while (lastCol < cols && (offsets[lastCol] as number) < scroll.left + size.width) lastCol += 1
  lastCol = Math.min(cols, lastCol + 1)

  const visibleRows: number[] = []
  for (let r = firstRow; r < lastRow; r += 1) visibleRows.push(r)
  const visibleCols: number[] = []
  for (let c = firstCol; c < lastCol; c += 1) visibleCols.push(c)

  // Keep the selection mounted even when scrolled out of the window, so the
  // outline and the formula bar never disagree about what is selected.
  if (selection.r < firstRow || selection.r >= lastRow) visibleRows.push(selection.r)
  if (selection.c < firstCol || selection.c >= lastCol) visibleCols.push(selection.c)

  const onKeyDown = (event: React.KeyboardEvent) => {
    const moves: Record<string, [number, number]> = {
      ArrowUp: [-1, 0],
      ArrowDown: [1, 0],
      ArrowLeft: [0, -1],
      ArrowRight: [0, 1],
      Enter: [1, 0],
      Tab: [0, 1],
    }
    const move = moves[event.key]
    if (!move) return
    event.preventDefault()
    const next = {
      r: Math.max(0, Math.min(rows - 1, selection.r + (move[0] as number))),
      c: Math.max(0, Math.min(cols - 1, selection.c + (move[1] as number))),
    }
    onSelect(next)
    const node = viewport.current
    if (!node) return
    const top = next.r * ROW_HEIGHT
    if (top < node.scrollTop) node.scrollTop = top
    if (top + ROW_HEIGHT > node.scrollTop + node.clientHeight) {
      node.scrollTop = top + ROW_HEIGHT - node.clientHeight
    }
  }

  const frozen = sheet.freeze

  return (
    <div className="os-grid-wrap">
      <div className="os-corner" style={{ width: GUTTER_WIDTH, height: HEADER_HEIGHT }} />
      <div className="os-colheads" style={{ left: GUTTER_WIDTH, height: HEADER_HEIGHT }}>
        <div style={{ width: totalWidth, transform: `translateX(${-scroll.left}px)` }}>
          {visibleCols.map((c) => (
            <div
              key={c}
              className={`os-head${c === selection.c ? ' is-active' : ''}`}
              style={{ left: offsets[c], width: widthOf(sheet, c), height: HEADER_HEIGHT }}
            >
              {columnName(c)}
            </div>
          ))}
        </div>
      </div>
      <div className="os-rowheads" style={{ top: HEADER_HEIGHT, width: GUTTER_WIDTH }}>
        <div style={{ height: totalHeight, transform: `translateY(${-scroll.top}px)` }}>
          {visibleRows.map((r) => (
            <div
              key={r}
              className={`os-head${r === selection.r ? ' is-active' : ''}`}
              style={{ top: r * ROW_HEIGHT, height: ROW_HEIGHT, width: GUTTER_WIDTH }}
            >
              {r + 1}
            </div>
          ))}
        </div>
      </div>

      {/* biome-ignore lint/a11y/useSemanticElements: a <table> cannot be
          virtualized with absolute positioning; the grid roles carry the
          semantics instead. */}
      <div
        ref={measure}
        className="os-viewport"
        style={{ top: HEADER_HEIGHT, left: GUTTER_WIDTH }}
        role="grid"
        aria-rowcount={rows}
        aria-colcount={cols}
        aria-label={`${sheet.name} grid`}
        tabIndex={0}
        onKeyDown={onKeyDown}
        onScroll={(event) =>
          setScroll({
            top: (event.target as HTMLDivElement).scrollTop,
            left: (event.target as HTMLDivElement).scrollLeft,
          })
        }
      >
        <div className="os-canvas" style={{ width: totalWidth, height: totalHeight }}>
          {frozen && frozen.r > 0 ? (
            <div
              className="os-freeze-row"
              style={{ top: frozen.r * ROW_HEIGHT, width: totalWidth }}
            />
          ) : null}
          {frozen && frozen.c > 0 ? (
            <div
              className="os-freeze-col"
              style={{ left: offsets[frozen.c], height: totalHeight }}
            />
          ) : null}
          {visibleRows.map((r) =>
            visibleCols.map((c) => (
              <GridCell
                key={`${r},${c}`}
                sheet={sheet}
                values={values}
                r={r}
                c={c}
                left={offsets[c] as number}
                width={widthOf(sheet, c)}
                selected={selection.r === r && selection.c === c}
                onSelect={onSelect}
              />
            )),
          )}
        </div>
      </div>
    </div>
  )
}

interface CellProps {
  sheet: CompiledSheet
  values: ValueMap
  r: number
  c: number
  left: number
  width: number
  selected: boolean
  onSelect: (selection: Selection) => void
}

function GridCell({ sheet, values, r, c, left, width, selected, onSelect }: CellProps) {
  const cell = sheet.cells.get(`${r},${c}`)
  const computed = readsFromValues(cell)
    ? values.get(`${sheet.name}!${r},${c}`)
    : (cell?.value ?? null)

  let text = ''
  let tone = ''
  if (isNotEvaluated(computed)) {
    text = '#NOT_EVALUATED'
    tone = ' is-skipped'
  } else if (isExcelError(computed)) {
    text = computed.code
    tone = ' is-error'
  } else if (computed !== null && computed !== undefined) {
    text = formatValue(computed, cell?.format)
    if (typeof computed === 'number') tone = ' is-number'
  }

  const style = cell
    ? toStyleObject(
        mergeStyle(resolveStyle(DEFAULT_THEME, undefined), resolveStyle(DEFAULT_THEME, cell.style)),
      )
    : {}

  const span = cell?.span
  const boxWidth = span && span.cols > 1 ? spanWidth(sheet, c, span.cols) : width

  return (
    // biome-ignore lint/a11y/useSemanticElements: see the grid container above
    <div
      role="gridcell"
      // Roving tabindex: only the selected cell is in the tab order, and arrow
      // keys move the selection — the standard grid keyboard pattern.
      tabIndex={selected ? 0 : -1}
      aria-selected={selected}
      aria-rowindex={r + 1}
      aria-colindex={c + 1}
      className={`os-cell${tone}${selected ? ' is-selected' : ''}`}
      style={{
        ...(style as React.CSSProperties),
        left,
        top: r * ROW_HEIGHT,
        width: boxWidth,
        height: (span?.rows ?? 1) * ROW_HEIGHT,
      }}
      onMouseDown={() => onSelect({ r, c })}
    >
      <span className="os-cell-text">{text}</span>
    </div>
  )
}

function spanWidth(sheet: CompiledSheet, c: number, cols: number): number {
  let total = 0
  for (let i = 0; i < cols; i += 1) total += widthOf(sheet, c + i)
  return total
}

export { GUTTER_WIDTH, HEADER_HEIGHT, ROW_HEIGHT }

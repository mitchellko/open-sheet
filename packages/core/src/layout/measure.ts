import type { Block, TableNode } from '../compile/nodes.js'
import type { Size } from '../model/geometry.js'

const EMPTY: Size = { rows: 0, cols: 0 }

export function tableRowCount(table: TableNode): number {
  const title = table.title ? 1 : 0
  const header = table.showHeader ? 1 : 0
  const total = table.total ? 1 : 0
  return title + header + table.data.length + total
}

export function measure(block: Block): Size {
  switch (block.kind) {
    case 'cell':
      return block.span ?? { rows: 1, cols: 1 }
    case 'note':
      return { rows: 1, cols: block.cols }
    case 'spacer':
      return { rows: block.rows, cols: block.cols }
    case 'spill':
      return { rows: block.rows, cols: block.cols }
    case 'chart':
      return { rows: block.rows, cols: block.cols }
    case 'kpiBand':
      return { rows: 2, cols: block.items.length }
    case 'table':
      return { rows: tableRowCount(block), cols: block.columns.length }
    case 'stack':
      return stackSize(block.children, block.gap)
    case 'row':
      return rowSize(block.children, block.gap)
  }
}

function nonEmpty(children: readonly Block[]): Size[] {
  return children.map(measure).filter((size) => size.rows > 0 && size.cols > 0)
}

function stackSize(children: readonly Block[], gap: number): Size {
  const sizes = nonEmpty(children)
  if (sizes.length === 0) return EMPTY
  let rows = 0
  let cols = 0
  for (const size of sizes) {
    rows += size.rows
    if (size.cols > cols) cols = size.cols
  }
  return { rows: rows + gap * (sizes.length - 1), cols }
}

function rowSize(children: readonly Block[], gap: number): Size {
  const sizes = nonEmpty(children)
  if (sizes.length === 0) return EMPTY
  let rows = 0
  let cols = 0
  for (const size of sizes) {
    cols += size.cols
    if (size.rows > rows) rows = size.rows
  }
  return { rows, cols: cols + gap * (sizes.length - 1) }
}

export function measureAll(children: readonly Block[], gap: number, axis: 'stack' | 'row'): Size {
  return axis === 'stack' ? stackSize(children, gap) : rowSize(children, gap)
}

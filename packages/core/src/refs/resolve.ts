import type { DefinedName } from '../compile/emit.js'
import {
  type Registry,
  requireAnchor,
  requireColumn,
  type TableAnchor,
} from '../compile/registry.js'
import { qualify, rangeToA1, toA1 } from '../model/a1.js'
import type { Rect } from '../model/geometry.js'
import type { CellRef, NameRef, RangeRef, Ref } from './ref.js'

export interface ResolveContext {
  registry: Registry
  definedNames: Map<string, DefinedName>
  sheet: string
  /**
   * The row this formula is being written into. Only used inside an Excel
   * Table, where a reference to the same row is written `[@Amount]` — which is
   * both the idiomatic form and, being row-independent, the form that lets one
   * stored formula fill every appended row.
   */
  row?: number
}

export interface ResolvedRef {
  sheet: string
  rect: Rect
  name?: string
}

export function resolveRef(ref: Ref, context: ResolveContext): ResolvedRef {
  switch (ref.kind) {
    case 'cell':
      return resolveCell(ref, context)
    case 'range':
      return resolveRange(ref, context)
    case 'name':
      return resolveName(ref, context)
  }
}

function asTable(anchor: ReturnType<typeof requireAnchor>, ref: Ref): TableAnchor {
  if (anchor.kind !== 'table') {
    throw new Error(
      `"${anchor.name}" is a key-value block; use ref('${anchor.name}').get(key) instead of ` +
        `${ref.kind === 'range' ? 'column()' : 'cell()/total()'}`,
    )
  }
  return anchor
}

function resolveCell(ref: CellRef, context: ResolveContext): ResolvedRef {
  const anchor = asTable(requireAnchor(context.registry, ref.block), ref)
  const c = requireColumn(anchor, ref.column)

  if (ref.part === 'header') {
    if (anchor.headerRow === undefined) {
      throw new Error(`table "${anchor.name}" has no header row to reference`)
    }
    return { sheet: anchor.sheet, rect: { r: anchor.headerRow, c, rows: 1, cols: 1 } }
  }

  if (ref.part === 'total') {
    if (anchor.totalRow === undefined) {
      throw new Error(
        `table "${anchor.name}" has no total row; add a total={{ ${ref.column}: 'sum' }} prop ` +
          'before referencing ref().total()',
      )
    }
    return { sheet: anchor.sheet, rect: { r: anchor.totalRow, c, rows: 1, cols: 1 } }
  }

  const row = ref.row ?? 0
  if (row < 0) {
    throw new Error(
      `row ${row} is before the first data row of table "${anchor.name}". ` +
        'A formula using r.prev() must guard with r.isFirst.',
    )
  }
  if (row >= anchor.rowCount) {
    throw new Error(
      `row ${row} is past the last data row of table "${anchor.name}" (${anchor.rowCount} rows). ` +
        'A formula using r.next() must guard with r.isLast.',
    )
  }
  return { sheet: anchor.sheet, rect: { r: anchor.firstDataRow + row, c, rows: 1, cols: 1 } }
}

function resolveRange(ref: RangeRef, context: ResolveContext): ResolvedRef {
  const anchor = asTable(requireAnchor(context.registry, ref.block), ref)
  if (anchor.rowCount === 0) {
    throw new Error(`table "${anchor.name}" has no data rows, so it has no range to reference`)
  }
  if (ref.part === 'body') {
    return {
      sheet: anchor.sheet,
      rect: {
        r: anchor.firstDataRow,
        c: anchor.rect.c,
        rows: anchor.rowCount,
        cols: anchor.rect.cols,
      },
    }
  }
  if (!ref.column) throw new Error(`ref('${ref.block}').column() requires a column key`)
  const c = requireColumn(anchor, ref.column)
  return {
    sheet: anchor.sheet,
    rect: { r: anchor.firstDataRow, c, rows: anchor.rowCount, cols: 1 },
  }
}

function resolveName(ref: NameRef, context: ResolveContext): ResolvedRef {
  const anchor = requireAnchor(context.registry, ref.block)
  if (anchor.kind !== 'keyValue') {
    throw new Error(`"${ref.block}" is a table; ref().get() only works on kind="keyValue" blocks`)
  }
  const addr = anchor.keys.get(ref.key)
  if (!addr) {
    const known = [...anchor.keys.keys()]
    throw new Error(`no key "${ref.key}" in "${ref.block}" (keys: ${known.join(', ')})`)
  }
  const defined = context.definedNames.get(ref.key)
  const resolved: ResolvedRef = {
    sheet: anchor.sheet,
    rect: { r: addr.r, c: addr.c, rows: 1, cols: 1 },
  }
  // The column was missing from this check, so a name that had been overwritten
  // by a block in another column still serialized as the bare name — pointing
  // Excel at a different cell than the one evaluated. Compile-time collision
  // detection makes that unreachable; this stays as the second lock.
  if (
    defined &&
    defined.sheet === anchor.sheet &&
    defined.addr.r === addr.r &&
    defined.addr.c === addr.c
  ) {
    resolved.name = ref.key
  }
  return resolved
}

const ABSOLUTE = { absoluteRow: true, absoluteCol: true } as const

/**
 * `costs[Amount]` rather than `B2:B13`. The point is what happens on the
 * recipient's screen: Excel takes a row typed below the table into the table,
 * and every structured reference follows. A plain range would still end where
 * it did, and the total would silently stop including the new row.
 *
 * Only the whole-column range gets this. A single cell inside a row stays A1,
 * because the structured form for it (`[@Amount]`) is relative to the row the
 * formula sits in, and ours are written per cell at absolute addresses.
 */
function structuredRef(ref: Ref, context: ResolveContext): string | undefined {
  if (ref.kind !== 'range' || ref.part !== 'column' || ref.column === undefined) return undefined
  const anchor = context.registry.get(ref.block)
  if (anchor?.kind !== 'table' || !anchor.table) return undefined
  const header = anchor.table.headers.get(ref.column)
  if (header === undefined) return undefined
  // A header holding [ ] # ' or @ has to be escaped with a single quote, or the
  // reference parses as something else entirely.
  const escaped = header.replace(/([[\]#'@])/g, "'$1")
  return `${anchor.name}[${escaped}]`
}

/**
 * `costs[[#This Row],[Amount]]` — this row's cell of that column. Only valid for
 * a formula that itself sits in the table's data rows, which is why it needs the
 * writer to say which row it is serialising.
 *
 * **Not `[@Amount]`.** That is the shorthand Excel shows in the formula bar, not
 * a form the file may contain: written into `<f>` it is read back as `#REF!`,
 * and every derived column in the workbook is broken the moment it opens.
 * LibreOffice computes the long form correctly too, so nothing is traded away.
 */
function sameRowRef(ref: Ref, context: ResolveContext): string | undefined {
  if (ref.kind !== 'cell' || ref.absolute === true) return undefined
  if (context.row === undefined) return undefined
  const anchor = context.registry.get(ref.block)
  if (anchor?.kind !== 'table' || !anchor.table || anchor.sheet !== context.sheet) return undefined
  if (context.row < anchor.firstDataRow || context.row > anchor.lastDataRow) return undefined

  const resolved = resolveRef(ref, context)
  // A reference to another row — `r.prev()` — stays A1, and Excel adjusts it
  // relatively when it fills an appended row, which is the behaviour we want.
  if (resolved.rect.r !== context.row) return undefined

  const header = anchor.table.headers.get(ref.column)
  if (header === undefined) return undefined
  return escapeHeader(anchor.name, header)
}

function escapeHeader(name: string, header: string): string {
  return `${name}[[#This Row],[${header.replace(/([[\]#'@])/g, "'$1")}]]`
}

export function refToA1(ref: Ref, context: ResolveContext): string {
  const sameRow = sameRowRef(ref, context)
  if (sameRow) return sameRow

  const structured = structuredRef(ref, context)
  if (structured) return structured

  const resolved = resolveRef(ref, context)
  if (resolved.name) return resolved.name

  const absolute = ref.kind !== 'name' && ref.absolute === true
  const options = absolute ? ABSOLUTE : {}
  const reference =
    resolved.rect.rows === 1 && resolved.rect.cols === 1
      ? toA1({ r: resolved.rect.r, c: resolved.rect.c }, options)
      : rangeToA1(resolved.rect, options)

  return qualify(resolved.sheet === context.sheet ? undefined : resolved.sheet, reference)
}

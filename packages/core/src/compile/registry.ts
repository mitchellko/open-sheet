import type { Addr, Rect } from '../model/geometry.js'

export interface TableAnchor {
  kind: 'table'
  name: string
  sheet: string
  rect: Rect
  titleRow?: number
  headerRow?: number
  firstDataRow: number
  lastDataRow: number
  rowCount: number
  totalRow?: number
  columns: Map<string, number>
  /**
   * Set when the table is emitted as an Excel Table. Holds each column's header
   * text, which is how a structured reference names it — `costs[Amount]`, not
   * `costs[amount]`.
   */
  table?: {
    headers: Map<string, string>
    /** Which aggregate each totalled column uses, so the table's own totals row agrees with ours. */
    totals: Map<string, string>
    /**
     * Headers of derived columns that cannot fill an appended row, because
     * their formula reads another row and no single stored formula serves
     * every row. Recorded here so the CLI can say so instead of the file
     * quietly doing less than the documentation promises.
     */
    noFillDown: string[]
  }
}

export interface KeyValueAnchor {
  kind: 'keyValue'
  name: string
  sheet: string
  rect: Rect
  keys: Map<string, Addr>
}

export type Anchor = TableAnchor | KeyValueAnchor

export type Registry = Map<string, Anchor>

export function requireAnchor(registry: Registry, name: string): Anchor {
  const anchor = registry.get(name)
  if (anchor) return anchor
  const known = [...registry.keys()]
  const suggestion = nearest(name, known)
  throw new Error(
    `no block named "${name}"` +
      (suggestion ? `; did you mean "${suggestion}"?` : '') +
      (known.length ? ` (known: ${known.join(', ')})` : ''),
  )
}

export function requireColumn(anchor: TableAnchor, column: string): number {
  const index = anchor.columns.get(column)
  if (index !== undefined) return index
  const known = [...anchor.columns.keys()]
  const suggestion = nearest(column, known)
  throw new Error(
    `no column "${column}" in table "${anchor.name}"` +
      (suggestion ? `; did you mean "${suggestion}"?` : '') +
      ` (columns: ${known.join(', ')}).\n` +
      `If "${column}" is in the data but not shown, either add col('${column}') so the ` +
      `formula can point at a cell, or read the raw value with r.data.${column} — the ` +
      'raw value is baked into the exported formula as a literal, so the recipient cannot change it.',
  )
}

function nearest(target: string, candidates: readonly string[]): string | undefined {
  let best: string | undefined
  let bestDistance = Number.POSITIVE_INFINITY
  for (const candidate of candidates) {
    const distance = editDistance(target.toLowerCase(), candidate.toLowerCase())
    if (distance < bestDistance) {
      bestDistance = distance
      best = candidate
    }
  }
  return bestDistance <= Math.max(2, Math.floor(target.length / 3)) ? best : undefined
}

function editDistance(a: string, b: string): number {
  const prev = new Array<number>(b.length + 1)
  const curr = new Array<number>(b.length + 1)
  for (let j = 0; j <= b.length; j += 1) prev[j] = j
  for (let i = 1; i <= a.length; i += 1) {
    curr[0] = i
    for (let j = 1; j <= b.length; j += 1) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1
      curr[j] = Math.min(
        (curr[j - 1] as number) + 1,
        (prev[j] as number) + 1,
        (prev[j - 1] as number) + cost,
      )
    }
    for (let j = 0; j <= b.length; j += 1) prev[j] = curr[j] as number
  }
  return prev[b.length] as number
}

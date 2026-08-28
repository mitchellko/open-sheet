type Cell = unknown
type Grid = Cell[][]

/**
 * The function library's `SORT` compares as strings — `960000` sorts above
 * `95000` above `81000` above `210000`, which is right only while every value
 * has the same number of digits. Its `UNIQUE` does not deduplicate at all, and
 * it has no `SORTBY`. A whitelisted function that returns a plausible wrong
 * answer is worse than an absent one, so these three are ours.
 */

/** Excel orders numbers before text before logicals, and text case-insensitively. */
function rank(value: Cell): number {
  if (typeof value === 'number') return 0
  if (typeof value === 'string') return 1
  if (typeof value === 'boolean') return 2
  return 3
}

export function compareCells(a: Cell, b: Cell): number {
  const ra = rank(a)
  const rb = rank(b)
  if (ra !== rb) return ra - rb
  if (typeof a === 'number' && typeof b === 'number') return a - b
  if (typeof a === 'string' && typeof b === 'string') {
    const la = a.toLowerCase()
    const lb = b.toLowerCase()
    return la < lb ? -1 : la > lb ? 1 : 0
  }
  if (typeof a === 'boolean' && typeof b === 'boolean') {
    return Number(a) - Number(b)
  }
  return 0
}

function asGrid(value: unknown): Grid {
  if (!Array.isArray(value)) return [[value]]
  return value.map((row) => (Array.isArray(row) ? row : [row]))
}

function order(value: unknown): number {
  return Number(value) === -1 ? -1 : 1
}

/** `SORT(array, [sortIndex], [sortOrder], [byCol])` — sortIndex is 1-based. */
export function SORT(
  array: unknown,
  sortIndex?: unknown,
  sortOrder?: unknown,
  byCol?: unknown,
): Grid {
  const grid = asGrid(array)
  const direction = order(sortOrder)
  const index = Math.max(1, Math.trunc(Number(sortIndex ?? 1) || 1)) - 1

  if (byCol === true) {
    const columns = grid[0]?.map((_, c) => grid.map((row) => row[c])) ?? []
    columns.sort((a, b) => direction * compareCells(a[index], b[index]))
    return columns[0]?.map((_, r) => columns.map((column) => column[r])) ?? []
  }
  return [...grid].sort((a, b) => direction * compareCells(a[index], b[index]))
}

/** `SORTBY(array, by1, [order1], by2, [order2], …)` — the keys are parallel ranges. */
export function SORTBY(array: unknown, ...rest: unknown[]): Grid {
  const grid = asGrid(array)
  const keys: { values: Cell[]; direction: number }[] = []
  for (let i = 0; i < rest.length; i += 2) {
    const flat = asGrid(rest[i]).map((row) => row[0])
    keys.push({ values: flat, direction: order(rest[i + 1]) })
  }
  const positions = grid.map((_, i) => i)
  positions.sort((a, b) => {
    for (const key of keys) {
      const result = key.direction * compareCells(key.values[a], key.values[b])
      if (result !== 0) return result
    }
    return a - b
  })
  return positions.map((i) => grid[i] as Cell[])
}

/** `UNIQUE(array, [byCol], [exactlyOnce])`. Rows compare by their whole contents. */
export function UNIQUE(array: unknown, byCol?: unknown, exactlyOnce?: unknown): Grid {
  const grid = asGrid(array)
  const rows = byCol === true ? (grid[0]?.map((_, c) => grid.map((row) => row[c])) ?? []) : grid

  const counts = new Map<string, number>()
  for (const row of rows) {
    const key = JSON.stringify(row)
    counts.set(key, (counts.get(key) ?? 0) + 1)
  }

  const seen = new Set<string>()
  const kept: Cell[][] = []
  for (const row of rows) {
    const key = JSON.stringify(row)
    if (seen.has(key)) continue
    seen.add(key)
    if (exactlyOnce === true && (counts.get(key) ?? 0) > 1) continue
    kept.push(row)
  }

  if (byCol !== true) return kept
  return kept[0]?.map((_, r) => kept.map((column) => column[r])) ?? []
}

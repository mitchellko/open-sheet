import * as formulajs from '@formulajs/formulajs'
import { formatValue } from '../style/css.js'
import { SORT, SORTBY, UNIQUE } from './arrays.js'
import { isWhitelisted } from './expr.js'

type Implementation = (...args: any[]) => unknown

const LIBRARY = formulajs as unknown as Record<string, Implementation>

/**
 * Dispatch is whitelist-only and resolved through this map — never by looking up
 * an arbitrary name on the library object at call time.
 */
/**
 * Excel has both `RANK` and `RANK.EQ`; the library exposes the family as a
 * namespace object, so `RANK` alone is not callable. Excel's own `RANK` is the
 * legacy alias of `RANK.EQ`, so that is what it resolves to.
 */
/**
 * The library's TEXT handles number codes and returns a date code's input
 * untouched, so `TEXT(EDATE(…),"yyyy-mm")` showed the reader a serial. We
 * already render a value under a format code — for every cell, in both
 * renderers — so TEXT is that same function rather than a second one that can
 * disagree with it.
 */
function TEXT(value: unknown, format: unknown): string {
  return formatValue(value, typeof format === 'string' ? format : undefined)
}

const OURS: Record<string, Implementation> = {
  SORT: SORT as Implementation,
  SORTBY: SORTBY as Implementation,
  TEXT: TEXT as Implementation,
  UNIQUE: UNIQUE as Implementation,
}

const ALIASES: Record<string, [string, string]> = {
  RANK: ['RANK', 'EQ'],
  'RANK.EQ': ['RANK', 'EQ'],
  'RANK.AVG': ['RANK', 'AVG'],
  VAR: ['VAR', 'S'],
  VARP: ['VAR', 'P'],
  'VAR.S': ['VAR', 'S'],
  'VAR.P': ['VAR', 'P'],
  STDEV: ['STDEV', 'S'],
  'STDEV.S': ['STDEV', 'S'],
  'STDEV.P': ['STDEV', 'P'],
  PERCENTILE: ['PERCENTILE', 'INC'],
  'PERCENTILE.INC': ['PERCENTILE', 'INC'],
  'PERCENTILE.EXC': ['PERCENTILE', 'EXC'],
  QUARTILE: ['QUARTILE', 'INC'],
  'QUARTILE.INC': ['QUARTILE', 'INC'],
  'QUARTILE.EXC': ['QUARTILE', 'EXC'],
  MODE: ['MODE', 'SNGL'],
  'MODE.SNGL': ['MODE', 'SNGL'],
  'MODE.MULT': ['MODE', 'MULT'],
}

export function lookup(name: string): Implementation | undefined {
  const upper = name.toUpperCase()
  if (!isWhitelisted(upper)) return undefined

  // Ours wins: the library's SORT compares as strings and its UNIQUE does not
  // deduplicate, and a whitelisted function that answers wrongly is worse than
  // one that is absent.
  const own = OURS[upper]
  if (own) return own

  const alias = ALIASES[upper]
  if (alias) {
    const [group, member] = alias
    const namespace = LIBRARY[group] as unknown as Record<string, Implementation> | undefined
    const implementation = namespace?.[member]
    return typeof implementation === 'function' ? implementation : undefined
  }

  const implementation = LIBRARY[upper]
  return typeof implementation === 'function' ? implementation : undefined
}

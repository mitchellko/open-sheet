import { numberFormat } from '../export/formats.js'
import type { BorderSide, CellStyle } from './types.js'

const WEIGHT_PX: Record<BorderSide['weight'], string> = {
  hair: '0.5px',
  thin: '1px',
  medium: '2px',
}

function side(border: BorderSide | undefined): string | undefined {
  if (!border) return undefined
  return `${WEIGHT_PX[border.weight]} solid ${border.color ?? 'currentColor'}`
}

const VERTICAL: Record<string, string> = { top: 'flex-start', middle: 'center', bottom: 'flex-end' }

export function toCssDeclarations(style: CellStyle): Record<string, string> {
  const out: Record<string, string> = {}

  if (style.font?.family) out['font-family'] = `${style.font.family}, system-ui, sans-serif`
  if (style.font?.size) out['font-size'] = `${style.font.size}px`
  if (style.font?.bold) out['font-weight'] = '700'
  if (style.font?.italic) out['font-style'] = 'italic'
  if (style.font?.color) out.color = style.font.color
  if (style.fill) out['background-color'] = style.fill

  if (style.align?.horizontal) out['text-align'] = style.align.horizontal
  if (style.align?.vertical) out['align-items'] = VERTICAL[style.align.vertical] as string
  if (style.align?.wrap) out['white-space'] = 'normal'
  if (style.align?.indent) out['padding-left'] = `${style.align.indent * 8}px`

  const top = side(style.border?.top)
  const bottom = side(style.border?.bottom)
  const left = side(style.border?.left)
  const right = side(style.border?.right)
  if (top) out['border-top'] = top
  if (bottom) out['border-bottom'] = bottom
  if (left) out['border-left'] = left
  if (right) out['border-right'] = right

  return out
}

export function toCssText(style: CellStyle): string {
  return Object.entries(toCssDeclarations(style))
    .map(([property, value]) => `${property}:${value}`)
    .join(';')
}

/**
 * Excel number formats drive both renderers. The HTML side cannot ask Excel to
 * format for it, so the common codes are interpreted here; anything else falls
 * back to the raw value rather than guessing at a format we do not understand.
 *
 * Sections matter more than they look. Excel reads `positive;negative;zero;text`
 * and the accounting format uses all of them — negatives in parentheses, zero as
 * a dash. Ignoring them made the viewer show `-84,500` where Excel showed
 * `(84,500)`: the same cell reading differently in the two places, which is the
 * one thing a "what you see is what exports" tool cannot afford.
 */
export function formatValue(value: unknown, format: string | undefined): string {
  if (value === null || value === undefined || value === '') return ''
  if (typeof value !== 'number') return String(value)

  const full = numberFormat(format)
  if (!full || full === 'General') return trimNumber(value)

  const { code, negated, literal } = section(full, value)
  if (literal !== undefined) return literal
  const magnitude = negated ? Math.abs(value) : value
  const rendered = renderSection(magnitude, code)
  return negated ? `(${rendered})` : rendered
}

interface Section {
  code: string
  /** The negative section is written for a positive number in parentheses. */
  negated: boolean
  /** A section that is nothing but literal text, e.g. the accounting zero dash. */
  literal?: string
}

function section(full: string, value: number): Section {
  const parts = splitSections(full)
  if (parts.length === 1) return { code: clean(parts[0] as string), negated: false }

  const chosen =
    value < 0 ? (parts[1] ?? parts[0]) : value === 0 ? (parts[2] ?? parts[0]) : parts[0]
  const code = clean(chosen as string)

  // A negative section written as `(#,##0)` already carries the sign visually.
  const parenthesised = /^\(.*\)$/.test(code)
  const bare = parenthesised ? code.slice(1, -1) : code

  if (!/[0#]/.test(bare)) {
    const text = bare.replace(/"/g, '').trim()
    return { code: bare, negated: false, literal: text }
  }
  return { code: bare, negated: value < 0 && parenthesised }
}

/** `;` inside quotes is literal text, not a section break. */
function splitSections(code: string): string[] {
  const out: string[] = []
  let current = ''
  let quoted = false
  for (const ch of code) {
    if (ch === '"') quoted = !quoted
    if (ch === ';' && !quoted) {
      out.push(current)
      current = ''
      continue
    }
    current += ch
  }
  out.push(current)
  return out
}

/** Strips Excel's alignment padding (`_(`, `* `) which has no HTML equivalent. */
function clean(code: string): string {
  return code.replace(/_./g, '').replace(/\*./g, '').trim()
}

function renderSection(value: number, code: string): string {
  // A date cell holds a serial; only the format says it is a date. Rendering the
  // number would show 46265 where Excel shows 2026-08-24.
  if (isDateCode(code)) return renderDate(value, code)
  if (code.includes('%')) {
    const decimals = decimalsIn(code)
    return `${(value * 100).toFixed(decimals)}%`
  }
  if (code.includes(',,')) return `${group((value / 1_000_000).toFixed(decimalsIn(code)))}M`
  if (code.includes(',"K"')) return `${group((value / 1_000).toFixed(decimalsIn(code)))}K`
  if (code.includes('#,##0')) return group(value.toFixed(decimalsIn(code)))
  if (code === '@') return String(value)

  return trimNumber(value)
}

/**
 * A date code is one built from date tokens with no numeric placeholders. The
 * earlier form required both a year and a day token, which excluded `ddd` and
 * `mmmm yyyy` — both perfectly ordinary date formats.
 */
function isDateCode(code: string): boolean {
  const body = code.replace(/"[^"]*"/g, '').replace(/AM\/PM|A\/P/gi, '')
  // A numeric placeholder, a percent or the text marker means it is not a date.
  if (/[#0?%@]/.test(body)) return false
  // Anything else is a literal, which is how Excel reads a format code — the
  // earlier character whitelist rejected `yyyy年m月` and every other code with a
  // word in it, and those fell through to number formatting as a bare serial.
  return /[ymdhs]/i.test(body)
}

const MONTHS_SHORT = [
  'Jan',
  'Feb',
  'Mar',
  'Apr',
  'May',
  'Jun',
  'Jul',
  'Aug',
  'Sep',
  'Oct',
  'Nov',
  'Dec',
]
const MONTHS_LONG = [
  'January',
  'February',
  'March',
  'April',
  'May',
  'June',
  'July',
  'August',
  'September',
  'October',
  'November',
  'December',
]
const DAYS_SHORT = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat']
const DAYS_LONG = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']

const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30)

function renderDate(serial: number, code: string): string {
  const date = new Date(EXCEL_EPOCH_UTC + Math.round(serial * 86_400_000))
  const pad = (n: number) => String(n).padStart(2, '0')

  // `m` means month or minute depending on whether it follows an hour token,
  // which is why this walks the code rather than running replacements.
  // An AM/PM marker anywhere in the code puts the hour on a 12-hour clock,
  // including the hour tokens that precede it.
  const meridiem = /AM\/PM|A\/P/i.test(code)
  const hour12 = (h: number) => h % 12 || 12

  let out = ''
  let i = 0
  let afterHour = false
  while (i < code.length) {
    const rest = code.slice(i)
    const marker = /^(AM\/PM|A\/P)/i.exec(rest)?.[0]
    if (marker) {
      const pm = date.getUTCHours() >= 12
      out += marker.length === 5 ? (pm ? 'PM' : 'AM') : pm ? 'P' : 'A'
      i += marker.length
      continue
    }
    const token = /^(yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s)/i.exec(rest)?.[0]
    if (!token) {
      if (rest[0] !== '"') out += rest[0]
      i += 1
      continue
    }
    const t = token.toLowerCase()
    if (t.startsWith('y'))
      out += t === 'yy' ? pad(date.getUTCFullYear() % 100) : date.getUTCFullYear()
    else if (t.startsWith('h')) {
      const h = meridiem ? hour12(date.getUTCHours()) : date.getUTCHours()
      out += t === 'hh' ? pad(h) : h
      afterHour = true
    } else if (t.startsWith('s'))
      out += t === 'ss' ? pad(date.getUTCSeconds()) : date.getUTCSeconds()
    else if (t.startsWith('d')) {
      out +=
        t === 'dddd'
          ? DAYS_LONG[date.getUTCDay()]
          : t === 'ddd'
            ? DAYS_SHORT[date.getUTCDay()]
            : t === 'dd'
              ? pad(date.getUTCDate())
              : date.getUTCDate()
      afterHour = false
    } else if (t.startsWith('m')) {
      if (afterHour) {
        out += t === 'mm' ? pad(date.getUTCMinutes()) : date.getUTCMinutes()
        afterHour = false
      } else {
        out +=
          t === 'mmmm'
            ? MONTHS_LONG[date.getUTCMonth()]
            : t === 'mmm'
              ? MONTHS_SHORT[date.getUTCMonth()]
              : t === 'mm'
                ? pad(date.getUTCMonth() + 1)
                : date.getUTCMonth() + 1
      }
    }
    i += token.length
  }
  return out
}

function decimalsIn(code: string): number {
  const match = /\.(0+)/.exec(code)
  return match ? (match[1] as string).length : code.includes('%') ? 1 : 0
}

function group(text: string): string {
  const [whole, fraction] = text.split('.')
  const sign = (whole as string).startsWith('-') ? '-' : ''
  const digits = (whole as string).replace('-', '').replace(/\B(?=(\d{3})+(?!\d))/g, ',')
  return sign + digits + (fraction ? `.${fraction}` : '')
}

function trimNumber(value: number): string {
  return Number.isInteger(value) ? String(value) : String(Number(value.toFixed(10)))
}

/**
 * React's inline `style` prop takes camelCase keys and silently drops kebab-case
 * ones, so the grid needs its own shape of the same declarations. Keeping this
 * beside the CSS adapter means the two cannot drift.
 */
export function toStyleObject(style: CellStyle): Record<string, string> {
  const out: Record<string, string> = {}
  for (const [property, value] of Object.entries(toCssDeclarations(style))) {
    out[property.replace(/-([a-z])/g, (_, ch: string) => ch.toUpperCase())] = value
  }
  return out
}

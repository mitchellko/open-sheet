import { describe, expect, it } from 'vitest'
import { compile } from '../compile/compile.js'
import { budget } from '../compile/fixtures.js'
import { toHtml } from '../export/html.js'
import { evaluateWorkbook } from '../formula/evaluate.js'
import { formatValue, toCssDeclarations, toStyleObject } from './css.js'
import { toArgb, toExcelStyle } from './excel.js'
import { DEFAULT_THEME, resolveStyle } from './theme.js'

describe('the style bridge', () => {
  it('renders one style into both targets consistently', () => {
    const header = resolveStyle(DEFAULT_THEME, 'tableHeader')
    expect(header).toBeDefined()

    const css = toCssDeclarations(header as never)
    const excel = toExcelStyle(header as never)

    expect(css['font-weight']).toBe('700')
    expect(excel.font?.bold).toBe(true)

    expect(css['background-color']).toBe('#0f172a')
    expect(excel.fill?.fgColor.argb).toBe('FF0F172A')

    expect(css.color).toBe('#ffffff')
    expect(excel.font?.color?.argb).toBe('FFFFFFFF')
  })

  it('converts colours to the ARGB Excel wants', () => {
    expect(toArgb('#1d4ed8')).toBe('FF1D4ED8')
    expect(toArgb('#abc')).toBe('FFAABBCC')
    expect(toArgb('FF00FF00')).toBe('FF00FF00')
  })

  it('falls back to the default theme when a theme omits a key', () => {
    const sparse = { ...DEFAULT_THEME, name: 'sparse', styles: { body: {} } }
    expect(resolveStyle(sparse, 'tableTotal')).toEqual(DEFAULT_THEME.styles.tableTotal)
  })
})

describe('number formats render the same way HTML and Excel do', () => {
  it('handles the named formats', () => {
    expect(formatValue(12_400_000, 'currency')).toBe('12,400,000')
    expect(formatValue(0.6029159, 'percent')).toBe('60.3%')
    expect(formatValue(0.6029159, 'percent2')).toBe('60.29%')
    expect(formatValue(12_400_000, 'millions')).toBe('12M')
    expect(formatValue(1234.5, 'decimal')).toBe('1,234.50')
    expect(formatValue(-1234, 'currency')).toBe('-1,234')
    expect(formatValue('text', 'currency')).toBe('text')
    expect(formatValue(null, 'currency')).toBe('')
  })

  it('honours the sections Excel reads, so the two renderers agree', () => {
    // `positive;negative;zero;text`. Ignoring these showed -84,500 in the viewer
    // where Excel showed (84,500) — the same cell reading differently in the two
    // places is the one thing a "what you see is what exports" tool cannot afford.
    expect(formatValue(-84_500, 'accounting')).toBe('(84,500)')
    expect(formatValue(12_400, 'accounting')).toBe('12,400')
    expect(formatValue(0, 'accounting')).toBe('-')

    expect(formatValue(-5, '#,##0;(#,##0);"nil"')).toBe('(5)')
    expect(formatValue(0, '#,##0;(#,##0);"nil"')).toBe('nil')
  })

  it('does not mistake a semicolon inside quotes for a section break', () => {
    expect(formatValue(5, '#,##0" a;b"')).toBe('5')
  })

  it('scales the way Excel does', () => {
    expect(formatValue(12_400_000_000, 'millions')).toBe('12,400M')
    expect(formatValue(12_400_000, 'thousands')).toBe('12,400K')
    // Excel rounds to the scale, so a small number really does read as 0M.
    expect(formatValue(-84_500, 'millions')).toBe('-0M')
  })
})

describe('html export', () => {
  const render = () => {
    const book = compile(budget())
    return toHtml(book, { title: 'Budget', values: evaluateWorkbook(book) })
  }

  it('writes a self-contained document', () => {
    const html = render()
    expect(html.startsWith('<!doctype html>')).toBe(true)
    expect(html).toContain('<style>')
    expect(html).not.toMatch(/<(script|link|img)\b/)
  })

  it('shows computed values, formatted', () => {
    const html = render()
    expect(html).toContain('12,400,000')
    expect(html).toContain('60.3%')
  })

  it('renders one section per sheet', () => {
    const html = render()
    expect(html).toContain('<h2>Assumptions</h2>')
    expect(html).toContain('<h2>P&amp;L</h2>')
  })

  it('escapes sheet names and text', () => {
    expect(render()).not.toContain('<h2>P&L</h2>')
  })

  it('merges spanned cells instead of repeating them', () => {
    const html = render()
    const notes = html.match(/Forecast beyond Q4/g) ?? []
    expect(notes).toHaveLength(1)
    expect(html).toMatch(/colspan="\d+"/)
  })

  it('keeps wide grids inside a scroll container', () => {
    expect(render()).toContain('class="os-scroll"')
  })

  it('carries print rules so the PDF matches', () => {
    const html = render()
    expect(html).toContain('@media print')
    expect(html).toContain('@page { size: A4 landscape')
    expect(html).toContain('display: table-header-group')
  })
})

describe('data bars land in both renderers', () => {
  it('records one rule per barred column, over the data range only', () => {
    const book = compile(budget())
    const pl = book.sheets[1]
    expect(pl?.conditionalFormats).toHaveLength(1)

    const anchor = book.registry.get('pl')
    if (anchor?.kind !== 'table') throw new Error('no pl table')
    const format = pl?.conditionalFormats[0]

    expect(format?.kind).toBe('dataBar')
    expect(format?.rect.r).toBe(anchor.firstDataRow)
    expect(format?.rect.rows).toBe(anchor.rowCount)
    expect(format?.rect.c).toBe(anchor.columns.get('revenue'))
  })

  it('draws the bar as a gradient in HTML, scaled to the range maximum', () => {
    const book = compile(budget())
    const html = toHtml(book, { values: evaluateWorkbook(book) })
    expect(html).toContain('linear-gradient(to right, #93c5fd 100%')
    expect(html).toMatch(/linear-gradient\(to right, #93c5fd 7[0-9]%/)
  })
})

describe('react needs its own shape of the same declarations', () => {
  it('camelCases every property', () => {
    const header = resolveStyle(DEFAULT_THEME, 'tableHeader')
    const object = toStyleObject(header as never)
    expect(object.fontWeight).toBe('700')
    expect(object.backgroundColor).toBe('#0f172a')
    expect(object.borderBottom).toBeDefined()
    expect(Object.keys(object).some((key) => key.includes('-'))).toBe(false)
  })

  it('keeps the same values as the CSS adapter', () => {
    const style = resolveStyle(DEFAULT_THEME, 'note') as never
    const css = toCssDeclarations(style)
    const object = toStyleObject(style)
    expect(Object.keys(object)).toHaveLength(Object.keys(css).length)
    expect(object.fontStyle).toBe(css['font-style'])
  })
})

describe('dates are serials with a format on top', () => {
  const serial = (y: number, m: number, d: number) =>
    (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86_400_000

  it('renders a date code as a date, not as the number underneath', () => {
    // A date cell holds 46258; only the format says it is 2026-08-24. Rendering
    // the number would show the serial where Excel shows the date.
    const d = serial(2026, 8, 24)
    expect(formatValue(d, 'date')).toBe('2026-08-24')
    expect(formatValue(d, 'dd/mm/yyyy')).toBe('24/08/2026')
    expect(formatValue(d, 'd mmm yyyy')).toBe('24 Aug 2026')
    expect(formatValue(d, 'mmmm yyyy')).toBe('August 2026')
    expect(formatValue(d, 'dddd')).toBe('Monday')
  })

  it('does not mistake a number format for a date one', () => {
    const d = serial(2026, 8, 24)
    expect(formatValue(d, 'currency')).toBe('46,258')
    expect(formatValue(d, 'number')).toBe('46,258')
    expect(formatValue(0.5, 'percent')).toBe('50.0%')
  })

  it('tells minutes from months by what precedes them', () => {
    const noon = serial(2026, 8, 24) + 0.5
    expect(formatValue(noon, 'yyyy-mm-dd hh:mm')).toBe('2026-08-24 12:00')
  })

  it('reads a code with a word in it as a date, since the words are literals', () => {
    // The earlier character whitelist rejected any code holding a letter that
    // was not a date token, so `yyyy年m月` fell through to number formatting and
    // showed the reader a bare serial. Excel treats unrecognised characters as
    // literals; so do we.
    const d = serial(2026, 8, 24)
    expect(formatValue(d, 'yyyy年m月')).toBe('2026年8月')
    expect(formatValue(d, 'yyyy年m月d日')).toBe('2026年8月24日')
    // A token with literals on *both* sides, which the walk has to carry
    // through rather than stopping at the first thing it does not recognise.
    expect(formatValue(d, '民國yy年')).toBe('民國26年')
  })

  it('still refuses a code that is numeric, percent or text', () => {
    const d = serial(2026, 8, 24)
    expect(formatValue(d, '#,##0.00')).toBe('46,258.00')
    expect(formatValue(0.6, '0.0%')).toBe('60.0%')
    expect(formatValue(1234, '@')).toBe('1234')
  })

  it('names months and days in English, where Excel follows the reader', () => {
    // A divergence we cannot close: `mmm` is "Aug" here and "8月" to a reader
    // whose Excel runs in Chinese. Numeric codes agree everywhere, which is why
    // the cross-engine harness asks only about those.
    expect(formatValue(serial(2026, 8, 24), 'd mmm yyyy')).toBe('24 Aug 2026')
  })

  it('puts the hour on a 12-hour clock when the code asks for AM/PM', () => {
    // `h:mm AM/PM` used to fail the date-code test on its letters, fall through
    // to number formatting, and show the reader a bare serial.
    const afternoon = serial(2026, 8, 24) + 13.5 / 24
    const morning = serial(2026, 8, 24) + 9.25 / 24
    expect(formatValue(afternoon, 'h:mm AM/PM')).toBe('1:30 PM')
    expect(formatValue(morning, 'hh:mm AM/PM')).toBe('09:15 AM')
    expect(formatValue(serial(2026, 8, 24), 'h:mm AM/PM')).toBe('12:00 AM')
    // and without the marker the same hour stays on a 24-hour clock
    expect(formatValue(afternoon, 'hh:mm')).toBe('13:30')
  })
})

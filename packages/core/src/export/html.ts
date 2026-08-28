import type { CompiledSheet, CompiledWorkbook, PlacedChart } from '../compile/emit.js'
import { type Computed, display, isExcelError, isNotEvaluated } from '../formula/value.js'
import { columnName } from '../model/a1.js'
import { parseCellKey, readsFromValues } from '../model/cell.js'
import { type HighlightTest, ICON_COLORS, ICON_GLYPHS, testOf } from '../model/highlight.js'
import { type ResolveContext, resolveRef } from '../refs/resolve.js'
import { formatValue, toCssText } from '../style/css.js'
import { themeFor } from '../style/design.js'
import { DEFAULT_THEME, resolveStyle } from '../style/theme.js'
import { mergeStyle, type Theme } from '../style/types.js'
import { marginBoxes } from './margin.js'
import { chartSvg, numberOf, seriesColor } from './svg-chart.js'

export interface HtmlOptions {
  title?: string
  theme?: Theme
  values?: Map<string, Computed>
  registry?: unknown
  definedNames?: unknown
  /** Landscape suits wide grids and is the default for print. */
  orientation?: 'portrait' | 'landscape'
  showGridHeaders?: boolean
}

interface Covered {
  master: string
  colspan: number
  rowspan: number
}

export function toHtml(book: CompiledWorkbook, options: HtmlOptions = {}): string {
  const theme = themeFor(book.design, options.theme ?? DEFAULT_THEME)
  // Charts resolve their ranges, so they need the workbook's registry.
  options = { registry: book.registry, definedNames: book.definedNames, ...options }
  const sheets = book.sheets.map((sheet) => renderSheet(sheet, theme, options)).join('\n')
  const title = escapeHtml(options.title ?? 'open-sheet')

  // A workbook of forms and a workbook of wide grids want opposite defaults, so
  // the sheets decide: any sheet asking for portrait wins, since a form printed
  // sideways is unusable while a grid merely wraps.
  const declared = book.sheets.map((sheet) => sheet.print?.orientation).filter(Boolean)
  const orientation =
    options.orientation ?? (declared.includes('portrait') ? 'portrait' : 'landscape')
  const size = book.sheets.find((sheet) => sheet.print?.size)?.print?.size ?? 'A4'
  const running = book.sheets.find((sheet) => sheet.print?.header || sheet.print?.footer)?.print
  const runningMargins = marginBoxes(running?.header, running?.footer)

  return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${title}</title>
<style>${stylesheet(theme, orientation, size, runningMargins)}</style>
</head>
<body>
<main class="os-workbook">
${sheets}
</main>
</body>
</html>
`
}

function stylesheet(
  theme: Theme,
  orientation: 'portrait' | 'landscape',
  size: string,
  runningMargins: string,
): string {
  const palette = Object.entries(theme.palette)
    .map(([name, value]) => `--os-${name}: ${value};`)
    .join('\n    ')

  return `
  :root {
    ${palette}
    --os-surface: #ffffff;
    --os-hairline: #e2e8f0;
  }
  * { box-sizing: border-box; }
  body {
    margin: 0;
    padding: 32px;
    background: #f1f5f9;
    color: var(--os-ink, #0f172a);
    font-family: Calibri, system-ui, -apple-system, sans-serif;
  }
  .os-workbook { display: flex; flex-direction: column; gap: 32px; max-width: 100%; }
  .os-sheet { background: var(--os-surface); border-radius: 8px; padding: 24px; box-shadow: 0 1px 3px rgba(15,23,42,.08); }
  .os-sheet > h2 {
    margin: 0 0 16px;
    font-size: 13px;
    font-weight: 700;
    letter-spacing: .08em;
    text-transform: uppercase;
    color: var(--os-muted, #64748b);
  }
  /* Wide grids scroll inside the sheet; the page itself never scrolls sideways. */
  .os-scroll { overflow-x: auto; }
  table.os-grid { border-collapse: collapse; font-variant-numeric: tabular-nums; }
  table.os-grid td { padding: 4px 8px; vertical-align: middle; white-space: nowrap; }
  table.os-grid td.os-num { text-align: right; }
  table.os-grid td.os-skip { color: #94a3b8; font-style: italic; }
  table.os-grid td.os-err { color: #b91c1c; }
  .os-head { background: #f8fafc; color: #94a3b8; font-size: 10px; text-align: center; font-weight: 600; }
  .os-chart-figure { margin: 20px 0 0; }
  .os-chart { max-width: 100%; height: auto; }
  .os-chart-title { font-size: 13px; font-weight: 700; fill: var(--os-ink, #0f172a); }
  .os-chart-grid { stroke: var(--os-hairline); stroke-width: 1; }
  .os-chart-tick { font-size: 10px; fill: var(--os-muted, #64748b); }
  @media print {
    body { background: #fff; padding: 0; }
    .os-sheet { box-shadow: none; border-radius: 0; padding: 0; break-after: page; }
    .os-sheet:last-child { break-after: auto; }
    thead { display: table-header-group; }
    tr { break-inside: avoid; }
    @page { size: ${size} ${orientation}; margin: 12mm; ${runningMargins} }
  }
  @media (prefers-color-scheme: dark) {
    body { background: #0b1220; color: #e2e8f0; }
    .os-sheet { background: #111827; box-shadow: none; border: 1px solid #1f2937; }
    .os-head { background: #0b1220; color: #475569; }
  }
`
}

interface BarScale {
  color: string
  max: number
}

/**
 * What one cell should look like once every conditional rule has had its say.
 * Computed here, in one place, from the same values the grid shows — so the
 * HTML and the xlsx cannot end up disagreeing about the same cell, which is a
 * failure this project has already had three times.
 */
interface Decoration {
  bar?: BarScale
  fill?: string
  color?: string
  bold?: boolean
  icon?: { glyph: string; color: string }
}

function lerp(from: string, to: string, t: number): string {
  const parse = (hex: string) => {
    const clean = hex.replace('#', '')
    const full =
      clean.length === 3
        ? clean
            .split('')
            .map((ch) => ch + ch)
            .join('')
        : clean
    return [0, 2, 4].map((i) => Number.parseInt(full.slice(i, i + 2), 16))
  }
  const a = parse(from)
  const b = parse(to)
  const mix = a.map((channel, i) => Math.round(channel + ((b[i] as number) - channel) * t))
  return `#${mix.map((channel) => channel.toString(16).padStart(2, '0')).join('')}`
}

function scaleColor(stops: readonly string[], t: number): string {
  if (stops.length === 2) return lerp(stops[0] as string, stops[1] as string, t)
  return t < 0.5
    ? lerp(stops[0] as string, stops[1] as string, t * 2)
    : lerp(stops[1] as string, stops[2] as string, (t - 0.5) * 2)
}

function matches(test: HighlightTest, value: Computed, others: readonly Computed[]): boolean {
  if ('above' in test) return typeof value === 'number' && value > test.above
  if ('below' in test) return typeof value === 'number' && value < test.below
  if ('atLeast' in test) return typeof value === 'number' && value >= test.atLeast
  if ('atMost' in test) return typeof value === 'number' && value <= test.atMost
  if ('equals' in test) return value === test.equals
  if ('between' in test) {
    return (
      typeof value === 'number' &&
      value >= (test.between[0] as number) &&
      value <= (test.between[1] as number)
    )
  }
  if ('contains' in test) {
    return typeof value === 'string' && value.toLowerCase().includes(test.contains.toLowerCase())
  }
  if ('duplicates' in test) return others.filter((other) => other === value).length > 1
  const numbers = others.filter((other): other is number => typeof other === 'number')
  if (typeof value !== 'number') return false
  if ('top' in test) {
    return [...numbers]
      .sort((a, b) => b - a)
      .slice(0, test.top)
      .includes(value)
  }
  return [...numbers]
    .sort((a, b) => a - b)
    .slice(0, test.bottom)
    .includes(value)
}

/** The HTML twin of an in-cell sparkline: the same numbers, drawn the same shape. */
function sparklineSvg(sheet: CompiledSheet, values: HtmlOptions['values']): Map<string, string> {
  const out = new Map<string, string>()
  for (const spark of sheet.sparklines) {
    const points: number[] = []
    for (let c = spark.source.c; c < spark.source.c + spark.source.cols; c += 1) {
      const cell = sheet.cells.get(`${spark.source.r},${c}`)
      const value = readsFromValues(cell)
        ? values?.get(`${sheet.name}!${spark.source.r},${c}`)
        : cell?.value
      points.push(typeof value === 'number' ? value : 0)
    }
    if (points.length < 2) continue

    const width = 64
    const height = 16
    const min = Math.min(...points)
    const max = Math.max(...points)
    const span = max - min || 1
    const step = width / (points.length - 1)
    const y = (value: number) => height - 2 - ((value - min) / span) * (height - 4)

    const body =
      spark.kind === 'column'
        ? points
            .map((value, i) => {
              const barWidth = Math.max((width / points.length) * 0.7, 1)
              const top = y(value)
              return `<rect x="${(i * width) / points.length}" y="${top}" width="${barWidth}" height="${Math.max(height - 2 - top, 1)}" fill="${spark.color}"/>`
            })
            .join('')
        : `<polyline points="${points.map((value, i) => `${i * step},${y(value)}`).join(' ')}" fill="none" stroke="${spark.color}" stroke-width="1.5"/>`

    out.set(
      `${spark.cell.r},${spark.cell.c}`,
      `<svg class="os-sparkline" viewBox="0 0 ${width} ${height}" width="${width}" height="${height}" role="img">${body}</svg>`,
    )
  }
  return out
}

function decorate(sheet: CompiledSheet, values: HtmlOptions['values']): Map<string, Decoration> {
  const out = new Map<string, Decoration>()
  const read = (r: number, c: number): Computed => {
    const cell = sheet.cells.get(`${r},${c}`)
    return (readsFromValues(cell) ? values?.get(`${sheet.name}!${r},${c}`) : cell?.value) ?? null
  }
  const at = (key: string): Decoration => {
    const existing = out.get(key)
    if (existing) return existing
    const fresh: Decoration = {}
    out.set(key, fresh)
    return fresh
  }

  for (const format of sheet.conditionalFormats) {
    const { r, c, rows, cols } = format.rect
    const cells: { key: string; value: Computed }[] = []
    for (let row = r; row < r + rows; row += 1) {
      for (let column = c; column < c + cols; column += 1) {
        cells.push({ key: `${row},${column}`, value: read(row, column) })
      }
    }
    const numbers = cells
      .map((cell) => cell.value)
      .filter((value): value is number => typeof value === 'number')

    if (format.kind === 'dataBar') {
      const max = Math.max(0, ...numbers.map(Math.abs))
      if (max === 0) continue
      for (const cell of cells) at(cell.key).bar = { color: format.color, max }
      continue
    }

    if (format.kind === 'colorScale' || format.kind === 'iconSet') {
      const min = Math.min(...numbers)
      const max = Math.max(...numbers)
      // A column where every value is the same has no scale; Excel paints the
      // first stop, so painting a midpoint here would be a visible divergence.
      const span = max - min
      for (const cell of cells) {
        if (typeof cell.value !== 'number') continue
        const t = span === 0 ? 0 : (cell.value - min) / span
        if (format.kind === 'colorScale') {
          at(cell.key).fill = scaleColor(format.scale, t)
        } else {
          // The same thirds the xlsx rule uses: 0%, 33%, 67%.
          const index = t >= 0.67 ? 2 : t >= 0.33 ? 1 : 0
          at(cell.key).icon = {
            glyph: ICON_GLYPHS[format.icons][index] as string,
            color: ICON_COLORS[format.icons][index] as string,
          }
        }
      }
      continue
    }

    const test = testOf(format.rule)
    const all = cells.map((cell) => cell.value)
    for (const cell of cells) {
      if (!matches(test, cell.value, all)) continue
      const decoration = at(cell.key)
      if (format.rule.fill) decoration.fill = format.rule.fill
      if (format.rule.color) decoration.color = format.rule.color
      if (format.rule.bold) decoration.bold = true
    }
  }
  return out
}

function renderSheet(sheet: CompiledSheet, theme: Theme, options: HtmlOptions): string {
  const covered = coverage(sheet)
  const decorations = decorate(sheet, options.values)
  const sparklines = sparklineSvg(sheet, options.values)
  const widths: string[] = []
  for (let c = 0; c < sheet.bounds.cols; c += 1) {
    const width = sheet.columnWidths.get(c) ?? theme.defaultColumnWidth
    widths.push(`<col style="width:${Math.round(width * 8)}px">`)
  }

  const rows: string[] = []
  if (options.showGridHeaders) {
    const heads = [`<td class="os-head"></td>`]
    for (let c = 0; c < sheet.bounds.cols; c += 1) {
      heads.push(`<td class="os-head">${columnName(c)}</td>`)
    }
    rows.push(`<tr>${heads.join('')}</tr>`)
  }

  for (let r = 0; r < sheet.bounds.rows; r += 1) {
    const cells: string[] = []
    if (options.showGridHeaders) cells.push(`<td class="os-head">${r + 1}</td>`)
    for (let c = 0; c < sheet.bounds.cols; c += 1) {
      const key = `${r},${c}`
      const hidden = covered.get(key)
      if (hidden && hidden.master !== key) continue
      cells.push(
        renderCell(sheet, r, c, theme, options, hidden, decorations.get(key), sparklines.get(key)),
      )
    }
    rows.push(`<tr>${cells.join('')}</tr>`)
  }

  const charts = sheet.charts
    .map(
      (chart) =>
        `<figure class="os-chart-figure">${renderChart(chart, sheet, theme, options)}</figure>`,
    )
    .join('')

  return `<section class="os-sheet">
  <h2>${escapeHtml(sheet.name)}</h2>
  <div class="os-scroll"><table class="os-grid"><colgroup>${
    options.showGridHeaders ? '<col style="width:36px">' : ''
  }${widths.join('')}</colgroup><tbody>
${rows.join('\n')}
  </tbody></table></div>
  ${charts}
</section>`
}

function renderCell(
  sheet: CompiledSheet,
  r: number,
  c: number,
  theme: Theme,
  options: HtmlOptions,
  span: Covered | undefined,
  decoration: Decoration | undefined,
  sparkline: string | undefined,
): string {
  const cell = sheet.cells.get(`${r},${c}`)
  const attrs: string[] = []
  if (span && span.colspan > 1) attrs.push(`colspan="${span.colspan}"`)
  if (span && span.rowspan > 1) attrs.push(`rowspan="${span.rowspan}"`)

  if (!cell) return `<td${attrs.length ? ` ${attrs.join(' ')}` : ''}></td>`

  const style = mergeStyle(resolveStyle(theme, undefined), resolveStyle(theme, cell.style))
  const computed = readsFromValues(cell)
    ? options.values?.get(`${sheet.name}!${r},${c}`)
    : (cell.value ?? null)

  let css = toCssText(style)
  if (cell.wrap) css += `${css ? ';' : ''}white-space:normal;vertical-align:top`
  // Order matters: a fill is a background colour and a bar is a background
  // image, so both can apply to one cell without either being lost.
  if (decoration?.fill) css += `${css ? ';' : ''}background-color:${decoration.fill}`
  if (decoration?.color) css += `${css ? ';' : ''}color:${decoration.color}`
  if (decoration?.bold) css += `${css ? ';' : ''}font-weight:600`
  const bar = decoration?.bar
  if (bar && typeof computed === 'number' && computed !== 0) {
    const pct = Math.min(100, Math.round((Math.abs(computed) / bar.max) * 100))
    css += `${css ? ';' : ''}background-image:linear-gradient(to right, ${bar.color} ${pct}%, transparent ${pct}%)`
  }
  if (css) attrs.push(`style="${css}"`)

  let className = ''
  let text: string
  if (isNotEvaluated(computed)) {
    className = 'os-skip'
    text = '#NOT_EVALUATED'
  } else if (isExcelError(computed)) {
    className = 'os-err'
    text = computed.code
  } else if (computed === undefined) {
    text = readsFromValues(cell) ? '' : display(cell.value ?? null)
  } else {
    if (typeof computed === 'number') className = 'os-num'
    text = formatValue(computed, cell.format)
  }

  // A note the reader cannot see in the browser is a note that only half
  // exists. `title` is the hover the xlsx gives them, in the medium at hand.
  if (cell.note) {
    attrs.push(`title="${escapeHtml(cell.note)}"`)
    className = className ? `${className} os-noted` : 'os-noted'
  }

  if (className) attrs.push(`class="${className}"`)
  if (sparkline) {
    return `<td${attrs.length ? ` ${attrs.join(' ')}` : ''}>${sparkline}</td>`
  }
  const icon = decoration?.icon
  const body = icon
    ? `<span class="os-icon" style="color:${icon.color}">${icon.glyph}</span>${escapeHtml(text)}`
    : escapeHtml(text)
  return `<td${attrs.length ? ` ${attrs.join(' ')}` : ''}>${body}</td>`
}

/** Reads the same evaluated values the grid shows, so the two cannot disagree. */
function renderChart(
  chart: PlacedChart,
  sheet: CompiledSheet,
  theme: Theme,
  options: HtmlOptions,
): string {
  void theme
  const context: ResolveContext = {
    registry: options.registry as never,
    definedNames: options.definedNames as never,
    sheet: sheet.name,
  }
  if (!context.registry) return ''

  const readRange = (ref: Parameters<typeof resolveRef>[0]): Computed[] => {
    const resolved = resolveRef(ref, context)
    const target = resolved.sheet === sheet.name ? sheet : undefined
    const out: Computed[] = []
    for (let r = resolved.rect.r; r < resolved.rect.r + resolved.rect.rows; r += 1) {
      for (let c = resolved.rect.c; c < resolved.rect.c + resolved.rect.cols; c += 1) {
        const cell = target?.cells.get(`${r},${c}`)
        out.push(
          cell?.expr
            ? (options.values?.get(`${resolved.sheet}!${r},${c}`) ?? null)
            : (cell?.value ?? null),
        )
      }
    }
    return out
  }

  try {
    return chartSvg(chart, {
      categories: readRange(chart.categories).map((value) => display(value ?? null)),
      series: chart.series.map((entry, index) => ({
        name: entry.name,
        values: readRange(entry.values).map(numberOf),
        color: seriesColor(index),
      })),
    })
  } catch {
    // A chart that cannot resolve is not worth failing the whole export over.
    return ''
  }
}

function coverage(sheet: CompiledSheet): Map<string, Covered> {
  const covered = new Map<string, Covered>()
  for (const [key, cell] of sheet.cells) {
    if (!cell.span || (cell.span.rows <= 1 && cell.span.cols <= 1)) continue
    const { r, c } = parseCellKey(key)
    for (let dr = 0; dr < cell.span.rows; dr += 1) {
      for (let dc = 0; dc < cell.span.cols; dc += 1) {
        covered.set(`${r + dr},${c + dc}`, {
          master: key,
          colspan: cell.span.cols,
          rowspan: cell.span.rows,
        })
      }
    }
  }
  return covered
}

const ESCAPES: Record<string, string> = {
  '&': '&amp;',
  '<': '&lt;',
  '>': '&gt;',
  '"': '&quot;',
  "'": '&#39;',
}

function escapeHtml(text: string): string {
  return text.replace(/[&<>"']/g, (ch) => ESCAPES[ch] as string)
}

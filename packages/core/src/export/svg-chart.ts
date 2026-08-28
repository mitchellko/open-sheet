import type { PlacedChart } from '../compile/emit.js'
import type { Computed } from '../formula/value.js'
import { formatValue } from '../style/css.js'

const PALETTE = ['#1d4ed8', '#0ea5e9', '#16a34a', '#d97706', '#dc2626', '#7c3aed']

export interface ChartData {
  categories: string[]
  series: { name: string; values: number[]; color: string }[]
}

function escapeXml(text: string): string {
  return text.replace(/[<>&]/g, (ch) => (ch === '<' ? '&lt;' : ch === '>' ? '&gt;' : '&amp;'))
}

/**
 * The HTML twin of the workbook's native chart. It reads the same evaluated
 * values the grid shows, so the two cannot disagree — but it is a picture, which
 * is exactly why the .xlsx gets real chart XML instead.
 */
export function chartSvg(chart: PlacedChart, data: ChartData, width = 520, height = 300): string {
  if (data.series.length === 0 || data.categories.length === 0) return ''

  const pad = { top: chart.title ? 34 : 16, right: 16, bottom: 42, left: 56 }
  const plotWidth = width - pad.left - pad.right
  const plotHeight = height - pad.top - pad.bottom

  const title = chart.title
    ? `<text x="${width / 2}" y="20" text-anchor="middle" class="os-chart-title">${escapeXml(chart.title)}</text>`
    : ''

  if (chart.chart === 'pie') {
    return wrap(width, height, title + pieBody(data, width, height, pad.top))
  }

  const stacked = chart.chart === 'stackedBar' || chart.chart === 'stackedArea'
  // A stacked chart's extent is the height of the stacks, not of one series.
  // Scaling to the tallest single value clips every stack above it.
  const all = stacked
    ? data.categories.map((_, i) =>
        data.series.reduce((sum, series) => sum + Math.max(series.values[i] ?? 0, 0), 0),
      )
    : data.series.flatMap((series) => series.values)
  const axes = chart.axes ?? {}
  const max = axes.max ?? Math.max(0, ...all)
  const min = axes.min ?? Math.min(0, ...all)
  const span = max - min || 1
  const y = (value: number) => pad.top + plotHeight - ((value - min) / span) * plotHeight

  const gridlines = [0, 0.25, 0.5, 0.75, 1]
    .map((fraction) => {
      const value = min + span * fraction
      return (
        `<line x1="${pad.left}" y1="${y(value)}" x2="${pad.left + plotWidth}" y2="${y(value)}" class="os-chart-grid"/>` +
        `<text x="${pad.left - 8}" y="${y(value) + 4}" text-anchor="end" class="os-chart-tick">${formatValue(value, axes.valueFormat ?? 'number')}</text>`
      )
    })
    .join('')

  const step = plotWidth / data.categories.length
  const labels = data.categories
    .map((label, i) => {
      const x = pad.left + step * (i + 0.5)
      return `<text x="${x}" y="${height - pad.bottom + 16}" text-anchor="middle" class="os-chart-tick">${escapeXml(label)}</text>`
    })
    .join('')

  const combo = chart.chart === 'combo'
  const barSeries = combo
    ? data.series.filter((_, i) => chart.series[i]?.as !== 'line')
    : data.series
  const lineSeries = combo ? data.series.filter((_, i) => chart.series[i]?.as === 'line') : []
  // A secondary series is on its own scale, so it gets its own extent —
  // otherwise a percentage target next to dollar revenue is a flat line at the
  // bottom of the plot, which is exactly what the second axis exists to avoid.
  const secondaryValues = combo
    ? data.series.filter((_, i) => chart.series[i]?.axis === 'secondary').flatMap((s2) => s2.values)
    : []
  const secondMax = Math.max(0, ...secondaryValues)
  const secondMin = Math.min(0, ...secondaryValues)
  const secondSpan = secondMax - secondMin || 1
  const y2 = (value: number) =>
    pad.top + plotHeight - ((value - secondMin) / secondSpan) * plotHeight

  const areaLike = chart.chart === 'area' || chart.chart === 'stackedArea'
  // Running totals per category, so each stacked segment sits on the one below.
  const base = data.categories.map(() => 0)

  const comboBody = !combo
    ? ''
    : barSeries
        .map((series, s) => {
          const barWidth = (step * 0.7) / Math.max(barSeries.length, 1)
          return series.values
            .map((value, i) => {
              const x = pad.left + step * i + step * 0.15 + barWidth * s
              const top = y(Math.max(value, 0))
              const bottom = y(Math.min(value, 0))
              return `<rect x="${x}" y="${top}" width="${barWidth}" height="${Math.max(bottom - top, 1)}" fill="${series.color}"/>`
            })
            .join('')
        })
        .join('') +
      lineSeries
        .map((series) => {
          const onSecond =
            chart.series.find((entry) => entry.name === series.name)?.axis === 'secondary'
          const scale = onSecond ? y2 : y
          const points = series.values
            .map((value, i) => `${pad.left + step * (i + 0.5)},${scale(value)}`)
            .join(' ')
          return `<polyline points="${points}" fill="none" stroke="${series.color}" stroke-width="2"/>`
        })
        .join('')

  const body = combo
    ? comboBody
    : chart.chart === 'scatter'
      ? data.series
          .map((series) =>
            series.values
              .map((value, i) => {
                const xValue = Number(data.categories[i])
                const t = Number.isFinite(xValue)
                  ? (xValue - Math.min(...data.categories.map(Number))) /
                    (Math.max(...data.categories.map(Number)) -
                      Math.min(...data.categories.map(Number)) || 1)
                  : i / Math.max(data.categories.length - 1, 1)
                return `<circle cx="${pad.left + t * plotWidth}" cy="${y(value)}" r="4" fill="${series.color}"/>`
              })
              .join(''),
          )
          .join('')
      : areaLike
        ? data.series
            .map((series) => {
              const tops = series.values.map((value, i) => {
                const bottom = stacked ? (base[i] as number) : 0
                const top = bottom + Math.max(value, 0)
                if (stacked) base[i] = top
                return { x: pad.left + step * (i + 0.5), top: y(top), bottom: y(bottom) }
              })
              const upper = tops.map((point) => `${point.x},${point.top}`).join(' ')
              const lower = [...tops]
                .reverse()
                .map((point) => `${point.x},${point.bottom}`)
                .join(' ')
              return `<polygon points="${upper} ${lower}" fill="${series.color}" fill-opacity="0.55" stroke="${series.color}" stroke-width="1.5"/>`
            })
            .join('')
        : chart.chart === 'line'
          ? data.series
              .map((series) => {
                const points = series.values
                  .map((value, i) => `${pad.left + step * (i + 0.5)},${y(value)}`)
                  .join(' ')
                return `<polyline points="${points}" fill="none" stroke="${series.color}" stroke-width="2"/>`
              })
              .join('')
          : data.series
              .map((series, s) => {
                const barWidth = stacked ? step * 0.7 : (step * 0.7) / data.series.length
                return series.values
                  .map((value, i) => {
                    const x = pad.left + step * i + step * 0.15 + (stacked ? 0 : barWidth * s)
                    const from = stacked ? (base[i] as number) : 0
                    const to = from + (stacked ? Math.max(value, 0) : value)
                    if (stacked) base[i] = to
                    const top = y(Math.max(to, from))
                    const bottom = y(Math.min(to, from))
                    return `<rect x="${x}" y="${top}" width="${barWidth}" height="${Math.max(bottom - top, 1)}" fill="${series.color}"/>`
                  })
                  .join('')
              })
              .join('')

  const legend =
    data.series.length > 1
      ? data.series
          .map((series, i) => {
            const x = pad.left + i * 110
            return (
              `<rect x="${x}" y="${height - 14}" width="9" height="9" fill="${series.color}"/>` +
              `<text x="${x + 14}" y="${height - 6}" class="os-chart-tick">${escapeXml(series.name)}</text>`
            )
          })
          .join('')
      : ''

  return wrap(width, height, title + gridlines + body + labels + legend)
}

function pieBody(data: ChartData, width: number, height: number, top: number): string {
  const values = data.series[0]?.values ?? []
  const total = values.reduce((sum, value) => sum + Math.max(value, 0), 0)
  if (total === 0) return ''

  const cx = width / 2
  const cy = top + (height - top - 30) / 2
  const radius = Math.min(cx, cy - top) * 0.8
  let angle = -Math.PI / 2

  return values
    .map((value, i) => {
      const slice = (Math.max(value, 0) / total) * Math.PI * 2
      const x1 = cx + radius * Math.cos(angle)
      const y1 = cy + radius * Math.sin(angle)
      angle += slice
      const x2 = cx + radius * Math.cos(angle)
      const y2 = cy + radius * Math.sin(angle)
      const large = slice > Math.PI ? 1 : 0
      const color = PALETTE[i % PALETTE.length] as string
      return `<path d="M ${cx} ${cy} L ${x1} ${y1} A ${radius} ${radius} 0 ${large} 1 ${x2} ${y2} Z" fill="${color}"/>`
    })
    .join('')
}

function wrap(width: number, height: number, body: string): string {
  return `<svg viewBox="0 0 ${width} ${height}" width="${width}" height="${height}" role="img" class="os-chart">${body}</svg>`
}

export function seriesColor(index: number): string {
  return PALETTE[index % PALETTE.length] as string
}

export function numberOf(value: Computed | undefined): number {
  return typeof value === 'number' ? value : 0
}

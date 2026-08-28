import { unzipSync, zipSync } from 'fflate'
import type { CompiledSheet, CompiledWorkbook, PlacedChart } from '../compile/emit.js'
import type { TableAnchor } from '../compile/registry.js'
import { serialize } from '../formula/serialize.js'
import { qualify, rangeToA1 } from '../model/a1.js'
import { cellKey } from '../model/cell.js'
import { type ResolveContext, resolveRef } from '../refs/resolve.js'
import { toArgb } from '../style/excel.js'
import { numberFormat } from './formats.js'

const NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'

const SERIES_COLORS = ['1D4ED8', '0EA5E9', '16A34A', 'D97706', 'DC2626', '7C3AED']

function escapeXml(text: string): string {
  return text.replace(/[<>&'"]/g, (ch) =>
    ch === '<'
      ? '&lt;'
      : ch === '>'
        ? '&gt;'
        : ch === '&'
          ? '&amp;'
          : ch === "'"
            ? '&apos;'
            : '&quot;',
  )
}

interface ResolvedSeries {
  name: string
  range: string
  color: string
  count: number
  as?: 'bar' | 'line'
  axis?: 'primary' | 'secondary'
}

function resolveRange(ref: Parameters<typeof resolveRef>[0], context: ResolveContext) {
  const resolved = resolveRef(ref, context)
  return {
    range: qualify(
      resolved.sheet,
      rangeToA1(resolved.rect, { absoluteRow: true, absoluteCol: true }),
    ),
    count: resolved.rect.rows * resolved.rect.cols,
  }
}

/** Values printed on the points themselves. Off unless asked for: on a dense series they collide. */
function labelsXml(on: boolean | undefined): string {
  if (!on) return ''
  return (
    '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/>' +
    '<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>'
  )
}

function seriesXml(
  series: ResolvedSeries[],
  categories: string,
  kind: PlacedChart['chart'],
  dataLabels: boolean | undefined,
): string {
  return series
    .map((entry, index) => {
      // A scatter's x axis is numeric, so its categories are a numRef and the
      // pair is xVal/yVal — a cat/val series renders as a straight diagonal.
      const points =
        kind === 'scatter'
          ? `<c:xVal><c:numRef><c:f>${escapeXml(categories)}</c:f></c:numRef></c:xVal>` +
            `<c:yVal><c:numRef><c:f>${escapeXml(entry.range)}</c:f></c:numRef></c:yVal>`
          : `<c:cat><c:strRef><c:f>${escapeXml(categories)}</c:f></c:strRef></c:cat>` +
            `<c:val><c:numRef><c:f>${escapeXml(entry.range)}</c:f></c:numRef></c:val>`
      const fill =
        kind === 'pie'
          ? ''
          : kind === 'scatter'
            ? // A scatter with a line drawn through it is a line chart with the
              // wrong x axis; markers only is what people mean by scatter.
              `<c:spPr><a:ln w="0"><a:noFill/></a:ln></c:spPr>` +
              `<c:marker><c:symbol val="circle"/><c:size val="6"/>` +
              `<c:spPr><a:solidFill><a:srgbClr val="${entry.color}"/></a:solidFill></c:spPr></c:marker>`
            : `<c:spPr><a:solidFill><a:srgbClr val="${entry.color}"/></a:solidFill></c:spPr>`
      const marker = kind === 'line' ? '<c:marker><c:symbol val="none"/></c:marker>' : ''
      return (
        `<c:ser><c:idx val="${index}"/><c:order val="${index}"/>` +
        `<c:tx><c:v>${escapeXml(entry.name)}</c:v></c:tx>` +
        fill +
        marker +
        labelsXml(dataLabels) +
        points +
        '</c:ser>'
      )
    })
    .join('')
}

const CAT_AX = '111111111'
const VAL_AX = '222222222'
/** A secondary axis needs its own pair; the extra category axis is deleted, not drawn twice. */
const CAT_AX_2 = '333333333'
const VAL_AX_2 = '444444444'

const AX_IDS = `<c:axId val="${CAT_AX}"/><c:axId val="${VAL_AX}"/>`
const AX_IDS_2 = `<c:axId val="${CAT_AX_2}"/><c:axId val="${VAL_AX_2}"/>`

/**
 * A combo is two plots in one plot area, each with its own series list. The
 * order matters: bars first, so the line draws over them rather than under.
 */
function comboXml(
  series: ResolvedSeries[],
  categories: string,
  dataLabels: boolean | undefined,
): string {
  const bars = series.filter((entry) => entry.as !== 'line')
  const lines = series.filter((entry) => entry.as === 'line')
  const secondary = series.some((entry) => entry.axis === 'secondary')
  // Whichever group holds the secondary series moves to the second axis pair.
  const lineOnSecondary = secondary && lines.some((entry) => entry.axis === 'secondary')

  const barPart = bars.length
    ? '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>' +
      seriesXml(bars, categories, 'bar', dataLabels) +
      `<c:gapWidth val="60"/>${lineOnSecondary ? AX_IDS : secondary ? AX_IDS_2 : AX_IDS}</c:barChart>`
    : ''
  const linePart = lines.length
    ? '<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>' +
      seriesXml(lines, categories, 'line', dataLabels) +
      `<c:marker val="1"/>${lineOnSecondary ? AX_IDS_2 : AX_IDS}</c:lineChart>`
    : ''
  return barPart + linePart
}

function plotXml(
  kind: PlacedChart['chart'],
  series: ResolvedSeries[],
  categories: string,
  dataLabels: boolean | undefined,
): string {
  if (kind === 'combo') return comboXml(series, categories, dataLabels)
  const inner = seriesXml(series, categories, kind, dataLabels)
  if (kind === 'pie') {
    return `<c:pieChart><c:varyColors val="1"/>${inner}</c:pieChart>`
  }
  if (kind === 'scatter') {
    return `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${inner}${AX_IDS}</c:scatterChart>`
  }
  if (kind === 'line') {
    return (
      '<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>' +
      inner +
      `<c:marker val="1"/>${AX_IDS}</c:lineChart>`
    )
  }
  if (kind === 'area' || kind === 'stackedArea') {
    return (
      `<c:areaChart><c:grouping val="${kind === 'stackedArea' ? 'stacked' : 'standard'}"/>` +
      `<c:varyColors val="0"/>${inner}${AX_IDS}</c:areaChart>`
    )
  }
  // `overlap` 100 is not cosmetic: without it a stacked bar draws its segments
  // side by side with a gap, which reads as a clustered chart with wrong values.
  const stacked = kind === 'stackedBar'
  return (
    `<c:barChart><c:barDir val="col"/><c:grouping val="${stacked ? 'stacked' : 'clustered'}"/>` +
    '<c:varyColors val="0"/>' +
    inner +
    `<c:gapWidth val="60"/>${stacked ? '<c:overlap val="100"/>' : ''}${AX_IDS}</c:barChart>`
  )
}

function axisTitleXml(title: string | undefined): string {
  if (!title) return ''
  return (
    '<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>' +
    escapeXml(title) +
    '</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>'
  )
}

function axesXml(chart: PlacedChart): string {
  if (chart.chart === 'pie') return ''
  const axes = chart.axes ?? {}
  // Order is fixed by the schema: scaling, delete, axPos, gridlines, title,
  // numFmt, then crossAx. Excel rejects the file outright if they are shuffled.
  const scaling =
    '<c:scaling><c:orientation val="minMax"/>' +
    (axes.max === undefined ? '' : `<c:max val="${axes.max}"/>`) +
    (axes.min === undefined ? '' : `<c:min val="${axes.min}"/>`) +
    '</c:scaling>'
  // A named format ("currency") resolves to its code; an unknown one resolves
  // to nothing, and an empty formatCode makes Excel reject the chart part.
  const format = numberFormat(axes.valueFormat)
  const valueFormat = format ? `<c:numFmt formatCode="${escapeXml(format)}" sourceLinked="0"/>` : ''

  // A scatter's x axis is numeric, so it is a second valAx, not a catAx.
  const first =
    chart.chart === 'scatter'
      ? `<c:valAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling>` +
        `<c:delete val="0"/><c:axPos val="b"/>${axisTitleXml(axes.category)}` +
        '<c:crossAx val="222222222"/></c:valAx>'
      : '<c:catAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
        `<c:delete val="0"/><c:axPos val="b"/>${axisTitleXml(axes.category)}` +
        '<c:crossAx val="222222222"/></c:catAx>'

  const primary =
    first +
    `<c:valAx><c:axId val="${VAL_AX}"/>` +
    scaling +
    '<c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/>' +
    axisTitleXml(axes.value) +
    valueFormat +
    `<c:crossAx val="${CAT_AX}"/></c:valAx>`

  if (chart.chart !== 'combo' || !chart.series.some((entry) => entry.axis === 'secondary')) {
    return primary
  }

  // The second category axis is required by the schema and deleted on sight:
  // drawing it would print the month labels twice, once along each edge.
  return (
    primary +
    `<c:valAx><c:axId val="${VAL_AX_2}"/><c:scaling><c:orientation val="minMax"/></c:scaling>` +
    '<c:delete val="0"/><c:axPos val="r"/>' +
    axisTitleXml(axes.secondary) +
    `<c:crossAx val="${CAT_AX_2}"/><c:crosses val="max"/></c:valAx>` +
    `<c:catAx><c:axId val="${CAT_AX_2}"/><c:scaling><c:orientation val="minMax"/></c:scaling>` +
    `<c:delete val="1"/><c:axPos val="b"/><c:crossAx val="${VAL_AX_2}"/></c:catAx>`
  )
}

function chartXml(chart: PlacedChart, context: ResolveContext): string {
  const categories = resolveRange(chart.categories, context)
  const series: ResolvedSeries[] = chart.series.map((entry, index) => {
    const resolved = resolveRange(entry.values, context)
    return {
      name: entry.name,
      range: resolved.range,
      count: resolved.count,
      color: SERIES_COLORS[index % SERIES_COLORS.length] as string,
      ...(entry.as === undefined ? {} : { as: entry.as }),
      ...(entry.axis === undefined ? {} : { axis: entry.axis }),
    }
  })

  const title = chart.title
    ? `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>${escapeXml(chart.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`
    : '<c:autoTitleDeleted val="1"/>'

  const legend =
    series.length > 1 || chart.chart === 'pie'
      ? '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>'
      : ''

  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}" xmlns:r="${NS_R}">` +
    '<c:chart>' +
    title +
    '<c:plotArea><c:layout/>' +
    plotXml(chart.chart, series, categories.range, chart.dataLabels) +
    axesXml(chart) +
    '</c:plotArea>' +
    legend +
    '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/>' +
    '</c:chart></c:chartSpace>'
  )
}

function drawingXml(charts: PlacedChart[]): string {
  const anchors = charts
    .map((chart, index) => {
      const { r, c, rows, cols } = chart.rect
      return (
        '<xdr:twoCellAnchor editAs="oneCell">' +
        `<xdr:from><xdr:col>${c}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${r}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
        `<xdr:to><xdr:col>${c + cols}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${r + rows}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>` +
        '<xdr:graphicFrame macro="">' +
        `<xdr:nvGraphicFramePr><xdr:cNvPr id="${index + 2}" name="Chart ${index + 1}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
        '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
        `<c:chart xmlns:c="${NS_C}" xmlns:r="${NS_R}" r:id="rId${index + 1}"/>` +
        '</a:graphicData></a:graphic></xdr:graphicFrame>' +
        '<xdr:clientData/></xdr:twoCellAnchor>'
      )
    })
    .join('')

  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<xdr:wsDr xmlns:xdr="${NS_XDR}" xmlns:a="${NS_A}">${anchors}</xdr:wsDr>`
  )
}

function drawingRels(count: number, firstChart: number): string {
  const rels = Array.from({ length: count }, (_, i) => {
    const target = `../charts/chart${firstChart + i}.xml`
    return `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="${target}"/>`
  }).join('')
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`
  )
}

const decoder = new TextDecoder()
const encoder = new TextEncoder()

/**
 * ExcelJS has no chart support, and an embedded image would go stale the moment
 * a number changed — so the chart parts are written into the package it
 * produced. Everything here is bound to cell ranges, which is what keeps the
 * chart live.
 */
const NS_X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
const NS_XM = 'http://schemas.microsoft.com/office/excel/2006/main'

/**
 * A sparkline is not a chart part — it lives in the worksheet's extension list,
 * which is why it survives being one cell wide. One group per (kind, colour)
 * pair, since the group carries both.
 */
function sparklineXml(sheet: CompiledSheet): string {
  if (sheet.sparklines.length === 0) return ''

  const groups = new Map<string, typeof sheet.sparklines>()
  for (const spark of sheet.sparklines) {
    const id = `${spark.kind}|${spark.color}`
    groups.set(id, [...(groups.get(id) ?? []), spark])
  }

  const body = [...groups.values()]
    .map((group) => {
      const first = group[0] as (typeof group)[number]
      const lines = group
        .map(
          (spark) =>
            '<x14:sparkline>' +
            `<xm:f>${escapeXml(qualify(sheet.name, rangeToA1(spark.source)))}</xm:f>` +
            `<xm:sqref>${rangeToA1({ ...spark.cell, rows: 1, cols: 1 })}</xm:sqref>` +
            '</x14:sparkline>',
        )
        .join('')
      return (
        `<x14:sparklineGroup displayEmptyCellsAs="gap" type="${first.kind === 'column' ? 'column' : 'line'}">` +
        `<x14:colorSeries rgb="${toArgb(first.color)}"/>` +
        `<x14:sparklines>${lines}</x14:sparklines></x14:sparklineGroup>`
      )
    })
    .join('')

  return (
    '<ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" ' +
    `xmlns:x14="${NS_X14}"><x14:sparklineGroups xmlns:xm="${NS_XM}">` +
    body +
    '</x14:sparklineGroups></ext>'
  )
}

/** `extLst` must be the last element of the worksheet, after pageSetup. */
function addExtension(sheetXml: string, ext: string): string {
  if (!ext) return sheetXml
  const existing = sheetXml.indexOf('<extLst>')
  if (existing >= 0) {
    return (
      sheetXml.slice(0, existing + '<extLst>'.length) +
      ext +
      sheetXml.slice(existing + '<extLst>'.length)
    )
  }
  return sheetXml.replace('</worksheet>', `<extLst>${ext}</extLst></worksheet>`)
}

/**
 * ExcelJS writes no `<calculatedColumnFormula>`, so Excel treats a derived
 * column as ordinary cell formulas and an appended row comes out blank — the
 * table grows, the ranges follow, and the reader is left to fill the formulas
 * in by hand. It also writes totals-row attributes on tables that have no
 * totals row. Both are repaired here.
 */
function rewriteTablePart(
  xml: string,
  anchor: TableAnchor,
  sheet: CompiledSheet,
  book: CompiledWorkbook,
): string {
  const context: ResolveContext = {
    registry: book.registry,
    definedNames: book.definedNames,
    sheet: anchor.sheet,
  }
  const hasTotals = anchor.totalRow !== undefined

  // Which columns can fill down was decided at compile time; the writer only
  // renders the decision, so the CLI and the file cannot disagree about it.
  const cannotFill = new Set(anchor.table?.noFillDown ?? [])
  const formulas = new Map<string, string>()
  for (const [key, header] of anchor.table?.headers ?? []) {
    if (cannotFill.has(header)) continue
    const column = anchor.columns.get(key)
    if (column === undefined) continue
    const cell = sheet.cells.get(cellKey(anchor.lastDataRow, column))
    if (!cell?.expr) continue
    formulas.set(header, serialize(cell.expr, { ...context, row: anchor.lastDataRow }))
  }

  const out = xml.replace(
    /<tableColumn([^>]*?)(\/>|>[\s\S]*?<\/tableColumn>)/g,
    (whole, attrs: string) => {
      const name = /\sname="([^"]*)"/.exec(attrs)?.[1]
      if (name === undefined) return whole
      let kept = attrs
      if (!hasTotals) kept = kept.replace(/\stotalsRow(Function|Label)="[^"]*"/g, '')
      const formula = formulas.get(decodeXml(name))
      if (!formula) return `<tableColumn${kept}/>`
      return `<tableColumn${kept}><calculatedColumnFormula>${escapeXml(formula.replace(/^=/, ''))}</calculatedColumnFormula></tableColumn>`
    },
  )
  return out
}

function decodeXml(text: string): string {
  return text
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
}

export function injectCharts(zip: Buffer, book: CompiledWorkbook): Buffer {
  const withCharts = book.sheets
    .map((sheet, index) => ({ sheet, index }))
    .filter(({ sheet }) => sheet.charts.length > 0)
  const withSparklines = book.sheets
    .map((sheet, index) => ({ sheet, index }))
    .filter(({ sheet }) => sheet.sparklines.length > 0)

  const hasTables = [...book.registry.values()].some(
    (anchor) => anchor.kind === 'table' && anchor.table !== undefined,
  )
  if (withCharts.length === 0 && withSparklines.length === 0 && !hasTables) return zip

  const files = unzipSync(new Uint8Array(zip))

  const appendable = [...book.registry.values()].filter(
    (anchor): anchor is TableAnchor => anchor.kind === 'table' && anchor.table !== undefined,
  )
  for (const anchor of appendable) {
    const sheet = book.sheets.find((candidate) => candidate.name === anchor.sheet)
    if (!sheet) continue
    for (const path of Object.keys(files)) {
      if (!/^xl\/tables\/table\d+\.xml$/.test(path)) continue
      const xml = text(files[path])
      if (!new RegExp(`name="${anchor.name}"`).test(xml)) continue
      files[path] = encoder.encode(rewriteTablePart(xml, anchor, sheet, book))
    }
  }

  for (const { sheet, index } of withSparklines) {
    const path = `xl/worksheets/sheet${index + 1}.xml`
    files[path] = encoder.encode(addExtension(text(files[path]), sparklineXml(sheet)))
  }
  let chartNumber = 1

  for (const { sheet, index } of withCharts) {
    const sheetNumber = index + 1
    const drawingNumber = sheetNumber
    const firstChart = chartNumber

    const context: ResolveContext = {
      registry: book.registry,
      definedNames: book.definedNames,
      sheet: sheet.name,
    }

    for (const chart of sheet.charts) {
      files[`xl/charts/chart${chartNumber}.xml`] = encoder.encode(chartXml(chart, context))
      chartNumber += 1
    }

    files[`xl/drawings/drawing${drawingNumber}.xml`] = encoder.encode(drawingXml(sheet.charts))
    files[`xl/drawings/_rels/drawing${drawingNumber}.xml.rels`] = encoder.encode(
      drawingRels(sheet.charts.length, firstChart),
    )

    const sheetPath = `xl/worksheets/sheet${sheetNumber}.xml`
    const relsPath = `xl/worksheets/_rels/sheet${sheetNumber}.xml.rels`
    const relId = addSheetRel(files, relsPath, drawingNumber)
    files[sheetPath] = encoder.encode(addDrawingElement(text(files[sheetPath]), relId))
  }

  if (withCharts.length > 0) {
    files['[Content_Types].xml'] = encoder.encode(
      addContentTypes(text(files['[Content_Types].xml']), chartNumber - 1, withCharts.length),
    )
  }

  return Buffer.from(zipSync(files))
}

function text(data: Uint8Array | undefined): string {
  if (!data) throw new Error('the workbook package is missing a part charts need')
  return decoder.decode(data)
}

function addSheetRel(
  files: Record<string, Uint8Array>,
  relsPath: string,
  drawingNumber: number,
): string {
  const existing = files[relsPath]
  const target = `../drawings/drawing${drawingNumber}.xml`
  const type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'

  if (!existing) {
    const id = 'rId1'
    files[relsPath] = encoder.encode(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        `<Relationship Id="${id}" Type="${type}" Target="${target}"/></Relationships>`,
    )
    return id
  }

  const xml = decoder.decode(existing)
  // Continue the sheet's own numbering; colliding with an existing id makes the
  // file open to an error dialog rather than a chart.
  const used = [...xml.matchAll(/Id="rId(\d+)"/g)].map((match) => Number(match[1]))
  const id = `rId${Math.max(0, ...used) + 1}`
  files[relsPath] = encoder.encode(
    xml.replace(
      '</Relationships>',
      `<Relationship Id="${id}" Type="${type}" Target="${target}"/></Relationships>`,
    ),
  )
  return id
}

/** `<drawing/>` has a fixed position in the sheet schema; misplacing it invalidates the file. */
const AFTER_DRAWING = ['</legacyDrawing>', '</picture>', '</oleObjects>']

function addDrawingElement(sheetXml: string, relId: string): string {
  const element = `<drawing r:id="${relId}"/>`
  if (sheetXml.includes('<drawing ')) return sheetXml

  for (const marker of AFTER_DRAWING) {
    const at = sheetXml.indexOf(marker)
    if (at !== -1)
      return sheetXml.slice(0, at + marker.length) + element + sheetXml.slice(at + marker.length)
  }
  return sheetXml.replace('</worksheet>', `${element}</worksheet>`)
}

function addContentTypes(xml: string, charts: number, drawings: number): string {
  const parts: string[] = []
  if (!xml.includes('Extension="xml"')) {
    parts.push('<Default Extension="xml" ContentType="application/xml"/>')
  }
  for (let i = 1; i <= charts; i += 1) {
    parts.push(
      `<Override PartName="/xl/charts/chart${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`,
    )
  }
  for (let i = 1; i <= drawings; i += 1) {
    if (xml.includes(`/xl/drawings/drawing${i}.xml"`)) continue
    parts.push(
      `<Override PartName="/xl/drawings/drawing${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`,
    )
  }
  return xml.replace('</Types>', `${parts.join('')}</Types>`)
}

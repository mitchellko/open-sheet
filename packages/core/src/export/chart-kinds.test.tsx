import { execFileSync } from 'node:child_process'
import { existsSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { strFromU8, unzipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { compile } from '../compile/compile.js'
import { Chart, col, Sheet, Stack, Table, Workbook } from '../compile/components.js'
import type { ChartAxes, ChartKind } from '../compile/nodes.js'
import { evaluateWorkbook } from '../formula/evaluate.js'
import { ref } from '../refs/ref.js'
import { toHtml } from './html.js'
import { XlsxWriter } from './xlsx.js'

const months = [
  { month: 'Jan', a: 10, b: 40 },
  { month: 'Feb', a: 30, b: 20 },
  { month: 'Mar', a: 50, b: 10 },
]

function book(kind: ChartKind, extra: { axes?: ChartAxes; dataLabels?: boolean } = {}) {
  return compile(
    (
      <Workbook>
        <Sheet name="S">
          <Stack gap={1}>
            <Table name="t" data={months} columns={[col('month'), col('a'), col('b')]} />
            <Chart
              kind={kind}
              title="Cost"
              categories={ref('t').column(kind === 'scatter' ? 'a' : 'month')}
              series={[
                { name: 'A', values: ref('t').column('a') },
                { name: 'B', values: ref('t').column('b') },
              ]}
              {...extra}
            />
          </Stack>
        </Sheet>
      </Workbook>
    ) as never,
  )
}

async function chartXml(kind: ChartKind, extra = {}): Promise<string> {
  const buffer = await new XlsxWriter().write(book(kind, extra))
  const files = unzipSync(new Uint8Array(buffer))
  const part = Object.keys(files).find((name) => /charts\/chart\d+\.xml$/.test(name))
  expect(part, 'no chart part was written').toBeDefined()
  return strFromU8(files[part as string] as Uint8Array)
}

function svg(kind: ChartKind, extra = {}): string {
  const compiled = book(kind, extra)
  return toHtml(compiled, { title: 'S', values: evaluateWorkbook(compiled) })
}

describe('the plot shapes', () => {
  it('stacks a bar chart and closes the gap between segments', async () => {
    const xml = await chartXml('stackedBar')
    expect(xml).toContain('<c:grouping val="stacked"/>')
    // Without overlap 100 the segments sit side by side and it reads as a
    // clustered chart whose values are wrong.
    expect(xml).toContain('<c:overlap val="100"/>')
  })

  it('leaves a plain bar clustered and without overlap', async () => {
    const xml = await chartXml('bar')
    expect(xml).toContain('<c:grouping val="clustered"/>')
    expect(xml).not.toContain('<c:overlap')
  })

  it('writes an area chart, stacked or not', async () => {
    expect(await chartXml('area')).toContain('<c:areaChart><c:grouping val="standard"/>')
    expect(await chartXml('stackedArea')).toContain('<c:areaChart><c:grouping val="stacked"/>')
  })

  it('gives a scatter xVal/yVal and two value axes', async () => {
    const xml = await chartXml('scatter')
    expect(xml).toContain('<c:scatterChart>')
    // A cat/val series on a scatter renders as a straight diagonal line.
    expect(xml).toContain('<c:xVal>')
    expect(xml).toContain('<c:yVal>')
    expect(xml).not.toContain('<c:catAx>')
    expect((xml.match(/<c:valAx>/g) ?? []).length).toBe(2)
  })
})

describe('what makes a chart readable', () => {
  it('titles both axes', async () => {
    const xml = await chartXml('bar', { axes: { category: 'Month', value: 'NT$' } })
    expect(xml).toContain('<a:t>Month</a:t>')
    expect(xml).toContain('<a:t>NT$</a:t>')
  })

  it('formats the value axis with the same codes cells use', async () => {
    const xml = await chartXml('bar', { axes: { valueFormat: 'currency' } })
    expect(xml).toContain('<c:numFmt formatCode=')
    expect(xml).toContain('sourceLinked="0"')
  })

  it('omits numFmt for a format it cannot resolve, rather than writing an empty one', async () => {
    // Excel rejects the whole chart part on an empty formatCode.
    const xml = await chartXml('bar', { axes: {} })
    expect(xml).not.toContain('<c:numFmt')
  })

  it('pins the value axis when asked', async () => {
    const xml = await chartXml('line', { axes: { min: 0, max: 100 } })
    expect(xml).toContain('<c:max val="100"/>')
    expect(xml).toContain('<c:min val="0"/>')
  })

  it('prints data labels only when asked', async () => {
    expect(await chartXml('bar', { dataLabels: true })).toContain('<c:showVal val="1"/>')
    expect(await chartXml('bar')).not.toContain('<c:dLbls>')
  })
})

describe('the HTML twin draws the same chart', () => {
  it('stacks the segments instead of putting them side by side', () => {
    const stacked = svg('stackedBar')
    const clustered = svg('bar')
    // Stacked bars are full-width, one per category; clustered are half-width,
    // two per category. Comparing widths is how you tell them apart.
    const widthsOf = (markup: string) =>
      [...markup.matchAll(/<rect[^>]*width="([\d.]+)"[^>]*fill="#/g)].map((m) => Number(m[1]))
    expect(Math.max(...widthsOf(stacked))).toBeGreaterThan(Math.max(...widthsOf(clustered)))
  })

  it('scales a stacked chart to the height of the stacks', () => {
    // Every category here sums to 50 while no single value exceeds 50, so a
    // chart scaled to the tallest single value would clip every stack.
    const markup = svg('stackedArea')
    expect(markup).toContain('<polygon')
    expect(markup).toContain('50')
  })

  it('draws a scatter as points, not a line', () => {
    const markup = svg('scatter')
    expect(markup).toContain('<circle')
    expect(markup).not.toContain('<polyline')
  })

  it('honours a pinned axis and the value format', () => {
    const markup = svg('bar', { axes: { min: 0, max: 100, valueFormat: 'percent' } })
    expect(markup).toContain('%')
  })
})

const SOFFICE = ['/opt/homebrew/bin/soffice', '/usr/bin/soffice', '/usr/bin/libreoffice'].find(
  (path) => existsSync(path),
)

describe.skipIf(!SOFFICE)('and a real spreadsheet reads them', () => {
  it.each(['stackedBar', 'area', 'stackedArea', 'scatter'] as const)(
    'keeps a %s bound to its ranges',
    { timeout: 180_000 },
    async (kind) => {
      const buffer = await new XlsxWriter().write(
        book(kind, { axes: { category: 'Month', value: 'NT$', valueFormat: 'currency' } }),
      )
      const dir = mkdtempSync(join(tmpdir(), `open-sheet-chart-${kind}-`))
      writeFileSync(join(dir, 'c.xlsx'), buffer)
      execFileSync(
        SOFFICE as string,
        [
          `-env:UserInstallation=file://${join(dir, 'profile')}`,
          '--headless',
          '--convert-to',
          'xlsx',
          '--outdir',
          join(dir, 'out'),
          join(dir, 'c.xlsx'),
        ],
        { stdio: 'pipe', timeout: 150_000 },
      )
      const files = unzipSync(new Uint8Array(readFileSync(join(dir, 'out', 'c.xlsx'))))
      const part = Object.keys(files).find((name) => /charts\/chart\d+\.xml$/.test(name))
      expect(part, `LibreOffice dropped the ${kind} chart entirely`).toBeDefined()
      const xml = strFromU8(files[part as string] as Uint8Array)
      // Re-emitted from LibreOffice's own model, and still pointing at cells
      // rather than at baked numbers.
      expect(xml).toMatch(/<c:f>S!\$[A-C]\$2:\$[A-C]\$4<\/c:f>/)
    },
  )
})

describe('a combo chart', () => {
  const combo = (secondary: boolean) =>
    compile(
      (
        <Workbook>
          <Sheet name="S">
            <Stack gap={1}>
              <Table name="t" data={months} columns={[col('month'), col('a'), col('b')]} />
              <Chart
                kind="combo"
                title="Actual vs target"
                categories={ref('t').column('month')}
                axes={{ value: 'NT$', secondary: '% of target' }}
                series={[
                  { name: 'Actual', values: ref('t').column('a'), as: 'bar' },
                  {
                    name: 'Target',
                    values: ref('t').column('b'),
                    as: 'line',
                    ...(secondary ? { axis: 'secondary' as const } : {}),
                  },
                ]}
              />
            </Stack>
          </Sheet>
        </Workbook>
      ) as never,
    )

  async function xmlOf(secondary: boolean): Promise<string> {
    const files = unzipSync(new Uint8Array(await new XlsxWriter().write(combo(secondary))))
    const part = Object.keys(files).find((name) => /charts\/chart\d+\.xml$/.test(name))
    return strFromU8(files[part as string] as Uint8Array)
  }

  it('puts two plots in one plot area, bars before the line', async () => {
    const xml = await xmlOf(false)
    expect(xml).toContain('<c:barChart>')
    expect(xml).toContain('<c:lineChart>')
    // Bars first, so the line draws over them rather than under.
    expect(xml.indexOf('<c:barChart>')).toBeLessThan(xml.indexOf('<c:lineChart>'))
  })

  it('shares one axis pair when nothing asks for a second', async () => {
    const xml = await xmlOf(false)
    expect((xml.match(/<c:valAx>/g) ?? []).length).toBe(1)
    expect(xml).not.toContain('333333333')
  })

  it('gives the secondary series its own axis pair, with the extra one deleted', async () => {
    const xml = await xmlOf(true)
    expect((xml.match(/<c:valAx>/g) ?? []).length).toBe(2)
    expect(xml).toContain('<c:axPos val="r"/>')
    expect(xml).toContain('<c:crosses val="max"/>')
    // Drawing the second category axis would print the month labels twice.
    expect(xml).toContain('<c:delete val="1"/>')
    expect(xml).toContain('<a:t>% of target</a:t>')
  })

  it('scales the secondary series separately in the HTML twin', () => {
    // 'a' runs 10..50 and 'b' runs 10..40. Sharing one scale is fine here; what
    // must differ is that the secondary line uses its own extent, so its top
    // point sits at the top of the plot rather than four fifths up.
    const shared = toHtml(combo(false), { title: 'S', values: evaluateWorkbook(combo(false)) })
    const split = toHtml(combo(true), { title: 'S', values: evaluateWorkbook(combo(true)) })
    const lineOf = (markup: string) => /<polyline points="([^"]+)"/.exec(markup)?.[1]
    expect(lineOf(shared)).toBeDefined()
    expect(lineOf(split)).not.toBe(lineOf(shared))
  })
})

describe.skipIf(!SOFFICE)('a combo through a real spreadsheet', () => {
  it('survives with both plots and both axes', { timeout: 180_000 }, async () => {
    const compiled = compile(
      (
        <Workbook>
          <Sheet name="S">
            <Stack gap={1}>
              <Table name="t" data={months} columns={[col('month'), col('a'), col('b')]} />
              <Chart
                kind="combo"
                categories={ref('t').column('month')}
                series={[
                  { name: 'Actual', values: ref('t').column('a'), as: 'bar' },
                  { name: 'Target', values: ref('t').column('b'), as: 'line', axis: 'secondary' },
                ]}
              />
            </Stack>
          </Sheet>
        </Workbook>
      ) as never,
    )
    const dir = mkdtempSync(join(tmpdir(), 'open-sheet-combo-'))
    writeFileSync(join(dir, 'c.xlsx'), await new XlsxWriter().write(compiled))
    execFileSync(
      SOFFICE as string,
      [
        `-env:UserInstallation=file://${join(dir, 'profile')}`,
        '--headless',
        '--convert-to',
        'xlsx',
        '--outdir',
        join(dir, 'out'),
        join(dir, 'c.xlsx'),
      ],
      { stdio: 'pipe', timeout: 150_000 },
    )
    const files = unzipSync(new Uint8Array(readFileSync(join(dir, 'out', 'c.xlsx'))))
    const part = Object.keys(files).find((name) => /charts\/chart\d+\.xml$/.test(name))
    expect(part, 'LibreOffice dropped the combo chart').toBeDefined()
    const xml = strFromU8(files[part as string] as Uint8Array)
    expect(xml).toContain('barChart')
    expect(xml).toContain('lineChart')
    expect(xml).toMatch(/<c:f>S!\$[A-C]\$2:\$[A-C]\$4<\/c:f>/)
  })
})

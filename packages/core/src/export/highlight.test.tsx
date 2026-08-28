import { existsSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { strFromU8, unzipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { compile } from '../compile/compile.js'
import { col, Sheet, Table, Workbook } from '../compile/components.js'
import type { ColumnSpec } from '../compile/nodes.js'
import { evaluateWorkbook } from '../formula/evaluate.js'
import { toHtml } from './html.js'
import { XlsxWriter } from './xlsx.js'

const data = [{ v: 10 }, { v: 50 }, { v: 90 }, { v: 50 }]

function book(options: Partial<ColumnSpec>) {
  return compile(
    (
      <Workbook>
        <Sheet name="S">
          <Table name="t" data={data} columns={[col('v', { header: 'V', ...options })]} />
        </Sheet>
      </Workbook>
    ) as never,
  )
}

function html(options: Partial<ColumnSpec>): string {
  const compiled = book(options)
  return toHtml(compiled, { title: 'S', values: evaluateWorkbook(compiled) })
}

async function sheetXml(options: Partial<ColumnSpec>): Promise<string> {
  const buffer = await new XlsxWriter().write(book(options))
  return strFromU8(unzipSync(new Uint8Array(buffer))['xl/worksheets/sheet1.xml'] as Uint8Array)
}

/** The rows of the rendered grid, so a cell can be checked by position. */
function cellsOf(markup: string): string[] {
  return markup.match(/<td[^>]*>.*?<\/td>/g) ?? []
}

describe('a colour scale', () => {
  it('uses percent, not percentile, so both renderers land on the same colour', async () => {
    // Excel's default midpoint is the median. The values here are 10/50/90/50,
    // whose median is 50 and whose linear midpoint is also 50 — but on any
    // skewed column the two differ, and the HTML could not reproduce a median
    // without re-implementing Excel's percentile. `percent` is reproducible.
    const xml = await sheetXml({ scale: ['#ff0000', '#ffffff', '#00ff00'] })
    expect(xml).toContain('type="percent" val="50"')
    expect(xml).not.toContain('percentile')
  })

  it('paints the lowest, middle and highest cells the stop colours', () => {
    const cells = cellsOf(html({ scale: ['#ff0000', '#ffffff', '#00ff00'] }))
    // header, then 10, 50, 90, 50
    expect(cells[1]).toContain('background-color:#ff0000')
    expect(cells[2]).toContain('background-color:#ffffff')
    expect(cells[3]).toContain('background-color:#00ff00')
  })

  it('paints the first stop when every value is identical', () => {
    // A flat column has no scale. Excel paints the first stop; a midpoint here
    // would be a divergence nobody would think to look for.
    const flat = compile(
      (
        <Workbook>
          <Sheet name="S">
            <Table
              name="t"
              data={[{ v: 5 }, { v: 5 }]}
              columns={[col('v', { scale: ['#ff0000', '#00ff00'] })]}
            />
          </Sheet>
        </Workbook>
      ) as never,
    )
    const markup = toHtml(flat, { title: 'S', values: evaluateWorkbook(flat) })
    expect(markup).toContain('background-color:#ff0000')
    expect(markup).not.toContain('background-color:#00ff00')
  })
})

describe('an icon set', () => {
  it('splits into the same thirds in both renderers', async () => {
    const xml = await sheetXml({ icons: 'arrows' })
    expect(xml).toContain('iconSet="3Arrows"')
    expect(xml).toContain('val="33"')
    expect(xml).toContain('val="67"')

    const cells = cellsOf(html({ icons: 'arrows' }))
    expect(cells[1]).toContain('▼')
    expect(cells[2]).toContain('▶')
    expect(cells[3]).toContain('▲')
  })
})

describe('a highlight rule', () => {
  it('writes cellIs for a comparison and styles the matching cells', async () => {
    const xml = await sheetXml({ highlight: { above: 40, fill: '#fee2e2', bold: true } })
    expect(xml).toContain('type="cellIs"')
    expect(xml).toContain('operator="greaterThan"')
    expect(xml).toContain('<formula>40</formula>')

    const cells = cellsOf(html({ highlight: { above: 40, fill: '#fee2e2', bold: true } }))
    expect(cells[1]).not.toContain('#fee2e2')
    expect(cells[2]).toContain('background-color:#fee2e2')
    expect(cells[2]).toContain('font-weight:600')
  })

  it('applies several rules in the order they were written', () => {
    const cells = cellsOf(
      html({
        highlight: [
          { above: 40, fill: '#dcfce7' },
          { above: 80, fill: '#fee2e2' },
        ],
      }),
    )
    // 50 matches the first only; 90 matches both and the later rule wins.
    expect(cells[2]).toContain('#dcfce7')
    expect(cells[3]).toContain('#fee2e2')
  })

  it('quotes a text comparand and leaves a number bare', async () => {
    const xml = await sheetXml({ highlight: { equals: 'done', fill: '#eee' } })
    expect(xml).toContain('&quot;done&quot;')
    const numeric = await sheetXml({ highlight: { equals: 50, fill: '#eee' } })
    expect(numeric).toContain('<formula>50</formula>')
  })

  it('finds duplicates with COUNTIF, since ExcelJS drops the built-in rule', async () => {
    // A `duplicateValues` rule is silently written as no element at all.
    const xml = await sheetXml({ highlight: { duplicates: true, fill: '#fee2e2' } })
    expect(xml).toContain('type="expression"')
    expect(xml).toContain('COUNTIF($A$2:$A$5,A2)&gt;1')

    const cells = cellsOf(html({ highlight: { duplicates: true, fill: '#fee2e2' } }))
    expect(cells[1]).not.toContain('#fee2e2')
    expect(cells[2]).toContain('#fee2e2')
    expect(cells[4]).toContain('#fee2e2')
  })

  it('ranks for top and bottom', async () => {
    const top = await sheetXml({ highlight: { top: 2, bold: true } })
    expect(top).toContain('type="top10"')
    expect(top).toContain('rank="2"')

    const cells = cellsOf(html({ highlight: { top: 2, bold: true } }))
    expect(cells[1]).not.toContain('font-weight:600')
    expect(cells[3]).toContain('font-weight:600')
  })
})

describe('a bar and a rule on the same column', () => {
  it('keeps both, since one is a background image and the other a colour', () => {
    const cells = cellsOf(html({ bar: true, highlight: { above: 80, fill: '#fee2e2' } }))
    expect(cells[3]).toContain('background-color:#fee2e2')
    expect(cells[3]).toContain('linear-gradient')
  })
})

const SOFFICE = ['/opt/homebrew/bin/soffice', '/usr/bin/soffice', '/usr/bin/libreoffice'].find(
  (path) => existsSync(path),
)

describe.skipIf(!SOFFICE)('and a real spreadsheet keeps the rules', () => {
  it('reads every kind back out of its own model', { timeout: 180_000 }, async () => {
    const { execFileSync } = await import('node:child_process')
    const compiled = compile(
      (
        <Workbook>
          <Sheet name="S">
            <Table
              name="t"
              data={data}
              columns={[
                col('v', { header: 'V', scale: ['#fee2e2', '#ffffff', '#dcfce7'] }),
                col('w', { header: 'W', icons: 'trafficLights', value: (row) => row.v }),
                col('x', {
                  header: 'X',
                  value: (row) => row.v,
                  highlight: [
                    { above: 40, fill: '#fee2e2' },
                    { duplicates: true, bold: true },
                  ],
                }),
              ]}
            />
          </Sheet>
        </Workbook>
      ) as never,
    )
    const dir = mkdtempSync(join(tmpdir(), 'open-sheet-cf-'))
    writeFileSync(join(dir, 'cf.xlsx'), await new XlsxWriter().write(compiled))
    execFileSync(
      SOFFICE as string,
      [
        `-env:UserInstallation=file://${join(dir, 'profile')}`,
        '--headless',
        '--convert-to',
        'xlsx',
        '--outdir',
        join(dir, 'out'),
        join(dir, 'cf.xlsx'),
      ],
      { stdio: 'pipe', timeout: 150_000 },
    )
    const xml = strFromU8(
      unzipSync(new Uint8Array(readFileSync(join(dir, 'out', 'cf.xlsx'))))[
        'xl/worksheets/sheet1.xml'
      ] as Uint8Array,
    )
    expect(xml).toContain('colorScale')
    expect(xml).toContain('iconSet')
    expect(xml).toContain('cellIs')
    expect(xml).toContain('COUNTIF')
  })
})

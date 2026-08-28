import { existsSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { strFromU8, unzipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { compile } from '../compile/compile.js'
import { col, Sheet, Table, Workbook } from '../compile/components.js'
import { evaluateWorkbook } from '../formula/evaluate.js'
import { toHtml } from './html.js'
import { XlsxWriter } from './xlsx.js'

const rows = [
  { name: 'A', q1: 10, q2: 30, q3: 20, q4: 50 },
  { name: 'B', q1: 40, q2: 20, q3: 30, q4: 10 },
]

function book(sparkline: Record<string, unknown>, order = ['q1', 'q2', 'q3', 'q4']) {
  return compile(
    (
      <Workbook>
        <Sheet name="S">
          <Table
            name="t"
            data={rows}
            columns={[
              col('name', { header: 'Name' }),
              ...order.map((key) => col(key, { header: key.toUpperCase() })),
              col('trend', { header: 'Trend', sparkline: sparkline as never }),
            ]}
          />
        </Sheet>
      </Workbook>
    ) as never,
  )
}

async function sheetXml(sparkline: Record<string, unknown>): Promise<string> {
  const buffer = await new XlsxWriter().write(book(sparkline))
  return strFromU8(unzipSync(new Uint8Array(buffer))['xl/worksheets/sheet1.xml'] as Uint8Array)
}

describe('an in-cell trend', () => {
  it('writes one sparkline per data row, each reading its own row', async () => {
    const xml = await sheetXml({ of: ['q1', 'q2', 'q3', 'q4'] })
    expect(xml).toContain('x14:sparklineGroup')
    expect((xml.match(/<x14:sparkline>/g) ?? []).length).toBe(2)
    expect(xml).toContain('<xm:f>S!B2:E2</xm:f>')
    expect(xml).toContain('<xm:sqref>F2</xm:sqref>')
    expect(xml).toContain('<xm:f>S!B3:E3</xm:f>')
  })

  it('puts the extension list last, where the schema requires it', async () => {
    const xml = await sheetXml({ of: ['q1', 'q2', 'q3', 'q4'] })
    expect(xml.indexOf('<extLst>')).toBeGreaterThan(xml.indexOf('<pageSetup'))
    expect(xml).toContain('</extLst></worksheet>')
  })

  it('groups by kind and colour, since the group carries both', async () => {
    const xml = await sheetXml({ of: ['q1', 'q2', 'q3', 'q4'], kind: 'column', color: '#16a34a' })
    expect((xml.match(/<x14:sparklineGroup /g) ?? []).length).toBe(1)
    expect(xml).toContain('type="column"')
    expect(xml).toContain('<x14:colorSeries rgb="FF16A34A"/>')
  })

  it('refuses columns that are not next to each other', () => {
    // One range per sparkline, so a gap would silently pull in whatever sits
    // between them — here the name column, which is not a number at all.
    expect(() =>
      compile(
        (
          <Workbook>
            <Sheet name="S">
              <Table
                name="t"
                data={rows}
                columns={[
                  col('q1'),
                  col('name'),
                  col('q2'),
                  col('trend', { sparkline: { of: ['q1', 'q2'] } }),
                ]}
              />
            </Sheet>
          </Workbook>
        ) as never,
      ),
    ).toThrow(/not next to each other/)
  })

  it('refuses a column it cannot find, naming what is there', () => {
    expect(() => book({ of: ['q1', 'nope'] })).toThrow(/not a column of table "t"/)
  })

  it('refuses a trend of one point', () => {
    expect(() => book({ of: ['q1'] })).toThrow(/at least two/)
  })

  it('draws the same shape in HTML', () => {
    const compiled = book({ of: ['q1', 'q2', 'q3', 'q4'] })
    const html = toHtml(compiled, { title: 'S', values: evaluateWorkbook(compiled) })
    expect((html.match(/class="os-sparkline"/g) ?? []).length).toBe(2)
    expect(html).toContain('<polyline')

    const columns = book({ of: ['q1', 'q2', 'q3', 'q4'], kind: 'column' })
    const bars = toHtml(columns, { title: 'S', values: evaluateWorkbook(columns) })
    expect(bars).toContain('os-sparkline')
    expect(bars).not.toContain('<polyline')
  })
})

const SOFFICE = ['/opt/homebrew/bin/soffice', '/usr/bin/soffice', '/usr/bin/libreoffice'].find(
  (path) => existsSync(path),
)

describe.skipIf(!SOFFICE)('through a real spreadsheet', () => {
  it('opens without complaint, whether or not it keeps them', { timeout: 180_000 }, async () => {
    const { execFileSync } = await import('node:child_process')
    const dir = mkdtempSync(join(tmpdir(), 'open-sheet-spark-'))
    writeFileSync(
      join(dir, 's.xlsx'),
      await new XlsxWriter().write(book({ of: ['q1', 'q2', 'q3', 'q4'] })),
    )
    execFileSync(
      SOFFICE as string,
      [
        `-env:UserInstallation=file://${join(dir, 'profile')}`,
        '--headless',
        '--convert-to',
        'csv:Text - txt - csv (StarCalc):44,34,76,1,,0,false,true,false,false,false,-1',
        '--outdir',
        join(dir, 'out'),
        join(dir, 's.xlsx'),
      ],
      { stdio: 'pipe', timeout: 150_000 },
    )
    // The file has to still be readable — an extLst in the wrong place makes
    // the whole worksheet unparseable, and the failure is total, not partial.
    const { readdirSync } = await import('node:fs')
    const produced = readdirSync(join(dir, 'out')).find((name) => name.endsWith('.csv'))
    expect(produced, 'LibreOffice could not open the file at all').toBeDefined()
    const csv = readFileSync(join(dir, 'out', produced as string), 'utf8')
    expect(csv).toContain('A')
    expect(csv).not.toContain('Err:')
  })
})

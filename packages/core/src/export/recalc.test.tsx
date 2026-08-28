import { execFileSync } from 'node:child_process'
import { existsSync, mkdtempSync, readdirSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { unzipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { compile } from '../compile/compile.js'
import { Chart, col, Sheet, Spill, Stack, Table, Workbook } from '../compile/components.js'
import { budget } from '../compile/fixtures.js'
import { evaluateWorkbook } from '../formula/evaluate.js'
import { sort, transpose } from '../formula/expr.js'
import { isExcelError, isNotEvaluated } from '../formula/value.js'
import { parseCellKey } from '../model/cell.js'
import { ref } from '../refs/ref.js'
import { XlsxWriter } from './xlsx.js'

function chartFixture() {
  return (
    <Workbook>
      <Sheet name="Sales">
        <Stack gap={1}>
          <Table
            name="sales"
            data={[
              { month: 'Jan', units: 120 },
              { month: 'Feb', units: 150 },
              { month: 'Mar', units: 190 },
            ]}
            columns={[col('month', { header: 'Month' }), col('units', { header: 'Units' })]}
          />
          <Chart
            kind="bar"
            title="Units by month"
            categories={ref('sales').column('month')}
            series={[{ name: 'Units', values: ref('sales').column('units') }]}
          />
        </Stack>
      </Sheet>
    </Workbook>
  )
}

const SOFFICE = ['/opt/homebrew/bin/soffice', '/usr/bin/soffice', '/usr/bin/libreoffice'].find(
  (path) => existsSync(path),
)

/**
 * Field 9 is "save cell contents as shown" — it must be false, or LibreOffice
 * writes the *formatted* value (0.6029… under a 0.0% format becomes "60.3%") and
 * the comparison measures our number formats instead of its arithmetic.
 * Field 12 is -1: export every sheet, each to its own file.
 */
const CSV_ALL_SHEETS =
  'csv:Text - txt - csv (StarCalc):44,34,76,1,,0,false,true,false,false,false,-1'

function parseCsv(text: string): string[][] {
  const rows: string[][] = []
  let row: string[] = []
  let field = ''
  let quoted = false
  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i]
    if (quoted) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"'
          i += 1
        } else quoted = false
      } else field += ch
      continue
    }
    if (ch === '"') quoted = true
    else if (ch === ',') {
      row.push(field)
      field = ''
    } else if (ch === '\n') {
      row.push(field)
      rows.push(row)
      row = []
      field = ''
    } else if (ch !== '\r') field += ch
  }
  if (field.length > 0 || row.length > 0) {
    row.push(field)
    rows.push(row)
  }
  return rows
}

/**
 * The check the whole project rests on: LibreOffice opens the exported workbook,
 * recalculates the formulas we wrote, and its numbers must match the ones our
 * own evaluator produced. A failure means either the export baked values instead
 * of formulas, or our evaluator disagrees with a real spreadsheet engine.
 */
describe.skipIf(!SOFFICE)('cross-engine recalculation', () => {
  it('LibreOffice agrees with our evaluator', { timeout: 180_000 }, async () => {
    const book = compile(budget())
    const values = evaluateWorkbook(book)

    // cacheValues: false on purpose. LibreOffice's default for xlsx is "never
    // recalculate on load" and it ignores fullCalcOnLoad here, so with cached
    // results present it would report the numbers we put in the file — this
    // check would pass even if serialize() emitted nonsense.
    const buffer = await new XlsxWriter().write(book, { values, cacheValues: false })

    const dir = mkdtempSync(join(tmpdir(), 'open-sheet-recalc-'))
    const xlsx = join(dir, 'fixture.xlsx')
    writeFileSync(xlsx, buffer)

    execFileSync(
      SOFFICE as string,
      [
        `-env:UserInstallation=file://${join(dir, 'profile')}`,
        '--headless',
        '--convert-to',
        CSV_ALL_SHEETS,
        '--outdir',
        dir,
        xlsx,
      ],
      { stdio: 'pipe', timeout: 150_000 },
    )

    const produced = readdirSync(dir).filter((f) => f.endsWith('.csv'))
    expect(produced.length).toBeGreaterThan(0)

    const bySheet = new Map<string, string[][]>()
    for (const file of produced) {
      const name = file.replace(/^fixture-?/, '').replace(/\.csv$/, '')
      bySheet.set(name, parseCsv(readFileSync(join(dir, file), 'utf8')))
    }

    let compared = 0
    const skipped: string[] = []

    for (const sheet of book.sheets) {
      const grid =
        bySheet.get(sheet.name) ?? (bySheet.size === 1 ? [...bySheet.values()][0] : undefined)
      if (!grid) {
        skipped.push(`whole sheet ${sheet.name} (no csv produced)`)
        continue
      }
      for (const [key, cell] of sheet.cells) {
        if (!cell.expr) continue
        const ours = values.get(`${sheet.name}!${key}`)
        if (ours === undefined || isNotEvaluated(ours)) {
          skipped.push(`${sheet.name}!${key} (#NOT_EVALUATED)`)
          continue
        }
        if (isExcelError(ours) || typeof ours !== 'number') continue

        const { r, c } = parseCellKey(key)
        const raw = grid[r]?.[c]
        expect(raw, `${sheet.name} r${r} c${c} missing from LibreOffice output`).toBeDefined()

        const theirs = Number(String(raw).replace(/[,%\s]/g, ''))
        expect(Number.isNaN(theirs), `${sheet.name} r${r} c${c} not numeric: "${raw}"`).toBe(false)

        const scale = String(raw).includes('%') ? 100 : 1
        expect(theirs / scale, `${sheet.name} r${r} c${c}`).toBeCloseTo(ours, 6)
        compared += 1
      }
    }

    expect(compared, 'no formula cells were compared — the check proved nothing').toBeGreaterThan(5)
    if (skipped.length)
      console.info(`recalc: skipped ${skipped.length} cell(s):`, skipped.join(', '))
  })
})

/**
 * The product claim, tested directly: change one assumption and the numbers that
 * depend on it move. Nothing here reads a value open-sheet computed — LibreOffice
 * is the only thing doing arithmetic.
 */
describe.skipIf(!SOFFICE)('the exported workbook is a live model', () => {
  it('recalculates the whole P&L when one assumption changes', { timeout: 180_000 }, async () => {
    const dir = mkdtempSync(join(tmpdir(), 'open-sheet-live-'))

    const netIncomeFor = async (taxRate: number): Promise<number[]> => {
      const book = compile(budget())
      const assumptions = book.sheets[0]
      const cell = assumptions?.cells.get('2,1')
      expect(cell, 'fixture should have a taxRate assumption cell').toBeDefined()
      ;(cell as { value?: number }).value = taxRate

      const name = `tax-${String(taxRate).replace('.', '_')}`
      const file = join(dir, `${name}.xlsx`)
      writeFileSync(file, await new XlsxWriter().write(book, { cacheValues: false }))

      execFileSync(
        SOFFICE as string,
        [
          `-env:UserInstallation=file://${join(dir, 'profile')}`,
          '--headless',
          '--convert-to',
          CSV_ALL_SHEETS,
          '--outdir',
          join(dir, name),
          file,
        ],
        { stdio: 'pipe', timeout: 150_000 },
      )

      const out = readdirSync(join(dir, name)).find((f) => f.includes('P') && f.endsWith('.csv'))
      const grid = parseCsv(readFileSync(join(dir, name, out as string), 'utf8'))
      // Net income is the 7th column; data rows start after the KPI band, gap,
      // title and header — but we locate it by the header rather than by counting.
      const headerRow = grid.findIndex((row) => row.includes('Net income'))
      const column = grid[headerRow]?.indexOf('Net income') as number
      return grid.slice(headerRow + 1, headerRow + 5).map((row) => Number(row[column]))
    }

    const base = await netIncomeFor(0.2)
    const raised = await netIncomeFor(0.35)

    expect(base.every(Number.isFinite)).toBe(true)
    expect(raised.every(Number.isFinite)).toBe(true)
    for (let i = 0; i < base.length; i += 1) {
      // net income = operating income × (1 - taxRate), so 0.35 yields 0.65/0.80 of 0.20
      expect(raised[i] as number).toBeCloseTo((base[i] as number) * (0.65 / 0.8), 4)
    }
  })
})

/**
 * Hand-written OOXML is the easiest way to produce a file that opens to an error
 * dialog, so the chart parts are checked against a real spreadsheet application
 * rather than against our own reading of the spec.
 */
describe.skipIf(!SOFFICE)('native charts survive a real spreadsheet application', () => {
  it('LibreOffice reads the chart and keeps its ranges live', { timeout: 180_000 }, async () => {
    const book = compile(chartFixture())
    const buffer = await new XlsxWriter().write(book, { values: evaluateWorkbook(book) })

    const dir = mkdtempSync(join(tmpdir(), 'open-sheet-chart-'))
    const xlsx = join(dir, 'chart.xlsx')
    writeFileSync(xlsx, buffer)

    execFileSync(
      SOFFICE as string,
      [
        `-env:UserInstallation=file://${join(dir, 'profile')}`,
        '--headless',
        '--convert-to',
        'ods',
        '--outdir',
        dir,
        xlsx,
      ],
      { stdio: 'pipe', timeout: 150_000 },
    )

    const ods = unzipSync(new Uint8Array(readFileSync(join(dir, 'chart.ods'))))
    const objects = Object.keys(ods).filter(
      (name) => name.endsWith('content.xml') && name !== 'content.xml',
    )
    expect(objects.length, 'LibreOffice should have imported a chart object').toBeGreaterThan(0)

    const chart = new TextDecoder().decode(ods[objects[0] as string] as Uint8Array)
    expect(chart, 'the chart type we asked for').toContain('chart:class="chart:bar"')
    expect(chart, 'the title we set').toContain('Units by month')
    // Bound to ranges rather than to values — change a cell and the chart moves.
    expect(chart).toMatch(/chart:values-cell-range-address="Sales\.B2:Sales\.B4"/)
  })
})

/**
 * A spill is the one construct whose formula decides how many cells it occupies,
 * which is exactly what a placement engine that owns every coordinate cannot
 * allow. We emit a legacy array formula over the declared rectangle instead — so
 * the footprint is fixed at compile time and the file format enforces it. This
 * proves a real engine fills the same cells with the same values we do.
 */
describe.skipIf(!SOFFICE)('a declared footprint is filled by a real engine', () => {
  const revenue = [
    { rep: 'Ana', revenue: 960_000 },
    { rep: 'Ben', revenue: 102_000 },
    { rep: 'Cai', revenue: 95_000 },
  ]

  function recalculate(node: unknown, values: ReturnType<typeof evaluateWorkbook>) {
    return async (buffer: Buffer) => {
      void node
      void values
      const dir = mkdtempSync(join(tmpdir(), 'open-sheet-spill-'))
      const xlsx = join(dir, 'spill.xlsx')
      writeFileSync(xlsx, buffer)
      execFileSync(
        SOFFICE as string,
        [
          `-env:UserInstallation=file://${join(dir, 'profile')}`,
          '--headless',
          '--convert-to',
          CSV_ALL_SHEETS,
          '--outdir',
          dir,
          xlsx,
        ],
        { stdio: 'pipe', timeout: 150_000 },
      )
      const csv = readdirSync(dir).find((f) => f.endsWith('.csv'))
      expect(csv, 'LibreOffice produced no output').toBeDefined()
      return parseCsv(readFileSync(join(dir, csv as string), 'utf8'))
    }
  }

  function spilled(formula: ReturnType<typeof sort>, rows: number, cols: number) {
    return compile(
      <Workbook>
        <Sheet name="Top">
          <Stack gap={1}>
            <Table
              name="reps"
              data={revenue}
              columns={[col('rep', { header: 'Rep' }), col('revenue', { header: 'Revenue' })]}
            />
            <Spill formula={formula} rows={rows} cols={cols} />
          </Stack>
        </Sheet>
      </Workbook>,
    )
  }

  function originOf(book: ReturnType<typeof compile>) {
    const sheet = book.sheets[0] as (typeof book.sheets)[number]
    const found = [...sheet.cells.entries()].find(([, cell]) => cell.spill)
    expect(found, 'no spill cell was emitted').toBeDefined()
    return parseCellKey((found as [string, unknown])[0] as string)
  }

  /**
   * TRANSPOSE rather than SORT: the dynamic-array functions are recent, and the
   * LibreOffice that CI installs from apt does not implement them — it returns
   * #NAME? where a current one computes the answer. TRANSPOSE has been in every
   * version, so this proves the footprint mechanism itself on any engine.
   */
  it('fills the whole declared range, on any engine', { timeout: 180_000 }, async () => {
    const book = spilled(transpose(ref('reps').column('revenue')) as never, 1, 3)
    const values = evaluateWorkbook(book)
    const { r, c } = originOf(book)
    const ours = [0, 1, 2].map((i) => values.get(`Top!${r},${c + i}`))
    expect(ours).toEqual([960_000, 102_000, 95_000])

    const buffer = await new XlsxWriter().write(book, { values, cacheValues: false })
    const grid = await recalculate(book, values)(buffer)
    expect([0, 1, 2].map((i) => grid[r]?.[c + i])).toEqual(['960000', '102000', '95000'])
  })

  /**
   * The values SORT produces, where the engine has SORT at all. Its absence is
   * the harness's third outcome — "this engine cannot compute it" — which is not
   * a disagreement and must not read as one.
   */
  it('sorts numerically, where the engine implements SORT', { timeout: 180_000 }, async () => {
    const book = spilled(sort(ref('reps').column('revenue'), 1, -1), 3, 1)
    const values = evaluateWorkbook(book)
    const { r, c } = originOf(book)
    // Mixed digit widths on purpose: with 300/900/500 a lexicographic sort and a
    // numeric one are indistinguishable, and that is how a string-comparing
    // SORT passed this test for a release.
    expect([0, 1, 2].map((i) => values.get(`Top!${r + i},${c}`))).toEqual([
      960_000, 102_000, 95_000,
    ])

    const buffer = await new XlsxWriter().write(book, { values, cacheValues: false })
    const grid = await recalculate(book, values)(buffer)
    const theirs = [0, 1, 2].map((i) => grid[r + i]?.[c])

    if (theirs.every((cell) => cell === '#NAME?')) {
      console.warn('this LibreOffice does not implement SORT; its values are unverified here')
      return
    }
    expect(theirs).toEqual(['960000', '102000', '95000'])
  })
})

import { execFileSync } from 'node:child_process'
import { existsSync, mkdtempSync, readdirSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { strFromU8, unzipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { compile } from '../compile/compile.js'
import { Cell, col, Sheet, Stack, Table, Workbook } from '../compile/components.js'
import { evaluateWorkbook } from '../formula/evaluate.js'
import { sum } from '../formula/expr.js'
import { toFormula } from '../formula/serialize.js'
import { ref } from '../refs/ref.js'
import { XlsxWriter } from './xlsx.js'

const costs = [
  { item: '運算', amount: 12000 },
  { item: '儲存', amount: 3400 },
]

function book(props: Record<string, unknown> = {}) {
  return compile(
    (
      <Workbook>
        <Sheet name="Costs">
          <Stack gap={1}>
            <Table
              name="costs"
              appendable
              data={costs}
              total={{ amount: 'sum' }}
              columns={[
                col('item', { header: 'Item' }),
                col('amount', { header: 'Amount', format: 'currency' }),
              ]}
              {...props}
            />
            <Cell formula={sum(ref('costs').column('amount'))} format="currency" />
          </Stack>
        </Sheet>
      </Workbook>
    ) as never,
  )
}

async function parts(compiled: ReturnType<typeof compile>) {
  const files = unzipSync(new Uint8Array(await new XlsxWriter().write(compiled)))
  const table = Object.keys(files).find((name) => /tables\/table\d+\.xml$/.test(name))
  return {
    table: table ? strFromU8(files[table] as Uint8Array) : undefined,
    sheet: strFromU8(files['xl/worksheets/sheet1.xml'] as Uint8Array),
  }
}

describe('a table the recipient can append to', () => {
  it('names the column by its header, which is what a structured reference uses', () => {
    const compiled = book()
    const formula = toFormula(sum(ref('costs').column('amount')), {
      registry: compiled.registry,
      definedNames: compiled.definedNames,
      sheet: 'Costs',
    })
    // Not costs[amount] — Excel keys the reference by the header text.
    expect(formula).toBe('=SUM(costs[Amount])')
  })

  it('leaves a plain table on A1 ranges', () => {
    const plain = compile(
      (
        <Workbook>
          <Sheet name="Costs">
            <Table name="costs" data={costs} columns={[col('item'), col('amount')]} />
          </Sheet>
        </Workbook>
      ) as never,
    )
    expect(
      toFormula(sum(ref('costs').column('amount')), {
        registry: plain.registry,
        definedNames: plain.definedNames,
        sheet: 'Costs',
      }),
    ).toBe('=SUM(B2:B3)')
  })

  it('still evaluates, since the viewer resolves the ref not the string', () => {
    const compiled = book()
    const values = evaluateWorkbook(compiled)
    const anchor = compiled.registry.get('costs')
    if (anchor?.kind !== 'table' || anchor.totalRow === undefined) throw new Error('!')
    expect(values.get(`Costs!${anchor.totalRow},${anchor.columns.get('amount')}`)).toBe(15400)
  })

  it('writes a table part covering the header, the data and the totals row', async () => {
    const { table } = await parts(book())
    expect(table, 'no table part was written').toBeDefined()
    expect(table).toContain('name="costs"')
    expect(table).toContain('ref="A1:B4"')
    expect(table).toContain('totalsRowCount="1"')
  })

  it("makes our total row the table's own, so appending cannot grow into it", async () => {
    // Outside the table, Excel takes the total row in the first time someone
    // types a row below — and the total becomes a data row of itself.
    const { table } = await parts(book())
    expect(table).toContain('totalsRowFunction="sum"')
  })

  it('keeps the styles our writer put on the cells', async () => {
    // ExcelJS's addTable writes its own row values; declared after our cells it
    // silently drops every format, note and fill we had set.
    const { sheet } = await parts(book())
    const amount = /<c r="B2"[^>]*>/.exec(sheet)?.[0] ?? ''
    expect(amount).toMatch(/s="\d+"/)
  })

  it('refuses a name Excel would reject', () => {
    expect(() =>
      compile(
        (
          <Workbook>
            <Sheet name="S">
              <Table name="my costs" appendable data={costs} columns={[col('item')]} />
            </Sheet>
          </Workbook>
        ) as never,
      ),
    ).toThrow(/no spaces/)
  })

  it('refuses a name that reads as a cell address', () => {
    expect(() =>
      compile(
        (
          <Workbook>
            <Sheet name="S">
              <Table name="AB12" appendable data={costs} columns={[col('item')]} />
            </Sheet>
          </Workbook>
        ) as never,
      ),
    ).toThrow(/cell address/)
  })

  it('refuses two columns sharing a header', () => {
    // `costs[Amount]` would be ambiguous, and Excel repairs the file by renaming
    // one of them — silently changing what the formula means.
    expect(() =>
      compile(
        (
          <Workbook>
            <Sheet name="S">
              <Table
                name="costs"
                appendable
                data={costs}
                columns={[col('item', { header: 'X' }), col('amount', { header: 'X' })]}
              />
            </Sheet>
          </Workbook>
        ) as never,
      ),
    ).toThrow(/two columns headed "X"/)
  })

  it('refuses appendable together with filter', () => {
    expect(() => book({ filter: true })).toThrow(/brings its own filter/)
  })

  it('refuses appendable without headers', () => {
    expect(() => book({ showHeader: false })).toThrow(/names its columns by their headers/)
  })
})

const SOFFICE = ['/opt/homebrew/bin/soffice', '/usr/bin/soffice', '/usr/bin/libreoffice'].find(
  (path) => existsSync(path),
)

describe.skipIf(!SOFFICE)('and a real spreadsheet computes the same-row reference', () => {
  it('agrees on a derived column, in a workbook named in Chinese', {
    timeout: 180_000,
  }, async () => {
    // The function harness only ever asked about whole-column references, so a
    // same-row reference could be wrong in every engine and nothing would say
    // so. This is the case it was blind to.
    const compiled = compile(
      (
        <Workbook>
          <Sheet name="登記表">
            <Table
              name="register"
              appendable
              total={{ 預算: 'sum' }}
              data={[
                { 預算: 100, 實際: 200 },
                { 預算: 300, 實際: 400 },
              ]}
              columns={[
                col('預算', { header: '預算' }),
                col('實際', { header: '實際' }),
                col('合計', {
                  header: '合計',
                  formula: (r) => ({
                    k: 'fn',
                    name: 'SUM',
                    args: [r.cell('預算'), r.cell('實際')],
                  }),
                }),
              ]}
            />
          </Sheet>
        </Workbook>
      ) as never,
    )
    const dir = mkdtempSync(join(tmpdir(), 'open-sheet-thisrow-'))
    writeFileSync(
      join(dir, 't.xlsx'),
      await new XlsxWriter().write(compiled, { cacheValues: false }),
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
        join(dir, 't.xlsx'),
      ],
      { stdio: 'pipe', timeout: 150_000 },
    )
    const produced = readdirSync(join(dir, 'out')).find((name) => name.endsWith('.csv'))
    expect(produced, 'LibreOffice could not open the file').toBeDefined()
    const csv = readFileSync(join(dir, 'out', produced as string), 'utf8')
    const rows = csv.split('\n').map((line) => line.split(','))
    expect(rows[1]?.[2]).toBe('300')
    expect(rows[2]?.[2]).toBe('700')
    expect(rows[3]?.[0]).toBe('400')
    expect(csv).not.toContain('Err:')
    expect(csv).not.toContain('#NAME?')
  })
})

describe.skipIf(!SOFFICE)('and a real spreadsheet computes the structured reference', () => {
  it('agrees with our evaluator', { timeout: 180_000 }, async () => {
    // The whole trade: the formula string changes in every workbook that opts
    // in, so a second engine has to agree it still means the same thing.
    const compiled = book()
    const dir = mkdtempSync(join(tmpdir(), 'open-sheet-table-'))
    writeFileSync(
      join(dir, 't.xlsx'),
      await new XlsxWriter().write(compiled, { cacheValues: false }),
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
        join(dir, 't.xlsx'),
      ],
      { stdio: 'pipe', timeout: 150_000 },
    )
    const produced = readdirSync(join(dir, 'out')).find((name) => name.endsWith('.csv'))
    expect(produced, 'LibreOffice could not open the file at all').toBeDefined()
    const csv = readFileSync(join(dir, 'out', produced as string), 'utf8')
    const rows = csv.split('\n').map((line) => line.split(','))

    // The totals row, and the standalone SUM(costs[Amount]) below the table.
    expect(rows[3]?.[1]).toBe('15400')
    expect(csv).toContain('15400')
    expect(csv).not.toContain('#NAME?')
    expect(csv).not.toContain('Err:')
  })
})

/**
 * Both of these shipped in the first cut and were found by a tester driving
 * real Excel — neither is visible from the XML alone, and neither breaks any
 * check we had. They are the reason this file asserts on the file format rather
 * than on our own model.
 */
describe('what Excel does with the file, not what we meant', () => {
  it('writes SUBTOTAL in the totals row, not the plain aggregate', async () => {
    // Given `totalsRowFunction="sum"` and a cell holding `SUM(costs[Amount])`,
    // Excel decides the row is not a totals row, drops `totalsRowCount`, and
    // swallows it as data — at which point `costs[Amount]` includes the total
    // itself and the cell is a circular reference. It still displays the cached
    // number, so the file looks correct and computes nothing.
    const compiled = book()
    const anchor = compiled.registry.get('costs')
    if (anchor?.kind !== 'table' || anchor.totalRow === undefined) throw new Error('!')
    const cell = compiled.sheets[0]?.cells.get(`${anchor.totalRow},${anchor.columns.get('amount')}`)
    expect(
      toFormula(cell?.expr as never, {
        registry: compiled.registry,
        definedNames: compiled.definedNames,
        sheet: 'Costs',
      }),
    ).toBe('=SUBTOTAL(109,costs[Amount])')
  })

  it('leaves a plain table on the plain aggregate', () => {
    const plain = compile(
      (
        <Workbook>
          <Sheet name="Costs">
            <Table
              name="costs"
              data={costs}
              total={{ amount: 'sum' }}
              columns={[col('item'), col('amount')]}
            />
          </Sheet>
        </Workbook>
      ) as never,
    )
    const anchor = plain.registry.get('costs')
    if (anchor?.kind !== 'table' || anchor.totalRow === undefined) throw new Error('!')
    const cell = plain.sheets[0]?.cells.get(`${anchor.totalRow},${anchor.columns.get('amount')}`)
    expect(
      toFormula(cell?.expr as never, {
        registry: plain.registry,
        definedNames: plain.definedNames,
        sheet: 'Costs',
      }),
    ).toBe('=SUM(B2:B3)')
  })

  it('writes a same-row reference in the form a file may contain', async () => {
    const { sheet } = await parts(derived())
    // `[@Q1]` is the shorthand Excel shows in the formula bar, not a storable
    // form: written into <f> it reads back as #REF! and every derived column in
    // the workbook is broken the moment it opens. Found by a tester driving
    // real Excel; no check here or in LibreOffice saw it.
    expect(sheet).toContain('SUM(costs[[#This Row],[Q1]],costs[[#This Row],[Q2]])')
    expect(sheet).not.toContain('[@')
  })

  it('carries the column formula so an appended row is not left blank', async () => {
    // Without <calculatedColumnFormula> Excel treats a derived column as
    // ordinary cell formulas: the table grows, the ranges follow, and the new
    // row's computed cells come out empty for the reader to fill in by hand.
    const { table } = await parts(derived())
    expect(table).toContain(
      '<calculatedColumnFormula>SUM(costs[[#This Row],[Q1]],costs[[#This Row],[Q2]])</calculatedColumnFormula>',
    )
  })

  it('records the columns that cannot fill down, so the CLI can say so', () => {
    // `r.prev()` reads another row, so no single stored formula serves every
    // row. Silence would leave the docs promising something untrue. It is a
    // property of the compiled model, not of writing the file — so the CLI, the
    // viewer and the writer all read the same decision.
    const anchor = derived().registry.get('costs')
    if (anchor?.kind !== 'table') throw new Error('!')
    expect(anchor.table?.noFillDown).toEqual(['Delta'])
  })

  it('carries the fact that a total row changes how a reader adds one', () => {
    // Measured in Excel, not inferred: with a totals row present, typing below
    // it does not extend the table and neither does Tab from the last cell.
    // Only inserting above the total does. `open-sheet build` turns this pair —
    // appendable plus a total row — into a note telling the author to say so,
    // because the reader types before they read anything.
    const anchor = book().registry.get('costs')
    if (anchor?.kind !== 'table') throw new Error('!')
    expect(anchor.table).toBeDefined()
    expect(anchor.totalRow).toBeDefined()
  })

  it('omits the column formula for one that cannot fill down', async () => {
    const { table } = await parts(derived())
    expect(table).toContain('name="Delta"')
    expect(table).not.toContain('Delta"><calculatedColumnFormula>')
  })

  it('drops the totals attributes on a table that has no totals row', async () => {
    const { table } = await parts(
      compile(
        (
          <Workbook>
            <Sheet name="Costs">
              <Table
                name="costs"
                appendable
                data={costs}
                columns={[col('item', { header: 'Item' }), col('amount', { header: 'Amount' })]}
              />
            </Sheet>
          </Workbook>
        ) as never,
      ),
    )
    expect(table).not.toContain('totalsRowFunction')
    expect(table).not.toContain('totalsRowLabel')
  })
})

function derived() {
  return compile(
    (
      <Workbook>
        <Sheet name="Costs">
          <Table
            name="costs"
            appendable
            data={[
              { q1: 1, q2: 2 },
              { q1: 3, q2: 4 },
            ]}
            columns={[
              col('q1', { header: 'Q1' }),
              col('q2', { header: 'Q2' }),
              col('sum', {
                header: 'Sum',
                formula: (r) => ({ k: 'fn', name: 'SUM', args: [r.cell('q1'), r.cell('q2')] }),
              }),
              col('delta', {
                header: 'Delta',
                formula: (r) =>
                  r.isFirst ? null : { k: 'op', op: '-', l: r.cell('q1'), r: r.prev().cell('q1') },
              }),
            ]}
          />
        </Sheet>
      </Workbook>
    ) as never,
  )
}

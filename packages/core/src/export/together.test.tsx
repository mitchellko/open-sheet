import { execFileSync } from 'node:child_process'
import { existsSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { strFromU8, unzipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { compile } from '../compile/compile.js'
import { Cell, col, Sheet, Spill, Stack, Table, Workbook } from '../compile/components.js'
import { evaluateWorkbook } from '../formula/evaluate.js'
import { sort } from '../formula/expr.js'
import { ref } from '../refs/ref.js'
import { XlsxWriter } from './xlsx.js'

const SOFFICE = ['/opt/homebrew/bin/soffice', '/usr/bin/soffice', '/usr/bin/libreoffice'].find(
  (path) => existsSync(path),
)

const requests = [
  { item: '筆電', amount: 45000, status: '待審' },
  { item: '螢幕', amount: 12000, status: '核准' },
  { item: '鍵盤', amount: 3000, status: '核准' },
]

/**
 * Everything M8 added, on one sheet, plus a spill. Each was verified alone;
 * this is the only test that says they do not interfere — a protected sheet
 * that blocks its own filter, or a validation range that lands on a spill
 * footprint, would pass every other test in this repo.
 */
function workbook() {
  return compile(
    (
      <Workbook>
        <Sheet name="Lists">
          <Table
            name="statuses"
            showHeader={false}
            data={[{ name: '待審' }, { name: '核准' }, { name: '退回' }]}
            columns={[col('name')]}
          />
        </Sheet>
        <Sheet name="Requests" protect={{ allow: ['requests'] }}>
          <Stack gap={1}>
            <Table
              name="requests"
              filter
              data={requests}
              total={{ amount: 'sum' }}
              columns={[
                col('item', { header: '項目' }),
                col('amount', {
                  header: '金額',
                  format: 'currency',
                  note: (row) => (row.amount > 40000 ? '超過部門權限，需副總簽核' : undefined),
                }),
                col('status', {
                  header: '狀態',
                  validate: {
                    list: ref('statuses').column('name'),
                    prompt: '從清單挑一個',
                    error: '請從下拉選單選擇',
                  },
                }),
              ]}
            />
            <Cell value="金額由高到低" style="note" />
            <Spill formula={sort(ref('requests').column('amount'), 1, -1)} rows={3} cols={1} />
          </Stack>
        </Sheet>
      </Workbook>
    ) as never,
  )
}

describe('all of it on one sheet', () => {
  it('compiles without any two features landing on the same cell', () => {
    const book = workbook()
    const sheet = book.sheets[1] as (typeof book.sheets)[number]
    const spillOrigin = [...sheet.cells.entries()].find(([, cell]) => cell.spill)
    expect(spillOrigin).toBeDefined()
    // The spill must not have been placed over the validated column.
    const validated = [...sheet.cells.values()].filter((cell) => cell.validate)
    expect(validated).toHaveLength(3)
    expect(validated.every((cell) => cell.spill === undefined)).toBe(true)
  })

  it('keeps the total on SUBTOTAL and the filter clear of it', async () => {
    const book = workbook()
    const values = evaluateWorkbook(book)
    const anchor = book.registry.get('requests')
    if (anchor?.kind !== 'table' || anchor.totalRow === undefined) throw new Error('!')
    expect(values.get(`Requests!${anchor.totalRow},${anchor.columns.get('amount')}`)).toBe(60000)

    const buffer = await new XlsxWriter().write(book, { values })
    const xml = strFromU8(
      unzipSync(new Uint8Array(buffer))['xl/worksheets/sheet2.xml'] as Uint8Array,
    )
    // Filter covers header + 3 data rows and stops above the total on row 5.
    expect(xml).toContain('<autoFilter ref="A1:C4"/>')
    expect(xml).toContain('<sheetProtection')
    expect(xml).toContain('<dataValidation')
  })

  it('leaves the filter usable on the protected sheet', async () => {
    // Protection blocks sorting and filtering by default, which would take away
    // the arrows the same sheet just asked for.
    const buffer = await new XlsxWriter().write(workbook())
    const xml = strFromU8(
      unzipSync(new Uint8Array(buffer))['xl/worksheets/sheet2.xml'] as Uint8Array,
    )
    const protection = /<sheetProtection[^>]*>/.exec(xml)?.[0] ?? ''
    expect(protection).toContain('autoFilter="0"')
    expect(protection).toContain('sort="0"')
  })
})

describe.skipIf(!SOFFICE)('and a real spreadsheet opens the result', () => {
  it('recalculates it without complaint', { timeout: 180_000 }, async () => {
    const book = workbook()
    const values = evaluateWorkbook(book)
    const dir = mkdtempSync(join(tmpdir(), 'open-sheet-together-'))
    const xlsx = join(dir, 'all.xlsx')
    writeFileSync(xlsx, await new XlsxWriter().write(book, { values, cacheValues: false }))

    execFileSync(
      SOFFICE as string,
      [
        `-env:UserInstallation=file://${join(dir, 'profile')}`,
        '--headless',
        '--convert-to',
        'csv:Text - txt - csv (StarCalc):44,34,76,1,,0,false,true,false,false,false,-1',
        '--outdir',
        dir,
        xlsx,
      ],
      { stdio: 'pipe', timeout: 150_000 },
    )

    const csv = readFileSync(join(dir, 'all-Requests.csv'), 'utf8')
    const rows = csv.split('\n').map((line) => line.split(','))
    // The total, recalculated from SUBTOTAL by LibreOffice itself.
    expect(rows[4]?.[1]).toBe('60000')
    // And the spill, sorted descending, below the note row.
    expect(csv).toContain('45000')
    expect(csv).not.toContain('#NAME?')
    expect(csv).not.toContain('Err:')
  })
})

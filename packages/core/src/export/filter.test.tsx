import { strFromU8, unzipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { compile } from '../compile/compile.js'
import { col, Sheet, Stack, Table, Workbook } from '../compile/components.js'
import { evaluateWorkbook } from '../formula/evaluate.js'
import { toFormula } from '../formula/serialize.js'
import { XlsxWriter } from './xlsx.js'

const costs = [
  { item: 'A', amount: 100 },
  { item: 'B', amount: 250 },
  { item: 'C', amount: 40 },
]

const table = (filter: boolean, title?: string) => (
  <Workbook>
    <Sheet name="Costs">
      <Table
        name="costs"
        filter={filter}
        {...(title === undefined ? {} : { title })}
        data={costs}
        columns={[col('item', { header: 'Item' }), col('amount', { header: 'Amount' })]}
        total={{ amount: 'sum' }}
      />
    </Sheet>
  </Workbook>
)

function totalFormula(node: unknown): string {
  const book = compile(node as never)
  const sheet = book.sheets[0] as (typeof book.sheets)[number]
  const anchor = book.registry.get('costs')
  if (anchor?.kind !== 'table' || anchor.totalRow === undefined) {
    throw new Error('no total row')
  }
  const column = anchor.columns.get('amount')
  const cell = sheet.cells.get(`${anchor.totalRow},${column}`)
  return toFormula(cell?.expr as never, {
    registry: book.registry,
    definedNames: book.definedNames,
    sheet: 'Costs',
  })
}

describe('filters, and the total that has to agree with them', () => {
  it('totals with SUBTOTAL so hiding a row changes the total', () => {
    // A plain SUM keeps counting rows the reader can no longer see. The number
    // then disagrees with the rows above it and nothing on screen says why.
    expect(totalFormula(table(true))).toBe('=SUBTOTAL(109,B2:B4)')
  })

  it('leaves an unfiltered table on the plain aggregate', () => {
    expect(totalFormula(table(false))).toBe('=SUM(B2:B4)')
  })

  it('means the same thing as the plain aggregate when nothing is hidden', () => {
    const book = compile(table(true) as never)
    const values = evaluateWorkbook(book)
    const anchor = book.registry.get('costs')
    if (anchor?.kind !== 'table' || anchor.totalRow === undefined) throw new Error('!')
    const column = anchor.columns.get('amount')
    expect(values.get(`Costs!${anchor.totalRow},${column}`)).toBe(390)
  })

  it('puts the arrows on the header row and stops before the total', async () => {
    // If the total row were inside the filter range, Excel would treat it as
    // data and hide the total the first time anyone filtered.
    const buffer = await new XlsxWriter().write(compile(table(true) as never))
    const xml = strFromU8(
      unzipSync(new Uint8Array(buffer))['xl/worksheets/sheet1.xml'] as Uint8Array,
    )
    expect(xml).toContain('<autoFilter ref="A1:B4"/>')
  })

  it('starts below a title, since the arrows belong on the headers', async () => {
    const buffer = await new XlsxWriter().write(compile(table(true, 'Q3 costs') as never))
    const xml = strFromU8(
      unzipSync(new Uint8Array(buffer))['xl/worksheets/sheet1.xml'] as Uint8Array,
    )
    expect(xml).toContain('<autoFilter ref="A2:B5"/>')
  })

  it('refuses a second filtered table rather than silently dropping one', async () => {
    const two = (
      <Workbook>
        <Sheet name="Both">
          <Stack gap={1}>
            <Table name="a" filter data={costs} columns={[col('item'), col('amount')]} />
            <Table name="b" filter data={costs} columns={[col('item'), col('amount')]} />
          </Stack>
        </Sheet>
      </Workbook>
    )
    await expect(new XlsxWriter().write(compile(two as never))).rejects.toThrow(
      /one filter per sheet/,
    )
  })

  it('refuses a filter with no header row to put it on', () => {
    expect(() =>
      compile(
        (
          <Workbook>
            <Sheet name="S">
              <Table name="t" filter showHeader={false} data={costs} columns={[col('item')]} />
            </Sheet>
          </Workbook>
        ) as never,
      ),
    ).toThrow(/nowhere to put them/)
  })
})

import { strFromU8, unzipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { compile } from '../compile/compile.js'
import { col, Sheet, Stack, Table, Workbook } from '../compile/components.js'
import { ref } from '../refs/ref.js'
import { XlsxWriter } from './xlsx.js'

const model = (protect?: { allow: string[] }) => (
  <Workbook>
    <Sheet name="Model" {...(protect ? { protect } : {})}>
      <Stack gap={1}>
        <Table
          name="assumptions"
          kind="keyValue"
          data={[{ key: 'growth', label: 'Growth', value: 0.08, format: 'percent' }]}
        />
        <Table
          name="pl"
          data={[{ revenue: 100 }, { revenue: 120 }]}
          columns={[
            col('revenue', { header: 'Revenue' }),
            col('grown', {
              header: 'Grown',
              formula: (r) => ({
                k: 'op',
                op: '*',
                l: r.cell('revenue'),
                r: { k: 'ref', target: ref('assumptions').get('growth') },
              }),
            }),
          ]}
        />
      </Stack>
    </Sheet>
  </Workbook>
)

function cells(node: unknown) {
  const book = compile(node as never)
  return { book, sheet: book.sheets[0] as (typeof book.sheets)[number] }
}

describe('locking the formulas and leaving the inputs open', () => {
  it('unlocks the named block and nothing else', () => {
    const { book, sheet } = cells(model({ allow: ['assumptions'] }))
    const assumptions = book.registry.get('assumptions')
    if (assumptions?.kind !== 'keyValue') throw new Error('!')
    const growth = assumptions.keys.get('growth')
    expect(sheet.cells.get(`${growth?.r},${growth?.c}`)?.unlocked).toBe(true)

    const pl = book.registry.get('pl')
    if (pl?.kind !== 'table') throw new Error('!')
    const revenue = pl.columns.get('revenue')
    expect(sheet.cells.get(`${pl.firstDataRow},${revenue}`)?.unlocked).toBeUndefined()
  })

  it('leaves a formula locked even inside an allowed block', () => {
    // A derived cell is not an input. Unlocking it would let a reader type over
    // the very thing the protection exists to keep.
    const { book, sheet } = cells(model({ allow: ['pl'] }))
    const pl = book.registry.get('pl')
    if (pl?.kind !== 'table') throw new Error('!')
    expect(sheet.cells.get(`${pl.firstDataRow},${pl.columns.get('revenue')}`)?.unlocked).toBe(true)
    expect(
      sheet.cells.get(`${pl.firstDataRow},${pl.columns.get('grown')}`)?.unlocked,
    ).toBeUndefined()
  })

  it('leaves the header locked — it is a label, not an input', () => {
    // Unlocking it invites a reader to rename the column every formula in the
    // sheet refers to, which is not the edit protection was meant to permit.
    const { book, sheet } = cells(model({ allow: ['pl'] }))
    const pl = book.registry.get('pl')
    if (pl?.kind !== 'table') throw new Error('!')
    expect(pl.headerRow).toBeDefined()
    expect(
      sheet.cells.get(`${pl.headerRow},${pl.columns.get('revenue')}`)?.unlocked,
    ).toBeUndefined()
  })

  it('is off unless asked for', () => {
    const { sheet } = cells(model())
    expect(sheet.protect).toBeUndefined()
    expect([...sheet.cells.values()].some((cell) => cell.unlocked)).toBe(false)
  })

  it('names the sheet a block actually lives on', () => {
    expect(() =>
      compile(
        (
          <Workbook>
            <Sheet name="A">
              <Table name="inputs" kind="keyValue" data={[{ key: 'x', label: 'X', value: 1 }]} />
            </Sheet>
            <Sheet name="B" protect={{ allow: ['inputs'] }}>
              <Table name="t" data={[{ v: 1 }]} columns={[col('v')]} />
            </Sheet>
          </Workbook>
        ) as never,
      ),
    ).toThrow(/on sheet "A"/)
  })

  it('writes the protection and the unlocked style into the file', async () => {
    const buffer = await new XlsxWriter().write(compile(model({ allow: ['assumptions'] }) as never))
    const files = unzipSync(new Uint8Array(buffer))
    const xml = strFromU8(files['xl/worksheets/sheet1.xml'] as Uint8Array)
    expect(xml).toContain('<sheetProtection')
    expect(strFromU8(files['xl/styles.xml'] as Uint8Array)).toContain('<protection locked="0"/>')
  })
})

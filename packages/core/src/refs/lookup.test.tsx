import { describe, expect, it } from 'vitest'
import { compile } from '../compile/compile.js'
import { col, Sheet, Stack, Table, Workbook } from '../compile/components.js'
import { evaluateWorkbook } from '../formula/evaluate.js'
import { toFormula } from '../formula/serialize.js'
import { display } from '../formula/value.js'
import { lookup } from './lookup.js'

const book = (ifMissing?: string) =>
  compile(
    <Workbook>
      <Sheet name="S">
        <Stack gap={1}>
          <Table
            name="products"
            data={[
              { sku: 'a', name: 'apple', price: 10 },
              { sku: 'b', name: 'pear', price: 20 },
            ]}
            columns={[col('sku'), col('name'), col('price')]}
          />
          <Table
            name="orders"
            data={[{ sku: 'b' }, { sku: 'zz' }]}
            columns={[
              col('sku'),
              col('price', {
                formula: (r) =>
                  lookup({
                    value: r.cell('sku'),
                    from: 'products',
                    match: 'sku',
                    get: 'price',
                    ...(ifMissing === undefined ? {} : { ifMissing }),
                  }),
              }),
            ]}
          />
        </Stack>
      </Sheet>
    </Workbook>,
  )

describe('lookup', () => {
  it('compiles to INDEX/MATCH over named columns, not a counted index', () => {
    // VLOOKUP would carry a positional column number, and inserting a column in
    // the lookup table would silently repoint it — the failure this framework
    // exists to remove.
    const compiled = book()
    const formula = toFormula(compiled.sheets[0]?.cells.get('5,1')?.expr as never, {
      registry: compiled.registry,
      definedNames: compiled.definedNames,
      sheet: 'S',
    })
    expect(formula).toBe('=INDEX(C2:C3,MATCH(A6,A2:A3,0))')
    expect(formula).not.toContain('VLOOKUP')
  })

  it('finds the matching row', () => {
    expect(evaluateWorkbook(book()).get('S!5,1')).toBe(20)
  })

  it('reports #N/A when nothing matches, as Excel does', () => {
    expect(display(evaluateWorkbook(book()).get('S!6,1') as never)).toBe('#N/A')
  })

  it('uses ifMissing when given one', () => {
    expect(evaluateWorkbook(book('—')).get('S!6,1')).toBe('—')
    // …and still finds the row that does match
    expect(evaluateWorkbook(book('—')).get('S!5,1')).toBe(20)
  })

  it('moves with the table', () => {
    // The whole point: both ranges resolve after layout.
    const grown = compile(
      <Workbook>
        <Sheet name="S">
          <Stack gap={1}>
            <Table name="pad" data={[{ x: 1 }, { x: 2 }]} columns={[col('x')]} />
            <Table
              name="products"
              data={[{ sku: 'a', price: 10 }]}
              columns={[col('sku'), col('price')]}
            />
            <Table
              name="orders"
              data={[{ sku: 'a' }]}
              columns={[
                col('sku'),
                col('price', {
                  formula: (r) =>
                    lookup({ value: r.cell('sku'), from: 'products', match: 'sku', get: 'price' }),
                }),
              ]}
            />
          </Stack>
        </Sheet>
      </Workbook>,
    )
    const formula = toFormula(grown.sheets[0]?.cells.get('8,1')?.expr as never, {
      registry: grown.registry,
      definedNames: grown.definedNames,
      sheet: 'S',
    })
    // Pushed down by the padding table; nothing in the source changed.
    // A one-row range collapses to a single address, which is correct.
    expect(formula).toBe('=INDEX(B6,MATCH(A9,A6,0))')
  })
})

describe('errors the library reports are Excel errors', () => {
  it('does not mistake #N/A for something we failed to compute', () => {
    // The library returns Error objects whose message is the code. Treating
    // those as not-evaluated hid a real condition and made IFNA unable to catch
    // the #N/A a failed MATCH exists to produce.
    const values = evaluateWorkbook(book())
    expect(display(values.get('S!6,1') as never)).toBe('#N/A')
    expect(display(values.get('S!6,1') as never)).not.toBe('#NOT_EVALUATED')
  })
})

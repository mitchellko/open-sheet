import { describe, expect, it } from 'vitest'
import { evaluateWorkbook } from '../formula/evaluate.js'
import { toFormula } from '../formula/serialize.js'
import { cellKey } from '../model/cell.js'
import { ref } from '../refs/ref.js'
import { compile } from './compile.js'
import { col, KpiBand, Row, Sheet, Stack, Table, Workbook } from './components.js'
import { budget, QUARTERS, type Quarter, sideBySide } from './fixtures.js'
import type { TableAnchor } from './registry.js'

function at(sheet: { cells: Map<string, unknown> }, r: number, c: number) {
  return sheet.cells.get(cellKey(r, c))
}

describe('the JSX pipeline', () => {
  it('compiles a workbook without touching a DOM', () => {
    const book = compile(budget())
    expect(book.sheets.map((s) => s.name)).toEqual(['Assumptions', 'P&L'])
  })

  it('flattens fragments, conditionals, and inline markup', () => {
    const book = compile(budget())
    const pl = book.sheets[1]!
    const note = at(pl, 10, 0) as { value: string }
    expect(note.value).toBe('Forecast beyond Q4 uses growth from Assumptions.')
  })

  it('lays a table out under its KPI band with the requested gap', () => {
    const pl = compile(budget()).sheets[1]!
    expect(at(pl, 0, 0)).toMatchObject({ value: 'FY Revenue' })
    expect(at(pl, 1, 0)).toMatchObject({ style: 'kpiValue' })
    expect(at(pl, 3, 0)).toMatchObject({ value: 'Quarter' })
    expect(at(pl, 4, 0)).toMatchObject({ value: 'Q1' })
  })

  it('records where each named block landed', () => {
    const table = compile(budget()).registry.get('pl') as TableAnchor
    expect(table.headerRow).toBe(3)
    expect(table.firstDataRow).toBe(4)
    expect(table.lastDataRow).toBe(7)
    expect(table.totalRow).toBe(8)
    expect([...table.columns.keys()]).toEqual([
      'quarter',
      'revenue',
      'cogs',
      'grossProfit',
      'netIncome',
      'qoq',
    ])
  })

  it('registers key-value entries as defined names', () => {
    const book = compile(budget())
    expect(book.definedNames.get('growth')).toMatchObject({
      sheet: 'Assumptions',
      addr: { r: 1, c: 1 },
    })
    expect(book.definedNames.get('taxRate')).toMatchObject({
      sheet: 'Assumptions',
      addr: { r: 2, c: 1 },
    })
  })

  it('leaves a formula cell empty when the column returns null', () => {
    const pl = compile(budget()).sheets[1]!
    expect(at(pl, 4, 5)).toEqual({})
    expect(at(pl, 5, 5)).toMatchObject({ expr: { k: 'op', op: '-' } })
  })

  it('places a Row side by side and honours Spacer', () => {
    const sheet = compile(sideBySide()).sheets[0]!
    expect(at(sheet, 0, 0)).toMatchObject({ value: 'left' })
    expect(at(sheet, 0, 3)).toMatchObject({ value: 'A' })
    expect(at(sheet, 7, 0)).toMatchObject({ value: 'below' })
  })

  it('rejects a workbook with duplicate block names', () => {
    const dup = {
      kind: 'workbook',
      children: [
        {
          kind: 'sheet',
          name: 'A',
          children: [
            {
              kind: 'table',
              name: 'x',
              variant: 'grid',
              showHeader: true,
              data: [],
              columns: [{ key: 'a' }],
            },
            {
              kind: 'table',
              name: 'x',
              variant: 'grid',
              showHeader: true,
              data: [],
              columns: [{ key: 'a' }],
            },
          ],
        },
      ],
    }
    expect(() => compile(dup)).toThrow(/duplicate block name "x" — used twice on sheet "A"/)
  })

  it('rejects a module that does not export a Workbook', () => {
    expect(() => compile({ kind: 'sheet', name: 'x', children: [] })).toThrow(
      /default-export a <Workbook>/,
    )
  })
})

describe('the invariant: inserting a data row', () => {
  const extra: Quarter = { quarter: 'Q2.5', revenue: 14_500_000, cogs: 5_800_000 }
  const grown = [...QUARTERS.slice(0, 2), extra, ...QUARTERS.slice(2)]

  it('shifts everything below the table by exactly one row', () => {
    const before = compile(budget()).sheets[1]!
    const after = compile(budget(grown)).sheets[1]!

    const noteBefore = at(before, 10, 0) as { value: string }
    const noteAfter = at(after, 11, 0) as { value: string }
    expect(noteBefore.value).toMatch(/^Forecast/)
    expect(noteAfter.value).toBe(noteBefore.value)
    expect(at(after, 10, 0)).not.toMatchObject({ value: noteBefore.value })
  })

  it('moves the total row with the data', () => {
    const before = compile(budget()).registry.get('pl') as TableAnchor
    const after = compile(budget(grown)).registry.get('pl') as TableAnchor

    expect(after.firstDataRow).toBe(before.firstDataRow)
    expect(after.lastDataRow).toBe(before.lastDataRow + 1)
    expect(after.totalRow).toBe((before.totalRow as number) + 1)
    expect(after.rowCount).toBe(before.rowCount + 1)
  })

  it('leaves the header and column mapping untouched', () => {
    const before = compile(budget()).registry.get('pl') as TableAnchor
    const after = compile(budget(grown)).registry.get('pl') as TableAnchor
    expect(after.headerRow).toBe(before.headerRow)
    expect([...after.columns]).toEqual([...before.columns])
  })

  it('keeps every row formula pointing at its own row', () => {
    const after = compile(budget(grown)).sheets[1]!
    const anchor = compile(budget(grown)).registry.get('pl') as TableAnchor
    for (let i = 0; i < grown.length; i += 1) {
      const cell = at(after, anchor.firstDataRow + i, 3) as {
        expr: { l: { target: { row: number } } }
      }
      expect(cell.expr.l.target.row).toBe(i)
    }
  })
})

describe('a bare ref is accepted wherever the docs say it is', () => {
  // Reported: the reference table in sheet-authoring lists ref().total() and
  // ref().cell() as things to write, but the types rejected them — documentation
  // and compiler disagreeing is worse than either being wrong alone.
  it('as a KPI value', () => {
    const book = compile(
      <Workbook>
        <Sheet name="S">
          <Stack gap={1}>
            <Table name="t" data={[{ n: 1 }, { n: 2 }]} columns={[col('n')]} total={{ n: 'sum' }} />
            <KpiBand items={[{ label: 'Total', value: ref('t').total('n') }]} />
          </Stack>
        </Sheet>
      </Workbook>,
    )
    expect(evaluateWorkbook(book).get('S!6,0')).toBe(3)
  })

  it('as a column formula', () => {
    const book = compile(
      <Workbook>
        <Sheet name="S">
          <Table
            name="t"
            data={[{ n: 5 }, { n: 9 }]}
            columns={[col('n'), col('copy', { formula: (r) => r.cell('n') })]}
          />
        </Sheet>
      </Workbook>,
    )
    const values = evaluateWorkbook(book)
    expect(values.get('S!1,1')).toBe(5)
    expect(values.get('S!2,1')).toBe(9)
  })

  it('as a key-value entry', () => {
    const book = compile(
      <Workbook>
        <Sheet name="S">
          <Stack gap={1}>
            <Table name="t" data={[{ n: 4 }]} columns={[col('n')]} total={{ n: 'sum' }} />
            <Table
              name="kv"
              kind="keyValue"
              data={[{ key: 'grand', label: 'Grand total', value: ref('t').total('n') }]}
            />
          </Stack>
        </Sheet>
      </Workbook>,
    )
    expect(evaluateWorkbook(book).get('S!4,1')).toBe(4)
  })
})

describe('defined names must be usable as Excel defined names', () => {
  const twoBlocks = (leftKey: string, rightKey: string) =>
    compile(
      <Workbook>
        <Sheet name="S">
          <Row gap={2}>
            <Table name="left" kind="keyValue" data={[{ key: leftKey, label: 'G', value: 0.1 }]} />
            <Table
              name="right"
              kind="keyValue"
              data={[{ key: rightKey, label: 'G', value: 0.2 }]}
            />
          </Row>
        </Sheet>
      </Workbook>,
    )

  it('refuses two blocks claiming one name', () => {
    // Reported by @ericweichun (#52). The evaluator resolved through the block
    // the author named while the exported formula used a defined name the other
    // block had overwritten — so the viewer showed 0.1 and Excel computed 0.2.
    expect(() => twoBlocks('growth', 'growth')).toThrow(/duplicate defined name "growth"/)
    expect(() => twoBlocks('growth', 'growth')).toThrow(/"left" and "right"/)
  })

  it('treats names as case-insensitive, as Excel does', () => {
    expect(() => twoBlocks('growth', 'Growth')).toThrow(/duplicate defined name/)
  })

  it('allows distinct names', () => {
    const book = twoBlocks('growthA', 'growthB')
    expect([...book.definedNames.keys()].sort()).toEqual(['growthA', 'growthB'])
  })

  it.each([
    ['B5', /looks like a cell address/],
    ['my key', /letters, digits, underscore/],
    ['2024', /must start with a letter/],
    ['c', /reserved by Excel/],
    ['', /is empty/],
  ])('refuses %s as a key', (key, message) => {
    expect(() => twoBlocks(key, 'other')).toThrow(message)
  })
})

describe('a reference never serializes to a cell it did not resolve', () => {
  it('falls back to an address when the name points elsewhere', () => {
    // The second lock behind the compile-time check: even if a stale name
    // reached the serializer, it must not be used unless sheet, row *and*
    // column match. The column was the part missing.
    const book = compile(
      <Workbook>
        <Sheet name="S">
          <Table name="a" kind="keyValue" data={[{ key: 'rate', label: 'R', value: 1 }]} />
        </Sheet>
      </Workbook>,
    )
    const entry = book.definedNames.get('rate')
    if (!entry) throw new Error('no defined name')
    book.definedNames.set('rate', { ...entry, addr: { r: entry.addr.r, c: entry.addr.c + 9 } })

    const formula = toFormula(ref('a').get('rate'), {
      registry: book.registry,
      definedNames: book.definedNames,
      sheet: 'S',
    })
    expect(formula).not.toBe('=rate')
    expect(formula).toMatch(/^=[A-Z]+\d+$/)
  })
})

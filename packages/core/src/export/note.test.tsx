import { strFromU8, unzipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { compile } from '../compile/compile.js'
import { Cell, col, Sheet, Stack, Table, Workbook } from '../compile/components.js'
import { evaluateWorkbook } from '../formula/evaluate.js'
import { toHtml } from './html.js'
import { XlsxWriter } from './xlsx.js'

const months = [
  { month: '2026-07', revenue: 100 },
  { month: '2026-08', revenue: 61 },
]

const book = () =>
  compile(
    (
      <Workbook>
        <Sheet name="Revenue">
          <Stack gap={1}>
            <Table
              name="rev"
              data={months}
              columns={[
                col('month', { header: 'Month' }),
                col('revenue', {
                  header: 'Revenue',
                  note: (row) =>
                    row.month === '2026-08'
                      ? 'Only 19 days of data — the export ran mid-month'
                      : undefined,
                }),
              ]}
            />
            <Cell value={7} note="From the billing export" />
          </Stack>
        </Sheet>
      </Workbook>
    ) as never,
  )

describe('where a number came from', () => {
  it('attaches a note only to the rows that need one', () => {
    const compiled = book()
    const sheet = compiled.sheets[0] as (typeof compiled.sheets)[number]
    const anchor = compiled.registry.get('rev')
    if (anchor?.kind !== 'table') throw new Error('!')
    const column = anchor.columns.get('revenue')
    expect(sheet.cells.get(`${anchor.firstDataRow},${column}`)?.note).toBeUndefined()
    expect(sheet.cells.get(`${anchor.firstDataRow + 1},${column}`)?.note).toContain('19 days')
  })

  it('writes a legacy note, which every spreadsheet app reads', async () => {
    const buffer = await new XlsxWriter().write(book())
    const files = unzipSync(new Uint8Array(buffer))
    const comments = Object.keys(files).find((name) => /comments\d*\.xml$/.test(name))
    expect(comments, 'no comment part was written').toBeDefined()
    const xml = strFromU8(files[comments as string] as Uint8Array)
    expect(xml).toContain('19 days of data')
    expect(xml).toContain('From the billing export')
    // The VML drawing is what makes Excel show the marker; without it the note
    // is in the file and invisible.
    expect(Object.keys(files).some((name) => name.endsWith('.vml'))).toBe(true)
  })

  it('shows the note in HTML too, with a marker so it is discoverable', () => {
    const compiled = book()
    const html = toHtml(compiled, { title: 'Revenue', values: evaluateWorkbook(compiled) })
    expect(html).toContain('title="Only 19 days of data — the export ran mid-month"')
    expect(html).toContain('os-noted')
  })

  it('escapes a note rather than letting it close the attribute', () => {
    const compiled = compile(
      (
        <Workbook>
          <Sheet name="S">
            <Cell value={1} note={'has "quotes" & <angle>'} />
          </Sheet>
        </Workbook>
      ) as never,
    )
    const html = toHtml(compiled, { title: 'S' })
    expect(html).toContain('&quot;quotes&quot;')
    expect(html).not.toContain('note="has "quotes"')
  })
})

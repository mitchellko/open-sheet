import { describe, expect, it } from 'vitest'
import { compile } from '../compile/compile.js'
import { col, Sheet, Spill, Stack, Table, Workbook } from '../compile/components.js'
import { evaluateWorkbook } from '../formula/evaluate.js'
import { sort } from '../formula/expr.js'
import { ref } from '../refs/ref.js'
import { toCsv } from './csv.js'
import { toHtml } from './html.js'

const reps = [{ v: 300 }, { v: 900 }, { v: 500 }]

function book(rows: number) {
  return compile(
    (
      <Workbook>
        <Sheet name="S">
          <Stack gap={1}>
            <Table name="t" data={reps} columns={[col('v', { header: 'V' })]} />
            <Spill formula={sort(ref('t').column('v'), 1, -1)} rows={rows} cols={1} />
          </Stack>
        </Sheet>
      </Workbook>
    ) as never,
  )
}

/**
 * Every renderer chose for itself whether a cell reads from the value map, and
 * each of them asked the same wrong question — "does this cell carry a
 * formula?". A cell inside a spill footprint carries neither a formula nor a
 * literal: the origin's formula fills it and the result lives only in the
 * value map. So the viewer, the HTML and the CSV all showed the top-left corner
 * of a spill and blank everywhere else, while Excel showed all of it.
 *
 * Found by a tester opening the same file in both.
 */
describe('a spill is drawn everywhere it is filled', () => {
  it('fills the whole footprint in HTML, not just the first cell', () => {
    const compiled = book(3)
    const html = toHtml(compiled, { title: 'S', values: evaluateWorkbook(compiled) })
    for (const value of ['900', '500', '300']) expect(html).toContain(`>${value}<`)
  })

  it('shows #N/A where the result does not reach, as a spreadsheet does', () => {
    const compiled = book(5)
    const html = toHtml(compiled, { title: 'S', values: evaluateWorkbook(compiled) })
    expect((html.match(/#N\/A/g) ?? []).length).toBe(2)
  })

  it('stops at the declared footprint when the result is longer', () => {
    const compiled = book(2)
    const values = evaluateWorkbook(compiled)
    const sheet = compiled.sheets[0] as (typeof compiled.sheets)[number]
    const origin = [...sheet.cells.entries()].find(([, cell]) => cell.spill)
    const [key] = origin as [string, unknown]
    const { r, c } = { r: Number(key.split(',')[0]), c: Number(key.split(',')[1]) }
    expect(values.get(`S!${r},${c}`)).toBe(900)
    expect(values.get(`S!${r + 1},${c}`)).toBe(500)
    expect(values.get(`S!${r + 2},${c}`)).toBeUndefined()
  })

  it('carries into the CSV too', () => {
    const compiled = book(3)
    const csv = toCsv(
      compiled.sheets[0] as (typeof compiled.sheets)[number],
      evaluateWorkbook(compiled),
    )
    expect(csv).toContain('900')
    expect(csv).toContain('500')
  })
})

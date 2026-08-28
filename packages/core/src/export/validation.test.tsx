import { existsSync, mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { strFromU8, unzipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { compile } from '../compile/compile.js'
import { Cell, col, Sheet, Stack, Table, Workbook } from '../compile/components.js'
import type { Validation } from '../model/validation.js'
import { ref } from '../refs/ref.js'
import { XlsxWriter } from './xlsx.js'

async function sheetXml(node: unknown, index = 1): Promise<string> {
  const buffer = await new XlsxWriter().write(compile(node as never))
  const files = unzipSync(new Uint8Array(buffer))
  return strFromU8(files[`xl/worksheets/sheet${index}.xml`] as Uint8Array)
}

const register = (validate: Validation) => (
  <Workbook>
    <Sheet name="Register">
      <Table
        name="requests"
        data={[
          { item: 'A', status: '待審' },
          { item: 'B', status: '核准' },
        ]}
        columns={[col('item', { header: '項目' }), col('status', { header: '狀態', validate })]}
      />
    </Sheet>
  </Workbook>
)

describe('what the recipient is allowed to type', () => {
  it('writes one rule over the column, not one per row', async () => {
    const xml = await sheetXml(register({ list: ['待審', '核准', '退回'] }))
    const elements = xml.match(/<dataValidation /g) ?? []
    expect(elements).toHaveLength(1)
    expect(xml).toContain('sqref="B2:B3"')
    expect(xml).toContain('待審,核准,退回')
  })

  it('carries both messages, because a refusal with no reason is worse than none', async () => {
    const xml = await sheetXml(
      register({
        list: ['待審', '核准'],
        promptTitle: '選一個',
        prompt: '從清單挑',
        errorTitle: '不在清單裡',
        error: '請從下拉選單選擇',
      }),
    )
    expect(xml).toContain('showInputMessage="1"')
    expect(xml).toContain('promptTitle="選一個"')
    expect(xml).toContain('showErrorMessage="1"')
    expect(xml).toContain('errorTitle="不在清單裡"')
  })

  it('refuses a list item containing a comma rather than silently splitting it', async () => {
    // Excel reads an inline list as one comma-separated string, so "Taipei, TW"
    // would become two options and the recipient would never know why.
    await expect(
      new XlsxWriter().write(compile(register({ list: ['Taipei, TW', 'Tokyo'] }) as never)),
    ).rejects.toThrow(/comma/)
  })

  it('points a list at a range so the lookup sheet stays the source of truth', async () => {
    const book = (
      <Workbook>
        <Sheet name="Lists">
          <Table
            name="statuses"
            showHeader={false}
            data={[{ name: '待審' }, { name: '核准' }]}
            columns={[col('name')]}
          />
        </Sheet>
        <Sheet name="Register">
          <Table
            name="requests"
            data={[{ status: '待審' }]}
            columns={[col('status', { validate: { list: ref('statuses').column('name') } })]}
          />
        </Sheet>
      </Workbook>
    )
    const xml = await sheetXml(book, 2)
    expect(xml).toContain('Lists!$A$1:$A$2')
  })

  it('reads an appendable list through INDIRECT, so appending an option reaches it', async () => {
    // An absolute range is fixed: an option appended to the bottom of the
    // lookup sheet never reaches the dropdown, with no error and nothing for
    // whoever maintains the list to notice. INDIRECT over the table column
    // resolves to whatever the table has grown to.
    const xml = await sheetXml(
      (
        <Workbook>
          <Sheet name="Lists">
            <Table
              name="statuses"
              appendable
              data={[{ name: '待審' }, { name: '核准' }]}
              columns={[col('name', { header: '狀態' })]}
            />
          </Sheet>
          <Sheet name="Register">
            <Table
              name="requests"
              data={[{ status: '待審' }]}
              columns={[col('status', { validate: { list: ref('statuses').column('name') } })]}
            />
          </Sheet>
        </Workbook>
      ) as never,
      2,
    )
    expect(xml).toContain('INDIRECT(&quot;statuses[狀態]&quot;)')
    // Never the bare structured reference: written straight into a validation
    // it makes Excel refuse to open the workbook — not ignore the rule, refuse
    // the file.
    expect(xml).not.toContain('<formula1>statuses[')
  })

  it('turns bounds into the operator Excel expects', async () => {
    // `between` is the format's default operator, so it is written by omission —
    // what distinguishes it is the second bound.
    const both = await sheetXml(register({ whole: { min: 1, max: 10 } }))
    expect(both).toContain('<formula1>1</formula1><formula2>10</formula2>')
    expect(both).not.toContain('operator=')

    const lower = await sheetXml(register({ whole: { min: 1 } }))
    expect(lower).toContain('operator="greaterThanOrEqual"')
    expect(lower).not.toContain('formula2')
  })

  it('converts an ISO date to the serial the file format stores', async () => {
    const xml = await sheetXml(register({ date: { from: '2026-01-01', to: '2026-12-31' } }))
    expect(xml).toContain('type="date"')
    // 2026-01-01 is serial 46023; the format has no notion of an ISO string.
    expect(xml).toContain('<formula1>46023</formula1>')
  })

  it('writes the enum the file format defines, not the word the API uses', async () => {
    // ST_DataValidationErrorStyle is stop | warning | information. We were
    // writing `error` and `info`, which Excel tolerates and openpyxl refuses —
    // so every workbook using validate was unreadable from Python, which is
    // exactly the pipeline this project is for. Two of the three were wrong and
    // the broken one was the default.
    const legal = new Set(['stop', 'warning', 'information'])
    for (const [style, expected] of [
      [undefined, 'stop'],
      ['warning', 'warning'],
      ['info', 'information'],
    ] as const) {
      const xml = await sheetXml(
        register({ list: ['待審'], error: 'x', ...(style ? { style } : {}) }),
      )
      const written = /errorStyle="([^"]*)"/.exec(xml)?.[1]
      expect(written, `style ${style ?? '(default)'}`).toBe(expected)
      expect(legal.has(written as string)).toBe(true)
    }
  })

  it('allows a blank by default — a half-filled form must still be saveable', async () => {
    const xml = await sheetXml(register({ list: ['待審'] }))
    expect(xml).toContain('allowBlank="1"')
  })

  it('validates a single cell too', async () => {
    const xml = await sheetXml(
      <Workbook>
        <Sheet name="S">
          <Stack>
            <Cell value={5} validate={{ whole: { min: 1, max: 10 }, error: '1 到 10' }} />
          </Stack>
        </Sheet>
      </Workbook>,
    )
    expect(xml).toContain('type="whole"')
    expect(xml).toContain('sqref="A1"')
  })
})

const SOFFICE = ['/opt/homebrew/bin/soffice', '/usr/bin/soffice', '/usr/bin/libreoffice'].find(
  (path) => existsSync(path),
)

describe.skipIf(!SOFFICE)('a real spreadsheet keeps the rule', () => {
  it('survives a LibreOffice round trip', { timeout: 180_000 }, async () => {
    const { execFileSync } = await import('node:child_process')
    const buffer = await new XlsxWriter().write(
      compile(register({ list: ['待審', '核准', '退回'], error: '請從清單選' }) as never),
    )
    const dir = mkdtempSync(join(tmpdir(), 'open-sheet-dv-'))
    const xlsx = join(dir, 'dv.xlsx')
    writeFileSync(xlsx, buffer)

    // Converting to xlsx again makes LibreOffice re-emit the rule from its own
    // model — proof it parsed it, not just that it copied bytes through.
    execFileSync(
      SOFFICE as string,
      [
        `-env:UserInstallation=file://${join(dir, 'profile')}`,
        '--headless',
        '--convert-to',
        'xlsx',
        '--outdir',
        join(dir, 'out'),
        xlsx,
      ],
      { stdio: 'pipe', timeout: 150_000 },
    )

    const { readFileSync } = await import('node:fs')
    const files = unzipSync(new Uint8Array(readFileSync(join(dir, 'out', 'dv.xlsx'))))
    const xml = strFromU8(files['xl/worksheets/sheet1.xml'] as Uint8Array)
    expect(xml).toContain('<dataValidation')
    expect(xml).toContain('待審')
  })
})

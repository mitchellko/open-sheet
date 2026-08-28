import { existsSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { strFromU8, unzipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { compile } from '../compile/compile.js'
import { col, Sheet, Stack, Table, Workbook } from '../compile/components.js'
import { toHtml } from './html.js'
import { encodeMargin, marginTemplate, pageCount, pageNumber, printDate } from './margin.js'
import { XlsxWriter } from './xlsx.js'

const rows = [{ v: 1 }, { v: 2 }, { v: 3 }]

const report = (print: Record<string, unknown>) => (
  <Workbook>
    <Sheet name="Report" print={print as never}>
      <Stack gap={1}>
        <Table name="summary" data={rows} columns={[col('v', { header: 'V' })]} />
        <Table name="detail" data={rows} columns={[col('v', { header: 'V' })]} />
      </Stack>
    </Sheet>
  </Workbook>
)

async function xml(node: unknown): Promise<string> {
  const buffer = await new XlsxWriter().write(compile(node as never))
  return strFromU8(unzipSync(new Uint8Array(buffer))['xl/worksheets/sheet1.xml'] as Uint8Array)
}

describe('what a printed page carries', () => {
  it('turns named fields into the codes nobody remembers', () => {
    expect(
      encodeMargin({
        center: { bold: 'FY26 Budget' },
        right: ['Page ', pageNumber, ' / ', pageCount],
      }),
    ).toBe('&C&BFY26 Budget&B&RPage &P / &N')
  })

  it('escapes a literal ampersand so it does not start a code', () => {
    // "Profit & Loss" would otherwise read as "Profit " followed by the code &L.
    expect(encodeMargin({ center: 'Profit & Loss' })).toBe('&CProfit && Loss')
  })

  it('writes the header and footer into the file', async () => {
    const out = await xml(report({ header: { center: 'FY26' }, footer: { right: [pageNumber] } }))
    expect(out).toContain('<oddHeader>&amp;CFY26</oddHeader>')
    expect(out).toContain('<oddFooter>&amp;R&amp;P</oddFooter>')
  })
})

describe('where the pages break', () => {
  it('resolves a named block to a row so the break moves with the content', async () => {
    const one = compile(report({ breakBefore: ['detail'] }) as never)
    const grown = compile(
      (
        <Workbook>
          <Sheet name="Report" print={{ breakBefore: ['detail'] }}>
            <Stack gap={1}>
              <Table name="summary" data={[...rows, { v: 4 }]} columns={[col('v')]} />
              <Table name="detail" data={rows} columns={[col('v')]} />
            </Stack>
          </Sheet>
        </Workbook>
      ) as never,
    )
    const before = (one.sheets[0] as (typeof one.sheets)[number]).pageBreaks
    const after = (grown.sheets[0] as (typeof grown.sheets)[number]).pageBreaks
    expect(before).toHaveLength(1)
    // One more data row above it, so the break is one row lower. A break written
    // as "row 47" would still be pointing at row 47.
    expect(after[0]).toBe((before[0] as number) + 1)
  })

  it('never puts a break above the first row, where it means nothing', () => {
    const book = compile(report({ breakBefore: ['summary'] }) as never)
    expect((book.sheets[0] as (typeof book.sheets)[number]).pageBreaks).toEqual([])
  })

  it('drops a break that lands on the top of the print area', () => {
    // The page it would start has already started. Excel ignores it, and a
    // break in the file that does nothing reads as though it should.
    const book = compile(report({ printArea: ['detail'], breakBefore: ['detail'] }) as never)
    expect((book.sheets[0] as (typeof book.sheets)[number]).pageBreaks).toEqual([])
  })

  it('prints only the named blocks', async () => {
    // The print area is a defined name in the workbook part, not a sheet
    // property — which is also why it survives a sheet rename.
    const buffer = await new XlsxWriter().write(
      compile(report({ printArea: ['summary'] }) as never),
    )
    const out = strFromU8(unzipSync(new Uint8Array(buffer))['xl/workbook.xml'] as Uint8Array)
    expect(out).toContain('_xlnm.Print_Area')
    expect(out).toContain('$A$1:$A$4')
    expect(out).not.toContain('$A$9')
  })

  it('names the sheet a block actually lives on', () => {
    expect(() =>
      compile(
        (
          <Workbook>
            <Sheet name="A">
              <Table name="here" data={rows} columns={[col('v')]} />
            </Sheet>
            <Sheet name="B" print={{ printArea: ['here'] }}>
              <Table name="other" data={rows} columns={[col('v')]} />
            </Sheet>
          </Workbook>
        ) as never,
      ),
    ).toThrow(/on sheet "A"/)
  })
})

describe('the same fields in the other renderers', () => {
  it('compiles to Chromium classes for the PDF, where the date does work', () => {
    const template = marginTemplate({ left: 'FY26', right: [pageNumber, ' / ', pageCount] })
    expect(template).toContain('<span class="pageNumber"></span>')
    expect(template).toContain('<span class="totalPages"></span>')
    expect(marginTemplate({ left: [printDate] })).toContain('<span class="date"></span>')
  })

  it('emits CSS margin boxes for the engines that have them', () => {
    const html = toHtml(compile(report({ footer: { right: [pageNumber, ' / ', pageCount] } })), {
      title: 'Report',
    })
    expect(html).toContain('@bottom-right')
    expect(html).toContain('counter(page)')
    expect(html).toContain('counter(pages)')
  })

  it('drops the fields CSS cannot express rather than printing something wrong', () => {
    // `date` means "the day this was printed". CSS has no way to say that, and
    // baking the build date would be a different statement.
    const html = toHtml(compile(report({ header: { left: [printDate] } })), { title: 'R' })
    expect(html).toContain('@top-left')
    expect(html).toContain('content: ""')
  })
})

const SOFFICE = ['/opt/homebrew/bin/soffice', '/usr/bin/soffice', '/usr/bin/libreoffice'].find(
  (path) => existsSync(path),
)

describe.skipIf(!SOFFICE)('and a real spreadsheet agrees about the page', () => {
  it('reads the print area, the break and the header back', { timeout: 180_000 }, async () => {
    const { execFileSync } = await import('node:child_process')
    const buffer = await new XlsxWriter().write(
      compile(
        report({
          printArea: ['summary'],
          breakBefore: ['detail'],
          header: { center: 'FY26 Budget' },
          footer: { right: [pageNumber, ' / ', pageCount] },
        }) as never,
      ),
    )
    const dir = mkdtempSync(join(tmpdir(), 'open-sheet-print-'))
    writeFileSync(join(dir, 'p.xlsx'), buffer)
    execFileSync(
      SOFFICE as string,
      [
        `-env:UserInstallation=file://${join(dir, 'profile')}`,
        '--headless',
        '--convert-to',
        'xlsx',
        '--outdir',
        join(dir, 'out'),
        join(dir, 'p.xlsx'),
      ],
      { stdio: 'pipe', timeout: 150_000 },
    )

    // Re-emitted from LibreOffice's own model, so this is proof it parsed them.
    const files = unzipSync(new Uint8Array(readFileSync(join(dir, 'out', 'p.xlsx'))))
    const book = strFromU8(files['xl/workbook.xml'] as Uint8Array)
    expect(book).toContain('_xlnm.Print_Area')
    const sheet = strFromU8(files['xl/worksheets/sheet1.xml'] as Uint8Array)
    expect(sheet).toContain('FY26 Budget')
    expect(sheet).toContain('&amp;P')
    expect(sheet).toContain('<rowBreaks')
  })
})

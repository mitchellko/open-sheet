import { execFileSync } from 'node:child_process'
import { existsSync, mkdtempSync, readdirSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { XlsxWriter } from '../export/xlsx.js'
import { CASES } from './cases.js'
import { FUNCTIONS } from './expr.js'
import { compare, evaluateCases, layout, summarise } from './verify.js'

const SOFFICE = ['/opt/homebrew/bin/soffice', '/usr/bin/soffice', '/usr/bin/libreoffice'].find(
  (path) => existsSync(path),
)

const CSV_RAW = 'csv:Text - txt - csv (StarCalc):44,34,76,1,,0,false,true,false,false,false,-1'

function parseCsv(text: string): string[][] {
  const rows: string[][] = []
  let row: string[] = []
  let field = ''
  let quoted = false
  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i]
    if (quoted) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"'
          i += 1
        } else quoted = false
      } else field += ch
      continue
    }
    if (ch === '"') quoted = true
    else if (ch === ',') {
      row.push(field)
      field = ''
    } else if (ch === '\n') {
      row.push(field)
      rows.push(row)
      row = []
      field = ''
    } else if (ch !== '\r') field += ch
  }
  if (field.length > 0 || row.length > 0) {
    row.push(field)
    rows.push(row)
  }
  return rows
}

/**
 * Every case in one workbook, so verifying forty functions costs one LibreOffice
 * invocation. The engine is the authority; our evaluator and the case's own
 * `expect` are both checked against it.
 */
describe.skipIf(!SOFFICE)('function verification', () => {
  it('reports where we agree, disagree, and cannot compute', { timeout: 240_000 }, async () => {
    const { book, resultColumn } = layout(CASES)
    const ours = evaluateCases(book, CASES)

    const dir = mkdtempSync(join(tmpdir(), 'open-sheet-verify-'))
    const file = join(dir, 'cases.xlsx')
    // No cached results: the engine must compute, not read back our answers.
    writeFileSync(file, await new XlsxWriter().write(book, { cacheValues: false }))

    execFileSync(
      SOFFICE as string,
      [
        `-env:UserInstallation=file://${join(dir, 'profile')}`,
        '--headless',
        '--convert-to',
        CSV_RAW,
        '--outdir',
        dir,
        file,
      ],
      { stdio: 'pipe', timeout: 200_000 },
    )

    const csv = readdirSync(dir).find((name) => name.endsWith('.csv'))
    expect(csv, 'LibreOffice produced no output').toBeDefined()
    const grid = parseCsv(readFileSync(join(dir, csv as string), 'utf8'))

    const theirs = CASES.map((_, index) => grid[index + 1]?.[resultColumn])
    const results = compare(CASES, ours, theirs)

    console.info(`\n${summarise(results)}\n`)

    // Disagreement is the only outcome that fails the build: it means our
    // evaluator and a real spreadsheet read the same formula differently, which
    // is the one thing this project cannot ship. Gaps are reported, not fatal.
    const wrong = results.filter((result) => result.outcome === 'disagrees')
    expect(
      wrong.map((r) => `${r.fn}: ours ${r.ours} vs engine ${r.theirs}`),
      'our evaluator disagrees with the engine',
    ).toEqual([])
  })

  it('every whitelisted function has at least one case', () => {
    // This used to check eight hardcoded names under a title that claimed all of
    // them, and 30 functions were whitelisted with nothing verifying any of them.
    const covered = new Set(CASES.flatMap((testCase) => testCase.fn.split('+')))
    // The array-returning ones cannot be a single-cell case by construction:
    // their result is a rectangle. They are verified in export/recalc.test.tsx,
    // which compiles a <Spill> and reads the whole footprint back.
    const elsewhere = new Set(['SORT', 'SORTBY', 'UNIQUE', 'FILTER', 'SEQUENCE', 'TRANSPOSE'])
    const uncovered = FUNCTIONS.filter((name) => !covered.has(name) && !elsewhere.has(name))
    expect(uncovered).toEqual([])
  })
})

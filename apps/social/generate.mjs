import { mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { compile, evaluateWorkbook, toHtml } from '@open-sheet/core'
import book from './sheet.mjs'

const here = dirname(fileURLToPath(import.meta.url))
const OUT = join(here, '..', '..', '.github', 'assets', 'social')

/**
 * The left panel is lifted out of sheet.mjs rather than restated, so the code in
 * the picture is the code that produced the sheet beside it. A promotional image
 * that has drifted from the API it advertises is worse than none.
 */
function excerpt() {
  const source = readFileSync(join(here, 'sheet.mjs'), 'utf8')
  const body = source.split('// #region shot')[1]?.split('// #endregion shot')[0]
  if (!body) throw new Error('sheet.mjs has no // #region shot marker')
  return body
    .replace(/^\s*const columns = \[\n/, '')
    .replace(/\n\]\s*$/, '')
    .split('\n')
    .map((line) => line.replace(/^ {2}/, ''))
    .join('\n')
    .trim()
}

const KEYWORDS = [
  'col',
  'sub',
  'div',
  'formula',
  'header',
  'format',
  'scale',
  'highlight',
  'above',
  'fill',
  'bold',
  'bar',
  'total',
]

function highlight(code) {
  let out = code.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
  out = out.replace(/('[^']*')/g, '<em class="s">$1</em>')
  for (const word of KEYWORDS) {
    out = out.replace(new RegExp(`\\b${word}\\b(?![^<]*</em>)`, 'g'), `<em class="k">${word}</em>`)
  }
  return out.replace(/\b(r)\.(cell|prev|isFirst)\b/g, '<em class="v">$1</em>.<em class="m">$2</em>')
}

function page(sheetHtml, code) {
  const style = /<style>([\s\S]*?)<\/style>/.exec(sheetHtml)[1]
  const body = /<body>([\s\S]*)<\/body>/.exec(sheetHtml)[1]
  return `<!doctype html><html><head><meta charset="utf-8"><style>
${style}
* { box-sizing: border-box; }
html, body { margin: 0; }
body { background: #0b1020; font-family: -apple-system, "Helvetica Neue", sans-serif; }
.wrap { display: grid; grid-template-columns: 1fr 1fr; width: 1600px; }
.pane { padding: 46px 48px; display: flex; flex-direction: column; }
.left { background: #0b1020; }
.right { background: #f6f8fb; }
.tag { font-size: 12px; letter-spacing: .15em; text-transform: uppercase; font-weight: 700; }
.left .tag { color: #7c8db5; }
.right .tag { color: #7c8aa3; }
h3 { font-size: 21px; margin: 8px 0 26px; font-weight: 600; letter-spacing: -.01em; }
.left h3 { color: #eef2fb; }
.right h3 { color: #0f172a; }
pre { margin: 0; font-family: "SF Mono", Menlo, monospace; font-size: 15px;
      line-height: 1.62; color: #dbe4f5; white-space: pre; }
em { font-style: normal; }
.k { color: #7dd3fc; } .s { color: #fca5a5; } .v { color: #fbbf24; } .m { color: #a5b4fc; }
.right table { font-size: 18px !important; }
.right table td { padding: 9px 14px !important; }
.right h2, .right .os-sheet-name { display: none; }
.grow { flex: 1; min-height: 24px; }
/* The sheet is shorter than the source beside it, so it sits in the middle of
   its half rather than clinging to the top with a field of grey beneath. */
.right main, .right .os-workbook { margin: auto 0; }
.note { font-size: 14.5px; line-height: 1.55; margin: 0; }
.left .note { color: #8ea3c9; }
.right .note { color: #5b6b82; }
.note b { color: #e8eefb; font-weight: 600; }
.right .note b { color: #0f172a; }
</style></head><body>
<div class="wrap">
  <div class="pane left">
    <div class="tag">what you write</div>
    <h3>Columns have names. Nothing has an address.</h3>
    <pre>${highlight(code)}</pre>
    <div class="grow"></div>
    <p class="note">Add a quarter to the data array and every reference
      re&#8209;resolves at compile time. <b>There is no B2:B13 to get wrong.</b></p>
  </div>
  <div class="pane right">
    <div class="tag">what the recipient opens</div>
    <h3>A live model, not a picture of one.</h3>
    <div class="grow"></div>
    ${body}
    <div class="grow"></div>
    <p class="note">Gross profit is <b>=B2-C2</b> in the file, not 7,300,000.
      Change a number and the whole sheet recalculates — in Excel, in Sheets,
      in LibreOffice.</p>
  </div>
</div></body></html>`
}

const compiled = compile(book)
const values = evaluateWorkbook(compiled)
const html = page(toHtml(compiled, { title: 'P&L', values }), excerpt())

mkdirSync(OUT, { recursive: true })
writeFileSync(join(OUT, 'source-vs-output.html'), html)

// Optional by design, exactly as the PDF export is: Playwright is ~300MB and
// nobody should need it to read this repo.
const specifier = 'playwright'
const playwright = await import(/* @vite-ignore */ specifier).catch(() => undefined)
if (!playwright) {
  console.log('wrote source-vs-output.html — install playwright to render the png')
  process.exit(0)
}

// The package can be present while its browser is not — a separate download,
// and a separate failure worth naming rather than a stack trace.
let browser
try {
  browser = await playwright.chromium.launch()
} catch (error) {
  console.log('wrote source-vs-output.html')
  console.log(`could not launch a browser: ${String(error).split('\n')[0]}`)
  console.log('run `npx playwright install chromium` to render the png')
  process.exit(0)
}

try {
  const tab = await browser.newPage({
    viewport: { width: 1600, height: 900 },
    deviceScaleFactor: 2,
  })
  await tab.setContent(html, { waitUntil: 'load' })
  const frame = await tab.locator('.wrap').first()
  await frame.screenshot({ path: join(OUT, 'source-vs-output.png') })
  console.log(`wrote ${join(OUT, 'source-vs-output.png')}`)
} finally {
  await browser.close()
}

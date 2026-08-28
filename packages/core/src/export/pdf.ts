import type { CompiledWorkbook } from '../compile/emit.js'
import { type HtmlOptions, toHtml } from './html.js'
import { marginTemplate, unsupportedFields } from './margin.js'

export interface PdfOptions extends HtmlOptions {
  format?: 'A4' | 'A3' | 'Letter' | 'Legal'
}

interface PlaywrightPage {
  setContent(html: string, options?: { waitUntil?: string }): Promise<void>
  emulateMedia(options: { media: string }): Promise<void>
  pdf(options: Record<string, unknown>): Promise<Uint8Array>
}

interface PlaywrightBrowser {
  newPage(): Promise<PlaywrightPage>
  close(): Promise<void>
}

interface PlaywrightLike {
  chromium: { launch(options?: Record<string, unknown>): Promise<PlaywrightBrowser> }
}

export class PlaywrightMissingError extends Error {
  constructor() {
    super(
      'PDF export needs a browser. Install playwright in your workspace ' +
        '(`pnpm add -D playwright && pnpm exec playwright install chromium`), ' +
        'or export HTML and print it from the browser.',
    )
    this.name = 'PlaywrightMissingError'
  }
}

/**
 * Optional by design. Playwright is ~300MB installed, which is not a cost every
 * user of a spreadsheet framework should pay to get an .xlsx — so it is loaded
 * only when a PDF is actually asked for, and its absence is a clear message
 * rather than a resolution error.
 */
export async function toPdf(book: CompiledWorkbook, options: PdfOptions = {}): Promise<Buffer> {
  // Non-literal specifier on purpose: the dependency is optional, so the type
  // checker must not require it to be installed, and the bundler must leave the
  // import for runtime.
  const specifier = 'playwright'
  const playwright = (await import(/* @vite-ignore */ specifier).catch(() => undefined)) as
    | PlaywrightLike
    | undefined
  if (!playwright) throw new PlaywrightMissingError()

  const html = toHtml(book, options)
  const browser = await playwright.chromium.launch()
  try {
    const page = await browser.newPage()
    await page.setContent(html, { waitUntil: 'load' })
    await page.emulateMedia({ media: 'print' })
    // Every sheet is on one page sequence, so the running header comes from the
    // first sheet that declares one rather than changing partway down the file.
    const print = book.sheets.find((sheet) => sheet.print?.header || sheet.print?.footer)?.print
    for (const [margin, where] of [
      [print?.header, 'header'],
      [print?.footer, 'footer'],
    ] as const) {
      const lost = unsupportedFields(margin, 'pdf')
      if (lost.length > 0) {
        process.emitWarning?.(
          `open-sheet: the PDF ${where} drops ${lost.join(', ')} — Chromium has no equivalent. ` +
            'The xlsx export keeps them.',
        )
      }
    }

    const running = print?.header || print?.footer
    const pdf = await page.pdf({
      format: options.format ?? 'A4',
      landscape: (options.orientation ?? 'landscape') === 'landscape',
      printBackground: true,
      margin: {
        top: print?.header ? '20mm' : '12mm',
        bottom: print?.footer ? '20mm' : '12mm',
        left: '12mm',
        right: '12mm',
      },
      ...(running
        ? {
            displayHeaderFooter: true,
            headerTemplate: print?.header ? marginTemplate(print.header) : '<span></span>',
            footerTemplate: print?.footer ? marginTemplate(print.footer) : '<span></span>',
          }
        : {}),
    })
    return Buffer.from(pdf)
  } finally {
    await browser.close()
  }
}

import type { MarginContent, MarginField, PageMargin } from '../compile/nodes.js'

/**
 * Excel's header and footer format. `&P` is the page number, `&N` the count,
 * `&D` the date — a cryptic language nobody remembers, which is why authors
 * name the field and this produces the code.
 */
const CODES: Record<string, string> = {
  pageNumber: '&P',
  pageCount: '&N',
  date: '&D',
  time: '&T',
  sheetName: '&A',
  fileName: '&F',
}

function encodePart(part: string | MarginField): string {
  // A literal & would otherwise start a code; Excel escapes it by doubling.
  if (typeof part === 'string') return part.replace(/&/g, '&&')
  if ('field' in part) {
    const code = CODES[part.field]
    if (!code) throw new Error(`unknown header field "${part.field}"`)
    return code
  }
  // `&B` and `&I` toggle, so each has to be turned back off or it bleeds into
  // whatever follows — including the next section of the same line.
  if ('bold' in part) return `&B${encode(part.bold)}&B`
  return `&I${encode(part.italic)}&I`
}

export function encode(content: MarginContent): string {
  if (Array.isArray(content)) return content.map(encodePart).join('')
  return encodePart(content as string | MarginField)
}

export function encodeMargin(margin: PageMargin): string {
  let out = ''
  if (margin.left !== undefined) out += `&L${encode(margin.left)}`
  if (margin.center !== undefined) out += `&C${encode(margin.center)}`
  if (margin.right !== undefined) out += `&R${encode(margin.right)}`
  return out
}

/** Named so the common fields read as values rather than object literals. */
export const pageNumber: MarginField = { field: 'pageNumber' }
export const pageCount: MarginField = { field: 'pageCount' }
export const printDate: MarginField = { field: 'date' }
export const printTime: MarginField = { field: 'time' }
export const sheetName: MarginField = { field: 'sheetName' }
export const fileName: MarginField = { field: 'fileName' }

/**
 * Chromium is the PDF engine, and it does not implement CSS margin boxes
 * (`@top-center` and friends). It has its own header template instead, where
 * the running fields are spans with reserved class names — so that is what the
 * same authored fields compile to for PDF.
 *
 * `date` works here and does not in the HTML export, because Chromium fills it
 * at print time and CSS has no way to express "the day this was printed".
 */
const HTML_FIELDS: Record<string, string> = {
  pageNumber: '<span class="pageNumber"></span>',
  pageCount: '<span class="totalPages"></span>',
  date: '<span class="date"></span>',
  time: '',
  sheetName: '<span class="title"></span>',
  fileName: '<span class="title"></span>',
}

function escapeText(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

function templatePart(part: string | MarginField): string {
  if (typeof part === 'string') return escapeText(part)
  if ('field' in part) return HTML_FIELDS[part.field] ?? ''
  if ('bold' in part) return `<b>${toTemplate(part.bold)}</b>`
  return `<i>${toTemplate(part.italic)}</i>`
}

export function toTemplate(content: MarginContent): string {
  if (Array.isArray(content)) return content.map(templatePart).join('')
  return templatePart(content as string | MarginField)
}

/** Chromium renders the template at 0 font-size unless it is told otherwise. */
export function marginTemplate(margin: PageMargin): string {
  const cell = (content: MarginContent | undefined, align: string) =>
    `<div style="flex:1;text-align:${align}">${content === undefined ? '' : toTemplate(content)}</div>`
  return (
    `<div style="font-size:9px;font-family:system-ui,sans-serif;color:#64748b;` +
    `width:100%;padding:0 12mm;display:flex">` +
    cell(margin.left, 'left') +
    cell(margin.center, 'center') +
    cell(margin.right, 'right') +
    '</div>'
  )
}

/** Fields with no equivalent in a given renderer, so the loss can be reported rather than silent. */
export function unsupportedFields(
  margin: PageMargin | undefined,
  renderer: 'pdf' | 'html',
): string[] {
  if (!margin) return []
  const lost = new Set<string>()
  const walk = (content: MarginContent | undefined): void => {
    if (content === undefined) return
    if (Array.isArray(content)) {
      for (const part of content) walk(part)
      return
    }
    if (typeof content === 'string') return
    const field = content as MarginField
    if ('bold' in field) {
      walk(field.bold)
      return
    }
    if ('italic' in field) {
      walk(field.italic)
      return
    }
    if (!('field' in field)) return
    const supported =
      renderer === 'pdf'
        ? HTML_FIELDS[field.field] !== ''
        : field.field === 'pageNumber' || field.field === 'pageCount'
    if (!supported) lost.add(field.field)
  }
  walk(margin.left)
  walk(margin.center)
  walk(margin.right)
  return [...lost]
}

const CSS_FIELDS: Record<string, string> = {
  pageNumber: 'counter(page)',
  pageCount: 'counter(pages)',
  date: '',
  time: '',
  sheetName: '',
  fileName: '',
}

function cssPart(part: string | MarginField): string {
  if (typeof part === 'string') return JSON.stringify(part)
  if ('field' in part) return CSS_FIELDS[part.field] ?? ''
  if ('bold' in part) return toCssContent(part.bold)
  return toCssContent(part.italic)
}

/** A CSS `content` value: quoted strings and counters, space separated. */
export function toCssContent(content: MarginContent): string {
  const parts = (Array.isArray(content) ? content : [content as string | MarginField])
    .map(cssPart)
    .filter((part) => part !== '')
  return parts.length > 0 ? parts.join(' ') : '""'
}

/**
 * CSS margin boxes, for the print engines that implement them. Chromium does
 * not — which is why the PDF export goes through its header template instead —
 * but Firefox, Prince and paged.js do, and emitting nothing would mean printing
 * the HTML by hand loses the header everywhere rather than only in Chrome.
 */
export function marginBoxes(
  header: PageMargin | undefined,
  footer: PageMargin | undefined,
): string {
  const boxes: string[] = []
  const add = (margin: PageMargin | undefined, edge: 'top' | 'bottom') => {
    if (!margin) return
    for (const [side, content] of [
      ['left', margin.left],
      ['center', margin.center],
      ['right', margin.right],
    ] as const) {
      if (content === undefined) continue
      boxes.push(
        `@${edge}-${side} { content: ${toCssContent(content)}; font-size: 9pt; color: #64748b; }`,
      )
    }
  }
  add(header, 'top')
  add(footer, 'bottom')
  return boxes.join('\n      ')
}

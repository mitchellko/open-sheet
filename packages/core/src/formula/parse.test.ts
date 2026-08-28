import { describe, expect, it } from 'vitest'
import { compile } from '../compile/compile.js'
import { Cell } from '../compile/components.js'
import { evaluateWorkbook } from './evaluate.js'
import { FUNCTIONS, mul, raw } from './expr.js'
import { parseFormula } from './parse.js'
import { serialize } from './serialize.js'
import { isNotEvaluated } from './value.js'

/**
 * A function we do not implement and have no plan to. Asserted against the
 * whitelist so that whitelisting it later fails this test loudly instead of
 * quietly turning these into tests of nothing — which is what happened when
 * M7 whitelisted the LOG10 and XIRR this file used to name.
 */
const UNSUPPORTED = 'BESSELJ'

const context = { registry: new Map(), definedNames: new Map(), sheet: 'S' } as never

const round = (input: string) => serialize(parseFormula(input).expr, context)

describe('parsing hand-written formulas', () => {
  it('round-trips arithmetic with the right precedence', () => {
    expect(round('=A1+B2')).toBe('A1+B2')
    expect(round('=A1+B2*2')).toBe('A1+B2*2')
    expect(round('=(A1+B2)*2')).toBe('(A1+B2)*2')
    expect(round('=A1-(B2-1)')).toBe('A1-(B2-1)')
    expect(round('=-A1')).toBe('-A1')
  })

  it('parses ranges and whitelisted functions', () => {
    expect(round('=SUM(B2:B13)')).toBe('SUM(B2:B13)')
    expect(round('=ROUND(A1/B1, 2)')).toBe('ROUND(A1/B1,2)')
    expect(round('=IF(A1>0, A1, 0)')).toBe('IF(A1>0,A1,0)')
  })

  it('handles literals the way Excel writes them', () => {
    expect(round('="say ""hi"""')).toBe('"say ""hi"""')
    expect(round('=TRUE')).toBe('TRUE')
    expect(round('=1.5e3')).toBe('1500')
  })

  it('keeps defined names verbatim', () => {
    expect(round('=B4*growth')).toBe('B4*growth')
  })

  it('does not mistake a function for a cell address', () => {
    // LOG10 lexes like a cell reference until you see the parenthesis.
    expect(parseFormula('=LOG10(A1)').expr).toMatchObject({ k: 'fn', name: 'LOG10' })
  })

  it('degrades an unsupported function rather than reading it as a reference', () => {
    expect(FUNCTIONS).not.toContain(UNSUPPORTED)
    const parsed = parseFormula(`=${UNSUPPORTED}(A1)`)
    expect(parsed.degraded).toBe(true)
    expect(parsed.reason).toContain('unsupported function')
  })

  it('degrades rather than throwing on anything it cannot handle', () => {
    for (const input of [`=${UNSUPPORTED}(A1:A9)`, '=SUM(', '=@#$%', '=A1 A2']) {
      const parsed = parseFormula(input)
      expect(parsed.degraded, input).toBe(true)
      expect(parsed.expr.k).toBe('raw')
    }
  })

  it('treats an empty formula as an empty string', () => {
    expect(parseFormula('=').degraded).toBe(false)
  })

  it('keeps an address as an address rather than resolving it to a reference', () => {
    // It still evaluates, but it stays visibly the thing that breaks on a row
    // insert — resolving it would disguise that.
    expect(parseFormula('=A1').expr).toEqual({ k: 'addr', ref: 'A1' })
    expect(parseFormula('=B2:B13').expr).toEqual({ k: 'addr', ref: 'B2:B13' })
  })

  it('treats an out-of-bounds address as a defined name, which is what Excel does', () => {
    // ZZZZ1 is past column XFD, so it cannot be a cell reference — and that makes
    // it a legal defined name rather than an error.
    expect(parseFormula('=ZZZZ1')).toMatchObject({
      degraded: false,
      expr: { k: 'raw', src: 'ZZZZ1' },
    })
  })

  it('degrades the whole formula to one raw node, which needs no parentheses', () => {
    // XIRR is outside the whitelist, so nothing parses and the raw node *is* the
    // top-level expression.
    expect(round('=XIRR(A1:A9,B1:B9)*2')).toBe('XIRR(A1:A9,B1:B9)*2')
  })

  it('still parenthesises a raw expression nested inside another', () => {
    expect(serialize(mul(raw('=XIRR(A1:A9,B1:B9)'), 2), context)).toBe('(XIRR(A1:A9,B1:B9))*2')
    expect(serialize(mul(raw('growth'), 2), context)).toBe('growth*2')
  })
})

describe('formula strings in a workbook', () => {
  it('are parsed so the cell still evaluates', () => {
    const book = compile({
      kind: 'workbook',
      children: [
        {
          kind: 'sheet',
          name: 'S',
          children: [
            {
              kind: 'stack',
              gap: 0,
              children: [
                { kind: 'cell', value: 3 },
                { kind: 'cell', value: 4 },
                Cell({ formula: '=A1+A2' }),
              ],
            },
          ],
        },
      ],
    })
    const values = evaluateWorkbook(book)
    expect(values.get('S!2,0')).toBe(7)
  })

  it('degrade honestly when they cannot be parsed', () => {
    const book = compile({
      kind: 'workbook',
      children: [
        {
          kind: 'sheet',
          name: 'S',
          children: [Cell({ formula: `=${UNSUPPORTED}(B1:B9)` })],
        },
      ],
    })
    expect(isNotEvaluated(evaluateWorkbook(book).get('S!0,0'))).toBe(true)
  })
})

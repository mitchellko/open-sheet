import { describe, expect, it } from 'vitest'
import { compile } from '../compile/compile.js'
import { budget } from '../compile/fixtures.js'
import { ref } from '../refs/ref.js'
import type { ResolveContext } from '../refs/resolve.js'
import { evaluateWorkbook } from './evaluate.js'
import { add, div, mul, neg, raw, sub, sum } from './expr.js'
import { parseFormula } from './parse.js'
import { toFormula } from './serialize.js'

function contextFor(sheet: string): ResolveContext {
  const book = compile(budget())
  return { registry: book.registry, definedNames: book.definedNames, sheet }
}

describe('serialize', () => {
  const pl = () => contextFor('P&L')

  it('writes a column range the author never addressed', () => {
    expect(toFormula(sum(ref('pl').column('revenue')), pl())).toBe('=SUM(B5:B8)')
  })

  it('writes a same-row difference', () => {
    expect(toFormula(sub(ref('pl').cell('revenue', 2), ref('pl').cell('cogs', 2)), pl())).toBe(
      '=B7-C7',
    )
  })

  it('writes the total cell', () => {
    expect(toFormula(ref('pl').total('revenue'), pl())).toBe('=B9')
  })

  it('prefers a defined name over an address, and qualifies across sheets', () => {
    expect(
      toFormula(mul(ref('pl').cell('revenue', 0), ref('assumptions').get('growth')), pl()),
    ).toBe('=B5*growth')
  })

  it('qualifies a cross-sheet address when there is no defined name', () => {
    const context = contextFor('Assumptions')
    expect(toFormula(sum(ref('pl').column('revenue')), context)).toBe("=SUM('P&L'!B5:B8)")
  })

  it('omits parentheses it does not need', () => {
    const c = pl()
    const a = ref('pl').cell('revenue', 0)
    const b = ref('pl').cell('cogs', 0)
    expect(toFormula(mul(add(a, b), 2), c)).toBe('=(B5+C5)*2')
    expect(toFormula(add(mul(a, 2), b), c)).toBe('=B5*2+C5')
    expect(toFormula(sub(a, sub(b, 1)), c)).toBe('=B5-(C5-1)')
    expect(toFormula(div(a, mul(b, 2)), c)).toBe('=B5/(C5*2)')
  })

  it('escapes strings and renders booleans Excel-style', () => {
    const c = pl()
    expect(toFormula({ k: 'lit', v: 'say "hi"' }, c)).toBe('="say ""hi"""')
    expect(toFormula({ k: 'lit', v: true }, c)).toBe('=TRUE')
  })

  it('handles unary minus', () => {
    const c = pl()
    expect(toFormula(neg(ref('pl').cell('revenue', 0)), c)).toBe('=-B5')
    expect(toFormula(neg(add(1, 2)), c)).toBe('=-(1+2)')
  })

  it('passes raw() through verbatim', () => {
    expect(toFormula(raw('=XIRR(A1:A9,B1:B9)'), pl())).toBe('=XIRR(A1:A9,B1:B9)')
  })

  it('serializes every formula the demo workbook produced', () => {
    const book = compile(budget())
    const context: ResolveContext = {
      registry: book.registry,
      definedNames: book.definedNames,
      sheet: 'P&L',
    }
    const formulas = [...(book.sheets[1]?.cells.values() ?? [])]
      .filter((cell) => cell.expr)
      .map((cell) => toFormula(cell.expr as never, context))
    expect(formulas).toContain('=B6-C6')
    expect(formulas).toContain('=B6/B5-1')
    expect(formulas).toContain('=SUM(B5:B8)')
    expect(formulas.every((f) => f.startsWith('='))).toBe(true)
  })
})

describe('resolution errors name what the author can act on', () => {
  it('rejects r.prev() on the first row', () => {
    expect(() => toFormula(ref('pl').cell('revenue', -1), contextFor('P&L'))).toThrow(
      /must guard with r\.isFirst/,
    )
  })

  it('suggests the column the author meant', () => {
    expect(() => toFormula(sum(ref('pl').column('revenu')), contextFor('P&L'))).toThrow(
      /did you mean "revenue"/,
    )
  })

  it('suggests the block the author meant', () => {
    expect(() => toFormula(sum(ref('pnl').column('revenue')), contextFor('P&L'))).toThrow(
      /did you mean "pl"/,
    )
  })
})

describe('raw as a tagged template', () => {
  const pl = () => contextFor('P&L')

  it('resolves interpolated references, so the escape hatch needs no addresses', () => {
    // The framework's one rule is that you never write a cell address. A raw()
    // that only takes a string forces you to break it — and the address it
    // forces you to write is exactly what an inserted row invalidates.
    expect(toFormula(raw`=LARGE(${ref('pl').column('revenue')}, 1)`, pl())).toBe('=LARGE(B5:B8, 1)')
  })

  it('handles several references and a defined name', () => {
    expect(
      toFormula(
        raw`=INDEX(${ref('pl').column('quarter')},MATCH(${ref('pl').total('revenue')},${ref('pl').column('revenue')},0))*${ref('assumptions').get('growth')}`,
        pl(),
      ),
    ).toBe('=INDEX(A5:A8,MATCH(B9,B5:B8,0))*growth')
  })

  it('still exports without evaluating', () => {
    // Honest: we cannot compute LARGE, and we do not pretend to.
    const book = compile(budget())
    const values = evaluateWorkbook(book)
    void values
    expect(raw`=LARGE(${ref('pl').column('revenue')}, 1)`.k).toBe('rawTemplate')
  })

  it('interpolates an expression, not only a bare reference', () => {
    // r.cell() returns a RefExpr, so an expression is the *commonest*
    // interpolation. Typing the signature as `Ref | Expr` while implementing
    // only `Ref` made exactly that form crash inside the writer, with tsc happy.
    expect(toFormula(raw`=1+${ref('pl').cell('revenue', 0)}`, pl())).toBe('=1+B5')
    expect(toFormula(raw`=SQRT(${mul(ref('pl').cell('revenue', 0), 2)})`, pl())).toBe('=SQRT(B5*2)')
  })

  it('interpolates a literal', () => {
    expect(toFormula(raw`=ROW()+${5}`, pl())).toBe('=ROW()+5')
  })

  it('refuses an empty interpolation with a message naming the position', () => {
    expect(() => raw`=A${undefined as never}`).toThrow(/interpolation 1 is undefined/)
  })

  it('keeps the plain string form working', () => {
    expect(toFormula(raw('=XIRR(A1:A9,B1:B9)'), pl())).toBe('=XIRR(A1:A9,B1:B9)')
  })
})

describe('functions newer than Excel 2007', () => {
  const pl = () => contextFor('P&L')

  it('are stored with the _xlfn prefix the format requires', () => {
    // Verified against LibreOffice: `IFS(...)` gives #NAME?, `_xlfn.IFS(...)`
    // computes. The prefix is how the file stores it; Excel displays the bare
    // name, so the author never sees it.
    expect(toFormula({ k: 'fn', name: 'IFS', args: [{ k: 'lit', v: 1 }] }, pl())).toBe(
      '=_xlfn.IFS(1)',
    )
    expect(toFormula({ k: 'fn', name: 'SWITCH', args: [{ k: 'lit', v: 1 }] }, pl())).toBe(
      '=_xlfn.SWITCH(1)',
    )
  })

  it('leaves functions that predate it alone', () => {
    expect(toFormula({ k: 'fn', name: 'SUM', args: [{ k: 'lit', v: 1 }] }, pl())).toBe('=SUM(1)')
    expect(toFormula({ k: 'fn', name: 'IFERROR', args: [{ k: 'lit', v: 1 }] }, pl())).toBe(
      '=IFERROR(1)',
    )
  })

  it('reads the prefix back off, since the author never wrote it', () => {
    expect(parseFormula('=_xlfn.SUM(1,2)').expr).toMatchObject({ k: 'fn', name: 'SUM' })
    expect(parseFormula('=_xlfn.IFS(1,2)').expr).toMatchObject({ k: 'fn', name: 'IFS' })
  })

  it('names the bare function when an _xlfn one degrades, never the prefix', () => {
    // `_xlfn.` is a storage detail of the file format; an author who never
    // wrote it should never have to read it in an error message.
    const parsed = parseFormula('=_xlfn.BESSELJ(1,2)')
    expect(parsed.reason ?? '').toContain('BESSELJ')
    expect(parsed.reason ?? '').not.toContain('_xlfn')
  })
})

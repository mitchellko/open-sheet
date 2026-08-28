import { fromA1 } from '../model/a1.js'
import type { BinaryOp, Expr } from './expr.js'
import { isWhitelisted } from './expr.js'

type Token =
  | { k: 'num'; v: number }
  | { k: 'str'; v: string }
  | { k: 'bool'; v: boolean }
  | { k: 'ref'; a: string; b?: string }
  | { k: 'name'; v: string }
  | { k: 'op'; v: BinaryOp }
  | { k: 'punct'; v: '(' | ')' | ',' }

const OPERATORS = ['<>', '<=', '>=', '+', '-', '*', '/', '^', '&', '=', '<', '>'] as const
const CELL = /^\$?[A-Za-z]{1,3}\$?\d{1,7}/
const NAME = /^[A-Za-z_][A-Za-z0-9_.]*/

class ParseError extends Error {}

function tokenize(input: string): Token[] {
  const tokens: Token[] = []
  let i = 0

  while (i < input.length) {
    const ch = input[i] as string
    if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r') {
      i += 1
      continue
    }

    if (ch === '"') {
      let value = ''
      i += 1
      while (i < input.length) {
        if (input[i] === '"') {
          if (input[i + 1] === '"') {
            value += '"'
            i += 2
            continue
          }
          i += 1
          break
        }
        value += input[i]
        i += 1
      }
      tokens.push({ k: 'str', v: value })
      continue
    }

    if (ch === '(' || ch === ')' || ch === ',') {
      tokens.push({ k: 'punct', v: ch })
      i += 1
      continue
    }

    const operator = OPERATORS.find((candidate) => input.startsWith(candidate, i))
    if (operator) {
      tokens.push({ k: 'op', v: operator })
      i += operator.length
      continue
    }

    if (/[0-9.]/.test(ch)) {
      const match = /^\d*\.?\d+(?:[eE][+-]?\d+)?/.exec(input.slice(i))
      if (!match) throw new ParseError(`bad number at ${i}`)
      tokens.push({ k: 'num', v: Number(match[0]) })
      i += match[0].length
      continue
    }

    const rest = input.slice(i)
    const cell = CELL.exec(rest)
    if (cell) {
      const after = rest.slice(cell[0].length)
      // A cell immediately followed by `(` is a function name that happens to
      // look like an address, e.g. LOG10(…).
      if (!after.startsWith('(')) {
        const second = after.startsWith(':') ? CELL.exec(after.slice(1)) : null
        if (second) {
          tokens.push({ k: 'ref', a: cell[0], b: second[0] })
          i += cell[0].length + 1 + second[0].length
        } else {
          tokens.push({ k: 'ref', a: cell[0] })
          i += cell[0].length
        }
        continue
      }
    }

    const name = NAME.exec(rest)
    if (name) {
      // `_xlfn.IFS` and `_xlfn._xlws.SORT` are how the file stores them; the
      // author writes `IFS` and `SORT`.
      const raw = name[0]
      const bare = raw.replace(/^_xlfn\.(_xlws\.)?/i, '')
      const upper = bare.toUpperCase()
      if (upper === 'TRUE' || upper === 'FALSE') tokens.push({ k: 'bool', v: upper === 'TRUE' })
      else tokens.push({ k: 'name', v: bare })
      i += raw.length
      continue
    }

    throw new ParseError(`unexpected character "${ch}" at ${i}`)
  }

  return tokens
}

const PRECEDENCE: Record<BinaryOp, number> = {
  '^': 5,
  '*': 4,
  '/': 4,
  '+': 3,
  '-': 3,
  '&': 2,
  '=': 1,
  '<': 1,
  '>': 1,
  '<=': 1,
  '>=': 1,
  '<>': 1,
}

class Parser {
  private at = 0
  constructor(private readonly tokens: Token[]) {}

  parse(): Expr {
    const expr = this.expression(0)
    if (this.at < this.tokens.length) throw new ParseError('trailing input')
    return expr
  }

  private peek(): Token | undefined {
    return this.tokens[this.at]
  }

  private expression(minimum: number): Expr {
    let left = this.unary()
    for (;;) {
      const token = this.peek()
      if (token?.k !== 'op') break
      const precedence = PRECEDENCE[token.v]
      if (precedence < minimum) break
      this.at += 1
      const right = this.expression(precedence + 1)
      left = { k: 'op', op: token.v, l: left, r: right }
    }
    return left
  }

  private unary(): Expr {
    const token = this.peek()
    if (token?.k === 'op' && (token.v === '-' || token.v === '+')) {
      this.at += 1
      const operand = this.unary()
      return token.v === '-' ? { k: 'neg', e: operand } : operand
    }
    return this.primary()
  }

  private primary(): Expr {
    const token = this.peek()
    if (!token) throw new ParseError('unexpected end of formula')
    this.at += 1

    if (token.k === 'num') return { k: 'lit', v: token.v }
    if (token.k === 'str') return { k: 'lit', v: token.v }
    if (token.k === 'bool') return { k: 'lit', v: token.v }

    if (token.k === 'ref') {
      // Verify the addresses parse, then keep them verbatim: an address written
      // by hand is exactly what the reference API exists to avoid, so it is
      // preserved rather than dignified with a resolved Ref.
      fromA1(token.a)
      if (token.b) fromA1(token.b)
      return { k: 'addr', ref: token.b ? `${token.a}:${token.b}` : token.a }
    }

    if (token.k === 'punct' && token.v === '(') {
      const inner = this.expression(0)
      this.consume(')')
      return inner
    }

    if (token.k === 'name') {
      const next = this.peek()
      if (next?.k === 'punct' && next.v === '(') {
        this.at += 1
        const args: Expr[] = []
        if (!this.isPunct(')')) {
          args.push(this.expression(0))
          while (this.isPunct(',')) {
            this.at += 1
            args.push(this.expression(0))
          }
        }
        this.consume(')')
        if (!isWhitelisted(token.v)) throw new ParseError(`unsupported function ${token.v}`)
        return { k: 'fn', name: token.v.toUpperCase(), args }
      }
      // A bare name is a defined name; keep it verbatim.
      return { k: 'raw', src: token.v }
    }

    throw new ParseError(`unexpected token`)
  }

  private isPunct(value: ')' | ','): boolean {
    const token = this.peek()
    return token?.k === 'punct' && token.v === value
  }

  private consume(value: ')' | ','): void {
    if (!this.isPunct(value)) throw new ParseError(`expected ${value}`)
    this.at += 1
  }
}

export interface ParsedFormula {
  expr: Expr
  /** True when the whole formula fell back to raw() and will not be evaluated. */
  degraded: boolean
  reason?: string
}

/**
 * A compatibility shim for `"=A1+B2"` written by hand. It parses what it can so
 * the cell still evaluates in the viewer — but a literal address is precisely
 * what breaks when a row is inserted, so addresses are preserved verbatim rather
 * than resolved, and callers are expected to warn.
 *
 * Never throws: unparsed input degrades to raw(), which exports fine and shows
 * as #NOT_EVALUATED rather than as an invented number.
 */
export function parseFormula(input: string): ParsedFormula {
  const source = input.replace(/^\s*=/, '').trim()
  if (source === '') return { expr: { k: 'lit', v: '' }, degraded: false }

  try {
    return { expr: new Parser(tokenize(source)).parse(), degraded: false }
  } catch (error) {
    return {
      expr: { k: 'raw', src: source },
      degraded: true,
      reason: error instanceof Error ? error.message : String(error),
    }
  }
}

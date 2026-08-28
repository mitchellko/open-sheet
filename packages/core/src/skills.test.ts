import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { FUNCTIONS } from './formula/expr.js'
import * as entry from './index.js'

const skills = join(dirname(fileURLToPath(import.meta.url)), '..', 'skills')

function frontmatter(markdown: string): Record<string, string> {
  const match = /^---\n([\s\S]*?)\n---/.exec(markdown)
  if (!match) return {}
  const out: Record<string, string> = {}
  for (const line of (match[1] as string).split('\n')) {
    const at = line.indexOf(':')
    if (at > 0) out[line.slice(0, at).trim()] = line.slice(at + 1).trim()
  }
  return out
}

/**
 * Skills ship to users and are the only documentation an agent reads. A broken
 * one fails silently — nothing imports it, no build step touches it, and the
 * only symptom is worse workbooks written by someone else's agent.
 *
 * This exists because SKILL.md was overwritten with the contents of one of its
 * own reference files by a careless scripted edit, shipped that way, and was
 * caught by a tester noticing the checklist they were told about was missing.
 */
describe.skipIf(!existsSync(skills))('the shipped skills', () => {
  const names = readdirSync(skills).filter((name) => existsSync(join(skills, name, 'SKILL.md')))

  it('are all present', () => {
    expect(names.sort()).toEqual([
      'apply-comments',
      'create-sheet',
      'create-theme',
      'current-sheet',
      'sheet-authoring',
    ])
  })

  it('each declare a name matching their directory', () => {
    for (const name of names) {
      const meta = frontmatter(readFileSync(join(skills, name, 'SKILL.md'), 'utf8'))
      expect(meta.name, `${name}/SKILL.md frontmatter`).toBe(name)
      expect(meta.description?.length ?? 0, `${name} description`).toBeGreaterThan(40)
    }
  })

  it('sheet-authoring carries the contract, not a copy of a reference file', () => {
    const text = readFileSync(join(skills, 'sheet-authoring', 'SKILL.md'), 'utf8')
    expect(text).toContain('## The one rule')
    expect(text).toContain('## The file contract')
    expect(text).toContain('## The component surface')
    expect(text).toContain('## What the framework cannot check for you')
    expect(text).toContain('## Self-review before finishing')
  })

  it('keeps the self-review checks that were each added for a reason', () => {
    const text = readFileSync(join(skills, 'sheet-authoring', 'SKILL.md'), 'utf8')
    for (const check of [
      'No A1 address anywhere',
      'assumptions block',
      'Nothing that depends on the values is decided in the `data` array',
      'r.isFirst',
      'Nothing was invented',
      'covers the same ground',
      '#NOT_EVALUATED',
    ]) {
      expect(text, `self-review is missing: ${check}`).toContain(check)
    }
  })

  it('references every file it points at', () => {
    const text = readFileSync(join(skills, 'sheet-authoring', 'SKILL.md'), 'utf8')
    const referenced = [...text.matchAll(/references\/([a-z-]+\.md)/g)].map((m) => m[1])
    expect(new Set(referenced).size).toBeGreaterThanOrEqual(4)
    for (const file of referenced) {
      expect(
        existsSync(join(skills, 'sheet-authoring', 'references', file as string)),
        `missing references/${file}`,
      ).toBe(true)
    }
  })

  it('no skill is a duplicate of another file', () => {
    const seen = new Map<string, string>()
    for (const name of names) {
      const body = readFileSync(join(skills, name, 'SKILL.md'), 'utf8').slice(0, 400)
      const previous = seen.get(body)
      expect(previous, `${name}/SKILL.md duplicates ${previous}`).toBeUndefined()
      seen.set(body, name)
    }
  })

  /**
   * The reference told authors to reach for `raw()` on XIRR for two releases
   * after M7 whitelisted it — steering them to #NOT_EVALUATED for a function
   * that works. Docs that name functions rot in both directions.
   */
  describe('the formula reference', () => {
    const text = readFileSync(join(skills, 'sheet-authoring', 'references', 'formulas.md'), 'utf8')

    it('names only builders that exist', () => {
      const table = text.slice(text.indexOf('| Aggregate |'), text.indexOf('\n\nEvery one'))
      const named = [...table.matchAll(/`([a-z_][a-z_0-9]*)`/g)].map((m) => m[1] as string)
      expect(named.length).toBeGreaterThan(80)
      expect(named.filter((name) => !(name in entry))).toEqual([])
    })

    it('does not send authors to raw() for something we support', () => {
      for (const match of text.matchAll(/raw\('=([A-Z][A-Z0-9.]*)\(/g)) {
        expect(FUNCTIONS, `raw() example uses the supported ${match[1]}`).not.toContain(match[1])
      }
    })
  })
})

import { readdirSync, readFileSync, statSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const here = dirname(fileURLToPath(import.meta.url))
const source = join(here, '..', '..', 'core', 'skills')
const shipped = join(here, '..', 'template', 'skills')

function tree(root: string, prefix = ''): Map<string, string> {
  const out = new Map<string, string>()
  for (const entry of readdirSync(root)) {
    const path = join(root, entry)
    const key = prefix ? `${prefix}/${entry}` : entry
    if (statSync(path).isDirectory()) {
      for (const [name, body] of tree(path, key)) out.set(name, body)
    } else {
      out.set(key, readFileSync(path, 'utf8'))
    }
  }
  return out
}

/**
 * The template is a copy, synced by a build script. Turbo caches the CLI build
 * against this package's own files, and the skills it copies live in another
 * package — so editing a skill does not invalidate the cache and the template
 * ships whatever it had last time. That is how `printing.md` came to be missing
 * from a package that had already been built.
 *
 * Run `pnpm cli sync-skills` when this fails.
 */
describe('the skills the CLI ships', () => {
  const authored = tree(source)
  const copied = tree(shipped)

  it('has every file the framework has', () => {
    expect([...authored.keys()].filter((name) => !copied.has(name)).sort()).toEqual([])
  })

  it('ships nothing the framework does not have', () => {
    expect([...copied.keys()].filter((name) => !authored.has(name)).sort()).toEqual([])
  })

  it('matches them byte for byte', () => {
    const stale = [...authored].filter(([name, body]) => copied.get(name) !== body).map(([n]) => n)
    expect(stale.sort()).toEqual([])
  })
})

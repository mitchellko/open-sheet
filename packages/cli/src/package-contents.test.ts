import { execFileSync } from 'node:child_process'
import { existsSync, mkdtempSync, readdirSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const here = dirname(fileURLToPath(import.meta.url))
const packages = join(here, '..', '..')

function pack(name: string): { files: string[]; manifest: Record<string, any> } {
  const out = mkdtempSync(join(tmpdir(), 'open-sheet-pack-'))
  execFileSync('pnpm', ['pack', '--pack-destination', out], {
    cwd: join(packages, name),
    stdio: 'pipe',
    timeout: 120_000,
    shell: process.platform === 'win32',
  })
  const tarball = join(out, readdirSync(out).find((file) => file.endsWith('.tgz')) as string)

  const files = execFileSync('tar', ['tzf', tarball], { encoding: 'utf8' })
    .replace(/\r\n/g, '\n')
    .split('\n')
    .filter(Boolean)
    .map((path) => path.replace(/^package\//, '').replace(/\r$/, ''))

  const manifest = JSON.parse(
    execFileSync('tar', ['xzOf', tarball, 'package/package.json'], { encoding: 'utf8' }),
  )
  return { files, manifest }
}

/**
 * Three of these were real, and none of them was visible before packing:
 * the scaffolder shipped without skills, the framework shipped without the
 * viewer, and npm silently dropped the template's .gitignore. A test that only
 * runs the source tree cannot see any of them.
 */
describe.skipIf(!existsSync(join(packages, 'core', 'dist')))('what actually ships', () => {
  it('core ships the viewer, which is served as source through the user’s Vite', () => {
    const { files } = pack('core')
    expect(files).toContain('src/app/main.tsx')
    expect(files).toContain('dist/index.mjs')
    expect(files).toContain('bin.js')
  }, 120_000)

  it('core ships the skills, and no tests or fixtures', () => {
    const { files } = pack('core')
    expect(files).toContain('skills/create-sheet/SKILL.md')
    expect(files).toContain('skills/sheet-authoring/references/formulas.md')
    expect(files.filter((f) => f.includes('.test.'))).toEqual([])
    expect(files.filter((f) => f.includes('fixtures'))).toEqual([])
  }, 120_000)

  it('the scaffolder carries the skills itself', () => {
    // `npx @open-sheet/cli init` installs the scaffolder alone, so core is not
    // on disk to resolve them from.
    const { files } = pack('cli')
    expect(files).toContain('template/skills/create-sheet/SKILL.md')
    expect(files).toContain('template/sheets/getting-started/index.tsx')
  }, 120_000)

  it('the scaffolder ships its gitignore under a name npm will not drop', () => {
    const { files } = pack('cli')
    expect(files).toContain('template/gitignore')
    expect(files).not.toContain('template/.gitignore')
  }, 120_000)

  it('no package ships an unresolvable workspace: dependency', () => {
    // pnpm rewrites the protocol when packing; npm leaves it, and an installed
    // package carrying `workspace:*` cannot resolve. Publishing must go through
    // pnpm — this is what would catch it if that ever changed.
    for (const name of ['core', 'cli', 'mcp']) {
      const { manifest } = pack(name)
      const ranges = [
        ...Object.values(manifest.dependencies ?? {}),
        ...Object.values(manifest.peerDependencies ?? {}),
      ]
      expect(
        ranges.filter((range) => String(range).startsWith('workspace:')),
        name,
      ).toEqual([])
    }
  }, 200_000)
})

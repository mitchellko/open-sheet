import { execFileSync } from 'node:child_process'
import { existsSync, mkdtempSync, readdirSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { afterAll, describe, expect, it } from 'vitest'

const here = dirname(fileURLToPath(import.meta.url))
const repo = join(here, '..', '..', '..')

function chromium(): string | undefined {
  const cache = join(process.env.HOME ?? '', 'Library', 'Caches', 'ms-playwright')
  const linux = join(process.env.HOME ?? '', '.cache', 'ms-playwright')
  for (const root of [cache, linux]) {
    if (!existsSync(root)) continue
    for (const build of readdirSync(root).filter((d) => d.startsWith('chromium'))) {
      for (const rel of [
        join('chrome-headless-shell-mac-arm64', 'chrome-headless-shell'),
        join('chrome-linux', 'headless_shell'),
        join('chrome-linux', 'chrome'),
      ]) {
        const path = join(root, build, rel)
        if (existsSync(path)) return path
      }
    }
  }
  return undefined
}

const BROWSER = chromium()

/**
 * The only test that would have caught either of the two blank-screen bugs that
 * shipped.
 *
 * Both existed solely in the published shape — one from a devDependency being
 * bundled, one from npm's flat node_modules serving CommonJS the browser cannot
 * import — and both returned HTTP 200 for every module. Checking status codes
 * proved nothing; the page has to actually render.
 */
describe.skipIf(!BROWSER)('a published install renders', () => {
  const scratch = mkdtempSync(join(tmpdir(), 'open-sheet-e2e-'))
  let server: { kill: () => void } | undefined

  afterAll(() => {
    server?.kill()
  })

  it('scaffolds, installs from tarballs, and paints the grid', { timeout: 600_000 }, async () => {
    const run = (cmd: string, args: string[], cwd: string) =>
      execFileSync(cmd, args, {
        cwd,
        stdio: 'pipe',
        timeout: 240_000,
        encoding: 'utf8',
        shell: process.platform === 'win32',
      })

    for (const pkg of ['core', 'cli']) {
      run('pnpm', ['pack', '--pack-destination', scratch], join(repo, 'packages', pkg))
    }
    const tarball = (name: string) =>
      join(scratch, readdirSync(scratch).find((f) => f.startsWith(`open-sheet-${name}-`)) as string)

    const host = join(scratch, 'host')
    execFileSync('mkdir', ['-p', host])
    run('npm', ['init', '-y'], host)
    run('npm', ['install', '--silent', tarball('cli')], host)
    run('npx', ['create-open-sheet', 'init', 'w'], host)

    const workspace = join(host, 'w')
    run('npm', ['install', '--silent', tarball('core')], workspace)

    const { spawn } = await import('node:child_process')
    const port = 5461
    const child = spawn('npx', ['open-sheet', 'dev', '--port', String(port)], {
      cwd: workspace,
      stdio: 'ignore',
    })
    server = child

    // Wait for the server rather than guessing at a sleep.
    const base = `http://localhost:${port}`
    for (let i = 0; i < 60; i += 1) {
      const ok = await fetch(base)
        .then((r) => r.ok)
        .catch(() => false)
      if (ok) break
      await new Promise((done) => setTimeout(done, 1000))
    }

    const { chromium: launcher } = await import('playwright')
    const browser = await launcher.launch({ executablePath: BROWSER })
    try {
      const page = await browser.newPage()
      const errors: string[] = []
      page.on('pageerror', (error) => errors.push(error.message))
      page.on('console', (message) => {
        if (message.type() === 'error') errors.push(message.text())
      })

      await page.goto(base, { waitUntil: 'domcontentloaded', timeout: 60_000 })
      await page.waitForSelector('.os-cell', { timeout: 60_000 })

      const text = await page.$eval('body', (el) => (el as HTMLElement).innerText)

      expect(errors, 'the console must be clean').toEqual([])
      // The sidebar and toolbar mounted…
      expect(text).toContain('Workbooks')
      expect(text).toContain('Inspect')
      // …and the grid painted real, formatted values from the starter workbook.
      expect(text).toContain('Unit price')
      expect(text).toContain('49.00')
      expect(text).toContain('38.0%')
    } finally {
      await browser.close()
    }
  })
})

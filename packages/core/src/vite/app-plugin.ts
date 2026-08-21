import { existsSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import type { Plugin } from 'vite'
import type { ResolvedConfig } from './config.js'

const here = dirname(fileURLToPath(import.meta.url))

/**
 * Anchored on the package root rather than on hops relative to this file: the
 * bundler decides which chunk this code lands in, and that has already changed
 * once — a relative path here breaks silently the next time it does.
 */
export function packageRoot(): string {
  let dir = here
  for (let i = 0; i < 6; i += 1) {
    if (existsSync(join(dir, 'package.json')) && existsSync(join(dir, 'src'))) return dir
    const parent = dirname(dir)
    if (parent === dir) break
    dir = parent
  }
  return resolve(here, '..')
}

/**
 * The viewer lives inside this package while workbooks live in the user's
 * workspace, so Vite's root stays the workspace and the app entry is reached
 * through /@fs. Serving it any other way would put the user's `sheets/` outside
 * the root and break relative imports in their own workbooks.
 */
export function appEntry(): string {
  const root = packageRoot()
  const candidates = [join(root, 'src', 'app', 'main.tsx'), join(root, 'app', 'main.tsx')]
  const found = candidates.find((path) => existsSync(path))
  if (!found) {
    throw new Error(
      `cannot locate the open-sheet viewer entry (package root ${root}; looked in ${candidates.join(', ')})`,
    )
  }
  return found
}

function shell(entry: string): string {
  const normalized = entry.replace(/\\/g, '/')
  const src = normalized.startsWith('/') ? `/@fs${normalized}` : `/@fs/${normalized}`
  return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>open-sheet</title>
</head>
<body>
<div id="root"></div>
<script type="module" src="${src}"></script>
</body>
</html>
`
}

export function appPlugin(_config: ResolvedConfig): Plugin {
  const entry = appEntry()

  return {
    name: 'open-sheet:app',
    configureServer(server) {
      // after Vite's own middlewares, so asset and HMR requests win
      return () => {
        server.middlewares.use(async (req, res, next) => {
          if (!req.url || req.method !== 'GET') return next()
          const [path] = req.url.split('?')
          if (path?.includes('.') || path?.startsWith('/@') || path?.startsWith('/__open-sheet')) {
            return next()
          }
          const html = await server.transformIndexHtml(req.url, shell(entry))
          res.statusCode = 200
          res.setHeader('Content-Type', 'text/html')
          res.end(html)
        })
      }
    },
  }
}

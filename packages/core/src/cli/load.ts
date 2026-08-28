import { createServer, type ViteDevServer } from 'vite'
import { resolveConfig, viteConfigFor } from '../vite/index.js'
import type { DesignSystem } from '../style/design.js'

export interface WorkbookModule {
  default: unknown
  meta?: { title?: string }
  design?: DesignSystem
}

/**
 * Loads a `.tsx` workbook through Vite's SSR pipeline so authors get the same
 * transform the dev server uses. The compiler itself needs no browser — this is
 * only how TypeScript and JSX get turned into a module.
 */
export async function createLoader(root: string): Promise<{
  load: (file: string) => Promise<WorkbookModule>
  close: () => Promise<void>
}> {
  const config = resolveConfig(root, {})
  const server: ViteDevServer = await createServer({
    ...viteConfigFor(config),
    root,
    configFile: false,
    server: { middlewareMode: true },
    appType: 'custom',
    logLevel: 'warn',
  })

  return {
    load: (file) => server.ssrLoadModule(file) as Promise<WorkbookModule>,
    close: () => server.close(),
  }
}

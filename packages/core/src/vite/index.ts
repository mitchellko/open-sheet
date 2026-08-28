import { readFileSync, realpathSync } from 'node:fs'
import { createRequire } from 'node:module'
import { dirname, resolve } from 'node:path'
import react from '@vitejs/plugin-react'
import type { InlineConfig, PluginOption } from 'vite'
import { apiPlugin } from './api-plugin.js'
import { appPlugin, packageRoot } from './app-plugin.js'
import type { ResolvedConfig } from './config.js'
import { jsxPlugin } from './jsx-plugin.js'
import { manifestPlugin } from './manifest-plugin.js'
import { mcpPlugin } from './mcp-plugin.js'

const require = createRequire(import.meta.url)

const REACT_SPECIFIERS = [
  'react-dom/client',
  'react-dom',
  'react/jsx-dev-runtime',
  'react/jsx-runtime',
  'react',
] as const

const COMMONJS_SPECIFIERS = ['@formulajs/formulajs'] as const

/**
 * The viewer ships inside this package, so its React must resolve from here —
 * the user's workspace has no reason to depend on React to author a spreadsheet.
 *
 * Anchored regexes, not the object form: object aliases match by prefix, so a
 * `react-dom` entry rewrites `react-dom/client` into `.../react-dom/index.js/client`.
 */
function reactAliases(): { find: RegExp; replacement: string }[] {
  const out: { find: RegExp; replacement: string }[] = []
  for (const specifier of REACT_SPECIFIERS) {
    try {
      out.push({
        find: new RegExp(`^${specifier.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&')}$`),
        replacement: require.resolve(specifier),
      })
    } catch {
      // leave it to the workspace; a clear resolve error beats a wrong alias
    }
  }
  return out
}

function esmEntry(specifier: string): string {
  const pkgPath = require.resolve(`${specifier}/package.json`)
  const pkg = JSON.parse(readFileSync(pkgPath, 'utf8')) as {
    exports?: { '.': { import?: { default?: string }; default?: string } }
    module?: string
    main?: string
  }
  const candidate =
    pkg.exports?.['.']?.import?.default ?? pkg.exports?.['.']?.default ?? pkg.module ?? pkg.main
  if (!candidate) return require.resolve(specifier)
  return resolve(dirname(pkgPath), candidate)
}

function commonJsAliases(): { find: RegExp; replacement: string }[] {
  const out: { find: RegExp; replacement: string }[] = []
  for (const specifier of COMMONJS_SPECIFIERS) {
    try {
      out.push({
        find: new RegExp(`^${specifier.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&')}$`),
        replacement: esmEntry(specifier),
      })
    } catch {
      // leave it to the workspace; a clear resolve error beats a wrong alias
    }
  }
  return out
}

/** The workspace root is already real (see resolveConfig); this is for our own package. */
function realPath(path: string): string {
  try {
    return realpathSync(path)
  } catch {
    return path
  }
}

/**
 * One instance of the compiler. The viewer imports it from this package while
 * workbooks import it by name; without an alias those can be two copies, and two
 * copies means a workbook compiled by one and rendered by the other.
 */
function coreAliases(): { find: RegExp; replacement: string }[] {
  const out: { find: RegExp; replacement: string }[] = []
  const entries: [RegExp, string][] = [
    [/^@open-sheet\/core\/jsx-dev-runtime$/, '@open-sheet/core/jsx-runtime'],
    [/^@open-sheet\/core\/jsx-runtime$/, '@open-sheet/core/jsx-runtime'],
    [/^@open-sheet\/core$/, '@open-sheet/core'],
  ]
  for (const [find, specifier] of entries) {
    try {
      out.push({ find, replacement: require.resolve(specifier) })
    } catch {
      // not self-resolvable: leave it to the workspace's own install
    }
  }
  return out
}

export function openSheetPlugins(config: ResolvedConfig, mcp = false): PluginOption[] {
  return [
    jsxPlugin(config),
    react({ include: /\/src\/app\/.*\.[jt]sx?$/ }),
    manifestPlugin(config),
    apiPlugin(config),
    ...(mcp ? [mcpPlugin(config)] : []),
    appPlugin(config),
  ]
}

export function viteConfigFor(
  config: ResolvedConfig,
  extra: Partial<InlineConfig> & { mcp?: boolean } = {},
): InlineConfig {
  const { mcp, ...rest } = extra
  const root = config.root
  return {
    root,
    configFile: false,
    logLevel: 'warn',
    appType: 'custom',
    plugins: openSheetPlugins(config, mcp === true),
    resolve: {
      alias: [...coreAliases(), ...reactAliases(), ...commonJsAliases()],
      dedupe: ['react', 'react-dom'],
    },
    optimizeDeps: {
      // Everything the viewer imports that is, or pulls in, CommonJS. Naming
      // them makes Vite convert the whole graph to ESM in one pass. Left to be
      // discovered — or excluded, as React was — the browser receives raw
      // CommonJS, the named import fails, and the viewer dies before it mounts
      // with nothing on screen but a console error.
      include: [...REACT_SPECIFIERS, ...COMMONJS_SPECIFIERS],
      // Already ESM, and aliased to a real path — but more importantly it
      // reaches Node-only code (the optional `import('playwright')` behind PDF
      // export) that the dep optimizer cannot analyse and refuses to bundle.
      // The browser never calls it; serving the module directly keeps it that way.
      exclude: ['@open-sheet/core', '@open-sheet/core/jsx-runtime'],
    },
    server: {
      port: config.port,
      fs: { allow: [root, realPath(packageRoot())] },
    },
    ...rest,
  }
}

export type { ResolvedConfig } from './config.js'
export { DEFAULT_PORT, resolveConfig } from './config.js'

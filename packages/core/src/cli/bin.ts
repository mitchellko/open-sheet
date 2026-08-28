import { relative } from 'node:path'
import { build } from './build.js'
import { dev } from './dev.js'
import { preview } from './preview.js'

const USAGE = `open-sheet — the spreadsheet framework built for agents

Usage:
  open-sheet <command> [options]

Commands:
  dev       Start the viewer with hot reload (default port 5373)
  build     Compile every workbook under sheets/ and write .xlsx (and .csv)
  preview   Serve what build wrote

Options:
  --help         Show this, from any command
  --out <dir>    Output directory (default: dist)
  --root <dir>   Workspace root (default: cwd)
  --port <n>     Port for dev/preview
  --host <host>  Bind address for dev
  --open         Open a browser (dev)
  --mcp          Mount an MCP endpoint at /mcp (needs @open-sheet/mcp)
  --no-csv       Skip the per-sheet .csv files
  --html         Also write a self-contained, printable .html
  --pdf          Also write a .pdf (needs playwright installed)
`

function flag(argv: string[], name: string): string | undefined {
  const index = argv.indexOf(`--${name}`)
  if (index === -1) return undefined
  return argv[index + 1]
}

const HELP = ['--help', '-h', 'help']
const BUILD_FLAGS = ['--out', '--root', '--no-csv', '--html', '--pdf']

export async function run(argv: string[]): Promise<number> {
  const command = argv[0]

  // Anywhere, not only first. `open-sheet build --help` used to compile the
  // whole workspace and overwrite dist/ — reaching for --help should never be
  // the thing that destroys your output.
  if (!command || argv.some((arg) => HELP.includes(arg))) {
    process.stdout.write(USAGE)
    return command ? 0 : 1
  }

  const root = flag(argv, 'root')
  const port = flag(argv, 'port')

  if (command === 'dev') {
    const options: Parameters<typeof dev>[0] = {
      open: argv.includes('--open'),
      mcp: argv.includes('--mcp'),
    }
    if (root) options.root = root
    if (port) options.port = Number(port)
    const host = flag(argv, 'host')
    if (host) options.host = host
    await dev(options)
    return new Promise<number>(() => {})
  }

  if (command === 'preview') {
    const options: Parameters<typeof preview>[0] = {}
    if (root) options.root = root
    if (port) options.port = Number(port)
    const out = flag(argv, 'out')
    if (out) options.out = out
    await preview(options)
    return new Promise<number>(() => {})
  }

  if (command !== 'build') {
    process.stderr.write(`unknown command: ${command}\n\n${USAGE}`)
    return 1
  }

  const unknown = argv
    .slice(1)
    .filter((arg) => arg.startsWith('--'))
    .filter((arg) => !BUILD_FLAGS.includes(arg))
  if (unknown.length > 0) {
    process.stderr.write(
      `unknown option${unknown.length > 1 ? 's' : ''}: ${unknown.join(', ')}\n\n${USAGE}`,
    )
    return 1
  }

  const options: Parameters<typeof build>[0] = {
    csv: !argv.includes('--no-csv'),
    html: argv.includes('--html'),
    pdf: argv.includes('--pdf'),
  }
  const out = flag(argv, 'out')
  if (out) options.out = out
  if (root) options.root = root

  const results = await build(options)
  const cwd = process.cwd()

  for (const result of results) {
    process.stdout.write(`${result.title} (${result.id})\n`)
    for (const file of result.files) process.stdout.write(`  ${relative(cwd, file)}\n`)
    if (result.notEvaluated > 0) {
      process.stdout.write(
        `  note: ${result.notEvaluated} cell(s) exported as live formulas but not evaluated here\n`,
      )
    }
    for (const note of result.notes) process.stdout.write(`  note: ${note}\n`)
    for (const warning of result.warnings) process.stdout.write(`  warning: ${warning}\n`)
  }

  return 0
}

# @open-sheet/cli

## 0.2.0

### Minor Changes

- The scaffolded workspace carries the current `sheet-authoring` skill, including
  the new references for printing and for what the recipient can do with the
  exported file. A test now compares the shipped copy against the framework's own
  byte for byte — the build script that syncs them was invisible to Turbo's cache,
  so a skill edit did not invalidate the CLI build and the template shipped stale.


## 0.1.9

## 0.1.8

## 0.1.7

## 0.1.6

## 0.1.5

## 0.1.4

## 0.1.3

## 0.1.2

## 0.1.1

## 0.1.0

### Minor Changes

- The scaffolded workspace carries the current `sheet-authoring` skill, including
  the new references for printing and for what the recipient can do with the
  exported file. A test now compares the shipped copy against the framework's
  own byte for byte — the build script that syncs them was invisible to Turbo's
  cache, so a skill edit did not invalidate the CLI build and the template
  shipped stale.


- 6832304: First release. Author a workbook as JSX with named data columns and export a
  `.xlsx` containing live formulas — no cell address appears anywhere in the
  source, and every reference re-resolves when the data changes.
  
  Includes the viewer and dev server, the `/create-sheet` skills, an MCP server,
  inspect mode with source write-back, themes and a design panel, native Excel
  charts, and export to xlsx, csv, html, and pdf.

### Patch Changes

- 320fbd6: The scaffolder now carries the skills itself, and the framework ships the viewer,
  so `npx @open-sheet/cli init` and `open-sheet dev` work from a published install.

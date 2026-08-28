# Social assets

Generates the images used to talk about open-sheet, from a workbook the
framework actually compiled.

```bash
pnpm --filter @open-sheet/social generate
```

Writes into `.github/assets/social/`.

## Why it lives here rather than in a design file

The left panel of the image is lifted out of `sheet.mjs` at generation time —
the region between the `// #region shot` markers — so the code in the picture is
the code that produced the sheet next to it. A promotional image that has
drifted from the API it advertises is worse than no image.

The right panel is the real HTML export of the real compiled workbook. Nothing
in either half is mocked up.

## Changing what it shows

Edit `sheet.mjs` like any other open-sheet workbook and run the command again.
Move the `#region shot` markers to change which part of the source appears.

## Playwright

The PNG step needs a browser, which is a separate download from the package:

```bash
npx playwright install chromium
```

Without it the command still writes `source-vs-output.html`, which you can open
and screenshot by hand.

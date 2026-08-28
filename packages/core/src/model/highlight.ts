/**
 * Two or three stops, lowest value to highest. Three is the usual shape for a
 * variance column: bad, neutral, good.
 */
export type ColorScale = readonly [string, string] | readonly [string, string, string]

export type IconSet = 'arrows' | 'trafficLights'

export interface HighlightStyle {
  fill?: string
  color?: string
  bold?: boolean
}

/**
 * Exactly one test per rule. Every one is decidable from the cell's own value
 * plus the range it sits in, which is what lets the HTML export draw the same
 * thing Excel does rather than approximate it.
 */
export type HighlightTest =
  | { above: number }
  | { below: number }
  | { atLeast: number }
  | { atMost: number }
  | { equals: number | string | boolean }
  | { between: readonly [number, number] }
  | { contains: string }
  | { duplicates: true }
  | { top: number }
  | { bottom: number }

export type Highlight = HighlightTest & HighlightStyle

export function testOf(rule: Highlight): HighlightTest {
  return rule as HighlightTest
}

/** Excel's own three-arrow order runs lowest first, so ours does too. */
export const ICON_GLYPHS: Record<IconSet, readonly [string, string, string]> = {
  arrows: ['▼', '▶', '▲'],
  trafficLights: ['●', '●', '●'],
}

export const ICON_COLORS: Record<IconSet, readonly [string, string, string]> = {
  arrows: ['#dc2626', '#ca8a04', '#16a34a'],
  trafficLights: ['#dc2626', '#eab308', '#16a34a'],
}

export const EXCEL_ICON_SETS: Record<IconSet, string> = {
  arrows: '3Arrows',
  trafficLights: '3TrafficLights1',
}

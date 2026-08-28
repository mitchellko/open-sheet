import type { Ref } from '../refs/ref.js'

/**
 * A rejection with no explanation is worse than no validation at all — the
 * recipient sees only that the file will not take their answer. Both messages
 * are part of the OOXML element, so both are authorable here.
 */
export interface ValidationMessages {
  /** Shown when the cell is selected. Say what is expected, before they type. */
  prompt?: string
  promptTitle?: string
  /** Shown when the entry is refused. Say why, and what would be accepted. */
  error?: string
  errorTitle?: string
  /** `error` refuses the entry; `warning` and `info` let it through. */
  style?: 'error' | 'warning' | 'info'
  /** Whether an empty cell passes. Defaults to true — a blank is not a typo. */
  allowBlank?: boolean
}

export interface Bounds {
  min?: number
  max?: number
}

export type ValidationRule =
  /** A fixed set of answers, or a range holding them — a lookup sheet stays the source of truth. */
  | { list: readonly string[] | Ref }
  | { whole: Bounds }
  | { decimal: Bounds }
  | { textLength: Bounds }
  /** Excel serials or ISO strings; both are accepted and converted. */
  | { date: { from?: string | number; to?: string | number } }
  /** Anything else: an expression that must evaluate true for the entry to pass. */
  | { custom: import('../formula/expr.js').Expr }

export type Validation = ValidationMessages & ValidationRule

export function ruleOf(validation: Validation): ValidationRule {
  return validation as ValidationRule
}

export function isListRule(rule: ValidationRule): rule is { list: readonly string[] | Ref } {
  return 'list' in rule
}

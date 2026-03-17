import type { Answers, SurveyQuestion, SurveyOption, ScopeMapEntry } from './types';

// ─── Entry / Contact helpers ──────────────────────────────────────────────────

const ENTRY_CODES = new Set([
  'Q000_PROJECT_TYPE',
  'Q000_FLOW_MODE',
  'Q000_PRESET_RESI',
  'Q000_PRESET_COMM',
  'Q000_PRESET_TONE',
]);

export function isEntryCode(code: string): boolean {
  return ENTRY_CODES.has(code);
}

export function isContactCode(code: string): boolean {
  return code?.startsWith('Q90') ?? false;
}

// ─── Scope-map based visibility ───────────────────────────────────────────────

/**
 * Returns the set of question_codes that should be visible
 * based on current answers + scopeMap triggers.
 * Entry, contact, COMMON/GLOBAL-scoped questions are always included.
 */
export function getVisibleCodes(
  answers: Answers,
  questions: SurveyQuestion[],
  scopeMap: ScopeMapEntry[],
): Set<string> {
  const codes = new Set<string>();

  // Always include entry, contact, and globally-scoped questions
  for (const q of questions) {
    const scope = (q.exposure_scope || '').toUpperCase();
    if (
      isEntryCode(q.question_code) ||
      isContactCode(q.question_code) ||
      scope === 'COMMON' ||
      scope === 'GLOBAL'
    ) {
      codes.add(q.question_code);
    }
  }

  // Evaluate scope map triggers
  for (const m of scopeMap) {
    const triggerValues = m.trigger_values_json || [];
    const answerVal = answers[m.trigger_question_id];
    if (!answerVal) continue;

    const answerArr = Array.isArray(answerVal) ? answerVal : [String(answerVal)];
    const hit = triggerValues.some((tv) => answerArr.includes(tv));

    if (hit) {
      const targets = m.target_question_ids_json || [];
      for (const t of targets) {
        codes.add(t);
      }
    }
  }

  return codes;
}

// ─── visible_if_expr evaluation ───────────────────────────────────────────────

/**
 * Evaluates a question's visible_if_expr against current answers.
 * Supports:
 *   - `CODE == VALUE`
 *   - `CODE != VALUE`
 *   - `CODE includes VALUE`
 *   - `CODE in [VAL1, VAL2, ...]`
 *   - `&&` (AND) and `||` (OR) operators
 *   - Nested parentheses (stripped for evaluation)
 */
export function evalVisExpr(expr: string, answers: Answers): boolean {
  const trimmed = (expr || '').trim();
  if (!trimmed) return true;

  // Split on && (top-level AND)
  const andParts = trimmed.split(/\s*&&\s*/);
  for (const part of andParts) {
    const cleaned = part.replace(/^\s*\(+\s*/, '').replace(/\s*\)+\s*$/, '').trim();
    if (cleaned && !evalClause(cleaned, answers)) return false;
  }
  return true;
}

function evalClause(p: string, answers: Answers): boolean {
  // Handle OR
  if (p.includes('||')) {
    return p
      .split(/\s*\|\|\s*/)
      .some((x) =>
        evalClause(
          x.replace(/^\s*\(+\s*/, '').replace(/\s*\)+\s*$/, '').trim(),
          answers,
        ),
      );
  }

  let m: RegExpMatchArray | null;

  // CODE includes VALUE
  if ((m = p.match(/^(\w+)\s+includes\s+(\w+)$/))) {
    const v = answers[m[1]];
    if (!v) return false;
    return (Array.isArray(v) ? v : [String(v)]).includes(m[2]);
  }

  // CODE in [VAL1, VAL2, ...]
  if ((m = p.match(/^(\w+)\s+in\s+\[([^\]]*)\]$/))) {
    const v = answers[m[1]];
    if (!v) return false;
    const allowed = m[2].split(',').map((s) => s.trim());
    return allowed.includes(String(v));
  }

  // CODE == VALUE
  if ((m = p.match(/^(\w+)\s*==\s*(\w+)$/))) {
    return String(answers[m[1]] || '') === m[2];
  }

  // CODE != VALUE
  if ((m = p.match(/^(\w+)\s*!=\s*(\w+)$/))) {
    return String(answers[m[1]] || '') !== m[2];
  }

  // Unknown expression → default visible
  return true;
}

// ─── Space→Trade filtering ────────────────────────────────────────────────────

const TRADE_TO_SPACE: Record<string, string> = {
  R_BATH: 'BATH',
  R_KITCHEN: 'KITCHEN',
};

/**
 * Filters trade options for R012_TRADES based on which spaces
 * are selected in R010_SPACES.
 */
export function filterTradesBySpaces(
  options: SurveyOption[],
  answers: Answers,
): SurveyOption[] {
  const spaces = answers.R010_SPACES;
  if (!Array.isArray(spaces) || spaces.length === 0) return options;
  if (spaces.includes('ALL')) return options;

  return options.filter((o) => {
    const requiredSpace = TRADE_TO_SPACE[o.option_code];
    if (requiredSpace) return spaces.includes(requiredSpace);
    return true; // general trades always visible
  });
}

import type { Answers, SurveyQuestion, SurveyOption, ScopeMapEntry, QuickBundle, Step } from './types';
import { getVisibleCodes, evalVisExpr, isEntryCode, isContactCode } from './visibilityEngine';

// ─── Helpers ──────────────────────────────────────────────────────────────────

function sortBySortOrder<T extends { sort_order: number }>(a: T, b: T): number {
  return (Number(a.sort_order) || 0) - (Number(b.sort_order) || 0);
}

/** Push a single-question step if the question exists. */
function pushQuestion(steps: Step[], pool: SurveyQuestion[], code: string): void {
  const q = pool.find((x) => x.question_code === code);
  if (q) steps.push({ title: q.title || code, desc: q.description || '', questions: [q] });
}

/** Push a grouped multi-question step if any of the codes exist. */
function pushGroup(
  steps: Step[],
  pool: SurveyQuestion[],
  codes: string[],
  title: string,
  desc?: string,
): void {
  const found = codes.map((c) => pool.find((x) => x.question_code === c)).filter(Boolean) as SurveyQuestion[];
  if (found.length) steps.push({ title, desc: desc || '', questions: found });
}

/** Build the contact step from Q90x questions. */
function makeContactStep(allQ: SurveyQuestion[]): Step {
  return {
    title: '연락처를 남겨주세요',
    desc: '상담 접수에 필요한 정보입니다.',
    questions: allQ.filter((x) => isContactCode(x.question_code)),
  };
}

/** Filter out empty steps. */
function filterSteps(steps: Step[]): Step[] {
  return steps.filter((s) => s.questions && s.questions.length > 0);
}

// ─── Residential flow ─────────────────────────────────────────────────────────

function buildResidential(steps: Step[], q: SurveyQuestion[], detail: boolean): void {
  pushQuestion(steps, q, 'R001_HOUSING_TYPE');
  pushQuestion(steps, q, 'R002_AREA');
  pushGroup(steps, q, ['R003_CONDITION', 'R004_OCCUPANCY'], '현재 상태', '집 상태와 거주 상태를 알려주세요.');
  if (detail) pushQuestion(steps, q, 'R005_REASON');
  pushQuestion(steps, q, 'R010_SPACES');
  pushQuestion(steps, q, 'R011_SCOPE_LEVEL');
  pushQuestion(steps, q, 'R012_TRADES');
  if (detail) pushQuestion(steps, q, 'R013_STRUCTURE');
  pushQuestion(steps, q, 'R100_DEMO_SCOPE');
  pushGroup(steps, q, ['R110_BATH_SCOPE', 'R111_BATH_LEVEL'], '욕실 공사', '욕실 공사 범위를 선택해주세요.');
  pushQuestion(steps, q, 'R112_BATH_GRADE');
  pushGroup(steps, q, ['R120_KITCHEN_SCOPE', 'R121_KITCHEN_LAYOUT'], '주방 공사', '주방 공사 범위를 선택해주세요.');
  pushQuestion(steps, q, 'R122_KITCHEN_GRADE');
  pushGroup(steps, q, ['R130_FLOOR_SCOPE', 'R131_FLOOR_MATERIAL'], '바닥 공사', '바닥 범위와 선호 소재를 선택해주세요.');
  pushQuestion(steps, q, 'R132_FLOOR_GRADE');
  pushQuestion(steps, q, 'R140_WINDOW_SCOPE');
  pushQuestion(steps, q, 'R150_LIGHT_SCOPE');
  pushQuestion(steps, q, 'R151_LIGHT_TYPE');
  pushQuestion(steps, q, 'R160_BUILTIN_SCOPE');
  pushGroup(steps, q, ['R200_STYLE', 'R201_TONE'], '분위기와 톤', '원하시는 느낌을 선택해주세요.');
  pushGroup(steps, q, ['R210_FAMILY', 'R211_PET'], '생활 정보', '가족 구성과 반려동물을 알려주세요.');
  pushQuestion(steps, q, 'R300_AGE');
  if (detail) pushQuestion(steps, q, 'R301_RISKS');
  pushQuestion(steps, q, 'R302_MEP');
  pushQuestion(steps, q, 'R310_BUDGET');
  pushGroup(steps, q, ['R320_START', 'R321_END', 'R322_VISIT'], '일정', '희망 시작/완료 시점을 알려주세요.');
}

// ─── Commercial flow ──────────────────────────────────────────────────────────

function buildCommercial(steps: Step[], q: SurveyQuestion[], detail: boolean): void {
  pushQuestion(steps, q, 'C001_BIZ_TYPE');
  pushQuestion(steps, q, 'C005_CUISINE_TYPE');
  pushQuestion(steps, q, 'C002_AREA');
  pushGroup(steps, q, ['C003_SITE_STATE', 'C004_OPERATION'], '공간 상태', '현재 공간과 영업 상태를 알려주세요.');
  pushQuestion(steps, q, 'C010_ZONES');
  pushQuestion(steps, q, 'C011_SCOPE_LEVEL');
  pushQuestion(steps, q, 'C012_TRADES');
  pushQuestion(steps, q, 'C110_DEMO_SCOPE');
  pushQuestion(steps, q, 'C120_KITCHEN_EXHAUST');
  pushQuestion(steps, q, 'C130_RESTROOM_SCOPE');
  pushQuestion(steps, q, 'C140_ELECTRICAL');
  pushQuestion(steps, q, 'C150_HVAC');
  pushQuestion(steps, q, 'C160_SIGNAGE');
  pushQuestion(steps, q, 'C170_FIXTURE');
  if (detail) pushQuestion(steps, q, 'C180_PERMIT');
  pushQuestion(steps, q, 'C200_BRAND_STYLE');
  pushQuestion(steps, q, 'C300_BUDGET');
  pushGroup(steps, q, ['C320_START', 'C321_END', 'C322_VISIT'], '일정', '오픈/공사 희망 시점을 알려주세요.');
}

// ─── Main step builder ────────────────────────────────────────────────────────

export function buildSteps(
  answers: Answers,
  questions: SurveyQuestion[],
  _options: SurveyOption[],
  scopeMap: ScopeMapEntry[],
  _bundles: QuickBundle[],
): Step[] {
  const pt = (answers.Q000_PROJECT_TYPE as string) || '';
  const fm = (answers.Q000_FLOW_MODE as string) || '';

  const visCodes = getVisibleCodes(answers, questions, scopeMap);

  // Filter to visible + expr-passing questions
  const allQ = questions
    .filter(
      (q) =>
        (visCodes.has(q.question_code) || isEntryCode(q.question_code)) &&
        evalVisExpr(q.visible_if_expr, answers),
    )
    .sort((a, b) => (Number(a.step_no) || 0) - (Number(b.step_no) || 0));

  const steps: Step[] = [];

  // Phase 1: project type selection
  if (!pt) {
    pushQuestion(steps, allQ, 'Q000_PROJECT_TYPE');
    return steps;
  }

  // Phase 2: flow mode selection
  if (!fm) {
    pushQuestion(steps, allQ, 'Q000_FLOW_MODE');
    return steps;
  }

  // Phase 3: PRESET flow → bundle selection + tone + contact
  if (fm === 'PRESET') {
    const presetKey = pt === 'RESIDENTIAL' ? 'Q000_PRESET_RESI' : 'Q000_PRESET_COMM';
    pushQuestion(steps, allQ, presetKey);
    pushQuestion(steps, allQ, 'Q000_PRESET_TONE');
    steps.push(makeContactStep(allQ));
    return filterSteps(steps);
  }

  // Phase 4: SIMPLE or DETAIL flow
  const detail = fm === 'DETAIL';
  const coreQ = allQ.filter(
    (q) => !isEntryCode(q.question_code) && !isContactCode(q.question_code),
  );

  if (pt === 'RESIDENTIAL') {
    buildResidential(steps, coreQ, detail);
  } else {
    buildCommercial(steps, coreQ, detail);
  }

  steps.push(makeContactStep(allQ));
  return filterSteps(steps);
}

// ─── Bundle application ───────────────────────────────────────────────────────

export function applyBundle(
  bundleCode: string,
  answers: Answers,
  bundles: QuickBundle[],
): Answers {
  const bundle = bundles.find((b) => b.bundle_id === bundleCode);
  if (!bundle) return answers;

  const defaults = { ...(bundle.default_answers_json || {}) };

  // Derive area from recommended template filters
  const filters = bundle.recommended_template_filters_json || {};
  const areaBand = (filters.area_band || '').trim();
  if (areaBand) {
    const projectType = String(answers.Q000_PROJECT_TYPE || '').toUpperCase();
    if (projectType === 'COMMERCIAL') {
      defaults.C002_AREA = areaBand;
    } else {
      defaults.R002_AREA = areaBand;
    }
  }

  return { ...answers, ...defaults };
}

/** Get option label for a given question + option code. */
export function getOptionLabel(
  options: SurveyOption[],
  questionId: string,
  optionCode: string,
): string {
  if (!questionId || !optionCode) return '';
  const opt = options.find(
    (o) => String(o.question_id) === String(questionId) && String(o.option_code) === String(optionCode),
  );
  return opt?.option_label || '';
}

/** Get sorted options for a question. */
export function getOptionsForQuestion(
  options: SurveyOption[],
  questionId: string,
): SurveyOption[] {
  return options
    .filter((o) => o.question_id === questionId)
    .sort(sortBySortOrder);
}

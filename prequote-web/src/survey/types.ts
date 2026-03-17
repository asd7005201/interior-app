// ─── Survey DB Types ──────────────────────────────────────────────────────────

export interface SurveyQuestion {
  question_id: string;
  section_id: string;
  step_no: number;
  question_code: string;
  question_type:
    | 'SINGLE_SELECT'
    | 'MULTI_SELECT'
    | 'SHORT_TEXT'
    | 'LONG_TEXT'
    | 'PHONE'
    | 'EMAIL'
    | 'FILE';
  title: string;
  description: string;
  helper_text: string;
  placeholder: string;
  is_required: boolean;
  is_active: boolean;
  exposure_scope: string;
  visible_if_expr: string;
  option_source: string;
  allow_multiple: boolean;
  min_select: number;
  max_select: number;
  sort_order: number;
}

export interface SurveyOption {
  option_id: string;
  question_id: string;
  option_code: string;
  option_label: string;
  option_description: string;
  option_image_url: string;
  badge_text: string;
  is_active: boolean;
  sort_order: number;
}

export interface ScopeMapEntry {
  map_id: string;
  project_scope: string;
  trigger_question_id: string;
  trigger_values_json: string[];
  target_question_ids_json: string[];
  action_type: string;
}

export interface QuickBundle {
  bundle_id: string;
  property_scope: string;
  bundle_name: string;
  short_description: string;
  flow_mode: string;
  default_answers_json: Record<string, string | string[]>;
  locked_question_codes_json: string[];
  recommended_template_filters_json?: Record<string, string>;
  is_active: boolean;
  sort_order: number;
}

export interface SurveyConfig {
  questions: SurveyQuestion[];
  options: SurveyOption[];
  scopeMap: ScopeMapEntry[];
  bundles: QuickBundle[];
  uiContent: Array<Record<string, string>>;
  settings: Record<string, string>;
  app_url: string;
}

export type Answers = Record<string, string | string[]>;

export interface SurveyState {
  screen: 'landing' | 'survey' | 'submitting' | 'done';
  step: number;
  answers: Answers;
  files: File[];
  contact: {
    name: string;
    phone: string;
    email: string;
    method: string;
    address: string;
    note: string;
  };
}

export interface Step {
  questions: SurveyQuestion[];
  title: string;
  desc?: string;
  isBranchIntro?: boolean;
}

export interface SubmitResult {
  request_id: string;
  share_token: string;
  estimate: {
    min: number;
    max: number;
    adjustments: Array<Record<string, unknown>>;
    flags: Array<Record<string, unknown>>;
  };
  flags?: Array<Record<string, unknown>>;
}

export interface SerializedFile {
  file_name: string;
  file_mime: string;
  file_size: number;
  upload_type: 'CLIENT_METADATA_ONLY';
  sort_order: number;
}

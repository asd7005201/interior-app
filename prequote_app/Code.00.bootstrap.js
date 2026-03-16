/**
 * ============================================================
 * PREQUOTE APP server bootstrap
 * Public survey intake + estimate generation + admin operations.
 * File structure:
 *   Code.00-13.*.js split the former monolithic Code.js by domain section
 *   utils.js shared spreadsheet/settings helpers
 *   BaseLib.gs bridge to quote master app (material/template lookup)
 *
 * Sheet structure (prequote_DB):
 *   Settings, SurveyQuestions, SurveyOptions, SurveyLogic,
 *   SurveyTagRules, SurveyProfiles, QuestionScopeMap,
 *   QuickStartBundles, EstimateRules,
 *   Requests, RequestAnswers, RequestRecommendations,
 *   RequestFiles, RequestTimeline,
 *   MaterialsCache, TemplatesCache, SyncLog,
 *   UiContent, AdminUsers
 * ============================================================
 */
var APP_VERSION_ = "PREQUOTE_V6";
var BUILD_TAG_ = "2026-03-12T00:00:00Z";
var __HEADERS_CACHE_ = Object.create(null);
var __COLMAP_CACHE_ = Object.create(null);
var __REQUEST_ROW_INDEX_CACHE_ = Object.create(null);
var __SHEET_META_MEMO_ = Object.create(null);
var __SHEET_ROWS_MEMO_ = Object.create(null);
var __ROW_OBJECT_MEMO_ = Object.create(null);
var __COLUMN_MATCH_MEMO_ = Object.create(null);
var __REQUEST_DUPLICATE_LOOKUP_MEMO_ = Object.create(null);
var __ENSURED_SHEET_SCHEMA_MEMO_ = Object.create(null);
var STATIC_ROWS_CACHE_PREFIX_ = "PQ_STATIC_ROWS_V1";
var STATIC_ROWS_CACHE_TTL_SEC_ = 600;
var REQUEST_ROW_INDEX_SHEET_ = "RequestRowIndex";
var SYNC_STATE_SHEET_ = "SyncState";
var PERF_METRICS_SHEET_ = "PerfMetrics";
var REQUEST_POSTPROCESS_QUEUE_SHEET_ = "RequestPostprocessQueue";
var REQUEST_ROW_INDEX_VERSION_ = "2026-03-13-v1";
var SHEET_SCHEMA_CACHE_PREFIX_ = "PQ_SCHEMA_V2";
var SHEET_SCHEMA_CACHE_TTL_SEC_ = 600;
var SURVEY_CONFIG_CACHE_PREFIX_ = "PQ_SURVEY_CONFIG_V2";
var SURVEY_CONFIG_CACHE_TTL_SEC_ = 3600;
var RESULT_BOOTSTRAP_CACHE_PREFIX_ = "PQ_RESULT_BOOTSTRAP_V1";
var RESULT_BOOTSTRAP_CACHE_TTL_SEC_ = 900;
var SURVEY_SUBMIT_POSTPROCESS_TRIGGER_HANDLER_ = "runSurveySubmitPostprocessQueue_";
var SURVEY_SUBMIT_POSTPROCESS_TRIGGER_PREFIX_ = "PQ_SUBMIT_POSTPROCESS_TRIGGER";
var SURVEY_SUBMIT_POSTPROCESS_TRIGGER_DELAY_MS_ = 10000;
var SURVEY_SUBMIT_POSTPROCESS_BATCH_SIZE_ = 4;
var SURVEY_SUBMIT_POSTPROCESS_MAX_ATTEMPTS_ = 3;
var SYNC_KEY_MATERIALS_CACHE_ = "materials_cache";
var SYNC_KEY_TEMPLATES_CACHE_ = "templates_cache";
var SURVEY_CLIENT_QUESTION_FIELDS_ = [
  "question_id",
  "question_code",
  "question_type",
  "title",
  "description",
  "placeholder",
  "helper_text",
  "step_no",
  "sort_order",
  "is_required",
  "visible_if_expr",
  "exposure_scope"
];
var SURVEY_CLIENT_OPTION_FIELDS_ = [
  "question_id",
  "option_code",
  "option_label",
  "option_description",
  "sort_order"
];
var SURVEY_CLIENT_SCOPE_FIELDS_ = [
  "trigger_question_id",
  "trigger_values_json",
  "target_question_ids_json"
];
var SURVEY_CLIENT_BUNDLE_FIELDS_ = [
  "bundle_id",
  "bundle_label",
  "project_type",
  "sort_order",
  "default_answers_json",
  "locked_question_codes_json",
  "recommended_tags_json",
  "recommended_template_filters_json"
];
var SURVEY_CLIENT_UI_CONTENT_FIELDS_ = [
  "content_key",
  "content_type",
  "title",
  "body",
  "content_value",
  "sort_order"
];
var REQUEST_STATUS_OPTIONS_ = ["NEW", "CONTACTED", "SCHEDULED", "CONVERTED", "CLOSED"];
var ADMIN_REVIEW_STATUS_OPTIONS_ = ["AUTO", "IN_REVIEW", "OVERRIDDEN"];
var REQUEST_PRIORITY_OPTIONS_ = ["LOW", "NORMAL", "HIGH", "URGENT"];
var REQUEST_WORK_STATE_OPTIONS_ = ["NEW", "REVIEWING", "WAITING_CUSTOMER", "READY_FOR_QUOTE", "DONE"];
var REQUEST_ADMIN_REVIEW_SHEET_ = "RequestAdminReviews";
var REQUEST_ADMIN_REVIEW_HEADERS_ = [
  "review_id",
  "request_id",
  "edited_at",
  "actor_type",
  "actor_id",
  "review_status",
  "override_estimate_min",
  "override_estimate_max",
  "override_recommendations_json",
  "override_answers_json",
  "estimate_reason",
  "review_note",
  "review_summary",
  "final_estimate_min",
  "final_estimate_max",
  "final_recommendation_count",
  "source_label",
  "is_current",
  "linked_quote_id",
  "linked_quote_status",
  "slack_notified_at",
  "created_at",
  "updated_at",
  "payload_json"
];
var REQUEST_RECOMMENDATION_HEADERS_ = [
  "request_id",
  "rec_id",
  "rec_type",
  "rank",
  "template_id",
  "material_id",
  "trade_code",
  "material_type",
  "title",
  "subtitle",
  "reason_text",
  "score",
  "price_hint_min",
  "price_hint_max",
  "matched_tags_json",
  "snapshot_json",
  "is_selected",
  "created_at",
  "is_visible",
  "is_deleted",
  "is_manual",
  "image_url",
  "image_file_id",
  "brand",
  "spec",
  "source_type",
  "source_ref_id",
  "source_ref_key",
  "review_id",
  "updated_at",
  "hidden_at",
  "deleted_at",
  "note"
];
var REQUEST_ASSIGNMENT_HEADERS_ = [
  "assignment_id",
  "request_id",
  "assigned_to_name",
  "assigned_to_email",
  "assigned_to_slack_id",
  "assigned_by",
  "assigned_at",
  "is_active",
  "note",
  "payload_json"
];
var REQUEST_TASK_HEADERS_ = [
  "task_id",
  "request_id",
  "task_type",
  "task_label",
  "task_status",
  "priority",
  "assignee_name",
  "assignee_email",
  "assignee_slack_id",
  "due_at",
  "remind_at",
  "created_at",
  "updated_at",
  "completed_at",
  "note",
  "payload_json"
];
var REQUEST_CHECKLIST_TEMPLATE_HEADERS_ = [
  "template_id",
  "project_type",
  "checklist_code",
  "checklist_label",
  "is_required",
  "sort_order",
  "note"
];
var REQUEST_CHECKLIST_ITEM_HEADERS_ = [
  "item_id",
  "request_id",
  "template_id",
  "checklist_code",
  "checklist_label",
  "is_required",
  "is_completed",
  "completed_at",
  "completed_by",
  "note",
  "sort_order",
  "payload_json"
];
var REQUEST_NOTIFICATION_LOG_HEADERS_ = [
  "notification_id",
  "request_id",
  "channel",
  "event_type",
  "dedup_key",
  "recipient_key",
  "recipient_value",
  "status",
  "created_at",
  "sent_at",
  "last_error",
  "payload_json"
];
var REQUEST_QUOTE_DRAFT_LINK_HEADERS_ = [
  "link_id",
  "request_id",
  "review_id",
  "quote_id",
  "status",
  "created_at",
  "updated_at",
  "source_app",
  "target_app",
  "open_url",
  "note",
  "payload_json"
];
var DASHBOARD_SAVED_VIEW_HEADERS_ = [
  "view_id",
  "user_key",
  "view_name",
  "filter_json",
  "sort_json",
  "is_default",
  "created_at",
  "updated_at",
  "note"
];
var REQUEST_ROW_INDEX_HEADERS_ = [
  "request_id",
  "requests_row_no",
  "answers_row_nos_json",
  "timeline_row_nos_json",
  "files_row_nos_json",
  "admin_review_row_nos_json",
  "assignments_row_nos_json",
  "tasks_row_nos_json",
  "checklist_item_row_nos_json",
  "quote_link_row_nos_json",
  "latest_review_id",
  "latest_review_row_no",
  "updated_at",
  "index_version"
];
var SYNC_STATE_HEADERS_ = [
  "sync_key",
  "source_app",
  "source_spreadsheet_id",
  "shared_master_version",
  "last_synced_at",
  "last_status",
  "read_count",
  "write_count",
  "duration_ms",
  "note"
];
var PERF_METRICS_HEADERS_ = [
  "metric_id",
  "measured_at",
  "metric_name",
  "duration_ms",
  "request_id",
  "actor_type",
  "actor_id",
  "context_json",
  "success_yn",
  "note"
];
var QUOTE_DB_PERF_METRICS_HEADERS_ = [
  "metric_id",
  "measured_at",
  "metric_name",
  "duration_ms",
  "entity_id",
  "actor_type",
  "actor_id",
  "context_json",
  "success_yn",
  "note"
];
var REQUEST_POSTPROCESS_QUEUE_HEADERS_ = [
  "job_id",
  "request_id",
  "job_type",
  "status",
  "attempt_count",
  "created_at",
  "scheduled_at",
  "started_at",
  "processed_at",
  "last_error",
  "payload_json"
];
var REQUEST_REQUIRED_COLUMNS_ = [
  "priority",
  "consultant_email",
  "memo",
  "internal_memo",
  "lead_score",
  "quote_conversion_status",
  "linked_quote_id",
  "assignee_name",
  "assignee_email",
  "assignee_slack_id",
  "next_action_type",
  "next_action_note",
  "next_action_due_at",
  "reminder_at",
  "work_state",
  "latest_review_id",
  "latest_review_at",
  "final_estimate_min",
  "final_estimate_max",
  "final_estimate_source",
  "final_recommendation_count",
  "duplicate_phone_key",
  "duplicate_email_key",
  "duplicate_customer_count",
  "review_status_label",
  "quote_draft_status",
  "answers_snapshot_json"
];
var REQUEST_TIMELINE_REQUIRED_COLUMNS_ = [
  "event_label",
  "dedup_key",
  "notify_channel",
  "notify_status",
  "linked_ref_id"
];
var REQUEST_FILE_REQUIRED_COLUMNS_ = [
  "thumbnail_url",
  "caption",
  "is_primary"
];
var QUOTE_DB_REQUIRED_QUOTE_COLUMNS_ = [
  "source_request_id",
  "source_review_id",
  "source_app",
  "assignee_name",
  "assignee_slack_id",
  "draft_stage",
  "prequote_summary_json",
  "recommended_material_ids_json",
  "recommended_materials_snapshot_json"
];
var QUOTE_DB_REQUIRED_ITEM_COLUMNS_ = [
  "source_request_id",
  "source_review_id",
  "source_request_rec_id",
  "imported_from_prequote",
  "source_material_id",
  "source_material_image_url"
];
var QUOTE_DRAFT_QUEUE_HEADERS_ = [
  "queue_id",
  "request_id",
  "review_id",
  "created_at",
  "status",
  "customer_name",
  "contact_phone",
  "project_type",
  "assignee_name",
  "payload_json",
  "linked_quote_id",
  "processed_at",
  "last_error"
];
var QUOTE_DRAFT_ITEM_HEADERS_ = [
  "queue_id",
  "request_id",
  "review_id",
  "seq",
  "material_id",
  "name",
  "brand",
  "spec",
  "image_file_id",
  "image_url",
  "qty",
  "unit",
  "unit_price",
  "amount",
  "note",
  "payload_json"
];
var QUOTE_IMPORT_LOG_HEADERS_ = [
  "log_id",
  "created_at",
  "request_id",
  "review_id",
  "quote_id",
  "status",
  "message",
  "payload_json"
];
var STATUS_LABEL_MAP_ = {
  NEW: "신규 접수",
  CONTACTED: "1차 연락",
  SCHEDULED: "상담 예정",
  CONVERTED: "실견적 전환",
  CLOSED: "종료"
};
var REVIEW_STATUS_LABEL_MAP_ = {
  AUTO: "자동 산출",
  IN_REVIEW: "관리자 검토중",
  OVERRIDDEN: "관리자 수정완료"
};
var EVENT_TYPE_LABEL_MAP_ = {
  STATUS_CHANGE: "상태 변경",
  ADMIN_NOTE: "관리자 메모",
  REVIEW_SAVE: "검토 저장",
  RECALCULATE: "자동 재계산",
  ASSIGNMENT_UPDATE: "담당자 지정",
  TASK_UPDATE: "다음 액션 갱신",
  CHECKLIST_SAVE: "체크리스트 저장",
  BULK_UPDATE: "일괄 처리",
  QUOTE_DRAFT_CREATED: "견적 초안 생성",
  SLACK_NOTIFY: "슬랙 알림",
  RESULT_VIEWED: "결과 확인"
};
var SEVERITY_LABEL_MAP_ = {
  HIGH: "높음",
  MEDIUM: "중간",
  LOW: "낮음"
};
var PRIORITY_LABEL_MAP_ = {
  LOW: "낮음",
  NORMAL: "보통",
  HIGH: "높음",
  URGENT: "긴급"
};
var WORK_STATE_LABEL_MAP_ = {
  NEW: "신규",
  REVIEWING: "검토 중",
  WAITING_CUSTOMER: "고객 회신 대기",
  READY_FOR_QUOTE: "실견적 전환 대기",
  DONE: "처리 완료"
};
var TASK_STATUS_LABEL_MAP_ = {
  OPEN: "대기",
  IN_PROGRESS: "진행 중",
  DONE: "완료",
  CANCELLED: "취소"
};
var ANSWER_GROUP_LABEL_MAP_ = {
  ENTRY: "프로젝트 설정",
  BASIC: "기본 정보",
  SCOPE: "공간 정보",
  TRADE: "공정 범위",
  STYLE: "취향",
  LIFE: "생활 정보",
  RISK: "현장 리스크",
  BUDGET: "예산",
  SCHEDULE: "일정",
  CONTACT: "연락 선호",
  OTHER: "기타"
};
var ANSWER_GROUP_ORDER_MAP_ = {
  ENTRY: 10,
  BASIC: 20,
  SCOPE: 30,
  TRADE: 40,
  STYLE: 50,
  LIFE: 60,
  RISK: 70,
  BUDGET: 80,
  SCHEDULE: 90,
  CONTACT: 100,
  OTHER: 110
};

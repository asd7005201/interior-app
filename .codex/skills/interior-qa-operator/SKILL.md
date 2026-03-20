---
name: interior-qa-operator
description: Use when the user wants URL-based QA for the quote or prequote apps, including functional checks, UI/UX review, copy review, visual/design inspection, evidence capture, and a prioritized defect report without involving executive debate.
---

# Interior QA Operator

Use this skill for execution-focused QA only. Do not run persona debate here. This skill produces evidence and a defect list.

## Scope

- `quote_app`: admin quote workflow, materials, templates, dashboard, catalog, edit/view flows
- `prequote_app`: landing, survey, result, complete, admin/survey builder, versions, requests, sync
- UI review, UX friction review, copy/typo review, redundant element review
- Evidence capture: screenshots, console errors, network failures, reproduction steps

## Inputs to collect first

- Target URL
- App type: `quote_app` or `prequote_app`
- Auth requirement and test account if login is required
- Environment: prod, staging, test, or Apps Script deployment alias
- Must-check scenarios, ideally 3 to 5
- Whether code changes are allowed, or report-only

If any of these are missing, make the smallest safe assumption and state it.

## Workflow

1. Restate the target, environment, and app type in 2 to 4 lines.
2. Read [references/qa-checklists.md](references/qa-checklists.md) and use only the relevant section.
3. Build a compact scenario list:
   - smoke
   - core flow
   - edge cases
   - UI/UX and copy pass
4. Execute the checks with available tools.
5. Capture evidence for every failed or suspicious case.
6. Classify findings by severity:
   - `P0`: app blocked, data corruption risk, submit impossible
   - `P1`: major flow broken, misleading result, strong trust issue
   - `P2`: workaround exists, visual/UX defect, confusing copy
   - `P3`: polish issue, typo, spacing, low-impact redundancy
7. Recommend the smallest valid fix.
8. If the user asked for fixes and the repository contains the relevant code, implement the fix after the report.

## Review rules

- Distinguish evidence from inference.
- Prefer reproducible issues over vague impressions.
- Do not claim business decisions or roadmap priorities here.
- Do not escalate to persona debate unless the user explicitly asks for the council skill.
- For Apps Script UIs, test both desktop and a narrow mobile viewport when the page is public-facing.
- For copy review, check labels, placeholders, empty states, validation messages, and confirmation text.
- For design review, check hierarchy, spacing consistency, CTA prominence, visual noise, and trust cues.

## Output format

Produce sections in this order:

1. `Execution Summary`
2. `Environment and Assumptions`
3. `Scenario Coverage`
4. `Findings`
5. `Recommended Fix Order`
6. `Open Risks`

Each finding must include:

- ID
- Severity
- Area
- Reproduction
- Expected
- Actual
- Evidence
- Suggested fix

## Bundled resources

- Report skeleton: [assets/report-template.md](assets/report-template.md)
- App-specific checklist: [references/qa-checklists.md](references/qa-checklists.md)

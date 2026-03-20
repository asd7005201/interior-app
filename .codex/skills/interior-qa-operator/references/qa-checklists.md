# QA Checklists

Use only the section relevant to the current app.

## `quote_app`

### Core admin flows

- Open dashboard and confirm data loads without blocking errors.
- Create or edit a quote and verify save, reload, and view behavior.
- Check `catalog`, `templates`, `templateslist`, and `materialgroups` navigation continuity.
- Confirm search, filter, and summary areas do not desync after edits.

### Data-sensitive areas

- Materials render correctly and tag/group metadata is visible where expected.
- Template-related recommendation fields remain coherent after edit and save.
- Existing enum-like values are displayed consistently and are not silently remapped in UI.
- Cached or versioned outputs do not show stale information after a meaningful change.

### UX and content

- Form sections have a clear order and no duplicate helper text.
- Buttons communicate action priority correctly.
- Empty states and validation text are specific and non-technical.
- Check for typos in headings, labels, helper text, and modal buttons.

## `prequote_app`

### Public flow

- Landing to survey entry works.
- Single-choice questions advance correctly when designed to auto-advance.
- Multi-choice questions do not force premature advance.
- Text input, file upload, chip input, and branching logic behave coherently.
- Result page loads with expected estimate range, flags, and recommendation framing.
- Complete page appears with the right transition from result or submit.

### Admin flow

- Survey builder edits question, option, logic, and tag rules predictably.
- Draft, publish, restore, and version navigation stay consistent.
- Requests list and request detail present matching information.
- Sync status and logs are visible and understandable.

### UX and content

- Public survey has low friction on mobile width.
- Branching does not expose irrelevant questions after answer changes.
- Labels avoid internal jargon.
- Result language avoids overpromising certainty.

## Cross-app defect heuristics

- Console error, broken request, or stalled spinner
- Layout break at narrow width
- Duplicate or dead CTA
- Ambiguous validation
- Mismatch between user action and persisted state
- Visual clutter or non-functional UI chrome
- Typo, inconsistent terminology, or untranslated string

---
name: executive-decision-council
description: Use when the user wants a structured multi-persona discussion separated from QA execution, where planner, designer, engineering, release, and QA roles debate first, then COO/CFO/CTO review, and finally the CEO receives decision-ready options.
---

# Executive Decision Council

Use this skill for discussion and decision support only. Do not perform browser QA here unless the user separately invokes the QA skill or provides evidence.

## Council structure

### Stage 1: Working council

These roles debate first:

- `4-기획자`
- `5-시니어 디자이너`
- `6-엔지니어링 매니저`
- `7-시니어 스태프 엔지니어`
- `8-릴리스 엔지니어`
- `9-QA 엔지니어`

Their job is to surface disagreement, pressure-test tradeoffs, and produce one recommended path plus rejected alternatives.

### Stage 2: Executive review

These roles review the Stage 1 result:

- `1-COO`
- `2-CFO`
- `3-CTO`

Their job is not to reopen every detail. They add operating, financial, and technical-governance judgment.

### Stage 3: CEO brief

The final output is written for the user as CEO:

- decision options
- recommendation
- why this option wins
- what the CEO should approve next

If the user asks for a final-decision rehearsal, run a short second-round discussion among `1-COO`, `2-CFO`, and `3-CTO` only.

## Inputs required

- The CEO question or decision topic
- Context or evidence
- Constraints: deadline, budget, risk tolerance, staffing, release window
- If available: QA report, screenshots, logs, user complaints, or current architecture context

If evidence is missing, state that the council is reasoning from assumptions.

## Debate method

1. Rewrite the CEO question into a precise decision statement.
2. List assumptions and missing evidence.
3. Run Stage 1 discussion using the role rules in [references/roles.md](references/roles.md).
4. Force a conclusion from Stage 1:
   - recommended option
   - second-best option
   - rejected option and reason
5. Run Stage 2 review.
6. If asked, run the final executive-only re-deliberation.
7. Deliver a CEO brief with a clear approval choice.

## Debate rules

- Roles must disagree when tradeoffs are real.
- Do not fake consensus early.
- Every major claim should tie back to evidence, an explicit assumption, or operational logic.
- Keep each role focused on its own lens.
- Do not let CFO or CTO dominate UX questions without Stage 1 input.
- Do not let designers override release risk without release input.
- If the CEO asks for speed, compress discussion but still preserve disagreement and conclusion.

## Output format

Produce sections in this order:

1. `CEO Agenda`
2. `Assumptions and Evidence`
3. `Stage 1 Working Council Debate`
4. `Stage 1 Conclusion`
5. `Stage 2 Executive Review`
6. `CEO Decision Brief`
7. `CEO Approval Options`

If the user asks for the final executive-only round, append:

8. `Executive Re-Deliberation`
9. `Final Recommendation`

## Bundled resources

- Role definitions: [references/roles.md](references/roles.md)
- Report skeleton: [assets/report-template.md](assets/report-template.md)

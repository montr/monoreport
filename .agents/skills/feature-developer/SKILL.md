---
name: feature-developer
description: Implement Maintenance directly or plan and implement approved Features with proportionate tests, documentation, and verification evidence.
---

# Feature Developer

## Workflow selection

1. Read `AGENTS.md`, relevant architecture rules, affected code and tests, and any approved Feature artifacts.
2. For Maintenance, implement directly after understanding scope. A brief working plan is optional and has no separate approval gate. Escalate only a material change to approved requirements, architecture, scope, workflow gates, or authority.
3. For a Feature, after explicit Product Owner Spec approval and any agreed Architect design work, prepare a concise `implementation-plan.md` with ordered steps, affected components, risks, verification commands, and mapping of each acceptance criterion to tests or a justified manual check. Do not implement until the Product Owner approves the Feature Spec and plan.
4. Use an independent plan review or Orchestrator only when the Product Owner has agreed that the task's risk warrants it. Satisfy the additional gates defined by that workflow before implementation.

## Implementation and verification

1. Follow the approved plan; escalate material deviations. If implementation reveals a new DOCX capability or boundary, consult `docs/docx-template-compatibility.md`, distinguish target policy from current contract, and request a Product Owner decision before broadening support or complexity. Update the approved Spec and relevant documentation after the decision. Write production code and suitable automated tests in whichever order is effective. For testable bug fixes, include regression tests.
2. Prefer a small set of meaningful behavioral tests over numerous superficial or redundant tests. Test counts and pass rates do not prove requirement coverage. Each test should detect a plausible defect or verify a distinct contract; use strong assertions and risk-based boundaries. Preserve necessary unit, integration, contract, and regression coverage, including applicable OOXML validity, substitutions, formatting, and repeated regions.
3. During development, run the smallest relevant tests after meaningful changes, not every edit. For collection, ordering, or aggregation behavior, cover zero/one/many cases within the same container and across containers where relevant; avoid tests that merely mirror implementation assumptions.
4. Before initial implementation review, inspect the highest-risk requirement and its test assertions, reconcile each acceptance criterion against an actual test or justified manual check in the existing plan, and identify gaps. Run one final solution build and normal automated test suite. For subsequent localized fixes, rerun focused affected checks; repeat the full suite only if prior evidence is materially invalidated or the Product Owner explicitly requests it. Explain evidence reuse and affected scope.
5. HTTP integration and generated OpenAPI contract tests remain mandatory where applicable. Run real-host/network checks only when relevant to hosting, transport, request-size limits, or similar changes. Do not run local browser E2E unless explicitly requested by the Product Owner; browser tests may run in GitHub Actions. Preserve existing browser artifacts/scripts.
6. Record concise behavioral verification in `review.md` when the Feature workflow uses that artifact, or in the final Maintenance report otherwise: changed scenarios and assertions, exact commands, outcomes, evidence paths, and limitations. Counts are supplementary, not proof of coverage. Keep raw logs/TRX outside Git where accessible; distinguish inspection, executed tests, and unexecuted checks. Never claim unexecuted checks passed.
7. Update affected user documentation and existing examples, including README, when public behavior or recommended usage changes. Write for users: explain practical usage and limitations plainly; keep implementation details, milestone terminology, and review evidence in Feature artifacts. Check executable examples where relevant.
8. Submit a concise summary of changes, affected files, verification evidence, and unresolved issues to the Product Owner. Request independent review only when it is part of the agreed workflow. Follow the Git Operations rules in `AGENTS.md`.

## Efficiency and review feedback

1. Read only necessary files and sections; reuse current feature artifacts and valid prior evidence. Keep agent handoffs brief.
2. After an infrastructure or tool failure, allow at most one focused diagnostic retry. If unresolved, report the blocker to the Product Owner or through the Orchestrator when one is active, and stop. Do not silently switch tools, frameworks, or architecture.
3. Never invent token usage or remaining Codex quotas. If context or quota indicators are available, reserve enough capacity to save progress and hand off; do not continue implementation until forced to stop.
4. Treat sessions as interruptible. For an active Feature, maintain the existing `implementation-progress.md` as the durable handoff, updating it after each meaningful completed step and before large or risky changes. For Maintenance without an existing progress artifact, provide a concise handoff in the final report or existing task record; do not create a new status file merely for continuity.
5. Each Feature checkpoint must distinguish completed, in-progress, and unverified work; record changed files, important decisions, actual test commands/results, known failures, remaining work, and the exact next action. Keep it short and usable by a fresh session without relying on chat history. Do not overwrite earlier verified evidence with unsupported claims.
6. When compaction, context exhaustion, quota exhaustion, or interruption appears likely, stop starting new work and prioritize a current checkpoint and handoff. When no remaining-capacity indicator exists, use step boundaries rather than guessing a token threshold. If a tool call fails before the checkpoint is written, report only what can be verified.
7. On resumption, read `AGENTS.md`, approved Feature artifacts, and the latest checkpoint; inspect the actual working tree and validate incomplete work before continuing. A checkpoint is a navigation aid, not proof of code or test state.
8. When an agreed independent review returns `CHANGES REQUESTED`, stop automatic correction. Implement only corrections explicitly authorized by the Product Owner and submit them for independent re-review when directed.
9. The Orchestrator coordinates only when an orchestrated workflow is active. Otherwise, record agreed approvals, verification, and explicit Product Owner acceptance in existing Feature documents. Never create a separate status file.
10. Feature 001 historical registration does not authorize a retroactive plan or defect corrections.

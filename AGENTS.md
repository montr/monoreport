# Project Rules

## Development Workflow

1. Execute bug fixes, refactoring, documentation changes, and trivial UI or behavior adjustments directly without mandatory task classification, Feature artifacts, plans, or dedicated agents. A testable bug fix should normally include a regression test.
2. For new user-facing or programmatic capabilities beyond trivial adjustments, propose a Feature Spec and ask the Product Owner whether to create one before substantial implementation. Respect an explicit Product Owner decision to proceed without a spec, subject to the high-risk rule. Keep Feature artifacts under `.docs/features/<NNN>-<slug>/` with sequential, never-reused IDs.
3. Feature Specs document the goal and rationale, requirements, acceptance criteria, constraints, and significant decisions, including relevant rejected alternatives. Record architectural decisions with cross-feature impact in ADRs when appropriate. An Implementation Plan and independent review are optional, not default gates. Do not require repeated approval for work the Product Owner has already explicitly authorized.
4. Before implementing high-risk changes affecting architecture, security or permissions, persistence or migration, concurrency, irreversible operations, public contracts, release workflows, binary or document integrity, or multiple components, briefly explain the specific risks, recommend proportionate planning or independent review, and obtain Product Owner confirmation of the approach. Obtain explicit approval before materially changing approved requirements, architecture, or scope. Avoid unnecessary workflow gates for low-risk work.
5. Codex implements tasks directly by default. Architect, Developer, Orchestrator, and Reviewer Skills remain available in the repository but are not automatically invoked. Use them only when explicitly requested or agreed with the Product Owner. If an independent review is agreed, follow its Skill's review and stop rules; do not silently treat unresolved blocking findings as passed.
6. For an active Feature, maintain its existing `feature-status.json` contract containing only `phase`. When resuming, reconcile the recorded phase with the Feature documents, Git, and actual implementation. Do not invent approvals, verification, or completion. Keep state updates proportional to the task; an Orchestrator is not required to maintain status.
7. Minimize context loading. At the start of a task, inspect only relevant project rules, code, and directly applicable Feature or ADR documents. Use the `.docs/features/` directory listing to discover Features; do not read all historical Feature Specs, plans, reviews, or statuses by default. Open older documents only when the task, a dependency, or a specific decision requires them. Prefer current requirements and decisions over historical execution logs. Do not load every Skill merely to discover the available workflow.
8. Preserve meaningful behavior-focused verification, disclose failed or unexecuted checks, update affected documentation, and protect existing user changes. Keep checks proportional to change risk.
9. Maintain `.docs/features/backlog.md` as a concise list of unapproved future work. Add or reprioritize items with Product Owner approval; remove completed, obsolete, or duplicate items. Do not assign Feature IDs or copy Feature Specs into the backlog. Preserve deferred requirements for Product Owner review rather than silently discarding them.
10. For UI work, seek early Product Owner visual and behavioral feedback. Browser automation requires explicit Product Owner approval.
11. Do not modify controlling project Skills without explicit Product Owner authorization. Follow the Git Operations rules below.

## Git Operations

1. The Product Owner exclusively manages Git commits and repository publishing operations.
2. Agents must not stage files, create commits, amend or rewrite commits, push, merge, rebase, delete branches, create tags, or publish releases without separate, explicit Product Owner authorization for the specific action.
3. Feature approval, implementation approval, and acceptance do not imply authorization for any Git operation listed above.
4. Agents may inspect Git status, history, and diffs. Preserve existing staged and unstaged user changes.

## Coding Guidelines

1. Implement services as instances with externally supplied constructor dependencies. Do not use public/internal static service methods, construct service dependencies internally, or introduce parameterless dependency defaults. Private static helpers are allowed; constants-only classes must not contain service logic.
2. Store per-operation data in a separate context, not reusable service state.
3. Prefer primary constructors with null validation (`?? throw`); use regular constructors when required by correct initialization.
4. Isolate DOCX-specific logic from shared APIs in anticipation of other formats such as XLSX, without premature abstractions or unapproved formats.

## Test Quality

1. Favor meaningful behavior-focused tests over raw test count or pass-rate targets. Verify acceptance criteria with assertions capable of detecting plausible defects; avoid redundant tests and coverage-number chasing.
2. Code coverage metrics may inform gap analysis but are not quality targets. Before adding tests, inspect relevant existing coverage and add tests only for meaningful missing behavior or regression risks. Do not introduce coverage tooling or thresholds without separate Product Owner approval.
3. Run the narrowest relevant tests first. Expand to broader suites when change impact or failure evidence warrants it. Avoid repeated full-suite runs without a reason; do not omit necessary verification solely to save time or tokens.
4. Keep command output concise; inspect verbose logs only for failures or diagnostics. Do not run browser automation or unusually resource-intensive integration/end-to-end suites without explicit Product Owner approval. If a check is skipped, explain why and what remains unverified.

## Documentation Rules

1. Write Skills, project rules, specifications, plans, decisions, and other project documentation in English; Product Owner discussions may use another language.
2. Use explicitly sequential numbered lists, not bullets; keep references unambiguous.
3. Do not add code-formatting examples unless the Product Owner requests them.

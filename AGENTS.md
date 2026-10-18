# Project Rules

## Development Workflow

1. Distinguish Maintenance (bug fixes, refactoring, documentation, and trivial changes) from Features (nontrivial new or changed capabilities). Handle Maintenance directly with proportionate verification. For Features, work from a Product Owner-approved Feature Spec and Implementation Plan before implementation. Keep Feature artifacts under `.docs/features/<NNN>-<slug>/` with sequential, never-reused IDs.
2. The Product Owner selects the workflow, agent roles, and review gates for each Feature. Do not automatically invoke Architect, Developer, Reviewer, or Orchestrator roles. Follow the relevant Skill when a role is explicitly selected. The Architect may resolve architectural decisions before the Developer prepares the Implementation Plan; neither role may presume Product Owner approval.
3. Obtain Product Owner approval before materially changing approved requirements, product scope, architecture, or workflow. Raise consequential ambiguities with clear alternatives; do not interrupt for routine implementation choices. Record approved decisions and acceptance in existing Feature documents. Do not create separate Feature status files or invent approvals.
4. Read only relevant rules, code, Feature documents, and ADRs. Reuse valid evidence and keep verification proportional to risk; disclose failed or unexecuted checks. Protect existing user changes.
5. Maintain `.docs/features/backlog.md` as a concise list of unapproved future work, with Product Owner approval for additions or reprioritization. Do not assign Feature IDs in the backlog.
6. Seek early Product Owner feedback for UI changes. Browser automation requires explicit Product Owner approval. Do not modify controlling project Skills without explicit Product Owner authorization.

## Product Boundaries and DOCX Compatibility

1. Read `docs/docx-template-compatibility.md` when designing, implementing, or reviewing DOCX template behavior. It is a target compatibility policy for a future simplification Feature until adopted by an approved Feature Spec; do not mistake it for current implemented behavior or retroactively change accepted Feature criteria.
2. When a new case raises a product-scope decision, distinguish an actual defect against the approved contract from unsupported behavior or a proposed extension. Explain practical user value, simpler alternatives (including actionable rejection), implementation cost, and document-integrity risks. Ask the Product Owner before expanding supported behavior or architectural complexity.
3. Record approved decisions in the relevant Feature Spec and update compatibility documentation when its intended policy changes. Developers and Reviewers must not independently broaden the contract. Escalate material questions, not routine implementation details.
4. Preserve valid supported documents and detect corrupt output or data loss. Do not pursue exhaustive OOXML feature coverage or speculative edge-case frameworks without an approved product need.

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

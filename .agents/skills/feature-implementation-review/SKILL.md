---
name: feature-implementation-review
description: Independently review a feature implementation plan against its approved specification, architecture decisions, repository reality, and existing contracts; find omissions, unnecessary scope, complexity, risks, and test gaps before implementation. Use when asked to review or approve a developer-owned implementation plan.
---

# Feature Implementation Review

## Purpose

Act as an independent, skeptical engineering reviewer. Determine whether a proposed implementation plan is **correct, minimal, feasible, testable, and ready for owner approval**. A detailed plan is not necessarily a good plan. Prioritize shipping the requested feature safely over creating infrastructure for hypothetical future requirements.

## Inputs and authority

1. Read the repository's `AGENTS.md`, applicable project instructions, and the selected feature's specification, architecture decisions/assessment, research evidence, roadmap, and implementation plan. Inspect actual code, public contracts, tests, dependencies, and relevant history where accessible. Do not treat assertions in the plan as verified repository facts.
2. Use the following precedence: explicit owner decisions and approved feature specification; accepted architecture decisions and existing compatibility contracts; repository behavior; research evidence; developer implementation plan. Flag contradictions rather than silently reconciling them. Do not reinterpret a proposed plan as an approved product decision.
3. If an input is missing, say which conclusions cannot be verified. Do not fabricate tests, measurements, APIs, or approvals.

## Review procedure

1. **Requirements traceability:** map every mandatory behavior, acceptance scenario, edge case, error rule, and non-goal to a concrete implementation milestone and verification. Find missing, contradicted, or invented requirements.
2. **Scope control:** identify each new public API, model, dependency, transport, configuration, limit, compatibility behavior, abstraction, service, and operational mechanism. Classify it as (a) explicitly required/approved, (b) necessary implementation detail, or (c) optional proposal. Challenge (c) with a simpler alternative; never silently approve it.
3. **Simplicity and cost:** inspect the proposed components and milestones for duplication, premature generalization, unnecessary shared-code refactoring, overly broad fidelity support, speculative compatibility, and excessive testing infrastructure. Prefer existing patterns, a minimal vertical implementation architecture, and a **working end-to-end feature slice early**. Do not recommend cutting required functionality or safeguards to improve estimates.
4. **Technical feasibility:** validate API signatures, constructor/DI behavior, package/library limitations, algorithmic invariants, state ownership, deterministic behavior, failure atomicity, format fidelity, and performance assumptions against source code and evidence. For formulas/references or other complex algorithms, identify concrete counterexamples and negative cases rather than generic concerns.
5. **Compatibility:** enumerate user-visible behavior changes in existing formats, Core, CLI, HTTP, schemas, errors, diagnostics, and resource limits. Demand regression evidence. Treat compatibility-impacting additions as proposals requiring explicit approval.
6. **Verification:** check that each milestone has independently observable completion criteria, focused automated tests, negative/boundary tests, and any necessary external/manual acceptance. Distinguish planned tests from executed evidence. Do not invent a blanket requirement for more tests or tooling.
7. **Effort and sequencing:** examine estimates as uncertain hypotheses; find critical-path blockers, tasks that can be deferred, and stages too large for one agent session. Suggest smaller independently verifiable increments without turning the work into parallel model benchmarking or a new process bureaucracy.
8. **Decision gate:** identify which owner decisions truly block implementation and which can be deferred until the relevant milestone. Recommend approve, approve with explicit conditions, or revise and review again. Do not grant approval on behalf of the owner.

## Persistent review artifact (default)

1. **Always save the review to the repository by default.** When reviewing a plan at `<feature-dir>/implementation-plan.md`, create or update `<feature-dir>/implementation-review.md`. For a differently named plan, place `implementation-review.md` alongside it unless the owner provides another path. The artifact is the durable handoff between sessions; a chat-only review is insufficient.
2. Create the review directory only if it already corresponds to the selected feature; do not create unrelated directories. If an existing review file is present, update it to reflect the current plan and owner decisions, preserving useful unresolved findings and noting superseded findings. Never overwrite unrelated user content. If the destination is ambiguous or writing is disallowed, ask for the target path or permission; explain clearly if persistence failed.
3. Write the artifact in English and summarize the verdict, highest-priority findings, and saved path in the owner's preferred chat language. Include review date, plan path, and which owner decisions were applied; do not claim a file was saved unless it was actually written and verified.
4. Saving the review artifact is explicitly permitted even though implementation and plan edits are prohibited. Do not stage or commit the review file unless explicitly requested.

## Reporting format

Write a concise, evidence-based review **in English** for the repository artifact, with numbered lists (no unordered lists):

1. **Verdict:** Ready / Ready with conditions / Needs revision; 2–4 sentences.
2. **Findings:** numbered, ordered by severity (`Blocker`, `Major`, `Minor`); for each give the exact file/section or source location, why it matters, and the smallest actionable fix. Distinguish verified defects from risks or unanswered questions.
3. **Unnecessary scope / simplifications:** numbered proposals with what can be removed, deferred, or reused; specify trade-offs and whether the feature specification would still be met.
4. **Owner decisions:** a compact numbered list of explicit approval items, with your recommendation and the simplest viable alternative.
5. **Milestone/test gaps:** only concrete omissions, with the milestone and observable acceptance test.
6. **Suggested plan edits:** exact, minimal edits for the developer; do not rewrite the whole plan unless requested.

If no significant findings exist, say so; do not invent issues to fill sections. Cite repository file paths and line numbers when available. Avoid scores and arbitrary numeric confidence ratings.

## Operating constraints

1. **Review only:** do not modify production code, tests, approved specifications, implementation plan, Git index, or commits. Do not start implementation or delegate it. The only default repository modification permitted is creating/updating the review artifact described above.
2. Before finalizing, reconcile the report with explicit owner decisions supplied in the current request. Do not recommend reversing an approved decision without identifying new evidence and explicitly marking the recommendation as a proposed reconsideration.
3. Keep owner-facing discussion in Russian when the user prefers it; keep review documents and code suggestions in English unless repository instructions explicitly differ.
4. Do not launch expensive/full suites, external applications, or network operations merely to review a plan. Use narrow read-only checks where helpful and state what was not run.
5. A review must not create a new mandatory agent role, approval process, public feature, or architecture policy. It evaluates the existing plan against existing authority.

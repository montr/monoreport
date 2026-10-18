---
name: feature-orchestrator
description: Optionally coordinate an explicitly selected complex or high-risk feature workflow, its agreed roles and gates, and resumption from persisted state.
---

# Feature Orchestrator

1. Use this Skill only when the Product Owner explicitly selects an orchestrated workflow for a complex or high-risk Feature. It is not the default entry point for Maintenance or ordinary Feature work.
2. Agree with the Product Owner which Architect, Developer, independent Reviewer, approval, and re-review gates are needed. Every Feature has an approved Spec and Developer-authored Implementation Plan; use Architect analysis before Developer planning when design is complex. Do not add role cycles merely because the Skill is active. When independent review is agreed, do not substitute Developer self-review for it.
3. Assign sequential, never-reused IDs (`001`, `002`, etc.) after checking directories and available Git history. Store Feature artifacts in `.docs/features/<NNN>-<slug>/`; create only the documents needed by the approved workflow.
4. Preserve project rules and uncommitted work. Do not modify controlling Skills or workflow without explicit Product Owner approval. Follow the Git Operations rules in `AGENTS.md`.

## Workflow

1. `discussion`: prepare the Feature Spec and resolve product decisions with the Product Owner. Use a separate Architect before implementation planning when agreed for complex architecture. Enter `planning` only after the approval gate selected by the Product Owner is satisfied.
2. `planning`: prepare a concise Implementation Plan. Use an independent plan Reviewer only when the approved workflow calls for one. Resolve required findings and obtain the agreed Product Owner authorization before entering `implementation`.
3. `implementation`: implement approved requirements and suitable automated tests, then supply verification evidence. Enter `review` only when independent review is part of the approved workflow; otherwise proceed to `acceptance`, disclosing failed or unexecuted checks.
4. `review`: the independent Reviewer assesses conformance and evidence within the agreed scope. On the first `CHANGES REQUESTED`, record findings with severity, evidence, and rationale; immediately stop automatic correction and re-review and report to the Product Owner. Resume only on explicit Product Owner direction. Unresolved blocking findings may not silently become `PASS`.
5. `acceptance`: the Product Owner explicitly accepts or directs corrections. Use only the roles and re-verification required by the approved workflow. Enter `completed` only after explicit Product Owner acceptance.
6. Report newly discovered risk and propose proportionate checks without automatically stopping. Consult `docs/docx-template-compatibility.md` for DOCX boundaries, distinguishing target policy from accepted behavior; escalate new product support decisions before broadening scope. Obtain explicit Product Owner approval before materially changing approved requirements, architecture, scope, or workflow gates, or when additional authority is required.

## State and records

1. Do not create or maintain `feature-status.json` or any other separate phase/status file. Infer progress from existing Feature documents and verified work.
2. Record specification approval in `feature-spec.md`, plan approval and any agreed plan review in `implementation-plan.md`, and implementation review, verification, and acceptance in `review.md`. Never invent approvals or `PASS`; update phase only when the applicable gates are satisfied, atomically where practical.
3. On resume, inspect relevant Feature documents, available repository state, and implementation. Reconcile claims with evidence; resume only incomplete authorized work and respect stop/approval gates.
4. Feature 001 historical registration: retain the original spec in `.docs/features/001-initial-implementation/feature-spec.md`, version `0.0.1`, initially in `review`; preserve requirements, translate into English if needed, and update obsolete path references. Do not invent retrospective approvals or a plan, or perform an audit or corrections without explicit authorization.

## Resource discipline

1. Read only relevant files/sections/diffs and valid prior evidence. Keep handoffs and feature review updates delta-based: decision, changed behavior, findings, verification, limitations, next action. Record each approval once; retain audit history without repeating boilerplate.
2. After a tool/environment failure, allow at most one focused diagnostic retry; if unresolved, stop and consult Product Owner. Do not silently replace tooling or make material architecture changes.
3. Never invent token usage or remaining Codex quotas. Browser and automated test execution policy belongs to the Developer and Reviewer Skills.

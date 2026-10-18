---
name: feature-reviewer
description: Independently assess plans, implementation, tests, evidence, and documentation when a high-risk or explicitly requested review is part of the agreed workflow.
---

# Feature Reviewer

## Independent review

1. Use this Skill only when the Product Owner requests independent review or agrees that a high-risk boundary warrants it. Maintenance and ordinary Features do not require this role by default.
2. Review independently of the implementing agent. Judge findings against the approved Feature contract; consult `docs/docx-template-compatibility.md` for DOCX scope, distinguishing its future target policy from currently approved acceptance criteria. Read the approved Feature Spec when one exists, applicable plan, architecture rules, changes, and existing evidence; use focused context. For historical Feature 001, acknowledge absent historical approvals or plan rather than inventing them.
3. Plan review: assess coverage of requirements, feasible implementation, tests, risks, and architecture before Product Owner plan approval when plan review is within the agreed scope.
4. Implementation review: reconcile every acceptance criterion with concrete tests or justified checks; inspect correctness, errors, data integrity, boundaries, and regressions. Judge test quality by defect-detection value, not count or pass rate; look for weak assertions and shared implementation assumptions. For aggregation or ordering, inspect zero, one, and many cases within and across containers where relevant.
5. Reuse trustworthy Developer verification evidence when still applicable. Do **not** rerun successful tests by default or independently repeat the whole suite. Run focused checks only for a specific uncertainty or evidence gap. A full Reviewer suite run requires explicit Product Owner approval; lack of a redundant run is not a finding.
6. If evidence is missing, contradictory, stale, or invalidated, request targeted verification or identify the limitation. Judge evidence validity by affected behavior; classify unrelated flaky failures without calling them successes. Examine commands, outcomes, and raw artifacts where useful; never claim unexecuted tests passed or impose a test-writing order.
7. Browser tests are not part of automatic local review; run only on explicit Product Owner request or via GitHub Actions. Preserve existing browser artifacts and scripts. HTTP integration and generated OpenAPI contract coverage remain required when applicable; real-host checks depend on the affected behavior.

## Documentation and findings

1. For public API, CLI, configuration, DI, or recommended usage changes, read each affected user-facing document **in full**, including README and pre-existing examples. Flag outdated, duplicated, conflicting, misleading, or overly implementation-focused guidance.
2. Compile examples or execute relevant CLI commands when practical; clearly distinguish inspection from actual execution. Ensure documentation describes the intended user experience.
3. Classify reproducible violations of the approved contract, document corruption, or data loss as potential blockers; distinguish nonblocking simplification opportunities and unsupported/out-of-scope cases. Do not require support for speculative OOXML constructs or broaden requirements without a Product Owner decision. Record concise findings in `implementation-plan.md` (plan review) or `review.md` (implementation review), with severity, file/location, evidence, rationale, and requested correction. Summarize verification, not verbose tool logs.
4. Return `PASS` only when no blocking findings remain and required checks are evidenced. Otherwise return `CHANGES REQUESTED`; stop automatic correction and re-review and report to the Product Owner, directly or through an active Orchestrator. Never silently convert unresolved blocking findings into `PASS`.
5. The Orchestrator owns phase transitions only when an orchestrated workflow is active. Otherwise, report the result to the handling agent and Product Owner. Record explicit Product Owner decisions separately from independent review results.

## Efficiency

1. Reuse previously inspected unchanged materials; focus renewed review on changed files and affected requirements, without overlooking cross-cutting effects.
2. After an infrastructure failure, perform at most one useful focused diagnostic retry, then report the blocker and stop. Do not silently switch tools/frameworks or introduce architectural changes.
3. Do not invent token statistics or remaining usage quotas. Follow `AGENTS.md` documentation rules.

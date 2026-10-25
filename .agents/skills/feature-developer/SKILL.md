---
name: feature-developer
description: Plan and implement approved Features, following architectural decisions and verifying acceptance criteria.
---

# Feature Developer

## Planning

1. Prepare a concise Implementation Plan based on the approved Feature Spec and architectural decisions. Define implementation milestones, dependencies, and verification criteria. Obtain Product Owner approval before implementation.
2. Number implementation stages sequentially within each Feature (`Stage 1`, `Stage 2`, etc.). Keep stage numbers stable; never renumber or reuse them after implementation begins.

## Implementation and Verification

1. Follow the approved Implementation Plan. Track milestone status directly in the plan using `In progress`, `Review`, or `Accepted` (with date). Leave the status unspecified for work not started.
2. Report significant deviations from approved requirements or architecture to the Product Owner before proceeding with the affected work.
3. Before submitting work for review, verify implementation against the relevant acceptance criteria, perform risk-appropriate checks, and report results and remaining gaps.
4. Provide a concise summary of completed work, verification results, and unresolved issues.
5. Before an anticipated session interruption, leave a concise handoff in the current Implementation Plan when needed. On resumption, verify actual progress against the plan and working tree.
6. Address review findings only after Product Owner approval. Update finding status in `implementation-review.md` to `Fixed` after correction or `Rejected` following an explicit Product Owner decision, briefly recording the reason for rejection.

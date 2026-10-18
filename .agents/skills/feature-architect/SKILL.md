---
name: feature-architect
description: Optionally develop a Feature Spec or focused architecture analysis with the Product Owner when a separate architecture role is useful or agreed.
---

# Feature Architect

1. Use this Skill when the Product Owner requests an Architect or when a complex/high-risk Feature needs architecture decisions before the Developer prepares an Implementation Plan. The Architect analyzes design options and constraints; the Developer owns the implementation plan. Architect review of that plan is optional and separately agreed.
2. Read `AGENTS.md`, relevant product rules, existing Feature artifacts, and affected source before designing; avoid unrelated context.
3. Ask focused questions, preferably one decision at a time. Separate agreed requirements from assumptions; explain alternatives and tradeoffs for consequential decisions. Never silently decide product behavior.
4. Write or refine `.docs/features/<NNN>-<slug>/feature-spec.md`: goal, supported/preserved/rejected behavior where applicable, scope and exclusions, edge cases, testable acceptance criteria, and architecture constraints. Specify identity and boundary cases for collections or aggregation when important. Preserve meaning and examples during revisions; identify ambiguities explicitly.
5. For DOCX decisions consult `docs/docx-template-compatibility.md`, distinguishing target policy from current accepted behavior. If a new case would extend support or complexity, present simpler alternatives and ask the Product Owner before changing the Spec; record the decision in the Spec and update the compatibility policy if needed. Respect component boundaries. Put product-specific rules in project documentation rather than this Skill.
6. Present the completed specification for explicit Product Owner approval. Do not implement or presume approval. The Orchestrator owns phase transitions only when an orchestrated workflow was explicitly selected; otherwise follow the approved workflow recorded for the Feature.
7. Follow `AGENTS.md` documentation rules; keep the handoff concise and factual.

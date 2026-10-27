# Project Rules

## Development Workflow

1. Use the Feature workflow only when explicitly requested by the Product Owner. Otherwise, handle work as a regular task without creating Feature artifacts. Store Feature artifacts under `.docs/features/<NNN>-<slug>/`, using sequential, never-reused IDs. Read only relevant rules, code, documentation, and ADRs.
2. Obtain Product Owner approval before starting Feature implementation, making significant changes to approved requirements, scope, or architecture, or performing potentially disruptive operations, including browser automation and resource-intensive tests. Resolve routine implementation decisions independently.
3. Keep active Feature documents consistent with approved decisions and implementation. Do not modify completed Feature artifacts for subsequent maintenance tasks unless explicitly requested by the Product Owner.
4. Do not invoke other agents or delegate work to them without explicit Product Owner authorization.
5. Do not introduce functionality beyond the approved requirements without explicit Product Owner approval. Documenting a proposal does not constitute approval. Resolve routine implementation details independently.

## Git Operations

1. Do not stage, commit, push, merge, rebase, reset, or otherwise modify Git history, branches, tags, or the index without explicit Product Owner approval.
2. Git inspection is allowed. Preserve all existing staged and unstaged changes.

## Coding Guidelines

1. Keep data in models and business logic in services. Avoid placing business logic in model classes.
2. Implement services as instances with externally supplied dependencies. Prefer a single constructor, using a primary constructor where appropriate. Avoid static service methods and constructing dependencies internally.
3. Separate feature-specific logic from shared functionality. Introduce abstractions only when justified by actual requirements.
4. Prefer built-in platform, framework, and established library capabilities over custom implementations. Introduce custom infrastructure only when justified by concrete requirements or clear technical benefits.

## Test Quality

1. Write tests for meaningful behavior, critical requirements, and regression risks. Avoid redundant tests and coverage-driven testing.
2. Organize and name tests by the components or subsystems they verify, not by Feature IDs. Use scenario-based names for integration tests spanning multiple components.
3. Run relevant tests first and expand verification according to change impact and risk. Verify the requested behavior, not just successful compilation or passing tests. Avoid unnecessary repeated runs and report any significant unverified behavior.

## Documentation Rules

1. Write Skills, project rules, specifications, plans, decisions, and other project documentation in English; Product Owner discussions may use another language.
2. Use explicitly sequential numbered lists, not bullets; keep references unambiguous.
3. Do not add code-formatting examples unless the Product Owner requests them.


## UI Development

1. Prefer existing UI framework components, icons, layout mechanisms, and built-in behavior over custom implementations. Check the framework's capabilities before introducing custom UI code.
2. Prefer theme tokens and framework styling over custom CSS. Avoid component-specific overrides when existing theme settings are sufficient.
3. Keep UI behavior and visual appearance consistent with the existing application. Preserve accessibility and responsive layouts.
4. If legal, security, or technical obligations require additional UI, explain the need and propose a minimal solution before implementing it.

# Karpathy behavioral guidelines

## 1. Think Before Coding

**Don't assume. Don't hide confusion. Surface tradeoffs.**

Before implementing:
- State your assumptions explicitly. If uncertain, ask.
- If multiple interpretations exist, present them - don't pick silently.
- If a simpler approach exists, say so. Push back when warranted.
- If something is unclear, stop. Name what's confusing. Ask.

## 2. Simplicity First

**Minimum code that solves the problem. Nothing speculative.**

- No features beyond what was asked.
- No abstractions for single-use code.
- No "flexibility" or "configurability" that wasn't requested.
- No error handling for impossible scenarios.
- If you write 200 lines and it could be 50, rewrite it.

Ask yourself: "Would a senior engineer say this is overcomplicated?" If yes, simplify.

## 3. Surgical Changes

**Touch only what you must. Clean up only your own mess.**

When editing existing code:
- Don't "improve" adjacent code, comments, or formatting.
- Don't refactor things that aren't broken.
- Match existing style, even if you'd do it differently.
- If you notice unrelated dead code, mention it - don't delete it.

When your changes create orphans:
- Remove imports/variables/functions that YOUR changes made unused.
- Don't remove pre-existing dead code unless asked.

The test: Every changed line should trace directly to the user's request.

## 4. Goal-Driven Execution

**Define success criteria. Loop until verified.**

Transform tasks into verifiable goals:
- "Add validation" → "Write tests for invalid inputs, then make them pass"
- "Fix the bug" → "Write a test that reproduces it, then make it pass"
- "Refactor X" → "Ensure tests pass before and after"

For multi-step tasks, state a brief plan:
```
1. [Step] → verify: [check]
2. [Step] → verify: [check]
3. [Step] → verify: [check]
```

Strong success criteria let you loop independently. Weak criteria ("make it work") require constant clarification.

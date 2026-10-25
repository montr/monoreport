# Project Rules

## Development Workflow

1. Store Feature artifacts under `.docs/features/<NNN>-<slug>/`, using sequential, never-reused IDs. Read only the relevant project rules, code, Feature documents, and ADRs.
2. Obtain Product Owner approval before starting Feature implementation, making significant changes to approved requirements, scope, or architecture, or performing potentially disruptive operations, including browser automation and resource-intensive tests. Resolve routine implementation decisions independently.
3. Keep the Feature Spec and Implementation Plan consistent with approved decisions and actual implementation. Update them when requirements or implementation approach change.
4. Do not invoke other agents or delegate work to them without explicit Product Owner authorization.
5. Do not introduce user-visible features or permanent UI elements beyond the approved Feature Spec without explicit Product Owner approval. Highlight such proposals as separate decisions during planning; merely including them in the Implementation Plan does not constitute explicit approval. Routine layout, styling, and interaction details necessary to implement approved requirements do not need separate approval. If legal, security, or technical obligations require additional UI, explain the need and propose a minimal solution before implementing it.

## Git Operations

1. Do not stage, commit, push, merge, rebase, reset, or otherwise modify Git history, branches, tags, or the index without explicit Product Owner approval.
2. Git inspection is allowed. Preserve all existing staged and unstaged changes.

## Coding Guidelines

1. Keep data in models and business logic in services. Avoid placing business logic in model classes.
2. Implement services as instances with externally supplied dependencies. Prefer a single primary constructor. Avoid static service methods and constructing dependencies internally.
3. Separate feature-specific logic from shared functionality. Introduce abstractions only when justified by actual requirements.
4. Prefer built-in platform, framework, and established library capabilities over custom implementations. Introduce custom infrastructure only when justified by concrete requirements or clear technical benefits.
5. Prefer the simplest solution that fully satisfies the approved requirements. Before adding code, abstractions, dependencies, tests, or infrastructure, consider whether they are actually necessary. Avoid work that provides no meaningful value, without compromising correctness, security, or maintainability.

## Test Quality

1. Write tests for meaningful behavior, critical requirements, and regression risks. Avoid redundant tests and coverage-driven testing.
2. Organize and name tests by the components or subsystems they verify, not by Feature IDs. Use scenario-based names for integration tests spanning multiple components.
3. Run relevant tests first and expand verification according to change impact and risk. Avoid unnecessary repeated runs and report any significant unverified behavior.

## Documentation Rules

1. Write Skills, project rules, specifications, plans, decisions, and other project documentation in English; Product Owner discussions may use another language.
2. Use explicitly sequential numbered lists, not bullets; keep references unambiguous.
3. Do not add code-formatting examples unless the Product Owner requests them.

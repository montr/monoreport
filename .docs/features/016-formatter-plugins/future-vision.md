# Formatter and Extension Plugin Vision

**Status:** Future direction only, recorded on 2026-11-01. Not approved for implementation in Feature 016.

## 1. Intended future capability

1. Distribute extensions as `.nupkg` packages. A package may contain multiple formatters or other extension types.
2. Support dependencies and resources alongside extension assemblies. Design dependencies, resources, and loading together rather than assuming one copied DLL is sufficient.
3. Extract configured packages and dynamically load extensions at application startup. Acquisition, installation locations, startup composition, and failure handling need a separately approved design.
4. Allow loaded extensions to replace built-ins. Formatter selection follows ordinary .NET keyed DI; a loader must define deterministic registration order rather than another priority mechanism.

## 2. Foundation supplied by Feature 016

1. All seven existing formatters use public `IFormatter` contracts in Core. Six simple implementations stay in Core; `MonoReport.Formatter.Spellout` supplies RUB formatting. Core does not depend on external implementations.
2. Composition is explicit through `IServiceCollection`. `AddMonoReportSpelloutFormatter(string key = "spellout")` registers the supplied key as the template name; multiple keys are allowed. No custom builder, scanning, manifest, or plugin registry is needed.
3. Metadata has name/type/required descriptions and one optional formatter-level positional target. Syntax accepts positional/named double-quoted strings, replacing colon patterns. Analyze validates known literals without dataset evaluation; runtime formatting receives culture through context.
4. A small operand boundary allows a future distinct reference while preserving literal identity. Feature 016 implements no reference syntax/evaluation or expression hierarchy. Quoted dataset-looking paths are literals; unquoted values are invalid.
5. An independently referenced project is not a dynamically loaded plugin. Core-owned contracts resolve D5 without another contracts assembly; this does not promise future binary compatibility. See the [Feature Spec decisions](feature-spec.md#8-product-owner-decisions-and-remaining-architectural-concerns).
6. Do not generalize formatter contracts to every extension type without requirements. No new currencies, languages, or other Spellout capabilities are implied.

## 3. Decisions deferred to a future Feature

1. Package identity, contract/version compatibility, dependencies/conflicts, target frameworks, and multiple extension registrations per package.
2. Package sourcing/configuration, dependency acquisition, extraction/cache layout, resources, updates/cleanup, and startup failures.
3. Assembly loading, dependency isolation, contract identity, resource access, and shared versus isolated dependency versions.
4. Deterministic startup discovery/registration order and operator selection of built-in replacements.
5. Third-party code trust/execution boundaries, deployment permissions, and loading diagnostics.
6. Contracts for other extension types. No common hierarchy, manifest, manager, hot reload, unloading, or runtime installation is implied.
7. Dataset-reference syntax, binding/evaluation, and resulting analysis requirements. Use a distinct operand without reinterpreting quoted literals; this does not expand RUB functionality.

namespace MonoReport.Models;

internal sealed record TemplateInspection(IReadOnlyList<StoryPlan> Stories,
    IReadOnlyList<StructuralDiagnostic> Diagnostics,
    IReadOnlyList<TemplateValidationIssue> Warnings,
    IReadOnlyList<TemplateRequirement> Requirements,
    IReadOnlyList<TemplateOccurrence> Occurrences);

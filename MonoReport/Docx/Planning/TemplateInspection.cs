using MonoReport.Core.Diagnostics;
using MonoReport.Core.Templates;
using System.Collections.Generic;

namespace MonoReport.Docx.Planning;

internal sealed record TemplateInspection(
    IReadOnlyList<StoryPlan> Stories,
    IReadOnlyList<StructuralDiagnostic> Diagnostics,
    IReadOnlyList<TemplateValidationIssue> Warnings,
    IReadOnlyList<TemplateRequirement> Requirements,
    IReadOnlyList<TemplateOccurrence> Occurrences,
    StoryInspection Package,
    DocxIdentityPlan Identities);
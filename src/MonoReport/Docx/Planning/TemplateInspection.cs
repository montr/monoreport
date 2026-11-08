using System.Collections.Generic;
using MonoReport.Core.Diagnostics;
using MonoReport.Core.Templates;

namespace MonoReport.Docx.Planning;

internal sealed record TemplateInspection(
    IReadOnlyList<StoryPlan> Stories,
    IReadOnlyList<StructuralDiagnostic> Diagnostics,
    IReadOnlyList<TemplateValidationIssue> Warnings,
    IReadOnlyList<TemplateRequirement> Requirements,
    IReadOnlyList<TemplateOccurrence> Occurrences,
    StoryInspection Package,
    DocxIdentityPlan Identities);
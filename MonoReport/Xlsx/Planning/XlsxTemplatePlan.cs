using System.Collections.Immutable;
using MonoReport.Core.Templates;
using System.Collections.Generic;

namespace MonoReport.Xlsx.Planning;

internal sealed record XlsxTemplatePlan(
    XlsxSourceInventory Source,
    ImmutableArray<XlsxSheetTemplate> Sheets,
    ImmutableArray<PlannedDependency> Dependencies,
    IReadOnlyList<TemplateRequirement> Requirements,
    IReadOnlyList<TemplateOccurrence> Occurrences)
{
    internal ImmutableArray<XlsxDefinedName> DefinedNames { get; init; } = [];
    internal ImmutableArray<XlsxSharedFormulaGroup> SharedFormulas { get; init; } = [];
}
namespace MonoReport.Core.Templates;

internal sealed record PlannedDependency(
    string Kind,
    TemplateExpression Expression,
    string CanonicalPath,
    string RequirementType,
    string Part,
    string Location,
    long SourceOrder);
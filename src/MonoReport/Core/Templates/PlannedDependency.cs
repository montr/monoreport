namespace MonoReport.Core.Templates;

internal sealed record PlannedDependency(
    string Kind,
    TemplateExpression Expression,
    string CanonicalPath,
    string RequirementType,
    string Part,
    string Location,
    long SourceOrder)
{
    internal const string FieldKind = "field";
    internal const string ForKind = "for";
    internal const string IfKind = "if";
}

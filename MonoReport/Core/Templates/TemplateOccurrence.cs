namespace MonoReport.Core.Templates;

public sealed record TemplateOccurrence(
    string Kind,
    string Expression,
    string Path,
    string Part,
    string Location);
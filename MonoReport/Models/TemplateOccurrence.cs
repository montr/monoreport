namespace MonoReport.Models;

public sealed record TemplateOccurrence(
    string Kind,
    string Expression,
    string Path,
    string Part,
    string Location);

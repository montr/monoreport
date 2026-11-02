namespace MonoReport.Core.Templates;

public sealed record TemplateValidationIssue(
    string Code,
    string Message,
    string? Field = null,
    string? Part = null,
    string? Location = null);
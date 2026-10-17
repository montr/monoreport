namespace MonoReport.Models;

public sealed class TemplateAnalysisResult(IEnumerable<TemplateValidationIssue> errors,
    IEnumerable<TemplateValidationIssue>? warnings = null,
    IEnumerable<TemplateRequirement>? requirements = null,
    IEnumerable<TemplateOccurrence>? occurrences = null)
{
    public IReadOnlyList<TemplateValidationIssue> Errors { get; } =
        Array.AsReadOnly((errors ?? throw new ArgumentNullException(nameof(errors))).ToArray());
    public IReadOnlyList<TemplateValidationIssue> Warnings { get; } =
        Array.AsReadOnly(warnings?.ToArray() ?? []);
    public IReadOnlyList<TemplateRequirement> Requirements { get; } =
        Array.AsReadOnly(requirements?.ToArray() ?? []);
    public IReadOnlyList<TemplateOccurrence> Occurrences { get; } =
        Array.AsReadOnly(occurrences?.ToArray() ?? []);
    public bool Valid => Errors.Count == 0;
}

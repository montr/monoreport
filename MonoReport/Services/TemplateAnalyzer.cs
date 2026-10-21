using DocumentFormat.OpenXml.Packaging;
using MonoReport.Models;
using MonoReport.Services.Xlsx;

namespace MonoReport.Services;

public sealed class TemplateAnalyzer
{
    private readonly DocxStoryReader _storyReader;
    private readonly DocxTemplatePlanner _planner;
    private readonly PackageFormatDetector? _formats;
    private readonly XlsxTemplateHandler? _xlsx;

    public TemplateAnalyzer(DocxStoryReader storyReader, DocxTemplatePlanner planner)
    {
        _storyReader = storyReader ?? throw new ArgumentNullException(nameof(storyReader));
        _planner = planner ?? throw new ArgumentNullException(nameof(planner));
    }

    internal TemplateAnalyzer(DocxStoryReader storyReader, DocxTemplatePlanner planner,
        PackageFormatDetector formats, XlsxTemplateHandler xlsx) : this(storyReader, planner)
    {
        _formats = formats ?? throw new ArgumentNullException(nameof(formats));
        _xlsx = xlsx ?? throw new ArgumentNullException(nameof(xlsx));
    }

    public TemplateAnalysisResult Analyze(Stream template)
    {
        ArgumentNullException.ThrowIfNull(template);
        if (!template.CanRead) throw new ArgumentException("Template stream must be readable.", nameof(template));
        using var buffer = new MemoryStream();
        template.CopyTo(buffer);
        buffer.Position = 0;
        try
        {
            if (_formats?.Detect(buffer) == Models.DocumentFormat.Xlsx) return _xlsx!.Analyze(buffer);
        }
        catch (TemplateException error)
        {
            return new([new(error.Code, error.Errors[0].Message)]);
        }
        var document = _storyReader.OpenForValidation(buffer, false, out var packageDiagnostic);
        if (packageDiagnostic is not null) return new([Issue(packageDiagnostic)]);
        using (document)
        {
            var inspection = Inspect(document!);
            return new(inspection.Diagnostics.Select(Issue), inspection.Warnings,
                inspection.Requirements, inspection.Occurrences);
        }
    }

    internal TemplateInspection Inspect(WordprocessingDocument document)
    {
        var storyInspection = _storyReader.Inspect(document);
        var stories = storyInspection.Stories.Select(story => _planner.Plan(story.Part, story.Xml)).ToArray();
        var identities = _planner.InspectIdentities(storyInspection, stories);
        var builder = new TemplateRequirementBuilder();
        foreach (var dependency in stories.SelectMany(story => story.Dependencies)) builder.Add(dependency);
        var diagnostics = storyInspection.Diagnostics.Concat(stories.SelectMany(story => story.Diagnostics))
            .Concat(builder.Diagnostics).Concat(identities.Diagnostics)
            .OrderBy(item => DiagnosticPartOrder(item, storyInspection.PartOrder))
            .ThenBy(item => item.SourceOrder)
            .ToArray();
        return new(stories, diagnostics, [], builder.Requirements, builder.Occurrences, storyInspection, identities);
    }

    private static int DiagnosticPartOrder(StructuralDiagnostic diagnostic,
        IReadOnlyDictionary<string, int> partOrder)
    {
        var part = diagnostic.Location?.Split(',', 2)[0];
        return part is not null && partOrder.TryGetValue(part, out var order) ? order : int.MaxValue;
    }

    private static TemplateValidationIssue Issue(StructuralDiagnostic diagnostic)
    {
        var part = diagnostic.Location?.Split(',', 2)[0];
        return new(diagnostic.Code, diagnostic.ToException().ToDiagnostic(null).Message, diagnostic.Field, part, diagnostic.Location);
    }

}

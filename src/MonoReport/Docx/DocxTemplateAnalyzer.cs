using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using MonoReport.Core.Diagnostics;
using MonoReport.Core.Templates;
using MonoReport.Docx.Planning;

namespace MonoReport.Docx;

public sealed class DocxTemplateAnalyzer(DocxStoryReader storyReader, DocxTemplatePlanner planner)
{
    private readonly DocxTemplatePlanner _planner = planner ?? throw new ArgumentNullException(nameof(planner));
    private readonly DocxStoryReader _storyReader = storyReader ?? throw new ArgumentNullException(nameof(storyReader));

    internal TemplateAnalysisResult Analyze(Stream template)
    {
        var document = _storyReader.OpenForValidation(template, false, out var packageDiagnostic);
        if (packageDiagnostic is not null) return new TemplateAnalysisResult([Issue(packageDiagnostic)]);
        using (document)
        {
            var inspection = Inspect(document!);
            return new TemplateAnalysisResult(inspection.Diagnostics.Select(Issue), inspection.Warnings,
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
        return new TemplateInspection(stories, diagnostics, [], builder.Requirements, builder.Occurrences,
            storyInspection, identities);
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
        return new TemplateValidationIssue(diagnostic.Code, diagnostic.ToException().ToDiagnostic(null).Message,
            diagnostic.Field, part, diagnostic.Location);
    }
}
using System.Text.Json;
using System.Xml.Linq;
using MonoReport.Core.Templates;
using System.IO;
using System;
using System.Linq;

namespace MonoReport.Docx;

public sealed class DocxGenerator(
    DocxStoryReader storyReader,
    DocxTemplateAnalyzer analyzer,
    DocxRenderer renderer)
{
    private readonly DocxTemplateAnalyzer _analyzer = analyzer ?? throw new ArgumentNullException(nameof(analyzer));
    private readonly DocxRenderer _renderer = renderer ?? throw new ArgumentNullException(nameof(renderer));
    private readonly DocxStoryReader _storyReader = storyReader ?? throw new ArgumentNullException(nameof(storyReader));

    internal GenerationResult Generate(MemoryStream package, JsonDocument data, GenerationOptions options)
    {
        var context = new DocxGenerationContext(data.RootElement, options);
        using (var document = _storyReader.Open(package, true))
        {
            var inspection = _analyzer.Inspect(document);
            if (inspection.Diagnostics.Count != 0)
            {
                var structural = inspection.Diagnostics.Take(DocxGenerationContext.ErrorLimit).ToArray();
                var first = structural[0].ToException();
                throw new TemplateException(first,
                    structural.Select(error => TemplateException.ToDiagnostic(error.ToException(), null)),
                    structural.Length == DocxGenerationContext.ErrorLimit);
            }

            context.InitializePackageIds(document, inspection);
            foreach (var story in inspection.Stories)
            {
                if (!_renderer.Render(story, context)) break;
                using var target = story.Part.GetStream(FileMode.Create, FileAccess.Write);
                story.Xml.Save(target, SaveOptions.DisableFormatting);
            }

            if (context.Errors.Count != 0)
                throw new TemplateException(context.FirstError!, context.Errors, context.LimitReached);
            context.SaveDeletedCommentDefinitions();
        }

        return new GenerationResult(package.ToArray(), context.Warnings.AsReadOnly());
    }
}
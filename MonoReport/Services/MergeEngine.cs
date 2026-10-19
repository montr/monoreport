using System.Text.Json;
using System.Xml.Linq;
using MonoReport.Models;

namespace MonoReport.Services
{
    public sealed class MergeEngine(DocumentRenderer renderer, DocxStoryReader storyReader,
        TemplateAnalyzer analyzer)
    {
        private readonly DocumentRenderer _renderer = renderer ?? throw new ArgumentNullException(nameof(renderer));
        private readonly DocxStoryReader _storyReader = storyReader ?? throw new ArgumentNullException(nameof(storyReader));
        private readonly TemplateAnalyzer _analyzer = analyzer ?? throw new ArgumentNullException(nameof(analyzer));

        /// <summary>Reads the template from its current position. The caller retains ownership of both inputs.</summary>
        public MergeResult Merge(Stream template, JsonDocument data, MergeOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(template);
            ArgumentNullException.ThrowIfNull(data);
            if (template.CanRead == false)
            {
                throw new ArgumentException("Template stream must be readable.", nameof(template));
            }

            options ??= new();
            var context = new MergeContext(data.RootElement, options);
            using (var output = new MemoryStream())
            {
                template.CopyTo(output);
                output.Position = 0;
                using (var document = _storyReader.Open(output, true))
                {
                    var analysis = _analyzer.Inspect(document);
                    if (analysis.Diagnostics.Count != 0)
                    {
                        var structural = analysis.Diagnostics.Take(MergeContext.ErrorLimit).ToArray();
                        var first = structural[0].ToException();
                        throw new TemplateException(first,
                            structural.Select(error => TemplateException.ToDiagnostic(error.ToException(), null)),
                            structural.Length == MergeContext.ErrorLimit);
                    }
                    context.InitializePackageIds(document, analysis);
                    foreach (var story in analysis.Stories)
                    {
                        if (!_renderer.Render(story, context)) break;
                        using (var target = story.Part.GetStream(FileMode.Create, FileAccess.Write))
                        {
                            story.Xml.Save(target, SaveOptions.DisableFormatting);
                        }
                    }

                    if (context.Errors.Count != 0)
                        throw new TemplateException(context.FirstError!, context.Errors, context.LimitReached);
                    context.SaveDeletedCommentDefinitions();
                }

                return new(output.ToArray(), context.Warnings.AsReadOnly());
            }
        }
    }
}

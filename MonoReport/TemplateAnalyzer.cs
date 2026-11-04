using System;
using System.IO;
using MonoReport.Core.Templates;
using MonoReport.Docx;
using MonoReport.Xlsx;

namespace MonoReport;

public sealed class TemplateAnalyzer(
    PackageFormatDetector formats,
    DocxTemplateAnalyzer docx,
    XlsxTemplateAnalyzer xlsx)
{
    private readonly DocxTemplateAnalyzer _docx = docx ?? throw new ArgumentNullException(nameof(docx));
    private readonly PackageFormatDetector _formats = formats ?? throw new ArgumentNullException(nameof(formats));
    private readonly XlsxTemplateAnalyzer _xlsx = xlsx ?? throw new ArgumentNullException(nameof(xlsx));

    public TemplateAnalysisResult Analyze(Stream template)
    {
        ArgumentNullException.ThrowIfNull(template);
        if (!template.CanRead) throw new ArgumentException("Template stream must be readable.", nameof(template));

        using var buffer = new MemoryStream();
        template.CopyTo(buffer);
        buffer.Position = 0;
        try
        {
            return _formats.Detect(buffer) switch
            {
                DocumentFormat.Docx => _docx.Analyze(buffer),
                DocumentFormat.Xlsx => _xlsx.Analyze(buffer),
                _ => throw new ArgumentOutOfRangeException()
            };
        }
        catch (TemplateException error)
        {
            return new TemplateAnalysisResult([new TemplateValidationIssue(error.Code, error.Errors[0].Message)]);
        }
    }
}
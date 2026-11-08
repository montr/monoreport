using System;
using System.IO;
using System.Text.Json;
using MonoReport.Docx;
using MonoReport.Xlsx;

namespace MonoReport;

public sealed class DocumentGenerator(PackageFormatDetector formats, DocxGenerator docx, XlsxGenerator xlsx)
{
    private readonly DocxGenerator _docx = docx ?? throw new ArgumentNullException(nameof(docx));
    private readonly PackageFormatDetector _formats = formats ?? throw new ArgumentNullException(nameof(formats));
    private readonly XlsxGenerator _xlsx = xlsx ?? throw new ArgumentNullException(nameof(xlsx));


    /// <summary>Reads the template from its current position. The caller retains ownership of both inputs.</summary>
    public GenerationResult Generate(Stream template, JsonDocument data, GenerationOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(template);
        ArgumentNullException.ThrowIfNull(data);
        if (!template.CanRead) throw new ArgumentException("Template stream must be readable.", nameof(template));

        using var package = new MemoryStream();
        template.CopyTo(package);
        package.Position = 0;
        options ??= new GenerationOptions();

        return _formats.Detect(package) switch
        {
            DocumentFormat.Docx => _docx.Generate(package, data, options),
            DocumentFormat.Xlsx => _xlsx.Generate(package, data, options),
            _ => throw new ArgumentOutOfRangeException()
        };
    }
}
using System;
using System.IO;
using System.Numerics;
using System.Text.Json;
using MonoReport.Core.Diagnostics;
using MonoReport.Core.Templates;
using MonoReport.Docx;
using MonoReport.Xlsx;

namespace MonoReport;

public sealed class SampleJsonGenerator(PackageFormatDetector formats, DocxSampleJsonGenerator docx,
    XlsxSampleJsonGenerator xlsx)
{
    private readonly DocxSampleJsonGenerator _docx = docx ?? throw new ArgumentNullException(nameof(docx));
    private readonly PackageFormatDetector _formats = formats ?? throw new ArgumentNullException(nameof(formats));
    private readonly XlsxSampleJsonGenerator _xlsx = xlsx ?? throw new ArgumentNullException(nameof(xlsx));

    public SampleJsonResult Generate(Stream template, SampleJsonOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(template);
        if (!template.CanRead) throw new ArgumentException("Template stream must be readable.", nameof(template));

        options ??= new SampleJsonOptions();
        if (options.ItemsPerCollection <= BigInteger.Zero)
            throw new ArgumentOutOfRangeException(nameof(options), "ItemsPerCollection must be a positive integer.");
        if (!Enum.IsDefined(options.Values))
            throw new ArgumentOutOfRangeException(nameof(options), "Values must be Filled or Empty.");

        using var buffer = new MemoryStream();
        template.CopyTo(buffer);
        buffer.Position = 0;
        return _formats.Detect(buffer) switch
        {
            DocumentFormat.Docx => _docx.Generate(buffer, options),
            DocumentFormat.Xlsx => _xlsx.Generate(buffer, options),
            _ => throw new ArgumentOutOfRangeException()
        };
    }
}

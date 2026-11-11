using System;
using System.Linq;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

var results = args.Select(path =>
{
    using var document = WordprocessingDocument.Open(path, false);
    var errors = new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(document)
        .Select(error => new
        {
            error.Id,
            error.Description,
            Part = error.Part?.Uri.ToString(),
            Path = error.Path?.XPath
        }).ToArray();
    return new { File = path, Errors = errors };
}).ToArray();
Console.WriteLine(JsonSerializer.Serialize(results, new JsonSerializerOptions { WriteIndented = true }));
return results.Any(result => result.Errors.Length != 0) ? 1 : 0;

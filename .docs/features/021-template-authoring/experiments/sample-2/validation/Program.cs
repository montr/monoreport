using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

var results = new List<object>();
foreach (var path in args)
{
    using var document = WordprocessingDocument.Open(path, false);
    var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
    var errors = validator.Validate(document).Select(error => new
    {
        error.Id,
        Type = error.ErrorType.ToString(),
        error.Description,
        Part = error.Part?.Uri.ToString(),
        Path = error.Path?.XPath
    }).ToArray();
    results.Add(new { Path = path, Errors = errors, ErrorCount = errors.Length });
}
Console.WriteLine(JsonSerializer.Serialize(results, new JsonSerializerOptions { WriteIndented = true }));

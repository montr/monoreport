using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

var reports = new List<object>();
var failed = false;
foreach (var path in args)
{
    using var document = WordprocessingDocument.Open(path, false);
    var validator = new OpenXmlValidator(FileFormatVersions.Office2007);
    var errors = validator.Validate(document).Select(error => new
    {
        error.Id,
        error.Description,
        Type = error.ErrorType.ToString(),
        Part = error.Part?.Uri.ToString(),
        Path = error.Path?.XPath
    }).ToArray();
    reports.Add(new { File = path, ErrorCount = errors.Length, Errors = errors });
    failed |= errors.Length != 0;
}
Console.WriteLine(JsonSerializer.Serialize(reports, new JsonSerializerOptions { WriteIndented = true }));
return failed ? 1 : 0;

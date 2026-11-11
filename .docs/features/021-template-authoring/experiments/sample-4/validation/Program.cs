using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
var results = new List<object>();
foreach (var path in args)
{
    using var document = SpreadsheetDocument.Open(path, false);
    var errors = validator.Validate(document).Select(error => new
    {
        error.Id,
        error.Description,
        Part = error.Part?.Uri.ToString(),
        Path = error.Path?.XPath
    }).ToArray();
    results.Add(new { File = Path.GetFileName(path), ErrorCount = errors.Length, Errors = errors });
}
Console.WriteLine(JsonSerializer.Serialize(results, new JsonSerializerOptions { WriteIndented = true }));

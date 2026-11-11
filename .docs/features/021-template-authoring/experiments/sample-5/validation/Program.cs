using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

var results = new Dictionary<string, object>();
foreach (var path in args)
{
    using var document = SpreadsheetDocument.Open(path, false);
    results[path] = new OpenXmlValidator().Validate(document).Select(error => new
    {
        error.Id,
        error.Description,
        Part = error.Part?.Uri.ToString(),
        Path = error.Path?.XPath
    }).ToArray();
}
Console.WriteLine(JsonSerializer.Serialize(results, new JsonSerializerOptions { WriteIndented = true }));

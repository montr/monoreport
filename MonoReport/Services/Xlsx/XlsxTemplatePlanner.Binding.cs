using System.Collections.Immutable;
using System.Globalization;
using System.Text.Json;
using System.Xml;
using MonoReport.Models;
using MonoReport.Models.Xlsx;

namespace MonoReport.Services.Xlsx;

internal sealed partial class XlsxTemplatePlanner
{
    internal XlsxWorkbookPlan AttachValues(XlsxWorkbookPlan geometry, ImmutableDictionary<SourceCellId, XlsxPreparedValue> values) =>
        geometry with { Sheets = geometry.Sheets.Select(sheet => sheet with
        {
            Rows = sheet.Rows.Select(row => row with
            {
                Cells = row.Cells.Select(cell => cell with { Value = values.GetValueOrDefault(cell.Id.Source) }).ToImmutableArray()
            }).ToImmutableArray()
        }).ToImmutableArray() };

    internal ImmutableDictionary<SourceCellId, XlsxPreparedValue> Bind(XlsxTemplatePlan plan, JsonElement data, XlsxOperationContext context)
    {
        if (data.ValueKind != JsonValueKind.Object)
            throw new StructuralDiagnostic(DiagnosticCodes.GenerationInvalidData, "XLSX requires a JSON object.",
                Explanation: "Supply one JSON object as the XLSX report data.").ToException();
        var scope = new DataScope(data, [data], ["$"]);
        var result = ImmutableDictionary.CreateBuilder<SourceCellId, XlsxPreparedValue>();
        foreach (var cell in plan.Sheets.SelectMany(sheet => sheet.Cells))
        {
            if (context.LimitReached) break;
            if (cell.Tokens.Length > 0) result.Add(cell.Source.Id, PrepareValue(cell, scope, context));
        }
        context.ThrowIfErrors();
        return result.ToImmutable();
    }

    private static string SafeText(string value, TemplateExpression expression, string location)
    {
        try { return XmlConvert.VerifyXmlChars(value); }
        catch (XmlException)
        {
            throw new StructuralDiagnostic(DiagnosticCodes.GenerationInvalidData, "The field contains characters that XML cannot store.",
                expression.Source, Location: location,
                Explanation: "Remove XML control characters from the field value before generating the workbook.").ToException();
        }
    }

    private static string DecodeScalarString(JsonElement value, TemplateExpression expression, string location)
    {
        try { return value.GetString()!; }
        catch (InvalidOperationException)
        {
            // This catch covers only JSON string decoding, never formatting, package work or renderer logic.
            throw new StructuralDiagnostic(DiagnosticCodes.GenerationInvalidData, "Invalid Unicode in the JSON string.",
                expression.Source, Location: location, Formatter: expression.Formatter?.Name,
                Explanation: "Supply valid Unicode text with paired surrogate characters in the JSON string.").ToException();
        }
    }

    private static string SafeNumber(string raw, TemplateExpression expression, string location)
    {
        // Compare normalized source decimal digits, before binary conversion; Excel stores 15 significant digits.
        var original = DecimalIdentity(raw);
        if (original is { Digits: "0" }) return "0";
        if (original is null || original.Value.Digits.Length > 15 ||
            !double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) || !double.IsFinite(number) ||
            number == 0 || Math.Abs(number) < 2.2250738585072014E-308 ||
            original != DecimalIdentity(number.ToString("G15", CultureInfo.InvariantCulture)))
            throw new TemplateException("Unsafe Excel numeric conversion.", expression.Source, expression.Source, location,
                code: DiagnosticCodes.XlsxUnsafeNumber);
        return raw;
    }

    private static (bool Negative, string Digits, long Exponent)? DecimalIdentity(string text)
    {
        var negative = text.StartsWith('-');
        if (negative || text.StartsWith('+')) text = text[1..];
        var exponentAt = text.IndexOfAny(['e', 'E']);
        var mantissa = exponentAt < 0 ? text : text[..exponentAt];
        var dot = mantissa.IndexOf('.');
        var fractionLength = dot < 0 ? 0 : mantissa.Length - dot - 1;
        var digits = mantissa.Replace(".", string.Empty, StringComparison.Ordinal).TrimStart('0');
        if (digits.Length == 0) return (false, "0", 0);
        var trimmed = digits.TrimEnd('0');
        if (exponentAt >= 0 && !long.TryParse(text[(exponentAt + 1)..], NumberStyles.AllowLeadingSign,
                CultureInfo.InvariantCulture, out _)) return null;
        var exponent = exponentAt < 0 ? 0 : long.Parse(text[(exponentAt + 1)..], CultureInfo.InvariantCulture);
        try { return (negative, trimmed, checked(exponent - fractionLength + digits.Length - trimmed.Length)); }
        catch (OverflowException) { return null; }
    }
}

using System.Collections.Immutable;
using System.Text;
using System.Text.Json;
using MonoReport.Models;
using MonoReport.Models.Xlsx;

namespace MonoReport.Services.Xlsx;

internal sealed partial class XlsxTemplatePlanner(TemplateLanguageParser language, DataResolver resolver, ValueFormatter formatter,
    XlsxFormulaReferences references)
{
    private readonly TemplateLanguageParser _language = language ?? throw new ArgumentNullException(nameof(language));
    private readonly DataResolver _resolver = resolver ?? throw new ArgumentNullException(nameof(resolver));
    private readonly ValueFormatter _formatter = formatter ?? throw new ArgumentNullException(nameof(formatter));
    private readonly XlsxFormulaReferences _references = references ?? throw new ArgumentNullException(nameof(references));

    internal XlsxTemplatePlan Plan(XlsxSourceInventory source, XlsxOperationContext context)
    {
        var plan = Prepare(source, context);
        return plan with { Sheets = plan.Sheets.Select(Structure).ToImmutableArray() };
    }

    // Prepares lexical scopes/dependencies only. Recursive grammar/geometry validation remains M2/M3.
    internal XlsxTemplatePlan Prepare(XlsxSourceInventory source, XlsxOperationContext context)
    {
        var diagnostics = new List<StructuralDiagnostic>();
        // Requirement conflicts can refer back to earlier cells, including across worksheets.
        // Apply the published cap only after merging/sorting source diagnostics; Bind keeps its evaluation cap.
        var partOrder = source.Sheets.Select((sheet, index) => (sheet.Part, index)).ToDictionary(item => item.Part, item => item.index, StringComparer.Ordinal);
        void Report(StructuralDiagnostic diagnostic)
        {
            diagnostics.Add(diagnostic);
        }
        var sheets = ImmutableArray.CreateBuilder<XlsxSheetTemplate>();
        var dependencies = ImmutableArray.CreateBuilder<PlannedDependency>();
        foreach (var sheet in source.Sheets)
        {
            var controls = ReadControls(sheet, context);
            var scopes = CellScopes(sheet, controls);
            var sheetDependencies = new List<PlannedDependency>();
            foreach (var marker in controls.Where(marker => marker.Collection is not null))
                sheetDependencies.Add(new("for", marker.Collection!, marker.CollectionPath!, "array", sheet.Part,
                    marker.Location, Order(marker.Cell.Id.Address)));
            var cells = ImmutableArray.CreateBuilder<XlsxCellPlan>();
            foreach (var cell in sheet.Cells.Where(cell => sheet.PrintArea.Contains(cell.Id.Address)))
            {
                var location = Location(sheet, cell.Id.Address);
                XlsxParsedFormula? formula = null;
                if (cell.HasFormula)
                {
                    try
                    {
                        var elements = cell.Xml.Elements().Where(element => element.Name.LocalName == "f").ToArray();
                        if (elements.Length != 1 || elements[0].Attributes().Any(attribute => attribute.Name.LocalName != "t" || attribute.Value != "normal"))
                            throw new StructuralDiagnostic(DiagnosticCodes.XlsxUnsafeFormula, "Use ordinary scalar formulas; shared groups require M3 and array/data-table/spill structures are unsupported.", Location: location,
                                Explanation: "Re-save the affected formula as an ordinary scalar formula using explicit retained A1 operands.").ToException();
                        formula = _references.Read(elements[0].Value, location);
                        if (formula.References.Any(reference => reference.Prefix.Length > 0))
                            throw new StructuralDiagnostic(DiagnosticCodes.XlsxUnsafeFormula, "Sheet-qualified formula binding requires M3.",
                                Location: location, Explanation: "Use explicit local retained A1 operands in M2; sheet-qualified formula binding requires M3.").ToException();
                    }
                    catch (TemplateException error)
                    {
                        Report(new(error.Code, error.Message, Location: location, Explanation: error.Message, SourceOrder: Order(cell.Id.Address)));
                    }
                }
                var scope = scopes.TryGetValue(cell.Id.Address.Row, out var found) ? found : ImmutableArray<string>.Empty;
                var parsed = cell.Text is null ? new TemplateParseResult([], []) : _language.Parse(cell.Text, location);
                foreach (var error in parsed.Diagnostics)
                {
                    Report(error with { SourceOrder = Order(cell.Id.Address) + error.SourceOrder });
                }
                if (cell.RichText && parsed.Tokens.Count > 0)
                    Report(new StructuralDiagnostic(DiagnosticCodes.XlsxUnsupportedFeature,
                        "Rich-text field substitution is not implemented in M1.", Location: location,
                        Explanation: "Use a plain cell string for fields until rich-run substitution is implemented; unchanged rich text can be retained.",
                        SourceOrder: Order(cell.Id.Address)));
                foreach (var token in parsed.Tokens.Where(token => token.Expression is not null))
                {
                    if (token.Expression!.Kind != TemplateExpressionKind.Field)
                        Report(new StructuralDiagnostic(DiagnosticCodes.XlsxUnsupportedFeature,
                            "Use XLSX worksheet control markers for repetition; scalar cells accept fields only.",
                            token.Expression.Source, Location: location,
                            Explanation: "Use XLSX worksheet control markers for repetition; scalar cells accept fields only.",
                            SourceOrder: Order(cell.Id.Address) + token.Start));
                    else if (token.Expression.Path!.ParentDepth > scope.Length)
                        Report(new StructuralDiagnostic(DiagnosticCodes.TemplateInvalidContext, "The field has no corresponding parent scope.",
                            token.Expression.Source, Location: location, SourceOrder: Order(cell.Id.Address) + token.Start));
                    else
                    {
                        var path = token.Expression.Path!;
                        var basis = path.Origin == TemplatePathOrigin.Root || path.ParentDepth == scope.Length
                            ? "$" : scope[path.ParentDepth] + "[]";
                        var canonical = basis + (path.CurrentElement || path.Segments.Count == 0 ? string.Empty : "." + string.Join('.', path.Segments));
                        sheetDependencies.Add(new("field", token.Expression, canonical, "value", sheet.Part,
                            location, Order(cell.Id.Address) + token.Start));
                    }
                }
                cells.Add(new(cell, parsed.Tokens.ToImmutableArray(), location) { Formula = formula });
            }
            sheets.Add(new(sheet, cells.ToImmutable(), controls));
            dependencies.AddRange(sheetDependencies.OrderBy(item => item.SourceOrder));
        }
        var requirements = new TemplateRequirementBuilder();
        foreach (var dependency in dependencies) requirements.Add(dependency);
        foreach (var diagnostic in requirements.Diagnostics)
        {
            Report(diagnostic);
        }
        foreach (var diagnostic in diagnostics.OrderBy(item => partOrder.GetValueOrDefault(item.Location?.Split(',', 2)[0] ?? string.Empty, int.MaxValue))
            .ThenBy(item => item.SourceOrder).Take(GenerationDiagnosticPolicy.ErrorLimit)) context.Report(diagnostic.ToException());
        context.ThrowIfErrors();
        return new(source, sheets.ToImmutable(), dependencies.ToImmutable(), requirements.Requirements, requirements.Occurrences);
    }

    internal XlsxPreparedValue PrepareValue(XlsxCellPlan cell, DataScope scope, XlsxOperationContext context)
    {
        if (cell.Source.Text is null) return new(cell.Source.NativeType, null);
        var text = new StringBuilder();
        var offset = 0;
        foreach (var token in cell.Tokens)
        {
            if (context.LimitReached) break;
            text.Append(cell.Source.Text, offset, token.Start - offset);
            offset = token.Start + token.Length;
            if (token.Kind != TemplateTokenKind.Expression)
            {
                text.Append(token.Kind == TemplateTokenKind.LiteralOpenBrace ? '{' : '}');
                continue;
            }
            var expression = token.Expression!;
            var path = expression.Path!;
            context.ResolvedFields = checked(context.ResolvedFields + 1);
            var value = _resolver.Resolve(path, scope);
            var dataPath = _resolver.DataPath(path, scope);
            try
            {
                if (value.State == ResolutionState.Missing)
                {
                    if (context.Options.MissingValues == MissingValueBehavior.Error)
                        throw new TemplateException("The scalar field is missing.", expression.Source, location: cell.Location);
                    if (context.Options.MissingValues == MissingValueBehavior.Warning)
                        context.Warnings.Add(new("A scalar field is missing; an empty value was used.", expression.Source,
                            expression.Source, cell.Location, cell.Location.Split(',', 2)[0], dataPath) { Code = DiagnosticCodes.XlsxMissingValue });
                }
                if (value.State != ResolutionState.Present)
                {
                    if (cell.NativeField) return new(null, null);
                    continue;
                }
                // Validate string decoding before either native binding or the shared textual formatter.
                var scalarString = value.Value.ValueKind == JsonValueKind.String
                    ? DecodeScalarString(value.Value, expression, cell.Location) : null;
                if (cell.NativeField)
                    return value.Value.ValueKind switch
                    {
                        JsonValueKind.String => new("inlineStr", SafeText(scalarString!, expression, cell.Location)),
                        JsonValueKind.Number => new("n", SafeNumber(value.Value.GetRawText(), expression, cell.Location)),
                        JsonValueKind.True => new("b", "1"),
                        JsonValueKind.False => new("b", "0"),
                        _ => throw new TemplateException("Objects and arrays cannot be used as scalar fields.", expression.Source,
                            location: cell.Location)
                    };
                text.Append(SafeText(_formatter.Format(value.Value, expression, context.Options, cell.Location), expression, cell.Location));
            }
            catch (TemplateException error) { context.Report(error, dataPath); }
        }
        text.Append(cell.Source.Text, offset, cell.Source.Text.Length - offset);
        return new("inlineStr", text.ToString());
    }
}

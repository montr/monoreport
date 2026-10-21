namespace MonoReport.Models
{
    public sealed class TemplateException : Exception
    {
        public TemplateException(string message, string? field = null, string? fieldCode = null,
            string? location = null, string? formatter = null, Exception? inner = null,
            string code = DiagnosticCodes.GenerationInvalidData) : base(message, inner)
        {
            Code = code;
            Field = field;
            FieldCode = fieldCode;
            Location = location;
            Formatter = formatter;
            Errors = Array.AsReadOnly(new[] { ToDiagnostic(this, null) });
        }

        internal TemplateException(TemplateException first, IEnumerable<GenerationDiagnostic> errors,
            bool limitReached) : base(first.Message, first.InnerException)
        {
            Code = first.Code;
            Field = first.Field;
            FieldCode = first.FieldCode;
            Location = first.Location;
            Formatter = first.Formatter;
            SourceContext = first.SourceContext;
            Explanation = first.Explanation;
            Errors = Array.AsReadOnly(errors.ToArray());
            LimitReached = limitReached;
        }

        internal TemplateException(StructuralDiagnostic diagnostic)
            : this(diagnostic.Detail, diagnostic.Field, diagnostic.FieldCode, diagnostic.Location,
                diagnostic.Formatter, code: diagnostic.Code)
        {
            SourceContext = diagnostic.SourceContext;
            Explanation = diagnostic.Explanation;
            Errors = Array.AsReadOnly(new[] { ToDiagnostic(this, null) });
        }

        public string Code { get; }
        internal string? SourceContext { get; }
        internal string? Explanation { get; }
        public string? Field { get; }

        public string? FieldCode { get; }

        public string? Location { get; }

        public string? Formatter { get; }

        public IReadOnlyList<GenerationDiagnostic> Errors { get; }

        public bool LimitReached { get; }

        internal GenerationDiagnostic ToDiagnostic(string? dataPath, string? location = null) =>
            ToDiagnostic(this, dataPath, location);

        internal static GenerationDiagnostic ToDiagnostic(TemplateException error, string? dataPath,
            string? location = null)
        {
            var resolvedLocation = location ?? error.Location;
            var message = SafeMessage(error.Code);
            if (error.Code.StartsWith("template.", StringComparison.Ordinal) && error.Field is not null)
                message += $" Token: «{Excerpt(error.Field, 100)}».";
            if (error.Explanation is not null) message += " " + Excerpt(error.Explanation, 1600);
            if (error.SourceContext is not null)
            {
                message += $" Nearby source text: «{Excerpt(error.SourceContext, 162)}».";
            }
            return new(error.Code, message, error.Field, Part(resolvedLocation),
                resolvedLocation, error.Formatter, dataPath);
        }

        private static string Excerpt(string value, int limit) =>
            (value.Length > limit ? value[..limit] + "…" : value).Replace('\r', ' ').Replace('\n', ' ').Replace('\t', ' ');

        private static string? Part(string? location) => location?.Split(',', 2)[0];

        private static string SafeMessage(string code) => code switch
        {
            DiagnosticCodes.TemplateInvalidPackage => "The content is not a readable supported Word or ordinary XLSX package. Re-save the document; for Excel, use an ordinary .xlsx workbook.",
            DiagnosticCodes.TemplateInvalidXml => "A document part contains invalid XML.",
            DiagnosticCodes.TemplateUnsupportedStructure => "The DOCX document structure is unsupported.",
            DiagnosticCodes.TemplateInvalidExpression => "A template expression has invalid syntax.",
            DiagnosticCodes.TemplateUnmatchedBrace => "Template braces are not balanced or escaped correctly.",
            DiagnosticCodes.TemplateInvalidContext => "A template path is invalid in its current loop context.",
            DiagnosticCodes.TemplateInvalidStructure => "A template block has an unsupported or ambiguous structure.",
            DiagnosticCodes.TemplateRequirementConflict => "The same dependency path is used with incompatible requirement types.",
            DiagnosticCodes.TemplateTrackedRevision => "Tracked revisions must be accepted or rejected before processing.",
            DiagnosticCodes.GenerationInvalidFormat => "A value cannot be formatted as requested.",
            DiagnosticCodes.XlsxInvalidPrintArea => "Set one finite rectangular PrintArea on every worksheet.",
            DiagnosticCodes.XlsxUnsupportedFeature => "The XLSX feature cannot be preserved safely; remove or re-author the indicated structure.",
            DiagnosticCodes.XlsxDimensionLimit => "The generated report exceeds Excel worksheet dimensions; reduce the report size.",
            DiagnosticCodes.XlsxUnsafeNumber => "The value cannot be stored faithfully as an Excel number. Supply it as a JSON string or render it with an explicit text formatter.",
            DiagnosticCodes.XlsxInvalidPackage => "The XLSX package is inconsistent. Repair or re-save the workbook with valid parts and relationships.",
            DiagnosticCodes.XlsxUnsafeFormula => "The XLSX formula reference cannot be transformed safely; re-author it using explicit retained A1 operands.",
            _ => "The data is incompatible with the template."
        };
    }
}

namespace MonoReport.Models
{
    public sealed class TemplateException : Exception
    {
        public TemplateException(string message, string? field = null, string? fieldCode = null,
            string? location = null, string? formatter = null, Exception? inner = null,
            string code = "generation.invalid_data") : base(message, inner)
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
            "template.invalid_package" => "The content is not a readable DOCX package.",
            "template.invalid_xml" => "A document part contains invalid XML.",
            "template.unsupported_structure" => "The DOCX document structure is unsupported.",
            "template.invalid_expression" => "A template expression has invalid syntax.",
            "template.unmatched_brace" => "Template braces are not balanced or escaped correctly.",
            "template.invalid_context" => "A template path is invalid in its current loop context.",
            "template.invalid_structure" => "A template block has an unsupported or ambiguous structure.",
            "template.requirement_conflict" => "The same dependency path is used with incompatible requirement types.",
            "template.tracked_revision" => "Tracked revisions must be accepted or rejected before processing.",
            "generation.invalid_format" => "A value cannot be formatted as requested.",
            _ => "The data is incompatible with the template."
        };
    }
}

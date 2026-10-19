namespace MonoReport.Models;

/// <summary>Canonical values of Core diagnostic-code identifiers published by MonoReport.</summary>
public static class DiagnosticCodes
{
    public const string TemplateInvalidPackage = "template.invalid_package";
    public const string TemplateInvalidXml = "template.invalid_xml";
    public const string TemplateUnsupportedStructure = "template.unsupported_structure";
    public const string TemplateInvalidExpression = "template.invalid_expression";
    public const string TemplateUnmatchedBrace = "template.unmatched_brace";
    public const string TemplateInvalidContext = "template.invalid_context";
    public const string TemplateInvalidStructure = "template.invalid_structure";
    public const string TemplateRequirementConflict = "template.requirement_conflict";
    public const string TemplateTrackedRevision = "template.tracked_revision";
    public const string GenerationInvalidData = "generation.invalid_data";
    public const string GenerationInvalidFormat = "generation.invalid_format";
}

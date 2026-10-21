namespace MonoReport.Models;

/// <summary>Canonical values of Core diagnostic-code identifiers published by MonoReport.</summary>
public static class DiagnosticCodes
{
    public const string XlsxMissingLevel = "generation.xlsx_missing_level";
    public const string XlsxUnsafeFormula = "template.xlsx_unsafe_formula";
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
    public const string XlsxInvalidPrintArea = "template.xlsx_invalid_print_area";
    public const string XlsxUnsupportedFeature = "template.xlsx_unsupported_feature";
    public const string XlsxDimensionLimit = "generation.xlsx_dimension_limit";
    public const string XlsxUnsafeNumber = "generation.xlsx_unsafe_number";
    public const string XlsxMissingValue = "generation.xlsx_missing_value";
    public const string XlsxInvalidPackage = "template.xlsx_invalid_package";
}

using System.Text.RegularExpressions;

namespace MonoReport.Core.Templates;

/// <summary>Validates identifiers used in template paths, formatter keys, and parameter names.</summary>
public static partial class TemplateIdentifier
{
    public static bool IsValid(string? identifier)
    {
        return identifier is not null && IdentifierRegex().IsMatch(identifier);
    }

    [GeneratedRegex(@"^[\p{L}_][\p{L}\p{Nd}_]*$", RegexOptions.CultureInvariant)]
    private static partial Regex IdentifierRegex();
}
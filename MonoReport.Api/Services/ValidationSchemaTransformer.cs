using Microsoft.AspNetCore.OpenApi;
using Microsoft.OpenApi;
using MonoReport.Models;

namespace MonoReport.Api.Services;

public sealed class ValidationSchemaTransformer : IOpenApiSchemaTransformer
{
    public Task TransformAsync(OpenApiSchema schema, OpenApiSchemaTransformerContext context, CancellationToken cancellationToken)
    {
        if (context.JsonTypeInfo.Type == typeof(TemplateAnalysisResult))
            schema.Required = new HashSet<string> { "valid", "errors", "warnings", "requirements", "occurrences" };
        if (context.JsonTypeInfo.Type == typeof(TemplateValidationIssue))
        {
            schema.Required = new HashSet<string> { "code", "message", "field", "part", "location" };
            foreach (var name in new[] { "field", "part", "location" })
                if (schema.Properties?[name] is OpenApiSchema property) property.Type = JsonSchemaType.String | JsonSchemaType.Null;
            schema.Description = "A location-bearing analyzer diagnostic. Syntax, context, structural, tracked-revision, and static format errors invalidate the template; warnings do not.";
        }
        if (context.JsonTypeInfo.Type == typeof(TemplateRequirement))
        {
            schema.Required = new HashSet<string> { "name", "type" };
            schema.Description = "A dependency node. Type is object, array, value, or boolean. Omitted fields means no child dependencies; array fields describe expressions used against elements, not an element schema.";
        }
        if (context.JsonTypeInfo.Type == typeof(TemplateOccurrence))
            schema.Required = new HashSet<string> { "kind", "expression", "path", "part", "location" };
        return Task.CompletedTask;
    }
}

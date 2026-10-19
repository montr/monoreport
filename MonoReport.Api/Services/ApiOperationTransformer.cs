using Microsoft.AspNetCore.OpenApi;
using Microsoft.OpenApi;
using MonoReport.Models;

namespace MonoReport.Api.Services;

public sealed class ApiOperationTransformer : IOpenApiOperationTransformer
{
    public Task TransformAsync(OpenApiOperation operation, OpenApiOperationTransformerContext context, CancellationToken cancellationToken)
    {
        var generation = context.Description.RelativePath == "api/documents/generate";
        var requestCodes = generation
            ? new[]
            {
                ApiDiagnosticCodes.RequestMissingInput, ApiDiagnosticCodes.RequestEmptyInput,
                ApiDiagnosticCodes.RequestDuplicateInput, ApiDiagnosticCodes.RequestInvalidInput,
                ApiDiagnosticCodes.RequestInvalidMultipart, ApiDiagnosticCodes.DataInvalidJson,
                ApiDiagnosticCodes.CultureInvalid
            }
            : new[]
            {
                ApiDiagnosticCodes.RequestMissingInput, ApiDiagnosticCodes.RequestEmptyInput,
                ApiDiagnosticCodes.RequestDuplicateInput, ApiDiagnosticCodes.RequestInvalidInput,
                ApiDiagnosticCodes.RequestInvalidMultipart, ApiDiagnosticCodes.RequestUnexpectedInput
            };
        var generationCodes = new[]
        {
            DiagnosticCodes.TemplateInvalidPackage, DiagnosticCodes.TemplateInvalidXml,
            DiagnosticCodes.TemplateUnsupportedStructure, DiagnosticCodes.TemplateInvalidExpression,
            DiagnosticCodes.TemplateUnmatchedBrace, DiagnosticCodes.TemplateInvalidContext,
            DiagnosticCodes.TemplateInvalidStructure, DiagnosticCodes.TemplateRequirementConflict,
            DiagnosticCodes.TemplateTrackedRevision, DiagnosticCodes.GenerationInvalidData,
            DiagnosticCodes.GenerationInvalidFormat
        };
        var properties = new Dictionary<string, IOpenApiSchema>
        {
            ["template"] = new OpenApiSchema { Type = JsonSchemaType.String, Format = "binary", Description = "Required nonempty DOCX template." }
        };
        if (generation)
        {
            properties["data"] = new OpenApiSchema { Type = JsonSchemaType.String, Description = "Required text field containing one raw JSON value." };
            properties["culture"] = new OpenApiSchema { Type = JsonSchemaType.String, Description = "Optional culture name; invariant when omitted or blank. Accept-Language is ignored." };
        }
        operation.RequestBody = new OpenApiRequestBody
        {
            Required = true,
            Content = new Dictionary<string, OpenApiMediaType>
            {
                ["multipart/form-data"] = new() { Schema = new OpenApiSchema
                {
                    Type = JsonSchemaType.Object, Properties = properties,
                    Required = generation ? new HashSet<string> { "template", "data" } : new HashSet<string> { "template" },
                    AdditionalPropertiesAllowed = generation
                } }
            }
        };
        operation.Summary = generation ? "Generate a DOCX document" : "Analyze a DOCX template";
        operation.Description = generation
            ? "Returns raw DOCX bytes after one parse-and-render operation. Template structure and actual JSON values are validated before output is produced."
            : "Returns diagnostics, an ordered recursive dependency tree, and every source occurrence without JSON input or generation. Invalid nonempty DOCX content returns 200 with valid=false.";
        if (generation)
        {
            var success = (OpenApiResponse)operation.Responses!["200"];
            success.Content = new Dictionary<string, OpenApiMediaType>
            {
                ["application/vnd.openxmlformats-officedocument.wordprocessingml.document"] = new() { Schema = new OpenApiSchema { Type = JsonSchemaType.String, Format = "binary" } }
            };
            success.Headers = new Dictionary<string, IOpenApiHeader>
            {
                ["Content-Disposition"] = new OpenApiHeader { Description = "Attachment with safe deterministic document.docx filename.", Schema = new OpenApiSchema { Type = JsonSchemaType.String } },
                ["X-MonoReport-Warning-Count"] = new OpenApiHeader { Description = "Present only when generation warnings exist.", Schema = new OpenApiSchema { Type = JsonSchemaType.Integer } }
            };
        }
        foreach (var (status, response) in operation.Responses!)
        {
            if (status == "200") continue;
            if (response is not OpenApiResponse concrete) continue;
            response.Description = status switch
            {
                "400" => string.Join(", ", requestCodes),
                "413" => ApiDiagnosticCodes.RequestTooLarge + " (pre-application host/proxy rejection may use a host-owned response)",
                "415" => ApiDiagnosticCodes.RequestUnsupportedMediaType,
                "422" => string.Join(", ", generationCodes),
                _ => ApiDiagnosticCodes.ServerInternalError
            };
            concrete.Content = new Dictionary<string, OpenApiMediaType>
            {
                ["application/problem+json"] = new() { Schema = new OpenApiSchema
                {
                    Type = JsonSchemaType.Object, Required = new HashSet<string> { "type", "title", "status", "detail", "code", "traceId" },
                    Properties = new Dictionary<string, IOpenApiSchema>
                    {
                        ["type"] = new OpenApiSchema { Type = JsonSchemaType.String }, ["title"] = new OpenApiSchema { Type = JsonSchemaType.String },
                        ["status"] = new OpenApiSchema { Type = JsonSchemaType.Integer }, ["detail"] = new OpenApiSchema { Type = JsonSchemaType.String },
                        ["code"] = new OpenApiSchema { Type = JsonSchemaType.String, Description = response.Description }, ["traceId"] = new OpenApiSchema { Type = JsonSchemaType.String },
                        ["field"] = new OpenApiSchema { Type = JsonSchemaType.String | JsonSchemaType.Null }
                    }
                } }
            };
            if (generation && status == "422")
            {
                var schema = (OpenApiSchema)concrete.Content["application/problem+json"].Schema!;
                schema.Required!.Add("errors");
                schema.Required.Add("limitReached");
                schema.Properties!["limitReached"] = new OpenApiSchema { Type = JsonSchemaType.Boolean };
                schema.Properties["errors"] = new OpenApiSchema
                {
                    Type = JsonSchemaType.Array,
                    Items = new OpenApiSchema
                    {
                        Type = JsonSchemaType.Object,
                        Required = new HashSet<string> { "code", "message" },
                        Properties = new Dictionary<string, IOpenApiSchema>
                        {
                            ["code"] = new OpenApiSchema { Type = JsonSchemaType.String },
                            ["message"] = new OpenApiSchema { Type = JsonSchemaType.String },
                            ["field"] = new OpenApiSchema { Type = JsonSchemaType.String | JsonSchemaType.Null },
                            ["part"] = new OpenApiSchema { Type = JsonSchemaType.String | JsonSchemaType.Null },
                            ["location"] = new OpenApiSchema { Type = JsonSchemaType.String | JsonSchemaType.Null },
                            ["formatter"] = new OpenApiSchema { Type = JsonSchemaType.String | JsonSchemaType.Null },
                            ["dataPath"] = new OpenApiSchema { Type = JsonSchemaType.String | JsonSchemaType.Null }
                        }
                    }
                };
            }
        }
        return Task.CompletedTask;
    }
}

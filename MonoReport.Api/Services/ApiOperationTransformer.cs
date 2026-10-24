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
            DiagnosticCodes.GenerationInvalidFormat, DiagnosticCodes.XlsxInvalidPrintArea,
            DiagnosticCodes.XlsxUnsupportedFeature, DiagnosticCodes.XlsxDimensionLimit,
            DiagnosticCodes.XlsxUnsafeNumber, DiagnosticCodes.XlsxInvalidPackage, DiagnosticCodes.XlsxUnsafeFormula
        };
        var properties = new Dictionary<string, IOpenApiSchema>
        {
            ["template"] = new OpenApiSchema { Type = JsonSchemaType.String, Format = "binary", Description = "Required nonempty Word or ordinary XLSX template. Actual OOXML content selects the handler; filename and upload MIME are not authoritative." }
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
        operation.Summary = generation ? "Generate a Word or XLSX document" : "Analyze a Word or XLSX template";
        operation.Description = generation
            ? "Returns raw DOCX or XLSX bytes by default after one operation. Explicit Accept: multipart/mixed opts into document and diagnostics.json attachments. Wildcards alone never opt in; q=0 excludes multipart. Among explicit supported document types and multipart, higher quality wins and binary wins ties. Other Accept values fall back to binary. Vary: Accept is always set. All generation and diagnostic preparation completes before success publication. XLSX supports static fields, nested vertical groups, multirow blocks and safe scalar formulas over finite A1 operands including supported sheet/name/shared references, with physical PrintArea crop; it requires JSON object data. Recursive horizontal expansion supports branching sibling blocks and multirow headers, with at most one horizontal region per worksheet. Template structure and values are validated before output is produced."
            : "Returns diagnostics, an ordered recursive dependency tree, and source occurrences without JSON input or generation. Invalid nonempty package content returns 200 with valid=false. XLSX source analysis supports nested templates and supported scalar formula references; generation additionally checks matrix cardinalities, ancestor availability and final reference images.";
        if (generation)
        {
            operation.Parameters ??= [];
            operation.Parameters.Add(new OpenApiParameter
            {
                Name = "Accept", In = ParameterLocation.Header, Required = false,
                Description = "Explicit multipart/mixed opts in to two attachments. Quality zero excludes multipart; higher explicit quality wins, binary wins ties. Wildcards alone and unsupported values retain binary fallback.",
                Schema = new OpenApiSchema { Type = JsonSchemaType.String }
            });
            var success = (OpenApiResponse)operation.Responses!["200"];
            var warningProperties = new Dictionary<string, IOpenApiSchema>();
            foreach (var name in new[] { "message", "field", "fieldCode", "part", "location", "dataPath", "code" })
                warningProperties[name] = new OpenApiSchema
                {
                    Type = name is "message" or "field" ? JsonSchemaType.String : JsonSchemaType.String | JsonSchemaType.Null,
                    Description = name == "code" ? "Optional stable XLSX diagnostic code; omitted when unset. DOCX warnings need no code." : null
                };
            success.Description = "Generated document; warnings do not make generation fail. Multipart has two attachments from the same merge, without base64 encoding.";
            success.Content = new Dictionary<string, OpenApiMediaType>
            {
                ["application/vnd.openxmlformats-officedocument.wordprocessingml.document"] = new() { Schema = new OpenApiSchema { Type = JsonSchemaType.String, Format = "binary" } },
                ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"] = new() { Schema = new OpenApiSchema { Type = JsonSchemaType.String, Format = "binary" } },
                ["multipart/mixed"] = new()
                {
                    Encoding = new Dictionary<string, OpenApiEncoding>
                    {
                        ["document"] = new() { ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
                        ["diagnostics"] = new() { ContentType = "application/json" }
                    },
                    Schema = new OpenApiSchema
                    {
                        Type = JsonSchemaType.Object, Required = new HashSet<string> { "document", "diagnostics" },
                        Properties = new Dictionary<string, IOpenApiSchema>
                        {
                            ["document"] = new OpenApiSchema { Type = JsonSchemaType.String, Format = "binary", Description = "Attachment document.docx or document.xlsx with the corresponding binary MIME type." },
                            ["diagnostics"] = new OpenApiSchema
                            {
                                Type = JsonSchemaType.Object, Description = "application/json attachment named diagnostics.json.",
                                Required = new HashSet<string> { "format", "fileName", "warnings" },
                                Properties = new Dictionary<string, IOpenApiSchema>
                                {
                                    ["format"] = new OpenApiSchema { Type = JsonSchemaType.String, Description = "docx or xlsx" },
                                    ["fileName"] = new OpenApiSchema { Type = JsonSchemaType.String, Description = "document.docx or document.xlsx" },
                                    ["warnings"] = new OpenApiSchema
                                    {
                                        Type = JsonSchemaType.Array, Description = "Original ordered MergeResult.Warnings, including an explicit empty array.",
                                        Items = new OpenApiSchema { Type = JsonSchemaType.Object, Required = new HashSet<string> { "message", "field", "fieldCode", "part", "location", "dataPath" }, Properties = warningProperties }
                                    }
                                }
                            }
                        }
                    }
                }
            };
            success.Headers = new Dictionary<string, IOpenApiHeader>
            {
                ["Content-Disposition"] = new OpenApiHeader { Description = "Binary response only: attachment with deterministic document.docx or document.xlsx filename selected from actual package content.", Schema = new OpenApiSchema { Type = JsonSchemaType.String } },
                ["Vary"] = new OpenApiHeader { Description = "Accept", Schema = new OpenApiSchema { Type = JsonSchemaType.String } },
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

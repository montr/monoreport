using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Http.Json;
using Microsoft.Extensions.Options;
using MonoReport.Api.Models;
using MonoReport.Core;
using MonoReport.Core.Formatting;

namespace MonoReport.Api.Services;

public sealed class FormatterEndpoints(
    KeyedServiceKeyProvider<string, IFormatter> formatterKeys,
    KeyedServiceProvider<IFormatter> formatters,
    FormatterParameterBinder parameterBinder,
    ValueFormatter valueFormatter,
    ApiProblemFactory problems,
    IOptions<JsonOptions> jsonOptions)
{
    private readonly JsonSerializerOptions _requestJsonOptions = new(jsonOptions.Value.SerializerOptions)
    {
        AllowDuplicateProperties = false
    };

    public IReadOnlyList<FormatterDescription> List()
    {
        return formatterKeys.Keys.Order(StringComparer.Ordinal).Select(name =>
        {
            var metadata = formatters.GetRequiredKeyedService(name).Metadata;
            return new FormatterDescription(name, metadata.Parameters.Select(parameter =>
                new FormatterParameterDescription(parameter.Name, parameter.Type.Name, parameter.Required)).ToArray(),
                metadata.PositionalParameter);
        }).ToArray();
    }

    public async Task<IResult> Execute(HttpContext context, string name)
    {
        var formatter = formatters.GetKeyedService(name);
        if (formatter is null)
            throw new RequestInputException(404, ApiDiagnosticCodes.FormatterNotFound, nameof(name));
        if (!context.Request.HasJsonContentType())
            throw new RequestInputException(415, ApiDiagnosticCodes.RequestUnsupportedMediaType);

        FormatterExecutionRequest? request;
        try
        {
            request = await context.Request.ReadFromJsonAsync<FormatterExecutionRequest>(_requestJsonOptions,
                context.RequestAborted);
        }
        catch (JsonException)
        {
            throw new RequestInputException(400, ApiDiagnosticCodes.RequestInvalidInput);
        }
        catch (InvalidOperationException error) when (error.InnerException is ArgumentException or NotSupportedException)
        {
            throw new RequestInputException(415, ApiDiagnosticCodes.RequestUnsupportedMediaType);
        }

        if (request is null)
            throw new RequestInputException(400, ApiDiagnosticCodes.RequestInvalidInput);
        if (request.Value.ValueKind == JsonValueKind.Undefined)
            throw new RequestInputException(400, ApiDiagnosticCodes.RequestMissingInput, "value");
        if (request.Parameters?.Values.Any(value => value is null) == true)
            throw new RequestInputException(400, ApiDiagnosticCodes.RequestInvalidInput, "parameters");

        var culture = CultureInfo.InvariantCulture;
        if (!string.IsNullOrWhiteSpace(request.Culture))
            try
            {
                culture = CultureInfo.GetCultureInfo(request.Culture);
            }
            catch (CultureNotFoundException)
            {
                throw new RequestInputException(400, ApiDiagnosticCodes.CultureInvalid, "culture");
            }

        var supplied = request.Parameters?.Select(parameter =>
            new KeyValuePair<string, FormatterOperand>(parameter.Key,
                new FormatterOperand(FormatterOperandKind.Literal, parameter.Value))) ?? [];
        var validation = parameterBinder.Bind(name, formatter, supplied, out var parameters);
        if (validation.Errors.Count != 0)
            return problems.Create(context, 400, ApiDiagnosticCodes.RequestInvalidInput, "parameters", validation.Errors);

        try
        {
            return Results.Ok(new FormatterExecutionResult(
                valueFormatter.Format(request.Value, name, parameters, new FormatterContext(culture))));
        }
        catch (Exception error) when (error is FormatException or ArgumentException or OverflowException
                                          or InvalidOperationException)
        {
            throw new RequestInputException(422, ApiDiagnosticCodes.FormatterInvalidValue, "value");
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc.Testing;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Api.Models;
using MonoReport.Api.Services;
using MonoReport.Core.Formatting;
using Xunit;

namespace MonoReport.Api.Tests;

public sealed class FormatterApiTests
{
    [Fact]
    public async Task List_returns_registered_formatter_metadata_with_serializable_parameter_types()
    {
        using var factory = new WebApplicationFactory<Program>();
        using var client = factory.CreateClient();
        var formatters = await client.GetFromJsonAsync<FormatterDescription[]>("/api/formatters");

        Assert.NotNull(formatters);
        Assert.Equal(new[] {"caps", "date", "firstcap", "lower", "number", "spellout", "upper"},
            formatters.Select(formatter => formatter.Name));
        var number = Assert.Single(formatters, formatter => formatter.Name == "number");
        Assert.Equal("pattern", number.PositionalParameter);
        Assert.Equal(new FormatterParameterDescription("pattern", "String", true), Assert.Single(number.Parameters));
        Assert.Empty(Assert.Single(formatters, formatter => formatter.Name == "upper").Parameters);
        var spellout = Assert.Single(formatters, formatter => formatter.Name == "spellout");
        Assert.Contains(new FormatterParameterDescription("capitalize", "Boolean", false), spellout.Parameters);
    }

    [Theory]
    [InlineData("upper", """{"value":"Hello world"}""", "HELLO WORLD")]
    [InlineData("lower", """{"value":"Hello WORLD"}""", "hello world")]
    [InlineData("caps", """{"value":"hello world"}""", "Hello World")]
    [InlineData("firstcap", """{"value":"hello WORLD"}""", "Hello WORLD")]
    [InlineData("number", """{"value":1234.5,"parameters":{"pattern":"N2"},"culture":"de-DE"}""", "1.234,50")]
    [InlineData("number", """{"value":"1234.5","parameters":{"pattern":"F2"}}""", "1234.50")]
    [InlineData("date", """{"value":"2026-11-04","parameters":{"pattern":"yyyy/MM/dd"}}""", "2026/11/04")]
    [InlineData("spellout", """{"value":1,"parameters":{"mode":"money","currency":"RUB"}}""", "One ruble 00 copecks")]
    [InlineData("upper", """{"value":true}""", "TRUE")]
    [InlineData("upper", """{"value":null}""", "")]
    [InlineData("number", """{"value":"","parameters":{"pattern":"F2"}}""", "")]
    public async Task Execute_uses_registered_formatters_and_existing_value_semantics(string name, string body,
        string expected)
    {
        using var factory = new WebApplicationFactory<Program>();
        using var client = factory.CreateClient();
        client.DefaultRequestHeaders.AcceptLanguage.ParseAdd("tr-TR");
        using var content = Json(body);
        using var response = await client.PostAsync($"/api/formatters/{name}/execute", content);

        response.EnsureSuccessStatusCode();
        var result = await response.Content.ReadFromJsonAsync<FormatterExecutionResult>();
        Assert.Equal(expected, result!.Result);
    }

    [Theory]
    [InlineData("missing", """{"value":1}""", 404, ApiDiagnosticCodes.FormatterNotFound)]
    [InlineData("UPPER", """{"value":"text"}""", 404, ApiDiagnosticCodes.FormatterNotFound)]
    [InlineData("upper", "{}", 400, ApiDiagnosticCodes.RequestMissingInput)]
    [InlineData("upper", "null", 400, ApiDiagnosticCodes.RequestInvalidInput)]
    [InlineData("upper", "{", 400, ApiDiagnosticCodes.RequestInvalidInput)]
    [InlineData("upper", """{"value":"text","culture":"invalid culture!"}""", 400, ApiDiagnosticCodes.CultureInvalid)]
    [InlineData("upper", """{"value":"text","parameters":{"extra":"value"}}""", 400, ApiDiagnosticCodes.RequestInvalidInput)]
    [InlineData("number", """{"value":1}""", 400, ApiDiagnosticCodes.RequestInvalidInput)]
    [InlineData("number", """{"value":null}""", 400, ApiDiagnosticCodes.RequestInvalidInput)]
    [InlineData("number", """{"value":1,"parameters":{"Pattern":"F2"}}""", 400, ApiDiagnosticCodes.RequestInvalidInput)]
    [InlineData("number", """{"value":1,"parameters":{"pattern":""}}""", 400, ApiDiagnosticCodes.RequestInvalidInput)]
    [InlineData("number", """{"value":1,"parameters":{"pattern":"Q2"}}""", 400, ApiDiagnosticCodes.RequestInvalidInput)]
    [InlineData("number", """{"value":1,"parameters":{"pattern":null}}""", 400, ApiDiagnosticCodes.RequestInvalidInput)]
    [InlineData("number", """{"value":1,"parameters":{"pattern":2}}""", 400, ApiDiagnosticCodes.RequestInvalidInput)]
    [InlineData("spellout", """{"value":1,"parameters":{"mode":"money"}}""", 400, ApiDiagnosticCodes.RequestInvalidInput)]
    [InlineData("number", """{"value":true,"parameters":{"pattern":"F2"}}""", 422, ApiDiagnosticCodes.FormatterInvalidValue)]
    [InlineData("number", """{"value":1e100,"parameters":{"pattern":"F2"}}""", 422, ApiDiagnosticCodes.FormatterInvalidValue)]
    [InlineData("date", """{"value":"yesterday","parameters":{"pattern":"O"}}""", 422, ApiDiagnosticCodes.FormatterInvalidValue)]
    [InlineData("upper", """{"value":{}}""", 422, ApiDiagnosticCodes.FormatterInvalidValue)]
    [InlineData("upper", """{"value":[]}""", 422, ApiDiagnosticCodes.FormatterInvalidValue)]
    public async Task Execute_returns_problem_details_for_invalid_requests_and_values(string name, string body,
        int status, string code)
    {
        using var factory = new WebApplicationFactory<Program>();
        using var client = factory.CreateClient();
        using var content = Json(body);
        using var response = await client.PostAsync($"/api/formatters/{name}/execute", content);

        Assert.Equal((HttpStatusCode) status, response.StatusCode);
        Assert.Equal("application/problem+json", response.Content.Headers.ContentType!.MediaType);
        using var problem = JsonDocument.Parse(await response.Content.ReadAsStringAsync());
        Assert.Equal(code, problem.RootElement.GetProperty("code").GetString());
        Assert.False(string.IsNullOrEmpty(problem.RootElement.GetProperty("traceId").GetString()));
    }

    [Fact]
    public async Task Execute_rejects_non_json_content()
    {
        using var factory = new WebApplicationFactory<Program>();
        using var client = factory.CreateClient();
        using var content = new StringContent("value");
        using var response = await client.PostAsync("/api/formatters/upper/execute", content);

        Assert.Equal(HttpStatusCode.UnsupportedMediaType, response.StatusCode);
    }

    [Theory]
    [InlineData("application/json; charset=unsupported-charset")]
    [InlineData("application/json; charset=\"unsupported-charset\"")]
    [InlineData("application/problem+json; charset=unsupported-charset")]
    [InlineData("application/json; charset=utf-7")]
    [InlineData("application/json; charset=windows-1251")]
    public async Task Execute_returns_unsupported_media_type_for_unsupported_json_charsets(string contentType)
    {
        using var factory = new WebApplicationFactory<Program>();
        using var client = factory.CreateClient();
        using var content = Json("""{"value":"Ada"}""");
        content.Headers.ContentType = MediaTypeHeaderValue.Parse(contentType);
        using var response = await client.PostAsync("/api/formatters/upper/execute", content);

        Assert.Equal(HttpStatusCode.UnsupportedMediaType, response.StatusCode);
        Assert.Equal("application/problem+json", response.Content.Headers.ContentType!.MediaType);
        using var problem = JsonDocument.Parse(await response.Content.ReadAsStringAsync());
        Assert.Equal(ApiDiagnosticCodes.RequestUnsupportedMediaType, problem.RootElement.GetProperty("code").GetString());
    }

    [Fact]
    public async Task Execute_accepts_supported_non_utf8_json_charsets()
    {
        using var factory = new WebApplicationFactory<Program>();
        using var client = factory.CreateClient();
        using var content = new StringContent("""{"value":"Ada"}""", Encoding.Unicode, "application/json");
        using var response = await client.PostAsync("/api/formatters/upper/execute", content);

        response.EnsureSuccessStatusCode();
        var result = await response.Content.ReadFromJsonAsync<FormatterExecutionResult>();
        Assert.Equal("ADA", result!.Result);
    }

    [Theory]
    [InlineData("""{"value":"first","value":"second","parameters":{"suffix":"!"}}""")]
    [InlineData("""{"value":"first","Value":"second","parameters":{"suffix":"!"}}""")]
    [InlineData("""{"value":"first","v\u0061lue":"second","parameters":{"suffix":"!"}}""")]
    [InlineData("""{"value":"Ada","culture":"en-US","culture":"tr-TR","parameters":{"suffix":"!"}}""")]
    [InlineData("""{"value":"Ada","parameters":{"suffix":"?"},"parameters":{"suffix":"!"}}""")]
    [InlineData("""{"value":"Ada","parameters":{"suffix":"?"},"Parameters":{"suffix":"!"}}""")]
    [InlineData("""{"value":"Ada","parameters":{"suffix":"?","suffix":"!"}}""")]
    [InlineData("""{"value":"Ada","parameters":{"suffix":"?","suff\u0069x":"!"}}""")]
    public async Task Execute_rejects_duplicate_fields_and_parameters_before_formatting(string body)
    {
        var formatter = new ProbeFormatter();
        using var factory = new WebApplicationFactory<Program>().WithWebHostBuilder(builder =>
            builder.ConfigureServices(services => services.AddKeyedSingleton<IFormatter>("probe", formatter)));
        using var client = factory.CreateClient();
        using var content = Json(body);
        using var response = await client.PostAsync("/api/formatters/probe/execute", content);

        Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
        Assert.Equal("application/problem+json", response.Content.Headers.ContentType!.MediaType);
        using var problem = JsonDocument.Parse(await response.Content.ReadAsStringAsync());
        Assert.Equal(ApiDiagnosticCodes.RequestInvalidInput, problem.RootElement.GetProperty("code").GetString());
        Assert.Equal(0, formatter.FormatCalls);
    }

    [Fact]
    public async Task Custom_keys_and_replacements_control_discovery_validation_and_execution()
    {
        var formatter = new ProbeFormatter();
        using var factory = new WebApplicationFactory<Program>().WithWebHostBuilder(builder =>
            builder.ConfigureServices(services =>
            {
                services.AddKeyedSingleton<IFormatter>("probe", formatter);
                services.AddKeyedSingleton<IFormatter>("number", formatter);
            }));
        using var client = factory.CreateClient();
        var formatters = await client.GetFromJsonAsync<FormatterDescription[]>("/api/formatters");
        foreach (var name in new[] {"number", "probe"})
        {
            var metadata = Assert.Single(formatters!, item => item.Name == name);
            Assert.Equal(new FormatterParameterDescription("suffix", "String", true), Assert.Single(metadata.Parameters));
            using var invalid = Json("""{"value":"Ada","parameters":{"suffix":"?"}}""");
            using var failure = await client.PostAsync($"/api/formatters/{name}/execute", invalid);
            Assert.Equal(HttpStatusCode.BadRequest, failure.StatusCode);
            using var problem = JsonDocument.Parse(await failure.Content.ReadAsStringAsync());
            Assert.Equal("suffix must be !", problem.RootElement.GetProperty("errors")[0].GetString());
            using var valid = Json("""{"value":"Ada","parameters":{"suffix":"!"},"culture":"en-US"}""");
            using var success = await client.PostAsync($"/api/formatters/{name}/execute", valid);
            success.EnsureSuccessStatusCode();
            var result = await success.Content.ReadFromJsonAsync<FormatterExecutionResult>();
            Assert.Equal("Ada!en-US", result!.Result);
        }

        Assert.Equal(2, formatter.FormatCalls);
    }

    [Fact]
    public async Task Discovery_and_execution_resolve_scoped_formatters_in_the_request_scope()
    {
        using var factory = new WebApplicationFactory<Program>().WithWebHostBuilder(builder =>
            builder.ConfigureServices(services => services.AddKeyedScoped<IFormatter, ScopedProbeFormatter>("scoped")));
        using var client = factory.CreateClient();
        var formatters = await client.GetFromJsonAsync<FormatterDescription[]>("/api/formatters");
        Assert.Single(formatters!, formatter => formatter.Name == "scoped");

        var results = new List<string>();
        for (var request = 0; request < 2; request++)
        {
            using var content = Json("""{"value":"Ada"}""");
            using var response = await client.PostAsync("/api/formatters/scoped/execute", content);
            response.EnsureSuccessStatusCode();
            var result = await response.Content.ReadFromJsonAsync<FormatterExecutionResult>();
            Assert.StartsWith("Ada:", result!.Result);
            results.Add(result.Result);
        }

        Assert.NotEqual(results[0], results[1]);
    }

    private static StringContent Json(string body) => new(body, Encoding.UTF8, "application/json");

    private sealed class ScopedProbeFormatter : IFormatter
    {
        private readonly Guid _id = Guid.NewGuid();
        private bool _validated;

        public FormatterMetadata Metadata { get; } = new([]);

        public FormatterValidationResult Validate(IReadOnlyDictionary<string, FormatterOperand> parameters)
        {
            _validated = true;
            return FormatterValidationResult.Valid;
        }

        public string Format(JsonElement value, IReadOnlyDictionary<string, FormatterOperand> parameters,
            FormatterContext context)
        {
            if (!_validated) throw new InvalidOperationException("The formatter was not validated in this scope.");
            return value.GetString() + ":" + _id;
        }
    }

    private sealed class ProbeFormatter : IFormatter
    {
        public int FormatCalls { get; private set; }
        public FormatterMetadata Metadata { get; } = new([new FormatterParameter("suffix", typeof(string), true)]);

        public FormatterValidationResult Validate(IReadOnlyDictionary<string, FormatterOperand> parameters)
        {
            return parameters["suffix"].Literal == "!"
                ? FormatterValidationResult.Valid
                : new FormatterValidationResult(["suffix must be !"]);
        }

        public string Format(JsonElement value, IReadOnlyDictionary<string, FormatterOperand> parameters,
            FormatterContext context)
        {
            FormatCalls++;
            return value.GetString() + parameters["suffix"].Literal + context.Culture.Name;
        }
    }
}

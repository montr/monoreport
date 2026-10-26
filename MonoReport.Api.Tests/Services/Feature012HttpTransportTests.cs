using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.AspNetCore.Mvc.Testing;
using Microsoft.AspNetCore.WebUtilities;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Net.Http.Headers;
using MonoReport.Models;

namespace MonoReport.Tests.Services;

public sealed class Feature012HttpTransportTests
{
    private const string Word = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    private const string Excel = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    [Fact]
    public async Task Test_owned_xlsx_generates_through_public_http()
    {
        var fixture = new XlsxTransportFixture();
        using var factory = new WebApplicationFactory<Program>(); using var client = factory.CreateClient();
        using var form = HttpApiTests.Form(fixture.Source, fixture.Data);
        using var response = await client.PostAsync("/api/documents/generate", form);
        Assert.Equal(HttpStatusCode.OK, response.StatusCode); AssertCount(response, 0);
        Assert.Equal(Excel, response.Content.Headers.ContentType!.MediaType);
        fixture.AssertGenerated(await response.Content.ReadAsByteArrayAsync());
    }

    [Theory]
    [InlineData(false, 0)] [InlineData(false, 1)] [InlineData(false, 3)]
    [InlineData(true, 0)] [InlineData(true, 1)] [InlineData(true, 3)]
    public async Task Multipart_preserves_document_bytes_and_original_core_warnings_with_exact_framing(bool spreadsheet, int count)
    {
        var fixture = new TransportFixture(); var source = fixture.Create(spreadsheet, count);
        var core = fixture.Merge(source); Assert.Equal(count, core.Warnings.Count);
        using var factory = new WebApplicationFactory<Program>(); using var client = factory.CreateClient();
        using var form = HttpApiTests.Form(source, TransportFixture.Data);
        using var request = new HttpRequestMessage(HttpMethod.Post, "/api/documents/generate") { Content = form };
        request.Headers.Accept.ParseAdd("multipart/mixed");
        using var response = await client.SendAsync(request);
        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        Assert.Equal("multipart/mixed", response.Content.Headers.ContentType!.MediaType);
        Assert.Contains("Accept", response.Headers.Vary);
        Assert.Null(response.Content.Headers.ContentDisposition);
        AssertCount(response, count);
        var boundary = response.Content.Headers.ContentType.Parameters.Single(p => p.Name == "boundary").Value!.Trim('"');
        Assert.StartsWith("monoreport-", boundary);
        var bytes = await response.Content.ReadAsByteArrayAsync();
        Assert.Equal(bytes.LongLength, response.Content.Headers.ContentLength);
        var reader = new MultipartReader(boundary, new MemoryStream(bytes));
        var document = (await reader.ReadNextSectionAsync())!;
        Assert.Equal(spreadsheet ? Excel : Word, document.ContentType);
        var extension = spreadsheet ? "xlsx" : "docx";
        Assert.Equal($"attachment; filename=\"document.{extension}\"", document.ContentDisposition);
        using var documentBody = new MemoryStream(); await document.Body.CopyToAsync(documentBody);
        var attachment = documentBody.ToArray();
        fixture.AssertPackage(attachment, spreadsheet);
        AssertPackageEntriesEqual(core.Document, attachment);
        var diagnostics = (await reader.ReadNextSectionAsync())!;
        Assert.Equal("application/json", diagnostics.ContentType);
        Assert.Equal("attachment; filename=\"diagnostics.json\"", diagnostics.ContentDisposition);
        using var diagnosticBody = new MemoryStream(); await diagnostics.Body.CopyToAsync(diagnosticBody);
        Assert.Null(await reader.ReadNextSectionAsync());
        using var json = JsonDocument.Parse(diagnosticBody.ToArray());
        Assert.Equal(new[] { "fileName", "format", "warnings" }, json.RootElement.EnumerateObject().Select(p => p.Name).Order());
        Assert.Equal(extension, json.RootElement.GetProperty("format").GetString());
        Assert.Equal("document." + extension, json.RootElement.GetProperty("fileName").GetString());
        var expectedWarnings = JsonSerializer.SerializeToElement(core.Warnings, new JsonSerializerOptions(JsonSerializerDefaults.Web));
        Assert.True(JsonElement.DeepEquals(expectedWarnings, json.RootElement.GetProperty("warnings")));
        Assert.Equal(count, json.RootElement.GetProperty("warnings").GetArrayLength());
        // Reconstruct the full wire body from independently specified framing and the parsed payloads.
        using var expected = new MemoryStream();
        expected.Write(Encoding.ASCII.GetBytes($"--{boundary}\r\nContent-Type: {(spreadsheet ? Excel : Word)}\r\nContent-Disposition: attachment; filename=\"document.{extension}\"\r\n\r\n"));
        expected.Write(attachment);
        expected.Write(Encoding.ASCII.GetBytes($"\r\n--{boundary}\r\nContent-Type: application/json\r\nContent-Disposition: attachment; filename=\"diagnostics.json\"\r\n\r\n"));
        expected.Write(diagnosticBody.ToArray());
        expected.Write(Encoding.ASCII.GetBytes($"\r\n--{boundary}--\r\n"));
        Assert.Equal(expected.ToArray(), bytes);
        using var binaryForm = HttpApiTests.Form(source, TransportFixture.Data);
        using var binary = await client.PostAsync("/api/documents/generate", binaryForm);
        Assert.Equal(spreadsheet ? Excel : Word, binary.Content.Headers.ContentType!.MediaType);
        Assert.Equal("document." + extension, binary.Content.Headers.ContentDisposition!.FileNameStar);
        Assert.Contains("Accept", binary.Headers.Vary); AssertCount(binary, count);
        AssertPackageEntriesEqual(attachment, await binary.Content.ReadAsByteArrayAsync());
        using var freshForm = HttpApiTests.Form(source, TransportFixture.Data);
        using var freshRequest = new HttpRequestMessage(HttpMethod.Post, "/api/documents/generate") { Content = freshForm };
        freshRequest.Headers.Accept.ParseAdd("multipart/mixed");
        using var fresh = await client.SendAsync(freshRequest);
        Assert.NotEqual(response.Content.Headers.ContentType.ToString(), fresh.Content.Headers.ContentType!.ToString());
    }

    public static IEnumerable<object?[]> NegotiationCases()
    {
        foreach (var spreadsheet in new[] { false, true })
        {
            var binary = spreadsheet ? Excel : Word;
            foreach (var (accept, mixed) in new (string?, bool)[]
            {
                (null, false), ("*/*", false), ("application/*", false), ("multipart/*", false),
                ("multipart/mixed", true), ("Multipart/Mixed;q=0.5", true),
                ("multipart/mixed;q=bogus", false), ("multipart/mixed;q=1.5", false),
                ("multipart/mixed;q=-1", false),
                ("multipart/mixed;q=0.5junk", false), ("multipart/mixed;q=1junk", false),
                ("multipart/mixed ; Q = 0.500 \t", true), ("multipart/mixed;q=0.001", true),
                ("multipart/mixed;q=1.000", true), ("multipart/mixed;q=0.000", false),
                ("multipart/mixed;q=1.001", false), ("multipart/mixed;q=-0.1", false),
                ("multipart/mixed;q=NaN", false), ("multipart/mixed;q=Infinity", false),
                ("multipart/mixed;q=1e0", false), ("multipart/mixed;q=+0.5", false),
                ("multipart/mixed;q=.5", false), ("multipart/mixed;q=0.5.0", false),
                ("multipart/mixed;q=\"0.5\"", false), ("multipart/mixed;q=0. 5", false),
                ("multipart/mixed;q=", false), ("multipart/mixed;q", false),
                ("multipart/mixed;q=0.5;q=bogus", false),
                ($"{binary};q=0.5junk, multipart/mixed;q=1", false),
                ($"{binary};q=0.500, multipart/mixed;q=0.5", false),
                ("multipart/mixed;q=0", false), ("multipart/mixed;q=0, */*;q=1", false),
                ("multipart/*;q=1, multipart/mixed;q=0", false), ("text/plain", false),
                ("not a media type", false), ("multipart/mixed;q=0.2, */*;q=1", true),
                ($"{binary};q=0.8, multipart/mixed;q=0.7", false),
                ($"{binary};q=0.7, multipart/mixed;q=0.8", true),
                ($"{binary}, multipart/mixed", false), ($"multipart/mixed;q=0.5, {binary};q=0.5", false),
                ($"{binary};q=0, multipart/mixed;q=0.1", true),
                ($"{binary};q=0, multipart/mixed;q=0", false)
            }) yield return [spreadsheet, accept, mixed];
        }
    }

    [Theory]
    [MemberData(nameof(NegotiationCases))]
    public async Task Accept_uses_explicit_opt_in_quality_binary_ties_and_fallback(bool spreadsheet, string? accept, bool mixed)
    {
        using var factory = new WebApplicationFactory<Program>(); using var client = factory.CreateClient();
        using var request = new HttpRequestMessage(HttpMethod.Post, "/api/documents/generate")
            { Content = HttpApiTests.Form(new TransportFixture().Create(spreadsheet, 0), TransportFixture.Data) };
        if (accept is not null) request.Headers.TryAddWithoutValidation("Accept", accept);
        using var response = await client.SendAsync(request);
        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        Assert.Equal(mixed ? "multipart/mixed" : spreadsheet ? Excel : Word, response.Content.Headers.ContentType!.MediaType);
        Assert.Contains("Accept", response.Headers.Vary);
    }

    [Theory]
    [InlineData(false, false)] [InlineData(true, false)]
    [InlineData(false, true)] [InlineData(true, true)]
    public async Task Multipart_generation_or_diagnostic_preparation_failure_returns_only_problem_details(bool spreadsheet, bool serializationFailure)
    {
        var converter = new FailingDiagnosticConverter();
        using var factory = new WebApplicationFactory<Program>().WithWebHostBuilder(builder => builder.ConfigureServices(services =>
        {
            if (serializationFailure) services.ConfigureHttpJsonOptions(options => options.SerializerOptions.Converters.Add(converter));
        }));
        using var client = factory.CreateClient();
        using var request = new HttpRequestMessage(HttpMethod.Post, "/api/documents/generate")
        {
            Content = HttpApiTests.Form(new TransportFixture().Create(spreadsheet, 1), serializationFailure ? TransportFixture.Data : "[]")
        };
        // DOCX permits array roots, so force a structural type error for its missing collection.
        if (!serializationFailure && !spreadsheet)
        {
            request.Content.Dispose(); request.Content = HttpApiTests.Form(new TransportFixture().Create(false, 1), "{\"First\":123}");
        }
        request.Headers.Accept.ParseAdd("multipart/mixed");
        using var response = await client.SendAsync(request);
        Assert.Equal(serializationFailure ? HttpStatusCode.InternalServerError : HttpStatusCode.UnprocessableEntity, response.StatusCode);
        Assert.Equal("application/problem+json", response.Content.Headers.ContentType!.MediaType);
        Assert.Null(response.Content.Headers.ContentDisposition); AssertCount(response, 0);
        var text = await response.Content.ReadAsStringAsync();
        Assert.DoesNotContain("monoreport-", text); Assert.DoesNotContain("private preparation sentinel", text);
        using var json = JsonDocument.Parse(text);
        Assert.Equal(serializationFailure ? "server.internal_error" : DiagnosticCodes.GenerationInvalidData, json.RootElement.GetProperty("code").GetString());
        Assert.Equal(serializationFailure ? 1 : 0, converter.Calls);
    }

    private static void AssertCount(HttpResponseMessage response, int count)
    {
        if (count == 0) Assert.False(response.Headers.Contains("X-MonoReport-Warning-Count"));
        else Assert.Equal(count.ToString(System.Globalization.CultureInfo.InvariantCulture), Assert.Single(response.Headers.GetValues("X-MonoReport-Warning-Count")));
    }

    private static void AssertPackageEntriesEqual(byte[] expected, byte[] actual)
    {
        // ZIP timestamps may differ between independent merges; every uncompressed package byte must agree.
        using var left = new System.IO.Compression.ZipArchive(new MemoryStream(expected));
        using var right = new System.IO.Compression.ZipArchive(new MemoryStream(actual));
        Assert.Equal(left.Entries.Select(e => e.FullName).Order(), right.Entries.Select(e => e.FullName).Order());
        foreach (var entry in left.Entries)
        {
            using var a = new MemoryStream(); using var b = new MemoryStream();
            using (var stream = entry.Open()) stream.CopyTo(a);
            using (var stream = right.GetEntry(entry.FullName)!.Open()) stream.CopyTo(b);
            Assert.Equal(a.ToArray(), b.ToArray());
        }
    }

    private sealed class FailingDiagnosticConverter : JsonConverter<MergeDiagnostic>
    {
        internal int Calls { get; private set; }
        public override MergeDiagnostic Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options) => throw new NotSupportedException();
        public override void Write(Utf8JsonWriter writer, MergeDiagnostic value, JsonSerializerOptions options)
        {
            Calls++; writer.WriteStartObject(); writer.WriteString("message", "private preparation sentinel");
            throw new InvalidOperationException("private preparation sentinel");
        }
    }
}

# MonoReport

MonoReport generates DOCX and XLSX documents from templates and JSON data — reports, contracts, invoices, spreadsheets, and other business documents.

**Use the editors you already know.** Create and maintain templates in Microsoft Word, Excel, or other compatible editors. No need to learn a new report designer or specialized editing tools.

**Preserve your document design.** Generated documents retain the template's layout, formatting, tables, styles, and supported Excel formulas.

**Separate data from presentation.** Templates define the appearance, while JSON provides the content. Reuse templates with different data to generate consistent documents.

## Table of Contents

1. [Ways to Use](#ways-to-use)
2. [Playground & Examples](#playground--examples)
3. [How It Works](#how-it-works)

## Ways to Use

MonoReport can be used as an HTTP API, a .NET library, or a command-line tool.

### HTTP API

Render documents or analyze templates through REST endpoints. Suitable for integration with web applications, backend services, and automation workflows.

**Analyze a template:**

```bash
curl -X POST http://localhost:8080/api/documents/analyze \
  -F "template=@Samples/sample.template.docx"
```

Returns template validation results, diagnostics, and required data fields as JSON.

**Render a document:**

```bash
curl -X POST http://localhost:8080/api/documents/render \
  -F "template=@Samples/sample.template.docx" \
  -F "data=<Samples/sample.json" \
  -o result.docx
```

The `data` form field contains the JSON text loaded from `Samples/sample.json`.

**Run with Docker:**

Run the API and Playground using the prebuilt image from GitHub Container Registry:

```bash
docker run -d --name monoreport -p 8080:8080 ghcr.io/montr/monoreport:latest
```

Open the Playground at [http://localhost:8080](http://localhost:8080). See the [repository](https://github.com/montr/monoreport) for deployment examples.

### .NET Library

Integrate document generation directly into .NET applications without an HTTP service.

Register MonoReport with dependency injection:

```csharp
services.AddMonoReport();
```

Inject `DocumentGenerator` and generate a document (excerpt; assumes `generator` and `json` are available):

```csharp
using System.Text.Json;

using var template = File.OpenRead("template.docx");
using var data = JsonDocument.Parse(json);

GenerationResult result = generator.Generate(template, data);

File.WriteAllBytes("result.docx", result.DocumentBytes);
```

### CLI

Generate documents from the terminal, scripts, or automated workflows.

Run from source:

```bash
dotnet run --project MonoReport.Cli -- \
  --template Samples/sample.template.docx \
  --json-file Samples/sample.json \
  --output result.docx
```

Alternatively, install MonoReport as a .NET tool and run it directly (confirm the published package and command names before using these commands):

```bash
dotnet tool install --global MonoReport.Cli

monoreport --template sample.template.docx \
  --json-file sample.json \
  --output result.docx
```

## Playground & Examples

Explore MonoReport in the browser. Try DOCX and XLSX templates with sample JSON data, analyze templates, preview generated documents, and download the results.

**[Try the Playground & Examples](https://demo.monoreport.net)**

## How It Works

MonoReport combines a DOCX or XLSX template with JSON data to generate a document.

1. **Create a template** in Word, Excel, or another compatible editor.
2. **Provide JSON data** containing values and collections.
3. **Generate the document** using the API, .NET library, CLI, or Playground.

### DOCX Templates

Insert JSON values into Word documents, repeat paragraphs and table rows, add conditional content, and format values while retaining supported document styling.

[DOCX Template Guide](docs/docx-template-guide.md) · [DOCX Compatibility](docs/docx-template-compatibility.md)

### XLSX Templates

Populate Excel cells, expand rows and columns from JSON collections, and preserve formatting and supported formulas.

[XLSX Template Guide](docs/xlsx-template-guide.md) · [XLSX Compatibility](docs/xlsx-template-compatibility.md)

## License

[Apache-2.0 license](LICENSE)

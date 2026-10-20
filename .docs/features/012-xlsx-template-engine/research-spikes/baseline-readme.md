# Feature 012 research spikes

1. These are **fixture-specific architecture experiments**, not a reusable XLSX engine, approved syntax, production code or production tests. `fixtures.py` supplies the row plan, table ownership, field bindings and reference scopes explicitly. `Program.cs` exercises parsed formula references and Open XML validation. `advanced.py` creates static feature-preservation and deliberately stale-reference fixtures.
2. Retained XLSX files are generated research artifacts. `unsafe-shift.xlsx` is **intentionally semantically incorrect**, although it passes schema validation. Do not use it as a sample generated report. Named `MR_` regions, repeated parent labels, control hiding and scope rules illustrate proposals awaiting Product Owner approval.
3. Actual execution workspace: `/private/tmp/feature012-research/`. Temporary files there include `downloads/`, `bin/`, `obj/`, `lo-profile/`, `evidence/pending.xlsx`, formula request/response intermediates, and `recalculated/`. Those caches/downloads/intermediates are not repository dependencies. Retained source, fixtures, calculation results and command results are copied here for review/reproduction.
4. Environment used: .NET SDK 10.0.102, Open XML SDK 3.5.1, ClosedXML.Parser 2.0.0/netstandard2.1, Python standard library, LibreOffice 26.8.0.3. No Python packages, production package changes, browser automation or Office Automation were used. Open XML validation uses the Office 2010 target.

## 1. Dependencies and provenance

1. Open XML SDK is an existing repository dependency. Its tested net10.0 closure is DocumentFormat.OpenXml 3.5.1 → DocumentFormat.OpenXml.Framework 3.5.1 → System.IO.Packaging 10.0.2; all three local NuGet manifests declare MIT. The net10.0 Packaging group has no dependencies.
2. ClosedXML.Parser 2.0.0 was checked before use: its NuGet manifest declares MIT and the selected netstandard2.1 target has no dependencies. The netstandard2.0 target has a separate System.Memory dependency and is not selected by the spike. Repository commit in the package is `658973aeedc2fe289e0ccba2bb9959f67692ded5`.
3. Download the pinned parser package from [NuGet](https://api.nuget.org/v3-flatcontainer/closedxml.parser/2.0.0/closedxml.parser.2.0.0.nupkg). Expected SHA-256: `85d40168dcb9fa5478c52c6ef26820beedc8e2cb764515906ed37185584ba37f`. Extract it to `downloads/parser/` so `downloads/parser/lib/netstandard2.1/ClosedXML.Parser.dll` exists. The project uses that assembly directly and does not compile downloaded source.
4. The standalone csproj restores OpenXml from the existing cache or configured package source. It is not included in `MonoReport.slnx`. Do not add the parser to production as part of reproducing this research.

## 2. Reproduction commands

1. Copy `Spike.csproj`, `Program.cs`, `fixtures.py`, and `advanced.py` into a fresh temporary workspace, for example `/private/tmp/feature012-research-replay/`. Run the following commands there after extracting the pinned parser package as described above. This keeps build outputs, fixtures and application profiles isolated from the repository.
2. Generate source XLSX/JSON and supplied geometry/binding requests: `python3 fixtures.py prepare`.
3. Compile and probe reference parsing/copy/relocation: `dotnet run --project Spike.csproj -- probe parser-results.json`.
4. Transform fixture formulas: `dotnet run --project Spike.csproj --no-build -- transform evidence/formula-requests.json evidence/formulas.json`.
5. Complete output and run geometry/reference/style/metadata assertions: `python3 fixtures.py finish`.
6. Create advanced static and intentionally unsafe fixtures: `python3 advanced.py`.
7. Validate all retained packages: `dotnet run --project Spike.csproj --no-build -- validate evidence/template.xlsx evidence/output.xlsx evidence/compatibility.xlsx evidence/compatibility-output.xlsx evidence/boundaries.xlsx evidence/advanced-compatibility.xlsx evidence/advanced-compatibility-output.xlsx evidence/unsafe-shift.xlsx`.
8. With application execution authorized, recalculate using the installed compatible application and a fresh isolated profile. For the example replay path: `/Applications/LibreOffice.app/Contents/MacOS/soffice -env:UserInstallation=file:///private/tmp/feature012-research-replay/lo-profile --headless --convert-to xlsx --outdir /private/tmp/feature012-research-replay/recalculated /private/tmp/feature012-research-replay/evidence/output.xlsx /private/tmp/feature012-research-replay/evidence/boundaries.xlsx /private/tmp/feature012-research-replay/evidence/advanced-compatibility-output.xlsx`. A different LibreOffice installation requires its corresponding executable path. No active personal profile is used.
9. Assert recalculated values and boundary counterexamples: `python3 fixtures.py recalc`.
10. Compare `parser-results.json` and `evidence/recalculation-results.json` with the retained records. Expected commands/results and the corrected early harness failures are recorded in `evidence/results.txt`. This is a small deterministic probe, not performance or comprehensive compatibility verification.

## 3. Artifact map

| File | Purpose |
| --- | --- |
| `evidence/template.xlsx`, `evidence/input.json` | Source two-level/two-axis case, four positional leaves and independent second table. Merges are explicitly authored fixed rectangles. |
| `evidence/output.xlsx` | Selectively copied/expanded output before external calculation; 61 transformed formulas, 19 merge copies, two sheets. |
| `evidence/compatibility.xlsx`, `evidence/compatibility-output.xlsx` | Static scalar edit with table, name, conditional formatting and validation. |
| `evidence/advanced-compatibility.xlsx`, `evidence/advanced-compatibility-output.xlsx` | Same static case with chart, valid PNG drawing and legacy comment. Unchanged parts/relationships are compared byte-for-byte. |
| `evidence/boundaries.xlsx` | Direct mapping versus unchanged/rebased helper formulas, including `ROWS` and `COUNTBLANK` differences. |
| `evidence/unsafe-shift.xlsx` | Negative semantic case: cells moved but table/name/CF/DV/chart/comment metadata remains stale. Zero schema errors must not be interpreted as correct output. |
| `evidence/recalculated-output.xlsx`, `evidence/recalculated-boundaries.xlsx`, `evidence/recalculated-advanced-compatibility-output.xlsx` | Actual LibreOffice import/export results; formula values are asserted on the first two. They are validation evidence, not authoritative preservation outputs because LibreOffice may normalize metadata. |
| `parser-results.json`, `evidence/recalculation-results.json`, `evidence/results.txt` | Observed capability forms, independently expected calculation values and execution record. |

1. The Python harness regenerates fixtures from its explicit case and writes `input.json`; it does not accept arbitrary user templates or data. Marker inference, generic hierarchy planning, field parser integration, empty inputs and comprehensive formula rejection are intentionally absent.
2. The C# mapper handles only supplied local rectangular images and a specific separation of scopes. It rejects tested disjoint images and ambiguous single-cell expansion; it does not implement production reference binding, absolute repeated-target semantics, special formula groups or general cross-sheet resolution. The cross-sheet workbook reference is relocated explicitly in the fixture generator.
3. Static feature preservation does not establish dynamic resizing/duplication support. The no-stretch geometry repeats parent labels per leaf and does not provide variable parent-spanning merges. Manual Excel rendering, repair prompts, pagination, shared formula packages and realistic size/performance checks remain unverified.

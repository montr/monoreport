# Product Backlog

Potential work for MonoReport. Items are not approved Features; scope and priorities require Product Owner approval.

## Product Direction

1. **Extended formatting** — additional field formats and currency support.
2. **MCP integration** — revisit when concrete consumers and use cases are identified.
3. **AI-Powered Template Generation** - Automatically convert an existing DOCX/XLSX report into a fully functional MonoReport template and JSON containing the original data, with AI generating all fields, loops, and conditions while preserving the original document’s structure, layout, and formatting.

## Engineering Backlog

1. **Release automation** — GitHub Releases and automatic release notes from commits.
2. **UI polish** — fix remaining minor visual issues, after confirming they still exist.
3. **Ant Design review** — consistent and appropriate use of components.
4. **Dependency audit** — review current npm audit findings.
5. **C# formatting and code style** — establish and automate consistent rules.
6. **UI performance** — measure bundle size and loading before optimizing.
7. **API documentation** — verify OpenAPI and interactive documentation coverage; address confirmed gaps only.
8. **Optional template data schema:** Allow templates to include an optional JSON data schema for validating input data and field references. Evaluate JSON Schema compatibility.
9. **Configurable diagnostic severity:** Allow individual diagnostic codes to be configured as warnings or errors, similar to compiler diagnostics.

## Data Source Catalog

**Status:** Backlog  
**Priority:** TBD

### Purpose

Define an open, lightweight, versioned protocol that allows independent systems to expose metadata about available datasets and optionally provide sample or real data.

Any system implementing this protocol can act as a Data Source. Consumers may include MonoReport Web Playground, Office Playground, and third-party applications.

### Architectural Principles

1. A Data Source is an independent HTTP API, with no dependency on MonoReport Core, Analyzer, or document generation.
2. MonoReport Service will provide the initial Data Source implementation but is not a mandatory intermediary.
3. Clients may connect directly to external Data Sources, such as iTender.
4. MonoReport Service may eventually connect to external Data Sources and expose them through its own API, acting as an optional proxy or aggregator.
5. Data Definitions describe available datasets rather than specific document templates. Multiple DOCX or XLSX templates may consume the same definition.
6. The protocol must support basic versioning and stable Data Definition identifiers.
7. Authentication is implementation-specific and may include anonymous access, SSO/OAuth, or service credentials.
8. Data Sources must remain independent of MonoReport Analyzer. They provide metadata and data only.

### Protocol Capabilities

1. **Capability Discovery:** Discover which operations a Data Source supports.
2. **List Definitions:** Retrieve available Data Definitions.
3. **Get Definition:** Retrieve a dataset's field structure, including nested objects and collections.
4. **Get Sample Data (optional):** Retrieve representative JSON data for testing and document generation.
5. **Get Data (optional):** Retrieve real JSON data, potentially requiring parameters and authorization.
6. **Extended Field Metadata (optional):** Provide descriptions, semantic data types, example values, and other authoring hints.

Capability Discovery allows clients to adapt their interfaces to the operations supported by each source.

Data Definition metadata describes what a source can provide. It is conceptually distinct from Analyzer `requirements`, which describe the data dependencies of a particular template.

### Initial Implementation

1. Implement the Data Source HTTP API in MonoReport Service.
2. Load Data Definitions from static JSON files.
3. Support configurable file locations, including externally mounted Docker volumes.
4. Optionally load sample JSON datasets from files.
5. Implement basic protocol versioning and capability discovery.
6. Integrate the protocol into Web Playground for dataset selection, field discovery, and sample data retrieval.

### Future Extensions

1. Independent Data Source implementation in iTender.
2. Multiple Data Source connections per client.
3. Parameterized real-data retrieval with authentication and authorization.
4. External Data Source integration through MonoReport Service.
5. Richer field metadata to improve template authoring.

### Open Questions

1. Data Definition schema and field metadata representation.
2. Capability discovery format and protocol version compatibility rules.
3. Parameter declaration and validation for real-data retrieval.
4. Multiple-source configuration, registration, and selection.
5. Proxying, authentication, and user identity propagation.

---

## MonoReport Office Playground (Word / Excel Add-in)

**Status:** Backlog  
**Priority:** TBD  
**Dependencies:** MonoReport Core, Data Source Protocol, document analysis and generation APIs

### Purpose

Provide template-authoring assistance, validation, and test document generation directly inside Microsoft Word and Excel.

Word and Excel remain the document editors. MonoReport provides expression authoring tools, field discovery, analysis, and generation without attempting to replace native Office editing capabilities.

### Core Capabilities

1. **Data Source Connections:** Connect to one or more Data Sources, including direct connections to systems such as iTender or connections through MonoReport Service.
2. **Data Definitions:** Browse available datasets and their field structures.
3. **Field Browser:** Search for fields and insert expressions such as `{Customer.Name}` into the current document.
4. **Expression Helpers:** Assist with inserting formatters and `For`, `If`, and `Else` constructs while respecting DOCX/XLSX structural constraints.
5. **Analyze:** Analyze the current document using MonoReport Analyzer, display diagnostics, and navigate to affected document locations.
6. **Data Compatibility Check:** Optionally compare template dependencies reported by Analyzer against the selected Data Definition. This is a separate compatibility check, not a Data Source responsibility or a syntax validation rule.
7. **Sample Data:** Retrieve sample JSON from a Data Source when supported.
8. **Real Data:** Retrieve real JSON when supported and authorized.
9. **Generate:** Generate a document from the current template and selected JSON using MonoReport generation APIs.
10. **Preview Result:** Open or preview the generated document in Word or Excel without modifying the original template.

### Architectural Principles

1. Office Playground is a client of MonoReport APIs and the Data Source Protocol, not a separate template engine.
2. Expression parsing, template analysis, and document generation must reuse existing MonoReport functionality.
3. Data Sources provide metadata and data only; they do not implement Analyzer.
4. Direct Data Source connections must not require MonoReport Service as an intermediary.
5. The UI must adapt to each Data Source's discovered capabilities.
6. Word and Excel should share the same conceptual experience and reusable UI components where practical, while allowing separate Office.js integrations.
7. Native Office capabilities should handle document editing, rendering, navigation, and zoom.
8. The Add-in must not depend on iTender or any other specific business system.
9. Document generation must not overwrite or damage the original template.
10. Authentication and authorization must follow the requirements of the selected Data Source.

### Example User Workflow

1. Open a DOCX template in Microsoft Word.
2. Open the MonoReport task pane.
3. Select a Data Source and Data Definition.
4. Browse available fields and insert expressions into the document.
5. Add loops, conditions, and formatters using authoring helpers.
6. Run Analyze and navigate to detected issues.
7. Retrieve sample data from the Data Source or provide custom JSON.
8. Run Generate using the current template and selected data.
9. Open the generated document in Word while preserving the original template.

An equivalent workflow should be supported in Excel, accounting for cell, range, and worksheet semantics and the eventual XLSX template syntax.

### Development Phases

1. **Word MVP:** Data Source connection, field browsing, expression insertion, and DOCX analysis.
2. **Word Playground:** Sample data retrieval, generation, and opening generated results in Word.
3. **Excel Add-in:** Equivalent functionality after XLSX template syntax and generation are implemented.
4. **Advanced Integrations:** Real-data retrieval, SSO, enterprise Data Sources, and additional template-authoring assistance.

### Open Questions

1. Office.js capabilities and limitations for accessing current documents and inserting structural expressions.
2. Mapping Analyzer diagnostics to precise Word/Excel document locations.
3. Safely opening generated DOCX/XLSX documents without modifying the source template.
4. Authentication for direct enterprise Data Source connections, including SSO, CORS, and Office WebView constraints.
5. Storage and management of Data Source connection settings and selected Data Definitions.
6. Feasibility of a shared Word/Excel Add-in architecture given differences between Office APIs.

## DOCX — Architecture Review: Data Hierarchy vs. Template Presentation

**Goal:** Verify that the MonoReport DOCX engine follows the principle **“Data defines the hierarchy; the template defines its presentation”**, consistent with the architectural direction adopted for XLSX.

### Review Scope

1. Verify that object hierarchy is determined exclusively by input data, independently of the DOCX template structure.
2. Verify that parent (`../`) and root references follow the data hierarchy, including ancestors not represented in the template.
3. Check whether templates can omit intermediate data levels while still rendering their descendants.
4. Verify that missing data levels do not require synthetic objects or placeholder sections.
5. Ensure input order and record multiplicity are preserved without implicit deduplication.
6. Review nested repetitions, scope resolution, and field binding for predictable behavior.

### Expected Behavior Example

**Data:** Company → Department → Employee.

**Template:** Renders companies and employees, omitting department sections, but displays each employee's department name using a parent reference (`../`).

**Expected result:** All employees appear in input order, with their actual department names, even though departments are not rendered as separate sections.

### Deliverables

1. Document the actual DOCX engine architecture and identify deviations from the data-authoritative principle.
2. Recommend minimal corrections and focused regression tests where necessary.
3. Propose architectural rules and worked examples for developer documentation.
4. Identify any DOCX-specific constraints that justify differences from XLSX behavior.

**Constraints:** Architecture review only. No implementation, speculative refactoring, or premature unification of DOCX and XLSX engines.

**Priority:** After Feature 012 stabilizes. Not a blocker for XLSX development.

## Documentation Refresh (README-first)

Review and rewrite MonoReport's user-facing documentation as a separate Feature.

1. Keep README as the main, coherent entry point for users. Move large standalone topics (DOCX/XLSX compatibility, architecture, etc.) into separate documents only when justified, with links from README.
2. Emphasize the data-first approach: explain structured input data and JSON, not just document templates. Clearly illustrate the workflow: template + data → generated document.
3. Restore and review illustrations, including the template, JSON input, and generated document. Prefer visual examples before detailed API documentation.
4. Review product positioning, supported capabilities, limitations, getting-started instructions, CLI/API examples, and documentation consistency across DOCX and XLSX.
5. Remove outdated information, unnecessary dependency/version details, duplication, and implementation-oriented explanations that are not useful to users. Preserve practical examples and working commands.

Start with a documentation audit and an agreed update plan. Do not rewrite documentation incrementally during unrelated Feature implementation.

## XLSX Engine Architecture Review

Review the architecture and implementation decisions of the XLSX Template Engine (Feature 012).

1. Assess the necessity and complexity of existing components, including the Logical Scheduler, Geometry Planner, formula handling, and their interactions.
2. Review data hierarchy semantics: data defines hierarchy, templates define presentation. Evaluate `MissingValues`, omitted hierarchy levels, repeated collections, and preservation of data identity.
3. Identify unnecessary abstractions, duplicated logic, implementation limitations that should not become product requirements, and opportunities to simplify the architecture.
4. Evaluate the use of Open XML SDK and opportunities to rely more on existing platform and library capabilities.
5. Document the resulting XLSX Engine architecture, significant decisions, trade-offs, and proposed improvements.
6. Include the deferred Microsoft Excel E1/A14 validation, to be performed after the Web UI becomes available. Record any issues discovered for subsequent correction.

Start with an independent architectural assessment. Do not refactor the engine without a separately approved implementation plan.

## XLSX Performance Benchmark

Create a dedicated, reproducible benchmark project for the XLSX Template Engine.

1. Evaluate workloads from small documents to approximately 1 million rows, varying columns, formulas, styles, and nested collections.
2. Measure execution time by processing stage, peak memory, allocations, generated file size, and throughput.
3. Ensure reproducibility through deterministic inputs, warm-up, repeated measurements, and process isolation where appropriate.
4. Compare performance across implementation versions and identify bottlenecks and scalability limits.
5. Use the existing `capacity-results.md` as historical reference data. Do not promote the temporary Feature 012 benchmark runner into permanent infrastructure without justification.
6. Support local benchmark execution; consider periodic CI benchmarking only if justified.

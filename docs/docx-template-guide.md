# DOCX Template Guide

Create Word document templates that MonoReport populates with JSON data while preserving supported document layout and formatting.

Use Microsoft Word or another DOCX-compatible editor. No special report designer is required.

## Contents

1. [Getting Started](#getting-started)
2. [Fields and Data Paths](#fields-and-data-paths)
3. [Loops](#loops)
4. [Conditions](#conditions)
5. [Formatters](#formatters)
6. [Literal Content](#literal-content)
7. [Working with Word Documents](#working-with-word-documents)
8. [Validation and Troubleshooting](#validation-and-troubleshooting)

For detailed support boundaries, see [DOCX Template Compatibility](docx-template-compatibility.md).

---

## Getting Started

A MonoReport template is an ordinary Word document containing expressions enclosed in curly braces. During generation, MonoReport replaces these expressions with values from JSON data.

**Template → Result**

Create a Word document using the content in the **Word template** column. The **Generated document** column shows the output for the JSON data below.

| Word template | Generated document |
|---|---|
| **INVOICE** | **INVOICE** |
| Invoice: `{Invoice.Number}` | Invoice: `INV-2026-001` |
| Customer: `{Customer.Name}` | Customer: `Acme Corporation` |
| Date: `{Invoice.Date date "MMM dd, yyyy"}` | Date: `Nov 07, 2026` |
| **Total:** `{Invoice.Total number "#,##0.00"}` | **Total:** `1,350.00` |

**JSON data**

```json
{
  "Invoice": {
    "Number": "INV-2026-001",
    "Date": "2026-11-07",
    "Total": 1350
  },
  "Customer": {
    "Name": "Acme Corporation"
  }
}
```

**Try it**

Save the Word template as `invoice.template.docx`. Open the [MonoReport Playground](https://demo.monoreport.net), upload the template, provide the JSON data, and generate the document.

The generated DOCX retains the template's supported layout and formatting. Reuse the template with different JSON data to produce another invoice.

---

## Fields and Data Paths

Fields insert values from JSON into the document. Use dot-separated paths to access nested objects.

**Template → Result**

| Word template | Generated document |
|---|---|
| Customer: `{Customer.Name}` | Customer: `Acme Corporation` |
| City: `{Customer.Address.City}` | City: `New York` |

**JSON data**

```json
{
  "Customer": {
    "Name": "Acme Corporation",
    "Address": { "City": "New York" }
  }
}
```

**Data paths**

| Expression | Meaning |
|---|---|
| `{Name}` | Property in the current context |
| `{Customer.Name}` | Property of a nested object |
| `{Customer.Address.City}` | Deeper nested property |
| `{./Name}` | Explicit property of the current object |
| `{../Number}` | Property in the parent context |
| `{$.Company.Name}` | Property from the JSON root |
| `{.}` | Current primitive array element, inside a loop |

**Contexts**

Inside a loop, a field such as `{Name}` refers to the current item. Use `./` for an explicit current-item reference, `../` for the parent context, and `$.` for the document root. The `{.}` expression inserts the current primitive array element.

**Missing values**

Field names are case-sensitive. Nulls and empty strings render as empty text; missing scalar properties follow the configured missing-value policy. Insert individual properties rather than attempting to render an object or array as a scalar value.

---

## Loops

Use `{For:Collection}` and `{/For}` to repeat content for each element of a JSON array. Loops support paragraphs, table rows, and other supported document structures.

**Repeating paragraphs**

In Word, put each marker and each service name expression in its own paragraph. The table below illustrates *paragraphs*, not a Word table.

| Word template — separate paragraphs | Generated document |
|---|---|
| **Services** | **Services** |
| `{For:Items}` | — |
| `{Name}` | `Consulting` |
| | `Development` |
| | `Support` |
| `{/For}` | — |

**JSON data**

```json
{
  "Items": [
    { "Name": "Consulting" },
    { "Name": "Development" },
    { "Name": "Support" }
  ]
}
```

MonoReport removes the two marker paragraphs and repeats the paragraph between them. The opening and closing markers must occupy dedicated paragraphs.

**Repeating table rows**

For invoice line items, create an actual Word table with a header, two dedicated marker rows, a data row between them, and a total row after them.

**Word template**

| Description | Amount |
|---|---:|
| `{For:Items}` | |
| `{Description}` | `{Amount number "#,##0.00"}` |
| `{/For}` | |
| **Total** | `{Total number "#,##0.00"}` |

**Generated document**

| Description | Amount |
|---|---:|
| `Website design` | `1,200.00` |
| `Hosting` | `150.00` |
| **Total** | **`1,350.00`** |

**JSON data**

```json
{
  "Items": [
    { "Description": "Website design", "Amount": 1200 },
    { "Description": "Hosting", "Amount": 150 }
  ],
  "Total": 1350
}
```

Each marker must occupy its own dedicated paragraph in one cell of its marker row. Other cells in that row must be empty. MonoReport removes the marker rows and repeats the data row. Keep headers and totals outside the loop.

**Nested loops**

A loop can contain another loop. The following template uses separate Word paragraphs for all markers and content.

| Word template — separate paragraphs | Generated document |
|---|---|
| `{For:Departments}` | — |
| **`{Name}`** | **`Engineering`** |
| `{For:Employees}` | — |
| `{Name}` | `Alice` |
| | `Bob` |
| `{/For}` | — |
| | **`Finance`** |
| | `Carol` |
| `{/For}` | — |

**JSON data**

```json
{
  "Departments": [
    {
      "Name": "Engineering",
      "Employees": [{ "Name": "Alice" }, { "Name": "Bob" }]
    },
    {
      "Name": "Finance",
      "Employees": [{ "Name": "Carol" }]
    }
  ]
}
```

Each nested loop establishes a new context. Parent (`../`) and root (`$.`) paths allow access to data outside the current item.

**Arrays of primitive values**

Use `{.}` instead of a property name when looping over strings or numbers.

| Word template — separate paragraphs | Generated document |
|---|---|
| `{For:Tags}` | — |
| `{.}` | `Urgent` |
| | `Approved` |
| | `Internal` |
| `{/For}` | — |

**JSON data**

```json
{ "Tags": ["Urgent", "Approved", "Internal"] }
```

Empty arrays produce no repeated content. Missing or null collections are treated as empty; incompatible non-null values produce an error.

---

## Conditions

Use `{If:Field}` and `{/If}` to include a region when a JSON Boolean is true. Add `{Else}` when you need an alternative for false.

**Simple condition — Template → Result**

The rows below represent separate Word paragraphs, not a Word table.

| Word template | Generated document (`IsPaid: true`) |
|---|---|
| `{If:IsPaid}` | — |
| **PAID** | **PAID** |
| `{/If}` | — |

**JSON data**

```json
{ "IsPaid": true }
```

For `false`, the conditional content is omitted.

**If / Else — Template → Result**

| Word template | Generated document (`IsPaid: false`) |
|---|---|
| `{If:IsPaid}` | — |
| Payment received. | — |
| `{Else}` | — |
| Payment pending. | Payment pending. |
| `{/If}` | — |

**JSON data**

```json
{ "IsPaid": false }
```

Conditions can be nested and combined with loops. Structural markers must be placed in dedicated paragraphs or supported marker rows, not inline in ordinary text. Conditions use Boolean values; missing or null conditions choose the false branch, while incompatible non-null values produce an error.

---

## Formatters

Formatters transform values before inserting them into a document. Use one formatter per expression, with double-quoted arguments where applicable.

**Text formatting — Template → Result**

For `"Name": "acme corporation"`:

| Word template | Generated document |
|---|---|
| `{Name upper}` | `ACME CORPORATION` |
| `{Name lower}` | `acme corporation` |
| `{Name caps}` | `Acme Corporation` |
| `{Name firstcap}` | `Acme corporation` |

**Numbers and dates — Template → Result**

| Word template | Generated document (English culture) |
|---|---|
| Amount: `{Amount number "#,##0.00"}` | Amount: `1,234.50` |
| Date: `{Date date "MMM dd, yyyy"}` | Date: `Nov 07, 2026` |

**JSON data**

```json
{
  "Amount": 1234.5,
  "Date": "2026-11-07"
}
```

Named formatter arguments are also supported: `{Amount number pattern="#,##0.00"}` and `{Date date pattern="MMM dd, yyyy"}`.

Number formatting follows the configured culture. Numeric data can be supplied as a JSON number or invariant-format numeric string. Dates use supported ISO date or date-time values.

**Numbers in words**

Use the optional `spellout` formatter to convert a number to words. Money mode accepts currency and language arguments.

Use `{Amount spellout}` for ordinary number words. For money, use the explicit currency and language options:

| Word template | Generated document (English culture) |
|---|---|
| `{Amount spellout mode="money" currency="USD" language="en"}` | `One thousand two hundred fifty dollars 75 cents` |

**JSON data**

```json
{
  "Amount": 1250.75
}
```

Money mode supports `USD`, `EUR`, and `RUB`; languages include `en` and `ru`. The default language depends on the configured culture. Applications must register the optional formatter using `AddMonoReportSpelloutFormatter()`.

**Custom formatters**

An application can register its own formatters. Their names and arguments depend on the application; available formatter metadata can be discovered through the MonoReport API.

---

## Literal Content

Use literal braces or Raw blocks when the generated document should contain text that looks like a MonoReport expression.

**Literal braces — Template → Result**

| Word template | Generated document |
|---|---|
| `{{` | `{` |
| `}}` | `}` |

**Raw blocks — Template → Result**

These template lines are separate Word paragraphs.

| Word template | Generated document |
|---|---|
| `{Raw}` | — |
| This text contains `{Customer.Name}`. | This text contains `{Customer.Name}`. |
| `{/Raw}` | — |

Raw preserves its body without evaluating expressions or brace escapes. Its markers disappear from the generated document.

Raw blocks may contain multiple paragraphs and supported complete tables; they do not nest. Delimiters must be dedicated sibling paragraphs under the same parent or supported dedicated sibling rows of one table. A Raw region may appear inside a loop or condition, subject to the enclosing operation's compatibility rules.

---

## Working with Word Documents

MonoReport uses the ordinary DOCX features of Word, preserving supported surrounding content and styling.

**Formatting and styles**

Create the layout directly in Word using fonts, colors, styles, paragraph spacing, borders, tables, page layout, images, and headers or footers. Insert expressions into the document text; generated values inherit supported template formatting.

**Paragraphs and table rows**

Structural expressions (`For`, `If`, `Else`, and `Raw`) operate on document blocks, not arbitrary fragments of text.

1. Put structural markers in dedicated paragraphs or valid marker rows.
2. Keep matching markers within supported structural boundaries.
3. Keep static headings, table headers, and totals outside repeated regions.
4. Avoid merged-cell arrangements that conflict with row-level operations.
5. For larger repeating sections, use paragraph-level delimiters around complete supported content, including tables.

**Headers, footers, and other content**

Scalar expressions are supported in ordinary document paragraphs and tables, including headers, footers, footnotes, and endnotes. Word fields, content controls, and specialized opaque content are preserved rather than evaluated.

**Images and document objects**

Existing images can be preserved in repeated content. Other Word structures have specific boundaries: comments and bookmarks may be preserved when unaffected, but local attachments inside `For`/`If` regions are restricted; affected internal hyperlinks require valid unchanged targets; unresolved tracked changes in inspected content are rejected.

See [DOCX Template Compatibility](docx-template-compatibility.md) for the exact supported cases and restrictions.

---

## Validation and Troubleshooting

Analyze a DOCX template before generation to inspect expressions, required data fields, and diagnostic messages. The [MonoReport Playground](https://demo.monoreport.net) provides an interface for template analysis and generation.

**Common problems**

| Symptom | What to check |
|---|---|
| Field is empty | JSON path, capitalization, current loop context, and missing-value policy |
| Loop produces no rows | Array contents and placement of dedicated markers |
| Condition has unexpected output | Boolean value and marker placement |
| Formatter fails | Formatter name, quoted arguments, and registration |
| Table structure is rejected | Dedicated marker rows, merged cells, and block boundaries |
| Word content is rejected | Tracked changes, comments, bookmarks, or internal links |
| Expression remains literal | Raw block or opaque Word structure |

**Before using a template**

1. Start with a simple working template and representative JSON.
2. Add loops and conditions incrementally; check marker placement.
3. Analyze the template after structural changes.
4. Test ordinary, empty, and missing-data cases as appropriate.
5. Open the generated DOCX in Word or another compatible editor to inspect final layout and pagination.

MonoReport validates supported structures but does not recalculate Word's visual layout. For advanced Word features and exact rejection behavior, consult [DOCX Template Compatibility](docx-template-compatibility.md).

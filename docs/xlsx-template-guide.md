# XLSX Template Guide

Create Excel reports from XLSX templates and JSON data using MonoReport. Design the report in Microsoft Excel or another XLSX-compatible editor; MonoReport fills cells and expands supported rows and columns while retaining supported workbook formatting.

## Contents

1. [Getting Started](#getting-started)
2. [Fields and Data Paths](#fields-and-data-paths)
3. [PrintArea and Control Cells](#printarea-and-control-cells)
4. [Vertical Expansion](#vertical-expansion)
5. [Horizontal Expansion](#horizontal-expansion)
6. [Combining Both Directions](#combining-both-directions)
7. [Cell Values, Formatting, and Formulas](#cell-values-formatting-and-formulas)
8. [Validation and Troubleshooting](#validation-and-troubleshooting)

For exact support boundaries, see [XLSX Template Compatibility](xlsx-template-compatibility.md).

---

## Getting Started

A basic XLSX template needs only ordinary cells containing expressions in curly braces. No repetition markers or `PrintArea` are needed until you want to expand rows or columns.

**Template → Result**

| Cell | Excel template | Generated document |
|---|---|---|
| A1 | **SALES SUMMARY** | **SALES SUMMARY** |
| A3 | Customer | Customer |
| B3 | `{Customer.Name}` | `Acme Corporation` |
| A4 | Orders | Orders |
| B4 | `{Orders}` | `12` |
| A5 | Revenue | Revenue |
| B5 | `{Revenue}` | `1350` |

**JSON data**

```json
{
  "Customer": { "Name": "Acme Corporation" },
  "Orders": 12,
  "Revenue": 1350
}
```

**Try it**

Enter the contents in the indicated Excel cells and save the workbook as `sales.template.xlsx`. Open the [MonoReport Playground](https://demo.monoreport.net), upload the workbook, supply the JSON data, and generate the result.

Cell formatting comes from Excel. For example, apply an Excel number format to B5 if you want the numeric result displayed as `1,350.00`.

---

## Fields and Data Paths

Expressions in spreadsheet cells use the same paths as DOCX templates.

**Template → Result**

| Excel template | Generated document |
|---|---|
| `{Customer.Name}` | `Acme Corporation` |
| `Order: {Order.Number}` | `Order: SO-1042` |
| `{$.Company.Name}` | `Northwind Supplies` |
| `{{example}}` | `{example}` |

**JSON data**

```json
{
  "Company": { "Name": "Northwind Supplies" },
  "Customer": { "Name": "Acme Corporation" },
  "Order": { "Number": "SO-1042" }
}
```


A complete, unformatted expression such as `{Revenue}` can insert a native Excel number or Boolean. An expression combined with other text, such as `Revenue: {Revenue}`, produces text. Explicit MonoReport formatters also produce text. This distinction matters for Excel formulas and numeric formatting.

For expression syntax, root and parent paths, and built-in formatters, see [DOCX Template Guide — Fields and Data Paths](docx-template-guide.md#fields-and-data-paths) and [Formatters](docx-template-guide.md#formatters). The expression language is shared; the surrounding spreadsheet behavior is different.

---

## PrintArea and Control Cells

**Set a rectangular Excel Print Area before using repeated rows or columns.** The Print Area defines the report region to retain. MonoReport removes cells outside it and moves the resulting report to **A1**. It is a processing boundary, not merely a print setting.

Use the space **above** and **to the left** of the Print Area for repetition controls. Those helper cells are intentionally excluded from the generated workbook.

**Template → Result**

| Template location | Excel template | Generated document |
|---|---|---|
| A column, outside PrintArea | `$level(1).Data` | *Removed* |
| B2, inside PrintArea | **Department** | **Department** |
| B3, inside PrintArea | `{Department}` | `Engineering` |
| C3, inside PrintArea | `{Name}` | `Avery Chen` |

The example above illustrates **which cells survive**, not a complete vertical-expansion template; the complete repeated block appears in the next section.

**How to set it in Excel**

1. Select the report's rectangular cell range, **excluding** the helper marker column and top control rows.
2. Use **Page Layout → Print Area → Set Print Area**.
3. Keep all content you want in the output *inside* that range, including titles, totals, and formulas.
4. For a repeating sheet, make sure the entire intended report is covered by one valid, rectangular Print Area.

The Print Area must be one finite rectangular range; disjoint, duplicate, malformed, or ambiguous definitions are rejected. For example, if row markers are in column A and horizontal markers are above row 3, the Print Area can begin at `B3`. In the generated workbook, the retained region starts at `A1`.

You can also define a Print Area on a sheet without repetitions. In that case, MonoReport still crops and relocates that sheet. A sheet with **no Print Area and no repetitions** keeps its original coordinates.

---

## Vertical Expansion

Vertical expansion repeats spreadsheet rows to display records and hierarchical groups. Unlike DOCX, XLSX does **not** use `{For:...}` markers. Instead, you put row controls in a helper column to the **left** of the Print Area.

**Template → Result**

This compact example follows the employees sample shown in the Playground. The row numbers on the left are **source worksheet rows**, not part of the generated report.

| Template row | Control cell (outside PrintArea) | Department | Employee | Position |
|---|---|---|---|---|
| 2 | | **Department** | **Employee** | **Position** |
| 3 | `$level(1).Data` | `{Department}` | | |
| 4 | `$level(2).Data` | | `{Name}` | `{Position}` |

**Generated document**

| Department | Employee | Position |
|---|---|---|
| **Engineering** | | |
| | `Avery Chen` | `Software Engineer` |
| | `Jordan Patel` | `QA Engineer` |
| **Human Resources** | | |
| | `Casey Nguyen` | `Recruiter` |

The two levels describe the hierarchy: the department row is produced for a parent group, followed by its employees. **The numbers `1` and `2` are level identifiers, not repeat counts or Excel row numbers.** Data order determines group order.

**JSON data**

```json
{
  "Data": [
    { "$level": 1, "Department": "Engineering" },
    { "$level": 2, "Name": "Avery Chen", "Position": "Software Engineer" },
    { "$level": 2, "Name": "Jordan Patel", "Position": "QA Engineer" },
    { "$level": 1, "Department": "Human Resources" },
    { "$level": 2, "Name": "Casey Nguyen", "Position": "Recruiter" }
  ]
}
```


This short example shows the level-ordered record structure. The Playground employee sample uses the same layout with more departments and employees.

**Rules to remember**

1. The marker column must be outside the Print Area.
2. Rows at different hierarchy levels can have different layouts and styles.
3. Records must appear in their intended hierarchy/order in JSON; an already-closed group cannot reappear later.
4. Level identifiers select template layouts. They are not necessarily consecutive, and the data hierarchy is not created by template indentation.

---

## Horizontal Expansion

Horizontal expansion repeats columns for a collection, such as suppliers in a bid comparison. Use helper markers **above** the Print Area to define the column geometry; header markers to the left identify the header rows. Row-level markers select the item data to display down the sheet.

**Template → Result**

This example expands one supplier value column into two columns. It uses a single horizontal dimension (`Bids`) and two item records (`Items`).

| Template location | Excel template |
|---|---|
| Above the supplier column, outside PrintArea | `$h(1)` |
| Left of the supplier header row, outside PrintArea | `$header(1).Bids` |
| Supplier header cell | `{Company}` |
| Left of the item row, outside PrintArea | `$level(1).Items` |
| Item cell | `{Name}` |
| Supplier value cell | `{Value}` |

**Generated document**

| Item | Alpha | Beta |
|---|---:|---:|
| `Laptop` | `1200` | `1150` |
| `Dock` | `80` | `75` |

**JSON data**

```json
{
  "Bids": [
    { "Company": "Alpha" },
    { "Company": "Beta" }
  ],
  "Items": [
    { "$level": 1, "$path": "laptop", "Name": "Laptop", "Value": 1200 },
    { "$level": 1, "$path": "laptop", "Value": 1150 },
    { "$level": 1, "$path": "dock", "Name": "Dock", "Value": 80 },
    { "$level": 1, "$path": "dock", "Value": 75 }
  ]
}
```

The two `Bids` produce **two dynamic value columns**. For each item, `Items` therefore contains **two consecutive records with the same `$path`**, in supplier order: Alpha, then Beta. The first record supplies the ordinary item label (`Name`); each record supplies the `Value` for its corresponding column. `$path` identifies the logical item shared by these records; it is not used to look up a supplier by name.

The marker `$h(1)` identifies the horizontal level, and `$header(1).Bids` supplies the matching column headings. The original single value column is replicated across the two suppliers.

---

## Combining Both Directions

Vertical and horizontal expansion can be combined—for example, **item groups and items** down the sheet, with **suppliers and comparison criteria** across it. This example uses the actual JSON structure of the tender sample.

**Template → Result**

| Template component | What it controls | Generated document |
|---|---|---|
| `$level(1).Items` | Group heading | `IT equipment` |
| `$l(2).Items` | Item detail row | `Laptop / IT equipment` |
| `$l(1).Items` | Group footer | `Lot total: IT equipment` |
| `$header(1).Bids` | Supplier heading | `Alpha`, `Beta` |
| `$header(2).Criteria` | Criterion heading | `Unit price`, `Lead days` |
| `$h(1)` and `$h(2)` | Two horizontal levels | One column per supplier–criterion combination |

**Generated document**

| Item | Quantity | Alpha: Unit price | Alpha: Lead days | Beta: Unit price | Beta: Lead days |
|---|---:|---:|---:|---:|---:|
| **IT equipment** | | | | | |
| `Laptop / IT equipment` | `2` | `1,200.00` | `10.00` | `1,150.00` | `12.00` |
| `Dock / IT equipment` | `3` | `80.00` | `5.00` | `75.00` | `7.00` |
| **Lot total: IT equipment** | | | | | |

This is a **simplified Markdown view** of the tender workbook. The actual Excel template uses separate supplier and criterion header rows, merges, and formatting. A literal total like `0.00` is not calculated by MonoReport: use a formula or data-driven value if a calculated total is required.

**JSON data**

```json
{
  "Title": "Office equipment",
  "Bids": [
    { "Company": "Alpha" },
    { "Company": "Beta" }
  ],
  "Criteria": [
    { "Name": "Unit price" },
    { "Name": "Lead days" }
  ],
  "Items": [
    { "$level": 1, "$path": "lot", "Name": "IT equipment" },
    { "$level": 1, "$path": "lot" },
    { "$level": 1, "$path": "lot" },
    { "$level": 1, "$path": "lot" },
    { "$level": 2, "$path": "laptop", "Name": "Laptop", "Quantity": 2, "Value": 1200 },
    { "$level": 2, "$path": "laptop", "Value": 10 },
    { "$level": 2, "$path": "laptop", "Value": 1150 },
    { "$level": 2, "$path": "laptop", "Value": 12 },
    { "$level": 2, "$path": "dock", "Name": "Dock", "Quantity": 3, "Value": 80 },
    { "$level": 2, "$path": "dock", "Value": 5 },
    { "$level": 2, "$path": "dock", "Value": 75 },
    { "$level": 2, "$path": "dock", "Value": 7 }
  ]
}
```

The horizontal dimensions contain **2 suppliers × 2 criteria = 4 dynamically expanded value columns**. Consequently, each logical item has **four consecutive `Items` records with the same `$path`**. They correspond, in order, to `Alpha / Unit price`, `Alpha / Lead days`, `Beta / Unit price`, and `Beta / Lead days`.

For `"$path": "laptop"`, the values `1200`, `10`, `1150`, and `12` populate those four columns. The same pattern produces Dock's `80`, `5`, `75`, and `7`.

The lot also has four records with `$path: "lot"`, even though only the first contains the displayed group label. This keeps the logical group aligned with the horizontal layout.

**How the records map to columns**

| Record within one `$path` | Supplier | Criterion | Laptop `Value` |
|---|---|---|---:|
| 1 | Alpha | Unit price | `1200` |
| 2 | Alpha | Lead days | `10` |
| 3 | Beta | Unit price | `1150` |
| 4 | Beta | Lead days | `12` |

**Rules to remember**

1. `$level(1).Items` and `$l(2).Items` are the long and short vertical level markers; the level number chooses the layout, not an iteration count.
2. `$header(1).Bids` and `$header(2).Criteria` describe the horizontal headings. `$h(1)` and `$h(2)` define their geometry above the Print Area.
3. **For every logical `$path`, provide as many consecutive records as the number of dynamically expanded leaf columns.** With nested horizontal dimensions, that count is the product of their collection sizes (four in this example). Each record follows the same column order.
4. The same `$path` must occur in a contiguous run; once a logical path closes, it cannot reappear later. Records must also have the appropriate consistent integer `$level`.
5. Matching is **positional**, not by the text of supplier, criterion, or item labels. Incorrect ordering can silently place values under the wrong headings.

**Sample files**

1. [`Samples/xlsx-tender.template.xlsx`](../Samples/xlsx-tender.template.xlsx)
2. [`Samples/xlsx-tender.json`](../Samples/xlsx-tender.json)

---

## Cell Values, Formatting, and Formulas

**Native Excel values**

| Excel template | JSON value | Generated document |
|---|---|---|
| `{Quantity}` | `2` | Native number `2` |
| `{Approved}` | `true` | Native Boolean `TRUE` |
| `Quantity: {Quantity}` | `2` | Text `Quantity: 2` |
| `{Amount number "#,##0.00"}` | `1234.5` | Text `1,234.50` (culture dependent) |

Native numbers retain the cell's Excel number format. An ISO date string is inserted as text; MonoReport does not automatically convert it to an Excel date serial. For a true numeric Excel date, supply an appropriate numeric value and apply the matching Excel format.

Use Excel cell formats for numeric display where possible. MonoReport formatters, shared with DOCX, are described in [DOCX Template Guide — Formatters](docx-template-guide.md#formatters). Unlike native Excel formatting, formatter output is text.

**Formulas**

MonoReport preserves and adjusts supported ordinary Excel formulas when cells are moved or repeated. You can author formulas using ordinary A1 cell and rectangular range references.

| Excel template | Generated document |
|---|---|
| `=B3*C3` | Formula references are adjusted when the row or column moves |
| `=SUM(D3:D8)` | A supported contiguous range follows the expanded report geometry |

MonoReport supports relative (`A1`), absolute (`$A$1`), and mixed (`$A1`, `A$1`) references when their final targets remain valid. It **does not calculate formulas**. It clears stale cached results and requests Excel to recalculate when the workbook opens. Unsupported or ambiguous references cause errors rather than silently producing misleading calculations.

`INDIRECT`, `OFFSET`, whole-column references, structured references, external workbook references, and dynamic-array formulas are **unsupported and rejected**. See the [compatibility contract](xlsx-template-compatibility.md) for other unsupported formula forms.

---

## Validation and Troubleshooting

Analyze the XLSX template in [MonoReport Playground](https://demo.monoreport.net) before generation. The Analyzer reports expressions, data requirements, and structural diagnostics, but it does not verify every data-dependent expansion or formula outcome without JSON input.

**Compatibility restrictions**

Several familiar Excel features are currently rejected rather than preserved: Excel Tables, AutoFilter, charts, drawings/images, conditional formatting, data validation, hyperlinks, comments, and Print Titles. Worksheets with repetition also reject panes and manual page breaks. On nonrepeating sheets, panes and page breaks are supported only when cropping leaves them valid. Merged cells must remain wholly compatible with cropping and repeated regions; partial or conflicting merges are rejected. See [XLSX Template Compatibility](xlsx-template-compatibility.md) for the precise conditions.

**Common problems**

| Problem | What to check |
|---|---|
| Repeating sheet is rejected | Confirm a single valid rectangular `PrintArea` and remove panes or manual page breaks |
| Controls appear in the result | Check that helper cells are outside `PrintArea` |
| Rows repeat incorrectly | Check `$level` values, contiguous groups, and input order |
| Supplier columns are missing | Check `$header(...)` controls, top `$h(...)` markers, and collection data |
| Matrix values are wrong | Verify Cartesian record order and `$path` identifiers; positions are not matched by labels |
| A number becomes text | Use a complete, unformatted scalar expression and Excel cell formatting |
| Formula cannot be relocated | Replace unsupported or ambiguous references with supported A1 cell/range formulas |
| Visual preview differs from Excel | Download and open the workbook in a compatible spreadsheet editor |

**Practical recommendations**

1. Start with scalar fields before adding repetition controls.
2. Define `PrintArea` deliberately and keep report content inside it.
3. Build vertical and horizontal expansion separately before combining them.
4. Check the output in Excel, especially formulas, merges, print layout, and page breaks.
5. Consult [XLSX Template Compatibility](xlsx-template-compatibility.md) when using advanced spreadsheet features.

XLSX templates do **not** support DOCX-style `If`/`Else` conditions. The Playground preview is approximate; actual Excel recalculation, layout, and printing must be checked in Excel.

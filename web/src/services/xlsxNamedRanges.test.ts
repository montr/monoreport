import type { DefinedName, Worksheet } from "@silurus/ooxml/xlsx";
import { describe, expect, it, vi } from "vitest";
import { XlsxNamedRanges } from "./xlsxNamedRanges";

function setup(definedNames: DefinedName[], sheetNames = ["First", "Hidden", "O'Brien, Q1!"], hidden = [1]) {
  const workbook = {
    sheetNames,
    isHidden: (index: number) => hidden.includes(index),
    getWorksheet: vi.fn(async () => ({ definedNames } as Worksheet))
  };
  return { workbook, names: new XlsxNamedRanges(workbook) };
}

describe("XLSX named range discovery", () => {
  it("resolves finite ranges and cells, quoted names and the active worksheet's Print_Area", async () => {
    const { names } = setup([
      { name: "Report", formula: "First!$B$2:$D$20" },
      { name: "Destination", formula: "'O''Brien, Q1!'!$a$10" },
      { name: "_xlnm.Print_Area", formula: "first!A1:XFD1048576" }
    ]);
    expect(await names.getForSheet(0)).toEqual([
      { id: "report", name: "Report", sheetIndex: 0, firstCell: "B2", range: "B2:D20" },
      { id: "destination", name: "Destination", sheetIndex: 2, firstCell: "A10", range: "A10" },
      { id: "_xlnm.print_area", name: "Print_Area", sheetIndex: 0, firstCell: "A1", range: "A1:XFD1048576" }
    ]);
  });

  it("uses the requested worksheet's inventory and omits hidden worksheets and foreign Print Areas", async () => {
    const { workbook, names } = setup([
      { name: "HiddenTarget", formula: "Hidden!$A$1" },
      { name: "_xlnm.Print_Area", formula: "First!$B$2:$C$3" }
    ]);
    expect(await names.getForSheet(2)).toEqual([]);
    expect(workbook.getWorksheet).toHaveBeenCalledExactlyOnceWith(2);
    expect(await names.getForSheet(1)).toEqual([]);
    expect(workbook.getWorksheet).toHaveBeenCalledOnce();
  });

  it.each([
    "OFFSET(First!$A$1,0,0,5,5)", "OtherName", "42", '"text"',
    "First!$A$1,First!$C$3", "First!$A:$C", "First!$1:$5", "$A$1:$B$2",
    "First!A1:B2", "First!$A1:$B$2", "First!$A$1:B$2", "First!$A$1:$B2",
    "'[external.xlsx]First'!$A$1", "'First:Hidden'!$A$1", "First:Hidden!$A$1",
    "Missing!$A$1", "First!$A$0", "First!$XFE$1", "First!$A$1048577",
    "First!$B$2:$A$3", "First!$A$3:$B$2", "First!$A$1#"
  ])("omits unsupported or unsafe ordinary definitions: %s", async formula => {
    const { names } = setup([{ name: "Unsupported", formula }]);
    expect(await names.getForSheet(0)).toEqual([]);
  });

  it("omits ambiguous names, even when only one conflicting definition is supported", async () => {
    const { names } = setup([
      { name: "Report", formula: "First!$A$1" },
      { name: "REPORT", formula: "First!$B$2" },
      { name: "Mixed", formula: "First!$A$1" },
      { name: "mixed", formula: "OFFSET(First!$A$1,0,0,5,5)" },
      { name: "_xlnm.Print_Area", formula: "First!$A$1:$B$2" },
      { name: "_xlnm.Print_Area", formula: "First!$C$1:$D$2" }
    ]);
    expect(await names.getForSheet(0)).toEqual([]);
  });

  it.each(["OFFSET(First!$A$1,0,0,2,2)", "First!$A$1,First!$C$3", "#REF!"])("does not invent a Print Area for %s", async formula => {
    const { names } = setup([{ name: "_xlnm.Print_Area", formula }]);
    expect(await names.getForSheet(0)).toEqual([]);
  });

  it("collapses equivalent references and ignores other built-in names", async () => {
    const { names } = setup([
      { name: "Report", formula: "First!$a$1:$b$2" },
      { name: "REPORT", formula: "'FIRST'!$A$1:$B$2" },
      { name: "_xlnm.Print_Titles", formula: "First!$A$1" },
      { name: "_xlnm._FilterDatabase", formula: "First!$A$1:$B$2" }
    ]);
    expect(await names.getForSheet(0)).toEqual([
      { id: "report", name: "Report", sheetIndex: 0, firstCell: "A1", range: "A1:B2" }
    ]);
  });

  it("handles worksheets without defined names", async () => {
    const { workbook, names } = setup([]);
    workbook.getWorksheet.mockResolvedValueOnce({} as Worksheet);
    expect(await names.getForSheet(0)).toEqual([]);
  });
});

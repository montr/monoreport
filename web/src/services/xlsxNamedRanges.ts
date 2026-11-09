import type { DefinedName, XlsxWorkbook } from "@silurus/ooxml/xlsx";

const printAreaName = "_xlnm.print_area";
const builtInNamePrefix = "_xlnm.";
const maxWorksheetRow = 1_048_576;
const maxWorksheetColumn = 16_384;

export interface XlsxNamedRange {
  id: string;
  name: string;
  sheetIndex: number;
  firstCell: string;
  range: string;
}

export class XlsxNamedRanges {
  constructor(private readonly workbook: Pick<XlsxWorkbook, "getWorksheet" | "sheetNames" | "isHidden">) {}

  async getForSheet(sheetIndex: number): Promise<XlsxNamedRange[]> {
    if (this.workbook.isHidden(sheetIndex)) return [];
    const worksheet = await this.workbook.getWorksheet(sheetIndex);
    const groups = new Map<string, DefinedName[]>();
    for (const definition of worksheet.definedNames ?? []) {
      const id = definition.name.toLowerCase();
      if (!id || (id.startsWith(builtInNamePrefix) && id !== printAreaName)) continue;
      const group = groups.get(id);
      if (group) group.push(definition);
      else groups.set(id, [definition]);
    }

    const ranges: XlsxNamedRange[] = [];
    for (const [id, definitions] of groups) {
      const targets = definitions.map(definition => this.resolve(definition.formula, id === printAreaName));
      const target = targets[0];
      // The published model loses scope. Conflicting duplicates cannot safely
      // distinguish local shadowing from an ambiguous or malformed definition.
      if (!target || targets.some(item => !item || item.sheetIndex !== target.sheetIndex || item.range !== target.range)) continue;
      if (id === printAreaName && target.sheetIndex !== sheetIndex) continue;
      ranges.push({ id, name: id === printAreaName ? "Print_Area" : definitions[0].name, ...target });
    }
    return ranges;
  }

  private resolve(formula: string, printArea: boolean): Omit<XlsxNamedRange, "id" | "name"> | undefined {
    const match = /^(?:'((?:[^']|'')+)'|([^'!,:\[\]]+))!(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]*)(?::(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]*))?$/.exec(formula);
    if (!match) return;
    if (!printArea && (!match[3] || !match[5] || (match[8] && (!match[7] || !match[9])))) return;
    const sheetName = match[1] !== undefined ? match[1].replace(/''/g, "'") : match[2];
    if (/[\[\]:]/.test(sheetName)) return;
    const sheetIndex = this.workbook.sheetNames.findIndex(name => name.toLowerCase() === sheetName.toLowerCase());
    if (sheetIndex < 0 || this.workbook.isHidden(sheetIndex)) return;

    const firstColumn = match[4].toUpperCase();
    const lastColumn = (match[8] ?? match[4]).toUpperCase();
    const firstRow = Number(match[6]);
    const lastRow = Number(match[10] ?? match[6]);
    const columnIndex = (letters: string) => [...letters].reduce((column, letter) => column * 26 + letter.charCodeAt(0) - 64, 0);
    const left = columnIndex(firstColumn);
    const right = columnIndex(lastColumn);
    if (firstRow > lastRow || lastRow > maxWorksheetRow || left > right || right > maxWorksheetColumn) return;
    const firstCell = `${firstColumn}${firstRow}`;
    return { sheetIndex, firstCell, range: match[8] ? `${firstCell}:${lastColumn}${lastRow}` : firstCell };
  }
}

"""Fixture-specific executable research. No general template syntax or production code."""
import copy
import json
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as E

BASE = Path(__file__).resolve().parent
OUT = BASE / "evidence"
OUT.mkdir(exist_ok=True)
S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P = "http://schemas.openxmlformats.org/package/2006/relationships"
E.register_namespace("", S)
E.register_namespace("r", R)

def el(tag, attrs=None, text=None):
    n = E.Element("{" + S + "}" + tag, attrs or {})
    n.text = text
    return n

def col(c):
    result = ""
    while c:
        c, rem = divmod(c - 1, 26)
        result = chr(65 + rem) + result
    return result

def xy(address):
    i = 0
    c = 0
    while i < len(address) and address[i].isalpha():
        c = c * 26 + ord(address[i]) - 64
        i += 1
    return int(address[i:]), c

def cell(address, value, style=0):
    n = el("c", {"r": address, "s": str(style)})
    if isinstance(value, str) and value.startswith("="):
        n.append(el("f", text=value[1:]))
    elif isinstance(value, (int, float)):
        n.append(el("v", text=str(value)))
    else:
        n.set("t", "inlineStr")
        n.append(el("is"))
        n[0].append(el("t", text=str(value)))
    return n

def sheet(rows, merges=(), advanced=False):
    root = el("worksheet")
    root.append(el("sheetViews"))
    root[-1].append(el("sheetView", {"workbookViewId": "0"}))
    root[-1][0].append(el("pane", {"xSplit": "1", "ySplit": "1", "topLeftCell": "B2", "activePane": "bottomRight", "state": "frozen"}))
    root.append(el("sheetFormatPr", {"defaultRowHeight": "15"}))
    root.append(el("cols"))
    for c in range(1, 7):
        root[-1].append(el("col", {"min": str(c), "max": str(c), "width": str(8 + c), "customWidth": "1", **({"hidden": "1"} if c == 1 else {})}))
    data = el("sheetData")
    for r, cells in rows.items():
        row = el("row", {"r": str(r), "ht": str(18 + r % 3), "customHeight": "1", **({"hidden": "1"} if r == 1 else {})})
        for c, value in sorted(cells.items()):
            row.append(cell(col(c) + str(r), value, 1 if c >= 4 else 0))
        data.append(row)
    root.append(data)
    if merges:
        root.append(el("mergeCells", {"count": str(len(merges))}))
        for merge in merges:
            root[-1].append(el("mergeCell", {"ref": merge}))
    if advanced:
        cf = el("conditionalFormatting", {"sqref": "B15:B16"})
        rule = el("cfRule", {"type": "cellIs", "priority": "1", "operator": "greaterThan"})
        rule.append(el("formula", text="0"))
        cf.append(rule)
        root.append(cf)
        dv = el("dataValidations", {"count": "1"})
        validation = el("dataValidation", {"type": "whole", "operator": "between", "sqref": "B15:B16"})
        validation.extend([el("formula1", text="0"), el("formula2", text="100")])
        dv.append(validation)
        root.append(dv)
    root.append(el("printOptions", {"horizontalCentered": "1"}))
    root.append(el("pageMargins", {k: "0.3" for k in ["left", "right", "top", "bottom", "header", "footer"]}))
    root.append(el("pageSetup", {"paperSize": "9", "orientation": "landscape", "fitToWidth": "1", "fitToHeight": "0"}))
    if advanced:
        root.append(el("tableParts", {"count": "1"}))
        root[-1].append(el("tablePart", {"{" + R + "}id": "table1"}))
    return root

STYLES = f'''<styleSheet xmlns="{S}"><numFmts count="1"><numFmt numFmtId="164" formatCode="0.00"/></numFmts><fonts count="1"><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left style="thin"/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="right" wrapText="1"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'''

def write_package(path, sheets, names=(), extra=None):
    wb = el("workbook")
    wb.append(el("bookViews"))
    wb[-1].append(el("workbookView"))
    wb.append(el("sheets"))
    for i, name in enumerate(sheets, 1):
        wb[-1].append(el("sheet", {"name": name, "sheetId": str(i), "{" + R + "}id": "s" + str(i)}))
    if names:
        wb.append(el("definedNames"))
        for name, text, scope in names:
            wb[-1].append(el("definedName", {"name": name, **({"localSheetId": str(scope)} if scope is not None else {})}, text))
    wb.append(el("calcPr", {"calcId": "0", "calcMode": "auto", "fullCalcOnLoad": "1", "forceFullCalc": "1"}))
    parts = {"xl/workbook.xml": E.tostring(wb), "xl/styles.xml": STYLES.encode()}
    parts["_rels/.rels"] = f'<Relationships xmlns="{P}"><Relationship Id="root" Type="{R}/officeDocument" Target="xl/workbook.xml"/></Relationships>'.encode()
    relations = f'<Relationships xmlns="{P}"><Relationship Id="styles" Type="{R}/styles" Target="styles.xml"/>'
    content = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
    content += '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    for i, root in enumerate(sheets.values(), 1):
        parts[f"xl/worksheets/sheet{i}.xml"] = E.tostring(root)
        relations += f'<Relationship Id="s{i}" Type="{R}/worksheet" Target="worksheets/sheet{i}.xml"/>'
        content += f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    if extra:
        parts.update(extra)
        content += '<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>'
    parts["xl/_rels/workbook.xml.rels"] = (relations + '</Relationships>').encode()
    parts["[Content_Types].xml"] = (content + '</Types>').encode()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)

def prepare():
    records = []
    for level, path, name, amount in [(10, "001", "Lot 1", 100), (20, "001/001", "A", 10), (20, "001/002", "B", 20), (10, "002", "Lot 2", 200), (20, "002/001", "C", 30)]:
        for i in range(4):
            records.append({"$level": level, "$path": path, "Name": name if i == 0 else "not the row name", "Price": amount + i, "Amount": (amount + i) * 2, "Delta": 1})
    data = {"Title": "Research", "Bids": [{"Company": "Alpha", "Address": "City A"}, {"Company": "Beta", "Address": "City B"}], "Criteria": [{"Name": "Price"}, {"Name": "Term"}], "Items": records, "Other": [{"Name": "X"}, {"Name": "Y"}, {"Name": "Z"}]}
    (OUT / "input.json").write_text(json.dumps(data, indent=2))
    rows = {
        1: {4: "h", 5: "h"},
        2: {2: "{Title}", 3: 2, 4: "=SUBTOTAL(9,D7:D10)", 5: "=SUBTOTAL(2,D7:D10)", 6: "=SUBTOTAL(9,D7:E10)"},
        3: {1: "$header(1).Bids", 4: "{Company}"}, 4: {1: "$h(1).Bids", 4: "{Address}"},
        5: {1: "$header(2).Criteria", 4: "{Name}"}, 6: {1: "$h(2).Criteria", 4: "Criterion"},
        7: {1: "$level(10).Items", 2: "{Name}", 4: "=SUBTOTAL(9,D8:D9)", 5: "=SUM(D8:D9)", 6: "=COUNT($D8:$D9)"},
        8: {1: "$l(20).Items", 2: "{Name}", 4: "{Price}", 5: "{Amount}"},
        9: {1: "$level(20).Items", 2: "", 4: "{Delta}", 5: "=D8+$C$2"},
        10: {1: "$level(10).Items", 2: "footer", 4: "=SUBTOTAL(9,D8:D9)", 5: "=COUNT(D8:D9)", 6: "=SUM($D8:$D9)"},
        11: {2: "grand total", 4: "=SUBTOTAL(9,D7:D10)"}, 12: {2: "static footer"},
        15: {1: "$level(5).Other", 2: "{Name}", 3: "second table"}, 16: {2: "after second table"}}
    source = sheet(rows, [f"D{r}:E{r}" for r in range(3, 7)] + ["B8:B9"])
    other = sheet({2: {2: "='Report'!D11", 3: "=2+2"}})
    names = [("_xlnm.Print_Area", "Report!$B$2:$F$16", 0), ("MR_Main", "Report!$B$2:$F$12", None), ("MR_Second", "Report!$B$15:$C$16", None)]
    write_package(OUT / "template.xlsx", {"Report": source, "Other Sheet": other}, names)

    groups = []
    seen = set()
    for record in records:
        path = record["$path"]
        if not groups or groups[-1][0]["$path"] != path:
            assert path not in seen, "noncontiguous path"
            seen.add(path)
            groups.append([])
        groups[-1].append(record)
    assert all(len(g) == 4 and len({r["$level"] for r in g}) == 1 for g in groups)
    assert [g[0]["$path"] for g in groups] == ["001", "001/001", "001/002", "002", "002/001"]
    # Explicit fixture row plan; parser/inference and generic hierarchy are intentionally excluded.
    plan = [(r, None, None) for r in range(1, 7)]
    parents = []
    for parent, children in [(groups[0], groups[1:3]), (groups[3], groups[4:5])]:
        start = len(plan) + 1
        plan.append((7, parent, parent[0]["$path"]))
        for child in children:
            plan.extend([(8, child, parent[0]["$path"]), (9, child, parent[0]["$path"])])
        plan.append((10, parent, parent[0]["$path"]))
        parents.append((start, len(plan)))
    plan.extend((r, None, None) for r in range(11, 15))
    plan.extend((15, [record], None) for record in data["Other"])
    plan.append((16, None, None))
    rowmap = {}
    for dest, (src, _, _) in enumerate(plan, 1):
        rowmap.setdefault(src, []).append(dest)
    final = copy.deepcopy(source)
    finaldata = final.find(f"{{{S}}}sheetData")
    finaldata.clear()
    requests, targets = [], []
    leaves = [(b, c) for b in data["Bids"] for c in data["Criteria"]]
    originalrows = {int(r.get("r")): r for r in source.find(f"{{{S}}}sheetData")}
    for dest, (src, group, parent) in enumerate(plan, 1):
        if src not in originalrows:
            continue
        row = copy.deepcopy(originalrows[src])
        row.set("r", str(dest))
        row.clear()  # retain the attributes explicitly below
        row.attrib.update(originalrows[src].attrib)
        row.set("r", str(dest))
        for original in originalrows[src]:
            _, c = xy(original.get("r"))
            for leaf in (range(4) if c in (4, 5) and src <= 12 else [None]):
                targetcol = c + 2 * leaf if leaf is not None else c + (6 if c >= 6 and src <= 12 else 0)
                node = copy.deepcopy(original)
                node.set("r", col(targetcol) + str(dest))
                textnode = node.find(f"{{{S}}}is/{{{S}}}t")
                if c == 1 or src == 1:
                    if textnode is not None:
                        textnode.text = ""  # provisional control-output policy: clear text, retain hidden bands
                elif textnode is not None and textnode.text.startswith("{"):
                    field = textnode.text[1:-1]
                    if src in (3, 4):
                        value = leaves[leaf][0][field]
                    elif src == 5:
                        value = leaves[leaf][1][field]
                    elif group:
                        value = group[leaf if leaf is not None else 0][field]
                    else:
                        value = data[field]
                    node = cell(node.get("r"), value, int(node.get("s")))
                formula = node.find(f"{{{S}}}f")
                if formula is not None:
                    rmap = dict(rowmap)
                    if parent:
                        rmap = {r: [pos for pos in positions if plan[pos - 1][2] == parent] if r in (7, 8, 9, 10) else positions for r, positions in rowmap.items()}
                    if src == 9:
                        rmap[8], rmap[9] = [dest - 1], [dest]
                    cmap = {4: [4 + 2 * (leaf or 0)], 5: [5 + 2 * (leaf or 0)], 6: [12]}
                    if src == 2 and c == 6:
                        cmap[4], cmap[5] = [4, 6, 8, 10], [5, 7, 9, 11]
                    requests.append({"Formula": formula.text, "Rows": rmap, "Cols": cmap})
                    targets.append(node.get("r"))
                row.append(node)
        row[:] = sorted(row, key=lambda n: xy(n.get("r"))[1])
        finaldata.append(row)
    merges = final.find(f"{{{S}}}mergeCells")
    merges.clear()
    for r in range(3, 7):
        for i in range(4):
            merges.append(el("mergeCell", {"ref": f"{col(4 + i * 2)}{r}:{col(5 + i * 2)}{r}"}))
    for r in rowmap[8]:
        merges.append(el("mergeCell", {"ref": f"B{r}:B{r+1}"}))
    merges.set("count", str(len(merges)))
    columns = final.find(f"{{{S}}}cols")
    columns.clear()
    for c in range(1, 13):
        src = (4 + (c - 4) % 2) if 4 <= c <= 11 else (6 if c == 12 else c)
        columns.append(el("col", {"min": str(c), "max": str(c), "width": str(8 + src), "customWidth": "1", **({"hidden": "1"} if c == 1 else {})}))
    other.find(f".//{{{S}}}f").text = "Report!D" + str(rowmap[11][0])
    finalnames = list(names)
    finalnames[0] = ("_xlnm.Print_Area", "Report!$B$2:$L$24", 0)
    finalnames[1] = ("MR_Main", "Report!$B$2:$L$18", None)
    finalnames[2] = ("MR_Second", "Report!$B$21:$C$24", None)
    write_package(OUT / "pending.xlsx", {"Report": final, "Other Sheet": other}, finalnames)
    (OUT / "formula-requests.json").write_text(json.dumps(requests))
    (OUT / "targets.json").write_text(json.dumps(targets))

    extra = {"xl/worksheets/_rels/sheet1.xml.rels": f'<Relationships xmlns="{P}"><Relationship Id="table1" Type="{R}/table" Target="../tables/table1.xml"/></Relationships>'.encode(), "xl/tables/table1.xml": f'<table xmlns="{S}" id="1" name="Table1" displayName="Table1" ref="B14:C16" totalsRowShown="0"><autoFilter ref="B14:C16"/><tableColumns count="2"><tableColumn id="1" name="Amount"/><tableColumn id="2" name="Label"/></tableColumns></table>'.encode()}
    compat = sheet({2: {2: "{Title}"}, 14: {2: "Amount", 3: "Label"}, 15: {2: 10, 3: "a"}, 16: {2: 20, 3: "b"}}, advanced=True)
    write_package(OUT / "compatibility.xlsx", {"Report": compat}, [("FixedValues", "Report!$B$15:$B$16", None)], extra)
    with zipfile.ZipFile(OUT / "compatibility.xlsx") as z, zipfile.ZipFile(OUT / "compatibility-output.xlsx", "w", zipfile.ZIP_DEFLATED) as result:
        for name in z.namelist():
            payload = z.read(name)
            result.writestr(name, payload.replace(b"{Title}", b"Research") if name == "xl/worksheets/sheet1.xml" else payload)
    # Pure boundary approach: unchanged formula now misses inserted data. Hybrid repairs reference.
    direct = sheet({2: {2: 10}, 3: {2: 20}, 7: {2: "=SUM(B2:B3)", 3: "=COUNT(B2:B3)", 4: "=SUBTOTAL(9,B2:B3)", 5: "=ROWS(B2:B3)", 6: "=COUNTBLANK(B2:B3)"}})
    helpers = sheet({2: {}, 3: {2: 10}, 4: {2: 20}, 5: {}, 7: {2: "=SUM(B2:B3)", 3: "=SUM(B2:B5)", 4: "=SUBTOTAL(9,B2:B5)", 5: "=ROWS(B2:B5)", 6: "=COUNTBLANK(B2:B5)"}})
    for row in helpers.find(f"{{{S}}}sheetData"):
        if row.get("r") in ("2", "5"):
            row.set("hidden", "1")
    write_package(OUT / "boundaries.xlsx", {"Direct": direct, "Boundaries": helpers})
    # Negative/data probes: these are fixture assertions, not implemented production diagnostics.
    bad = records[:4] + records[8:12] + records[:4]
    seen, previous, rejected = set(), None, False
    for record in bad:
        path = record["$path"]
        if path != previous and path in seen:
            rejected = True
        seen.add(path)
        previous = path
    assert rejected and len(records[:3]) != len(leaves)
    assert 2 != 2 * len(data["Criteria"]), "fixed parent merge cannot cover dynamic children without resizing"
    print("Prepared fixtures: 5 logical groups, 4 leaves, 2-column tile, 2 worksheets, 2 independent tables.")

def finish():
    formulas = json.loads((OUT / "formulas.json").read_text())
    targets = json.loads((OUT / "targets.json").read_text())
    with zipfile.ZipFile(OUT / "pending.xlsx") as z, zipfile.ZipFile(OUT / "output.xlsx", "w", zipfile.ZIP_DEFLATED) as result:
        for name in z.namelist():
            payload = z.read(name)
            if name == "xl/worksheets/sheet1.xml":
                root = E.fromstring(payload)
                cells = {n.get("r"): n for n in root.iter("{" + S + "}c")}
                assert len(targets) == len(formulas)
                for target, formula in zip(targets, formulas):
                    cells[target].find("{" + S + "}f").text = formula
                assert cells["D7"].find("{" + S + "}f").text == "SUBTOTAL(9,D8:D11)"
                assert cells["D13"].find("{" + S + "}f").text == "SUBTOTAL(9,D14:D15)"
                assert cells["J2"].find("{" + S + "}f").text == "SUBTOTAL(9,J7:J16)"
                assert cells["L2"].find("{" + S + "}f").text == "SUBTOTAL(9,D7:K16)"
                assert cells["K15"].find("{" + S + "}f").text == "J14+$C$2"
                assert cells["B8"].find(f"{{{S}}}is/{{{S}}}t").text == "A"
                assert cells["J8"].find("{" + S + "}v").text == "13"
                assert cells["B21"].find(f"{{{S}}}is/{{{S}}}t").text == "X"
                assert cells["B23"].find(f"{{{S}}}is/{{{S}}}t").text == "Z"
                assert cells["B24"].find(f"{{{S}}}is/{{{S}}}t").text == "after second table"
                assert len(root.find(f"{{{S}}}mergeCells")) == 19
                for r in [3, 4, 5, 6]:
                    for c in [4, 6, 8, 10]:
                        assert f"{col(c)}{r}:{col(c+1)}{r}" in [n.get("ref") for n in root.find(f"{{{S}}}mergeCells")]
                payload = E.tostring(root)
            result.writestr(name, payload)
    with zipfile.ZipFile(OUT / "template.xlsx") as a, zipfile.ZipFile(OUT / "output.xlsx") as b:
        assert a.read("xl/styles.xml") == b.read("xl/styles.xml")
        ar = E.fromstring(a.read("xl/worksheets/sheet1.xml"))
        br = E.fromstring(b.read("xl/worksheets/sheet1.xml"))
        for tag in ["printOptions", "pageMargins", "pageSetup", "sheetViews"]:
            assert E.tostring(ar.find(f"{{{S}}}{tag}")) == E.tostring(br.find(f"{{{S}}}{tag}"))
        assert [r.get("ht") for r in br.find(f"{{{S}}}sheetData") if r.get("r") in ["8", "10", "14"]] == ["20", "20", "20"]
    with zipfile.ZipFile(OUT / "compatibility.xlsx") as a, zipfile.ZipFile(OUT / "compatibility-output.xlsx") as b:
        assert a.namelist() == b.namelist()
        for name in a.namelist():
            if name != "xl/worksheets/sheet1.xml":
                assert a.read(name) == b.read(name)
        source, final = [E.fromstring(z.read("xl/worksheets/sheet1.xml")) for z in (a, b)]
        for tag in ["conditionalFormatting", "dataValidations", "tableParts", "cols", "pageSetup", "sheetViews"]:
            assert E.tostring(source.find(f"{{{S}}}{tag}")) == E.tostring(final.find(f"{{{S}}}{tag}"))
    print(f"Geometry/formula/fidelity assertions passed; {len(formulas)} formulas; 19 unchanged-size merge copies.")

def recalc():
    def values(filename, sheetid):
        with zipfile.ZipFile(BASE / "recalculated" / filename) as z:
            root = E.fromstring(z.read(f"xl/worksheets/sheet{sheetid}.xml"))
            return {c.get("r"): c.find(f"{{{S}}}v").text for c in root.iter(f"{{{S}}}c") if c.find(f"{{{S}}}v") is not None}
    vals = values("output.xlsx", 1)
    expected = {"D2": "63", "E2": "6", "L2": "1362", "D7": "32", "E7": "32", "L7": "4", "D12": "32", "E12": "4", "L12": "32", "D13": "31", "D16": "31", "D17": "63", "J2": "72", "K15": "35"}
    for address, value in expected.items():
        assert vals[address] == value, (address, vals[address], value)
    assert values("output.xlsx", 2)["B2"] == "63"
    direct, helper = values("boundaries.xlsx", 1), values("boundaries.xlsx", 2)
    assert [direct[col(c) + "7"] for c in range(2, 7)] == ["30", "2", "30", "2", "0"]
    assert [helper[col(c) + "7"] for c in range(2, 7)] == ["10", "30", "30", "4", "2"]
    (OUT / "recalculation-results.json").write_text(json.dumps({"output": expected, "crossSheetB2": "63", "direct": direct, "boundaries": helper}, indent=2))
    print("Recalculation assertions passed: group top/bottom SUM/COUNT/SUBTOTAL, grands, cross-sheet, two axes, helper counterexamples.")

if __name__ == "__main__":
    {"prepare": prepare, "finish": finish, "recalc": recalc}[sys.argv[1]]()

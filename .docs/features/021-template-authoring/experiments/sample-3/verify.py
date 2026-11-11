"""Focused source/result, business data, and package preservation assertions."""
import hashlib
import json
import posixpath
import zipfile
from decimal import Decimal, ROUND_HALF_UP
from urllib.parse import unquote
from xml.etree import ElementTree as E

from author import BASE, stamp

N = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{" + N["w"] + "}"


def load(path):
    with zipfile.ZipFile(path) as package:
        assert package.testzip() is None
        parts = {name: package.read(name) for name in package.namelist()}
    for name, value in parts.items():
        if name.endswith((".xml", ".rels")):
            E.fromstring(value)
        if name.endswith(".rels"):
            folder = posixpath.dirname(posixpath.dirname(name))
            for rel in E.fromstring(value):
                if rel.get("TargetMode") != "External":
                    target = unquote(rel.get("Target").split("#")[0])
                    resolved = posixpath.normpath(posixpath.join(folder, target)).lstrip("/")
                    assert resolved in parts, (name, target)
    return parts, E.fromstring(parts["word/document.xml"])


def text(node):
    return "".join(t.text or "" for t in node.iter(W + "t"))


def structure(node):
    return (node.tag, tuple(sorted(node.attrib.items())), node.text or "", tuple(structure(c) for c in node))


def formatting(row):
    tags = {W + tag for tag in ["trPr", "tcPr", "pPr", "rPr"]}
    return [structure(node) for node in row.iter() if node.tag in tags]


def number(value):
    return f"{Decimal(str(value)):,.2f}"


def expected(data):
    rows = [["Description", "Unit", "Quantity", "Unit Price, USD", "Total, USD"]]
    formats = [0]
    for project in data["Projects"]:
        subtotal = Decimal(0)
        rows.append([project["Name"], number(project["Total"])])
        formats.append(1)
        for group in project["Groups"]:
            total = Decimal(0)
            rows.append(["    " + group["Name"], number(group["Total"])])
            formats.append(2)
            for item in group["Items"]:
                calculated = (Decimal(str(item["Quantity"])) * Decimal(str(item["UnitPrice"]))).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                assert calculated == Decimal(str(item["Total"]))
                total += calculated
                rows.append(["        " + item["Description"], item["Unit"], *[number(item[key]) for key in ["Quantity", "UnitPrice", "Total"]]])
                formats.append(3)
            assert total == Decimal(str(group["Total"]))
            subtotal += total
        assert subtotal == Decimal(str(project["Total"]))
    assert sum(Decimal(str(p["Total"])) for p in data["Projects"]) == Decimal(str(data["GrandTotal"]))
    rows.append(["GRAND TOTAL", number(data["GrandTotal"])])
    formats.append(28)
    return rows, formats


def main():
    source_parts, source = load(BASE / "input/skill-sample-3.docx")
    source_table = source.find(".//w:tbl", N)
    source_rows = source_table.findall("w:tr", N)
    evidence = {"cases": {}, "preservation": {}, "hashes": {}}
    paths = {"template": BASE / "output/skill-sample-3.template.docx"}
    paths.update({name: BASE / ("output/skill-sample-3.result.docx" if name == "modified" else f"evidence/{name}.result.docx")
                  for name in ["original", "modified", "empty-items", "empty-groups", "empty-projects"]})
    for name, path in paths.items():
        parts, document = load(path)
        assert parts.keys() == source_parts.keys()
        unchanged = [part for part in parts if part != "word/document.xml" and parts[part] == source_parts[part]]
        assert len(unchanged) == len(parts) - 1
        table = document.find(".//w:tbl", N)
        for tag in ["tblPr", "tblGrid"]:
            assert structure(table.find("w:" + tag, N)) == structure(source_table.find("w:" + tag, N))
        assert structure(document.find(".//w:sectPr", N)) == structure(source.find(".//w:sectPr", N))
        evidence["preservation"][name] = {"byte_identical_non_main_parts": len(unchanged), "package_parts": len(parts), "relationships_resolve": True, "table_grid_properties_and_section": "identical"}
        evidence["hashes"][str(path.relative_to(BASE))] = hashlib.sha256(path.read_bytes()).hexdigest()
        if name == "template":
            assert [text(p) for p in document.findall(".//w:tr/w:tc/w:p", N) if text(p).startswith("{For:")] == ["{For:Projects}", "{For:Groups}", "{For:Items}"]
            continue
        data_path = BASE / ("output/skill-sample-3.json" if name == "original" else f"evidence/{name}-data.json")
        data = json.loads(data_path.read_text())
        wanted, formats = expected(data)
        rows = table.findall("w:tr", N)
        actual = [[text(cell) for cell in row.findall("w:tc", N)] for row in rows]
        assert actual == wanted, name
        for row, prototype in zip(rows, formats):
            assert formatting(row) == formatting(source_rows[prototype]), (name, prototype)
        body_paragraphs = document.find("w:body", N).findall("w:p", N)
        assert text(body_paragraphs[0]) == f'Project “{data["ProjectName"]}” Cost Estimate'
        assert text(body_paragraphs[1]) == "Period: " + data["Period"]
        for before, after in zip(source.find("w:body", N).findall("w:p", N), body_paragraphs):
            assert formatting(before) == formatting(after)
        assert "{" not in text(document) and "}" not in text(document)
        if name == "original":
            assert [text(p) for p in source.iter(W + "p")] == [text(p) for p in document.iter(W + "p")]
            for before, after in zip(source_rows, rows):
                assert formatting(before) == formatting(after)
        evidence["cases"][name] = {"rows": len(rows), "projects": len(data["Projects"]), "groups": sum(len(p["Groups"]) for p in data["Projects"]), "items": sum(len(g["Items"]) for p in data["Projects"] for g in p["Groups"]), "grand_total": data["GrandTotal"], "exact_cell_text_and_order": True, "all_totals_reconciled": True, "source_row_formatting_matches": True}
    for path in [BASE / "input/skill-sample-3.docx", BASE / "output/skill-sample-3.json"]:
        evidence["hashes"][str(path.relative_to(BASE))] = hashlib.sha256(path.read_bytes()).hexdigest()
    (BASE / "evidence/verification.json").write_text(json.dumps(evidence, indent=2) + "\n")
    stamp("focused_verification_completed")
    print(json.dumps(evidence["cases"], indent=2))


if __name__ == "__main__":
    main()

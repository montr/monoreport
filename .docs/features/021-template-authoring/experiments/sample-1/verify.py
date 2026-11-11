"""Focused value, repetition, OPC and preservation checks for this experiment."""

import copy
import datetime
from decimal import Decimal
import hashlib
import json
from pathlib import Path
import posixpath
import zipfile
from xml.etree import ElementTree as ET

BASE = Path(__file__).resolve().parent
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{" + NS["w"] + "}"
REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"


def signature(element):
    return (element.tag, sorted(element.attrib.items()), element.text or "",
            [signature(child) for child in element])


def text(element):
    return "".join(t.text or "" for t in element.iter(W + "t"))


def formatting(element):
    element = copy.deepcopy(element)
    for t in element.iter(W + "t"):
        t.text = ""
        t.attrib.pop("{http://www.w3.org/XML/1998/namespace}space", None)
    return signature(element)


def inspect(path):
    with zipfile.ZipFile(path) as z:
        assert z.testzip() is None
        assert len(z.namelist()) == len(set(z.namelist()))
        parts = {name: z.read(name) for name in z.namelist() if not name.endswith("/")}
    parsed = {name: ET.fromstring(value) for name, value in parts.items()}
    relationship_count = 0
    for name, root in parsed.items():
        if not name.endswith(".rels"):
            continue
        origin = "" if name == "_rels/.rels" else name.rsplit("/_rels/", 1)[0]
        ids = []
        for rel in root.findall(REL + "Relationship"):
            ids.append(rel.attrib["Id"])
            relationship_count += 1
            assert rel.attrib.get("TargetMode") != "External"
            target = posixpath.normpath(posixpath.join(origin, rel.attrib["Target"]))
            assert target in parts, (name, target)
        assert len(ids) == len(set(ids))
    types = parsed["[Content_Types].xml"]
    defaults = {item.attrib["Extension"] for item in types.findall(CT + "Default")}
    overrides = {item.attrib["PartName"].lstrip("/") for item in types.findall(CT + "Override")}
    assert overrides <= set(parts)
    assert all(name in overrides or name.rsplit(".", 1)[-1] in defaults
               for name in parts if name != "[Content_Types].xml")
    body = parsed["word/document.xml"].find("w:body", NS)
    assert body is not None and body[-1].tag == W + "sectPr"
    for cell in body.findall(".//w:tc", NS):
        assert cell[-1].tag == W + "p"
    return parts, body, relationship_count


source, source_body, count = inspect(BASE / "input/skill-sample-1.docx")
template, template_body, _ = inspect(BASE / "output/skill-sample-1.template.docx")
assert source.keys() == template.keys()
assert [name for name in source if source[name] != template[name]] == ["word/document.xml"]
source_rows = source_body.find("w:tbl", NS).findall("w:tr", NS)
template_rows = template_body.find("w:tbl", NS).findall("w:tr", NS)
assert len(template_rows) == 4
assert text(template_rows[1]) == "{For:Materials}"
assert text(template_rows[3]) == "{/For}"
assert signature(source_rows[0]) == signature(template_rows[0])
for row in source_rows[1:] + template_rows[1:]:
    assert formatting(row) == formatting(source_rows[1])
for row in [template_rows[1], template_rows[3]]:
    cells = row.findall("w:tc", NS)
    assert all(text(cell) == "" for cell in cells[1:])
    assert all(len(cell.findall("w:p", NS)) == 1 for cell in cells)
for i, child in enumerate(source_body):
    if child.tag == W + "tbl":
        for tag in ["tblPr", "tblGrid"]:
            assert signature(child.find(W + tag)) == signature(template_body[i].find(W + tag))
    elif i < 10 or child.tag == W + "sectPr":
        assert signature(child) == signature(template_body[i])
    else:
        assert formatting(child) == formatting(template_body[i])

results = []
for case, document_path, json_path in [
    ("original", "evidence/original.result.docx", "output/skill-sample-1.json"),
    ("two", "output/skill-sample-1.result.docx", "evidence/two-materials.json"),
    ("empty", "evidence/empty.result.docx", "evidence/empty-materials.json"),
]:
    data = json.loads((BASE / json_path).read_text(), parse_float=Decimal)
    parts, body, relationship_count = inspect(BASE / document_path)
    assert parts.keys() == source.keys()
    unchanged = [name for name in source if source[name] == parts[name]]
    assert set(source) - set(unchanged) <= {"word/document.xml"}
    table = body.find("w:tbl", NS)
    rows = table.findall("w:tr", NS)
    assert len(rows) == 1 + len(data["Materials"])
    assert signature(rows[0]) == signature(source_rows[0])
    for row, item in zip(rows[1:], data["Materials"]):
        cells = [text(cell) for cell in row.findall("w:tc", NS)]
        expected = [item["Reference"], datetime.date.fromisoformat(item["DeliveryDate"]).strftime("%d %b %Y"),
                    item["Name"], str(item["Quantity"]), item["Unit"],
                    f'{item["UnitPrice"]:,.2f}', f'{item["LineTotal"]:,.2f}', item["Currency"]]
        assert cells == expected, (case, cells, expected)
        assert formatting(row) == formatting(source_rows[1])
        assert Decimal(item["Quantity"]) * Decimal(item["UnitPrice"]) == Decimal(item["LineTotal"])
    assert data["MaterialCount"] == len(data["Materials"])
    assert Decimal(data["TotalEstimatedCost"]) == sum((Decimal(item["LineTotal"]) for item in data["Materials"]), Decimal(0))
    paragraphs = body.findall("w:p", NS)
    expected_summaries = [f'Material items: {data["MaterialCount"]}',
                          f'Total estimated cost: {data["TotalEstimatedCost"]:,.2f} USD',
                          'Amount in words: ' + data["AmountInWords"]]
    assert [text(p) for p in paragraphs[-3:]] == expected_summaries
    assert "{" not in text(body) and "}" not in text(body)
    for i, child in enumerate(body):
        if child.tag == W + "tbl":
            for tag in ["tblPr", "tblGrid"]:
                assert signature(child.find(W + tag)) == signature(source_body[i].find(W + tag))
        elif i < 10 or child.tag == W + "sectPr":
            assert signature(child) == signature(source_body[i])
        else:
            assert formatting(child) == formatting(source_body[i])
    identical = signature(body) == signature(source_body)
    if case == "original":
        assert identical, "The three-material generation must reproduce the original XML structure and values."
    results.append(dict(case=case, materialRows=len(rows)-1, totalRows=len(rows),
                        valuesMatch=True, formattingMatches=True, unrelatedContentMatches=True,
                        relationships=relationship_count, unchangedParts=unchanged,
                        originalBodyStructurallyIdentical=identical))

initial = json.loads((BASE / "evidence/initial-attempt.json").read_text())
assert hashlib.sha256((BASE / "input/skill-sample-1.docx").read_bytes()).hexdigest() == initial["sourceSha256"]
assert hashlib.sha256((BASE / "output/skill-sample-1.template.docx").read_bytes()).hexdigest() == initial["templateSha256"]
report = dict(sourceSha256=initial["sourceSha256"], initialTemplateUnchanged=True,
              templateChangedParts=["word/document.xml"],
              templateOtherEntriesByteIdentical=True, templateFormattingMatches=True,
              templateMarkerRowsValid=True, results=results)
(BASE / "evidence/verification.json").write_text(json.dumps(report, indent=2) + "\n")
print(json.dumps(report, indent=2))

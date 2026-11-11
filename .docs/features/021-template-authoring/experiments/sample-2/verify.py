"""Verify actual Sample 2 artifacts against data, source structure and package parts."""

import hashlib
import json
import posixpath
import re
import zipfile
from datetime import date
from decimal import Decimal
from pathlib import Path
from urllib.parse import unquote
from xml.dom import minidom
from xml.etree import ElementTree as ET

BASE = Path(__file__).resolve().parent
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"


def text(node):
    return "".join(t.text or "" for t in node.iter(W + "t"))


def signature(node):
    # Ignore only editing-session identifiers and text, retaining formatting/structure.
    attrs = tuple(sorted((k, v) for k, v in node.attrib.items()
                         if not k.startswith(W + "rsid")
                         and k.rsplit("}", 1)[-1] not in ["paraId", "textId"]))
    return (node.tag, attrs, None if node.tag == W + "t" else node.text,
            tuple(signature(child) for child in node))


def check_package(path):
    with zipfile.ZipFile(path) as package:
        assert package.testzip() is None
        names = package.namelist()
        assert len(names) == len(set(names))
        xml = {name: ET.fromstring(package.read(name)) for name in names
               if name.endswith((".xml", ".rels"))}
        types = xml["[Content_Types].xml"]
        defaults = {n.attrib["Extension"] for n in types.findall(CT + "Default")}
        overrides = {n.attrib["PartName"].lstrip("/") for n in types.findall(CT + "Override")}
        for name in names:
            assert name == "[Content_Types].xml" or name in overrides or name.rsplit(".", 1)[-1] in defaults
        rel_count = 0
        for name, relationships in xml.items():
            if not name.endswith(".rels"):
                continue
            parent = "" if name == "_rels/.rels" else name.split("/_rels/")[0]
            ids = [r.attrib["Id"] for r in relationships]
            assert len(ids) == len(set(ids))
            for relationship in relationships:
                if relationship.attrib.get("TargetMode") == "External":
                    continue
                target = unquote(relationship.attrib["Target"].split("#", 1)[0])
                resolved = posixpath.normpath(posixpath.join(parent, target)) if not target.startswith("/") else target[1:]
                assert resolved in names, (name, target, resolved)
                rel_count += 1
            if name != "_rels/.rels":
                owner = parent + "/" + name.rsplit("/", 1)[-1].removesuffix(".rels")
                if owner in xml:
                    for node in xml[owner].iter():
                        for key, value in node.attrib.items():
                            if key.startswith(R):
                                assert value in ids, (owner, key, value)
        main_bytes = package.read("word/document.xml")
        dom = minidom.parseString(main_bytes)
        root = dom.documentElement
        namespaces = {a.name: a.value for a in root.attributes.values() if a.name.startswith("xmlns")}
        for node in [root] + list(root.getElementsByTagName("*")):
            if node.hasAttribute("mc:Ignorable"):
                for prefix in node.getAttribute("mc:Ignorable").split():
                    assert "xmlns:" + prefix in namespaces
        return dict(parts={n: hashlib.sha256(package.read(n)).hexdigest() for n in names},
                    relationships=rel_count, namespaces=namespaces,
                    body=xml["word/document.xml"].find(W + "body"), root=xml["word/document.xml"])


def amount(value):
    return f"{Decimal(value):,.2f}"


def display_date(value):
    return date.fromisoformat(value).strftime("%d %b %Y")


def verify_result(name, path, data, source):
    actual = check_package(path)
    source_parts = source["parts"]
    assert actual["parts"].keys() == source_parts.keys()
    changed = [n for n in source_parts if source_parts[n] != actual["parts"][n]]
    assert changed == ["word/document.xml"]
    assert actual["namespaces"] == source["namespaces"]
    body = list(actual["body"])
    prototype = list(source["body"])
    orders = data["Orders"]
    assert len(body) == 9 + 5 * len(orders)
    assert len(actual["body"].findall(W + "tbl")) == len(orders)
    assert data["OrderCount"] == len(orders)
    assert Decimal(data["Total"]) == sum((Decimal(o["Total"]) for o in orders), Decimal(0))
    expected_intro = ["Sales Orders Report", f'Organization: {data["Organization"]}',
                      f'Warehouse: {data["Warehouse"]["Name"]}',
                      f'Warehouse address: {data["Warehouse"]["Address"]}',
                      f'Period: {display_date(data["Period"]["Start"])} – {display_date(data["Period"]["End"])}',
                      f'Orders: {len(orders)}     Total: {amount(data["Total"])} {data["Currency"]}',
                      "The following section repeats for each order."]
    assert [text(n) for n in body[:7]] == expected_intro
    for expected, current in zip(prototype[:7], body[:7]):
        assert signature(expected) == signature(current)
    row_counts = []
    for i, order in enumerate(orders):
        heading, metadata, summary, table, subtotal = body[7 + i * 5:12 + i * 5]
        assert text(heading) == f'Order {order["Number"]}'
        assert text(metadata) == f'Date: {display_date(order["Date"])}    Customer: {order["Customer"]}'
        assert text(summary) == f'Lines: {len(order["Lines"])}    Order total: {amount(order["Total"])} {data["Currency"]}'
        assert text(subtotal) == f'Order subtotal: {amount(order["Total"])} {data["Currency"]}'
        assert order["LineCount"] == len(order["Lines"])
        assert Decimal(order["Total"]) == sum((Decimal(line["Total"]) for line in order["Lines"]), Decimal(0))
        for template_node, current in zip(prototype[7:12], [heading, metadata, summary, table, subtotal]):
            if current.tag != W + "tbl":
                assert signature(template_node) == signature(current)
        for tag in ["tblPr", "tblGrid"]:
            assert ET.tostring(table.find(W + tag)) == ET.tostring(prototype[10].find(W + tag))
        rows = table.findall(W + "tr")
        source_rows = prototype[10].findall(W + "tr")
        assert ET.tostring(rows[0]) == ET.tostring(source_rows[0])
        assert len(rows) == len(order["Lines"]) + 1
        row_counts.append(len(rows))
        for index, (row, line) in enumerate(zip(rows[1:], order["Lines"]), 1):
            assert len(row.findall(W + "tc")) == 7
            assert line["Number"] == index
            assert Decimal(line["Total"]) == Decimal(line["UnitPrice"]) * line["Quantity"]
            assert [text(c) for c in row.findall(W + "tc")] == [
                str(index), line["Sku"], line["Description"], str(line["Quantity"]), line["Unit"],
                amount(line["UnitPrice"]), amount(line["Total"])]
            assert signature(row) == signature(source_rows[1])
    assert text(body[-2]) == f'REPORT TOTAL: {amount(data["Total"])} {data["Currency"]}'
    assert signature(body[-2]) == signature(prototype[-2])
    assert ET.tostring(body[-1]) == ET.tostring(prototype[-1])
    assert body[-1].tag == W + "sectPr"
    assert not re.search(r"\{[^}]*\}", text(actual["root"]))
    original_equal = ET.tostring(actual["root"]) == ET.tostring(source["root"])
    if name == "original":
        assert original_equal
    return dict(case=name, orders=len(orders), lines=sum(len(o["Lines"]) for o in orders),
                tableRowsIncludingHeaders=row_counts, total=data["Total"], currency=data["Currency"],
                changedParts=changed, unchangedParts=len(source_parts) - len(changed),
                allValuesHierarchyAndSummariesCorrect=True, sourceFormattingRetained=True,
                parsedMainDocumentEqualsSource=original_equal, relationships=actual["relationships"])


def main():
    source_path = BASE / "input/skill-sample-2.docx"
    assert hashlib.sha256(source_path.read_bytes()).hexdigest() == (BASE / "evidence/input-sha256.txt").read_text().strip()
    source = check_package(source_path)
    template = check_package(BASE / "output/skill-sample-2.template.docx")
    assert template["parts"].keys() == source["parts"].keys()
    assert template["namespaces"] == source["namespaces"]
    changed = [n for n in source["parts"] if source["parts"][n] != template["parts"][n]]
    assert changed == ["word/document.xml"]
    assert ET.tostring(list(template["body"])[-1]) == ET.tostring(list(source["body"])[-1])
    assert len(template["body"].findall(W + "tbl")) == 1
    assert len(template["body"].find(W + "tbl").findall(W + "tr")) == 4
    assert len(list(template["body"])) == 16
    assert ET.tostring(template["body"][0]) == ET.tostring(source["body"][0])
    assert ET.tostring(template["body"][6]) == ET.tostring(source["body"][6])
    results = []
    for name in ["original", "modified", "expanded", "empty-inner", "empty-outer"]:
        data_path = BASE / ("output/skill-sample-2.json" if name == "original" else f"evidence/{name}-data.json")
        output = BASE / ("output/skill-sample-2.result.docx" if name == "modified" else f"evidence/{name}.result.docx")
        results.append(verify_result(name, output, json.loads(data_path.read_text()), source))
    result = dict(inputUnchanged=True, template=dict(parts=len(template["parts"]), changedParts=changed,
                  unchangedParts=len(source["parts"]) - 1, validMarkerFootprint=True), cases=results)
    (BASE / "evidence/verification.json").write_text(json.dumps(result, indent=2) + "\n")
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()

import email.policy
from email.parser import BytesParser
import hashlib
import json
from pathlib import Path
import re
from zipfile import ZipFile
from xml.etree import ElementTree as ET

from author import BASE, NS, cell_value, mark, save_json


def unpack(case):
    headers = (BASE / f"evidence/{case}-render.headers").read_bytes()
    body = (BASE / f"evidence/{case}-response.mime").read_bytes()
    assert b"200 OK" in headers, (case, body[:1000])
    message = BytesParser(policy=email.policy.default).parsebytes(headers.split(b"\r\n", 1)[1] + body)
    assert message.is_multipart(), (case, headers)
    target = BASE / ("output/skill-sample-4.result.xlsx" if case == "modified" else f"evidence/{case}.result.xlsx")
    diagnostics = None
    for part in message.iter_parts():
        if part.get_content_type() == "application/json":
            diagnostics = json.loads(part.get_payload(decode=True))
        else:
            target.write_bytes(part.get_payload(decode=True))
    assert diagnostics is not None
    save_json(BASE / f"evidence/{case}-diagnostics.json", diagnostics)
    assert diagnostics["warnings"] == [], diagnostics
    return target


def load(path):
    with ZipFile(path) as z:
        assert z.testzip() is None
        assert len(z.namelist()) == len(set(z.namelist()))
        for name in z.namelist():
            if name.endswith((".xml", ".rels")):
                ET.fromstring(z.read(name))
        return {name: z.read(name) for name in z.namelist()}


def normalized(element):
    return (element.tag, sorted(element.attrib.items()), element.text or "", [normalized(c) for c in element])


def verify(case, path, data):
    source = load(BASE / "input/skill-sample-4.xlsx")
    result = load(path)
    sr = ET.fromstring(source["xl/worksheets/sheet1.xml"])
    rr = ET.fromstring(result["xl/worksheets/sheet1.xml"])
    assert set(source) == set(result)
    unchanged = [name for name in source if source[name] == result[name]]
    for name in source:
        if name not in ["xl/workbook.xml", "xl/worksheets/sheet1.xml"]:
            assert source[name] == result[name], name
    expected = [(3 if rec["$level"] == 1 else 4, rec) for rec in data["Data"]]
    expected = [(1, {}), (2, {})] + expected
    rows = rr.findall("s:sheetData/s:row", NS)
    assert len(rows) == len(expected), (case, len(rows), len(expected))
    source_rows = {int(row.get("r")): row for row in sr.findall("s:sheetData/s:row", NS)}
    expected_merges = []
    for number, (row, (source_number, rec)) in enumerate(zip(rows, expected), 1):
        assert row.get("r") == str(number)
        assert {k: v for k, v in row.attrib.items() if k != "r"} == {
            k: v for k, v in source_rows[source_number].attrib.items() if k != "r"}
        cells = list(row)
        originals = list(source_rows[source_number])
        assert len(cells) == len(originals), (case, number, len(cells), len(originals))
        for cell, original in zip(cells, originals):
            col = re.sub(r"\d", "", original.get("r"))
            assert cell.get("r") == f"{col}{number}"
            assert cell.get("s") == original.get("s")
            assert cell.get("t") == original.get("t")
            expected_value = cell_value(original)
            if source_number == 3 and col == "B":
                expected_value = rec["Department"]
            if source_number == 4 and col in ["C", "D"]:
                expected_value = rec["Name" if col == "C" else "Position"]
            assert cell_value(cell) == expected_value, (case, number, col, cell_value(cell), expected_value)
            assert "$level" not in cell_value(cell)
            assert not re.search(r"\{(?:Department|Name|Position)\}", cell_value(cell))
        if source_number == 3:
            expected_merges.append(f"B{number}:D{number}")
    assert [m.get("ref") for m in rr.findall("s:mergeCells/s:mergeCell", NS)] == expected_merges
    for tag in ["cols", "sheetFormatPr", "sheetViews", "pageMargins", "pageSetup"]:
        assert normalized(sr.find("s:" + tag, NS)) == normalized(rr.find("s:" + tag, NS)), (case, tag)
    assert rr.find("s:dimension", NS).get("ref") == f"A1:D{len(rows)}"
    workbook = ET.fromstring(result["xl/workbook.xml"])
    assert workbook.find("s:definedNames/s:definedName", NS).text == f"'Employees'!$A$1:$D${len(rows)}"
    source_workbook = ET.fromstring(source["xl/workbook.xml"])
    assert normalized(source_workbook.find("s:sheets", NS)) == normalized(workbook.find("s:sheets", NS))
    if case == "original":
        assert normalized(sr) == normalized(rr), "Original worksheet differs structurally"
        assert normalized(source_workbook) == normalized(workbook), "Original workbook differs structurally"
    return {"passed": True, "rows": len(rows), "departments": len(expected_merges),
            "employees": sum(rec["$level"] == 2 for rec in data["Data"]),
            "dimension": f"A1:D{len(rows)}", "merges": expected_merges,
            "byte_identical_parts": unchanged, "part_count": len(result),
            "worksheet_structurally_identical_to_input": case == "original",
            "sha256": hashlib.sha256(path.read_bytes()).hexdigest()}


def main():
    import sys
    report_path = BASE / "evidence/verification.json"
    report = json.loads(report_path.read_text()) if report_path.exists() else {}
    for case in sys.argv[1:]:
        data_path = BASE / ("output/skill-sample-4.json" if case == "original" else f"evidence/{case}-data.json")
        path = unpack(case)
        report[case] = verify(case, path, json.loads(data_path.read_text()))
        save_json(report_path, report)
        mark("first_verified_generated_document" if case == "original" else case + "_verified")
        print(case, "verified:", report[case]["rows"], "rows")


if __name__ == "__main__":
    main()

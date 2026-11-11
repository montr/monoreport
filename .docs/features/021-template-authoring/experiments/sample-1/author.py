"""One-off authoring of the supplied estimate; no sample templates are used."""

import copy
import datetime
import hashlib
import json
from pathlib import Path
import re
import zipfile
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

BASE = Path(__file__).resolve().parent
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
SOURCE = BASE / "input/skill-sample-1.docx"
source_hash = hashlib.sha256(SOURCE.read_bytes()).hexdigest()

with zipfile.ZipFile(SOURCE) as package:
    source_xml = package.read("word/document.xml")
    root = ET.fromstring(source_xml)
    body = root.find("w:body", NS)
    tables = body.findall("w:tbl", NS)
    assert len(tables) == 1
    rows = tables[0].findall("w:tr", NS)
    assert len(rows) == 4
    assert not root.findall(".//w:vMerge", NS)
    materials = []
    fields = ["Reference", "DeliveryDate", "Name", "Quantity", "Unit", "UnitPrice", "LineTotal", "Currency"]
    for row in rows[1:]:
        values = ["".join(t.text or "" for t in cell.findall(".//w:t", NS))
                  for cell in row.findall("w:tc", NS)]
        assert len(values) == len(fields)
        item = dict(zip(fields, values))
        item["DeliveryDate"] = datetime.datetime.strptime(item["DeliveryDate"], "%d %b %Y").date().isoformat()
        item["Quantity"] = int(item["Quantity"])
        for name in ["UnitPrice", "LineTotal"]:
            item[name] = float(item[name].replace(",", ""))
        materials.append(item)
    paragraphs = body.findall("w:p", NS)
    words = "".join(t.text or "" for t in paragraphs[-1].findall(".//w:t", NS))
    data = {"Materials": materials, "MaterialCount": len(materials),
            "TotalEstimatedCost": sum(item["LineTotal"] for item in materials),
            "AmountInWords": words.removeprefix("Amount in words: ")}

    # Replace only known text nodes and the contiguous three-row source region.
    # All surrounding XML bytes, prefixes, whitespace and package entries survive.
    xml = source_xml.decode("utf-8")
    raw_rows = re.findall(r"<w:tr>.*?</w:tr>", xml)
    assert len(raw_rows) == 4
    prototype = raw_rows[1]
    text_pattern = r"(<w:t(?:\s[^>]*)?>)(.*?)(</w:t>)"
    assert len(re.findall(text_pattern, prototype)) == 8
    expressions = ["{Reference}", '{DeliveryDate date "dd MMM yyyy"}', "{Name}",
                   "{Quantity}", "{Unit}", '{UnitPrice number "#,##0.00"}',
                   '{LineTotal number "#,##0.00"}', "{Currency}"]
    replacements = iter(expressions)
    repeated_row = re.sub(text_pattern, lambda m: m[1] + escape(next(replacements)) + m[3], prototype)
    marker_rows = []
    for marker in ["{For:Materials}", "{/For}"]:
        replacements = iter([marker] + [""] * 7)
        marker_rows.append(re.sub(text_pattern, lambda m: m[1] + next(replacements) + m[3], prototype))
    old_region = "".join(raw_rows[1:])
    assert xml.count(old_region) == 1
    xml = xml.replace(old_region, marker_rows[0] + repeated_row + marker_rows[1], 1)
    summaries = [
        ("<w:t>3</w:t>", "<w:t>{MaterialCount}</w:t>"),
        ("<w:t>125,430.50</w:t>", '<w:t>{TotalEstimatedCost number "#,##0.00"}</w:t>'),
        ("<w:t>" + data["AmountInWords"] + "</w:t>", "<w:t>{AmountInWords}</w:t>"),
    ]
    for old, new in summaries:
        assert xml.count(old) == 1
        xml = xml.replace(old, new, 1)
    ET.fromstring(xml.encode("utf-8"))
    template = BASE / "output/skill-sample-1.template.docx"
    assert not template.exists(), "Preserve an existing initial attempt rather than overwriting it."
    with zipfile.ZipFile(template, "w") as output:
        for entry in package.infolist():
            output.writestr(entry, xml.encode("utf-8") if entry.filename == "word/document.xml"
                            else package.read(entry.filename))

(BASE / "output/skill-sample-1.json").write_text(json.dumps(data, indent=2) + "\n")
changed = copy.deepcopy(data)
changed["Materials"] = [changed["Materials"][0], changed["Materials"][2]]
changed.update(MaterialCount=2, TotalEstimatedCost=60000,
               AmountInWords="Sixty thousand dollars 00 cents")
(BASE / "evidence/two-materials.json").write_text(json.dumps(changed, indent=2) + "\n")
empty = dict(Materials=[], MaterialCount=0, TotalEstimatedCost=0,
             AmountInWords="Zero dollars 00 cents")
(BASE / "evidence/empty-materials.json").write_text(json.dumps(empty, indent=2) + "\n")
(BASE / "evidence/initial-attempt.json").write_text(json.dumps({
    "sourceSha256": source_hash,
    "templateSha256": hashlib.sha256(template.read_bytes()).hexdigest(),
    "exampleTemplatesRead": False,
    "existingSampleJsonRead": False,
    "testFixturesRead": False,
    "otherExperimentalOutputsRead": False,
    "authoringAttempt": 1,
}, indent=2) + "\n")
assert hashlib.sha256(SOURCE.read_bytes()).hexdigest() == source_hash
print(json.dumps(data, indent=2))

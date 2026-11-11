"""Prepare focused data variations and extract REST MIME responses."""
import copy
import json
import zipfile
from decimal import Decimal, ROUND_HALF_UP
from email import policy
from email.parser import BytesParser
from io import BytesIO

from author import BASE, stamp


def save(name, value):
    (BASE / f"evidence/{name}-data.json").write_text(json.dumps(value, indent=2, ensure_ascii=False) + "\n")


def recalculate(data):
    for project in data["Projects"]:
        for group in project["Groups"]:
            for item in group["Items"]:
                item["Total"] = float((Decimal(str(item["Quantity"])) * Decimal(str(item["UnitPrice"])))
                    .quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))
            group["Total"] = float(sum(Decimal(str(item["Total"])) for item in group["Items"]))
        project["Total"] = float(sum(Decimal(str(group["Total"])) for group in project["Groups"]))
    data["GrandTotal"] = float(sum(Decimal(str(project["Total"])) for project in data["Projects"]))


def extract(name):
    headers = (BASE / f"evidence/{name}-render.headers").read_bytes()
    raw = (BASE / f"evidence/{name}-render.mime").read_bytes()
    assert b"200" in headers.splitlines()[0], headers.decode()
    content_type = next(line for line in headers.splitlines() if line.lower().startswith(b"content-type:"))
    message = BytesParser(policy=policy.default).parsebytes(content_type + b"\r\nMIME-Version: 1.0\r\n\r\n" + raw)
    assert message.is_multipart() and not message.defects
    parts = list(message.iter_parts())
    assert len(parts) == 2 and all(not part.defects for part in parts)
    attachments = {part.get_filename(): part for part in parts}
    document = attachments["document.docx"].get_payload(decode=True)
    diagnostics = json.loads(attachments["diagnostics.json"].get_payload(decode=True))
    assert diagnostics["format"] == "docx" and diagnostics["fileName"] == "document.docx"
    assert diagnostics["warnings"] == [], diagnostics
    with zipfile.ZipFile(BytesIO(document)) as package:
        assert package.testzip() is None
    output = BASE / ("output/skill-sample-3.result.docx" if name == "modified" else f"evidence/{name}.result.docx")
    output.write_bytes(document)
    (BASE / f"evidence/{name}-diagnostics.json").write_text(json.dumps(diagnostics, indent=2) + "\n")
    stamp("first_successful_document" if name == "original" else f"{name}_generated")
    print(f"{name}: valid DOCX extracted; zero warnings")


def prepare():
    original = json.loads((BASE / "output/skill-sample-3.json").read_text())
    changed = copy.deepcopy(original)
    changed["ProjectName"] = "Orion 7"
    changed["Period"] = "December 2027"
    changed["Projects"][0]["Name"] = "Depot Expansion"
    changed["Projects"][0]["Groups"][0]["Name"] = "Revised Foundation"
    changed["Projects"][0]["Groups"][0]["Items"] = [
        {"Description": "Concrete & aggregate", "Unit": "m³", "Quantity": 42.5, "UnitPrice": 8100.25},
        {"Description": "Steel <premium>", "Unit": "t", "Quantity": 4.75, "UnitPrice": 64000},
    ]
    changed["Projects"][0]["Groups"].pop(1)
    changed["Projects"][1]["Groups"][0]["Items"].append(
        {"Description": "Control Panels", "Unit": "pcs", "Quantity": 3, "UnitPrice": 1200.5})
    changed["Projects"][2]["Groups"][-1]["Items"] = [
        {"Description": "Maple Trees", "Unit": "pcs", "Quantity": 8, "UnitPrice": 85.5}]
    changed["Projects"].append({"Name": "Reserve Works", "Groups": [
        {"Name": "Commissioning", "Items": [{"Description": "Final Inspection", "Unit": "lot", "Quantity": 1, "UnitPrice": 2500}]},
        {"Name": "Deferred Work", "Items": []},
    ]})
    recalculate(changed)
    save("modified", changed)
    for name in ["empty-items", "empty-groups", "empty-projects"]:
        data = copy.deepcopy(original)
        if name == "empty-items":
            for project in data["Projects"]:
                for group in project["Groups"]:
                    group["Items"] = []
        elif name == "empty-groups":
            for project in data["Projects"]:
                project["Groups"] = []
        else:
            data["Projects"] = []
        recalculate(data)
        save(name, data)
    stamp("test_cases_prepared")


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        extract(sys.argv[1])
    else:
        prepare()

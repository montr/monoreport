"""Prepare bounded experimental datasets and extract captured REST MIME responses."""

import copy
import io
import json
import sys
import zipfile
from decimal import Decimal
from email import policy
from email.parser import BytesParser
from pathlib import Path

BASE = Path(__file__).resolve().parent
EVIDENCE = BASE / "evidence"
DOCX_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


def prepare():
    original = json.loads((BASE / "output/skill-sample-2.json").read_text())
    for name, counts in [("modified", [0, 1, 24]), ("expanded", [0, 1, 2, 12] * 6),
                         ("empty-inner", [0, 0]), ("empty-outer", [])]:
        data = copy.deepcopy(original)
        data["Currency"] = "EUR"
        data["Organization"] = "Sample 2 Verification Ltd."
        data["Orders"] = []
        for index, count in enumerate(counts):
            order = copy.deepcopy(original["Orders"][index % len(original["Orders"])])
            order["Number"] = f"VERIFY-{index + 1:03}"
            order["Customer"] = f"Customer {index + 1:03} & Partners"
            order["Lines"] = [copy.deepcopy(order["Lines"][j % len(order["Lines"])]) for j in range(count)]
            for j, line in enumerate(order["Lines"]):
                line["Number"] = j + 1
                line["Quantity"] = j + 2
                line["Total"] = str(Decimal(line["UnitPrice"]) * line["Quantity"])
            order["LineCount"] = count
            order["Total"] = str(sum((Decimal(line["Total"]) for line in order["Lines"]), Decimal("0.00")))
            data["Orders"].append(order)
        data["OrderCount"] = len(data["Orders"])
        data["Total"] = str(sum((Decimal(order["Total"]) for order in data["Orders"]), Decimal("0.00")))
        (EVIDENCE / f"{name}-data.json").write_text(json.dumps(data, indent=2) + "\n")


def extract(name):
    headers = (EVIDENCE / f"{name}-render.headers").read_bytes()
    raw = (EVIDENCE / f"{name}-render.mime").read_bytes()
    assert headers.splitlines()[0].split()[1] == b"200", headers[:500]
    message_headers = BytesParser(policy=policy.default).parsebytes(headers.split(b"\r\n", 1)[1])
    assert int(message_headers["Content-Length"]) == len(raw)
    assert message_headers.get_content_type() == "multipart/mixed"
    boundary = message_headers.get_boundary().encode("ascii")
    assert raw.endswith(b"--" + boundary + b"--\r\n")
    message = BytesParser(policy=policy.default).parsebytes(
        b"Content-Type: " + str(message_headers["Content-Type"]).encode("ascii") + b"\r\nMIME-Version: 1.0\r\n\r\n" + raw)
    assert not message.defects
    parts = list(message.iter_parts())
    assert len(parts) == 2 and all(not p.defects for p in parts)
    documents = [p for p in parts if p.get_content_type() == DOCX_TYPE]
    diagnostics = [p for p in parts if p.get_content_type() == "application/json"]
    assert len(documents) == len(diagnostics) == 1
    assert diagnostics[0].get_filename() == "diagnostics.json"
    info = json.loads(diagnostics[0].get_payload(decode=True))
    assert info["format"] == "docx" and info["fileName"] == documents[0].get_filename()
    assert isinstance(info["warnings"], list)
    payload = documents[0].get_payload(decode=True)
    with zipfile.ZipFile(io.BytesIO(payload)) as package:
        assert package.testzip() is None
        assert "word/document.xml" in package.namelist()
    destination = (BASE / "output/skill-sample-2.result.docx" if name == "modified"
                   else EVIDENCE / f"{name}.result.docx")
    destination.write_bytes(payload)
    (EVIDENCE / f"{name}-diagnostics.json").write_text(json.dumps(info, indent=2) + "\n")
    print(json.dumps(dict(case=name, document=str(destination.relative_to(BASE)),
                          bytes=len(payload), warnings=info["warnings"])))


if __name__ == "__main__":
    if sys.argv[1] == "prepare":
        prepare()
    else:
        for case in sys.argv[1:]:
            extract(case)

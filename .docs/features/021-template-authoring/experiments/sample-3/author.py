"""Sample-specific package editing; no document editor or template DSL."""
import json
import zipfile
from datetime import datetime, timezone
from decimal import Decimal
from pathlib import Path
from xml.dom import minidom

BASE = Path(__file__).resolve().parent
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def stamp(stage):
    path = BASE / "evidence/timings.json"
    timings = json.loads(path.read_text()) if path.exists() else {}
    timings[stage] = datetime.now(timezone.utc).isoformat()
    path.write_text(json.dumps(timings, indent=2) + "\n")


def text(node):
    return "".join(t.firstChild.data if t.firstChild else "" for t in node.getElementsByTagNameNS(W, "t"))


def children(node, name):
    return [c for c in node.childNodes if c.nodeType == c.ELEMENT_NODE and c.localName == name]


def money(value):
    return float(Decimal(value.replace(",", "")))


def main():
    stamp("authoring_started")
    source = BASE / "input/skill-sample-3.docx"
    with zipfile.ZipFile(source) as package:
        xml = package.read("word/document.xml")
        doc = minidom.parseString(xml)
        body = doc.getElementsByTagNameNS(W, "body")[0]
        table = children(body, "tbl")[0]
        rows = children(table, "tr")
        data = {"ProjectName": "Space 3", "Period": "October 2026", "Projects": []}
        for row in rows[1:-1]:
            cells = [text(c) for c in children(row, "tc")]
            if len(cells) == 5:
                group["Items"].append(dict(zip(
                    ["Description", "Unit", "Quantity", "UnitPrice", "Total"],
                    [cells[0].strip(), cells[1], *map(money, cells[2:])],
                )))
            elif cells[0].startswith("    "):
                group = {"Name": cells[0].strip(), "Total": money(cells[1]), "Items": []}
                project["Groups"].append(group)
            else:
                project = {"Name": cells[0], "Total": money(cells[1]), "Groups": []}
                data["Projects"].append(project)
        data["GrandTotal"] = money(text(children(rows[-1], "tc")[-1]))
        (BASE / "output/skill-sample-3.json").write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n")
        (BASE / "evidence/source-document.xml").write_bytes(xml)

        # Edit text nodes only; leave every existing run and its properties intact.
        title_nodes = children(body, "p")[0].getElementsByTagNameNS(W, "t")
        title_nodes[1].firstChild.data = "“{ProjectName}"
        title_nodes[2].firstChild.data = ""
        period_nodes = children(body, "p")[1].getElementsByTagNameNS(W, "t")
        period_nodes[1].firstChild.data = " {Period}"

        def prototype(index, values):
            row = rows[index].cloneNode(True)
            for cell, value in zip(children(row, "tc"), values):
                nodes = cell.getElementsByTagNameNS(W, "t")
                assert len(nodes) == 1
                nodes[0].firstChild.data = value
            # Word's editing-session identities must not be duplicated by loops.
            for node in [row, *row.getElementsByTagName("*")]:
                for attr in ["w14:paraId", "w14:textId"]:
                    if node.hasAttribute(attr):
                        node.removeAttribute(attr)
            return row

        def marker(value):
            row = doc.createElementNS(W, "w:tr")
            for i in range(5):
                cell = doc.createElementNS(W, "w:tc")
                para = doc.createElementNS(W, "w:p")
                cell.appendChild(para)
                row.appendChild(cell)
                if i == 0:
                    run = doc.createElementNS(W, "w:r")
                    token = doc.createElementNS(W, "w:t")
                    token.appendChild(doc.createTextNode(value))
                    run.appendChild(token)
                    para.appendChild(run)
            return row

        total = '{Total number "#,##0.00"}'
        replacement = [
            marker("{For:Projects}"), prototype(1, ["{Name}", total]),
            marker("{For:Groups}"), prototype(2, ["    {Name}", total]),
            marker("{For:Items}"), prototype(3, ["        {Description}", "{Unit}",
                '{Quantity number "#,##0.00"}', '{UnitPrice number "#,##0.00"}', total]),
            marker("{/For}"), marker("{/For}"), marker("{/For}"),
        ]
        for row in rows[1:-1]:
            table.removeChild(row)
        for row in replacement:
            table.insertBefore(row, rows[-1])
        nodes = children(rows[-1], "tc")[-1].getElementsByTagNameNS(W, "t")
        nodes[0].firstChild.data = '{GrandTotal number "#,##0.00"}'
        with zipfile.ZipFile(BASE / "output/skill-sample-3.template.docx", "w") as output:
            for entry in package.infolist():
                output.writestr(entry, doc.toxml(encoding="UTF-8") if entry.filename == "word/document.xml" else package.read(entry.filename))
    stamp("initial_authoring_completed")
    print("Authored three nested row loops; extracted 3 projects, 8 groups, 16 items.")


if __name__ == "__main__":
    main()

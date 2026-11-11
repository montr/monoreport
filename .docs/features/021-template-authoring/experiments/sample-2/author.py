"""One-off Sample 2 extraction and targeted authoring; no example templates used."""

import hashlib
import json
import re
import zipfile
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from xml.dom import minidom

BASE = Path(__file__).resolve().parent
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def children(node, name=None):
    return [n for n in node.childNodes if n.nodeType == n.ELEMENT_NODE
            and (name is None or n.tagName == name)]


def text(node):
    return "".join(t.firstChild.data if t.firstChild else ""
                   for t in node.getElementsByTagName("w:t"))


def change(node, old, new):
    # This source has one nonempty text node at every selected edit location.
    matches = [t for t in node.getElementsByTagName("w:t")
               if t.firstChild and old in t.firstChild.data]
    assert len(matches) == 1, (old, len(matches))
    matches[0].firstChild.data = matches[0].firstChild.data.replace(old, new, 1)


def money(value):
    return str(Decimal(value.replace(",", "")))


def save_json(path, value):
    path.write_text(json.dumps(value, indent=2, ensure_ascii=False) + "\n")


def main():
    source = BASE / "input/skill-sample-2.docx"
    with zipfile.ZipFile(source) as package:
        doc = minidom.parseString(package.read("word/document.xml"))
        body = doc.getElementsByTagName("w:body")[0]
        blocks = children(body)
        assert len(blocks) == 109
        orders = []
        for offset in range(7, 107, 5):
            heading, metadata, summary, table, subtotal = blocks[offset:offset + 5]
            date, customer = re.fullmatch(r"Date: (.*?)    Customer: (.*)", text(metadata)).groups()
            count, total = re.fullmatch(r"Lines: (\d+)    Order total: ([\d,.]+) USD", text(summary)).groups()
            lines = []
            for row in children(table, "w:tr")[1:]:
                number, sku, description, quantity, unit, price, amount = map(text, children(row, "w:tc"))
                line = dict(Number=int(number), Sku=sku, Description=description,
                            Quantity=int(quantity), Unit=unit, UnitPrice=money(price), Total=money(amount))
                assert Decimal(line["UnitPrice"]) * line["Quantity"] == Decimal(line["Total"])
                lines.append(line)
            assert len(lines) == int(count)
            assert sum(Decimal(line["Total"]) for line in lines) == Decimal(money(total))
            assert text(subtotal) == f"Order subtotal: {total} USD"
            orders.append(dict(Number=text(heading).removeprefix("Order "),
                               Date=datetime.strptime(date, "%d %b %Y").date().isoformat(),
                               Customer=customer, LineCount=int(count), Total=money(total), Lines=lines))
        start, end = text(blocks[4]).removeprefix("Period: ").split(" – ")
        data = dict(Organization=text(blocks[1]).removeprefix("Organization: "),
                    Warehouse=dict(Name=text(blocks[2]).removeprefix("Warehouse: "),
                                   Address=text(blocks[3]).removeprefix("Warehouse address: ")),
                    Period=dict(Start=datetime.strptime(start, "%d %b %Y").date().isoformat(),
                                End=datetime.strptime(end, "%d %b %Y").date().isoformat()),
                    Currency="USD", OrderCount=len(orders),
                    Total=str(sum(Decimal(order["Total"]) for order in orders)), Orders=orders)
        assert text(blocks[5]) == f'Orders: 20     Total: {Decimal(data["Total"]):,.2f} USD'
        assert text(blocks[107]) == f'REPORT TOTAL: {Decimal(data["Total"]):,.2f} USD'
        save_json(BASE / "output/skill-sample-2.json", data)

        # Preserve paragraphs, runs and their properties; only change selected text.
        replacements = {
            1: [(data["Organization"], "{Organization}")],
            2: [(data["Warehouse"]["Name"], "{Warehouse.Name}")],
            3: [(data["Warehouse"]["Address"], "{Warehouse.Address}")],
            4: [(start, '{Period.Start date "dd MMM yyyy"}'), (end, '{Period.End date "dd MMM yyyy"}')],
            5: [("20", "{OrderCount}"), ("60,206.90", '{Total number "#,##0.00"}'), ("USD", "{Currency}")],
            7: [(orders[0]["Number"], "{Number}")],
            8: [("01 Sep 2026", '{Date date "dd MMM yyyy"}'), (orders[0]["Customer"], "{Customer}")],
            9: [("8", "{LineCount}"), ("2,638.80", '{Total number "#,##0.00"}'), ("USD", "{$.Currency}")],
            11: [("2,638.80", '{Total number "#,##0.00"}'), ("USD", "{$.Currency}")],
            107: [("60,206.90", '{Total number "#,##0.00"}'), ("USD", "{Currency}")],
        }
        for index, edits in replacements.items():
            for old, new in edits:
                change(blocks[index], old, new)

        table = blocks[10]
        rows = children(table, "w:tr")
        prototype = rows[1]
        expressions = ["{Number}", "{Sku}", "{Description}", "{Quantity}", "{Unit}",
                       '{UnitPrice number "#,##0.00"}', '{Total number "#,##0.00"}']
        for cell, expression in zip(children(prototype, "w:tc"), expressions):
            change(cell, text(cell), expression)
        for row in rows[2:]:
            table.removeChild(row)
        for marker, before in [("{For:Lines}", prototype), ("{/For}", None)]:
            row = prototype.cloneNode(deep=True)
            for index, cell in enumerate(children(row, "w:tc")):
                change(cell, text(cell), marker if index == 0 else "")
            if before is None:
                table.appendChild(row)
            else:
                table.insertBefore(row, before)

        # Remove materialized orders 2–20, retaining the static report total and section.
        for block in blocks[12:107]:
            body.removeChild(block)
        for marker, before in [("{For:Orders}", blocks[7]), ("{/For}", blocks[107])]:
            paragraph = doc.createElementNS(WORD_NS, "w:p")
            run = doc.createElementNS(WORD_NS, "w:r")
            value = doc.createElementNS(WORD_NS, "w:t")
            value.appendChild(doc.createTextNode(marker))
            run.appendChild(value)
            paragraph.appendChild(run)
            body.insertBefore(paragraph, before)

        # minidom retains existing prefixes and every xmlns binding, including mc:Ignorable.
        edited = doc.toxml(encoding="UTF-8", standalone=True)
        target = BASE / "output/skill-sample-2.template.docx"
        with zipfile.ZipFile(target, "w") as output:
            for entry in package.infolist():
                output.writestr(entry, edited if entry.filename == "word/document.xml" else package.read(entry))
    save_json(BASE / "evidence/initial-attempt.json", dict(
        sourceSha256=hashlib.sha256(source.read_bytes()).hexdigest(),
        templateSha256=hashlib.sha256(target.read_bytes()).hexdigest(),
        outerLoop="Orders", innerLoop="Lines", orders=len(orders), lines=sum(len(o["Lines"]) for o in orders),
        total=data["Total"], consultedExamples=False,
        approach="Python standard-library ZIP/minidom, one-off targeted text and block edits"))
    print(json.dumps(dict(orders=len(orders), lines=sum(len(o["Lines"]) for o in orders), total=data["Total"])))


if __name__ == "__main__":
    main()

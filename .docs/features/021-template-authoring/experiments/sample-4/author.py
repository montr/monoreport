import copy
import datetime
import json
from pathlib import Path
import re
from zipfile import ZipFile
from xml.etree import ElementTree as ET

BASE = Path(__file__).resolve().parent
NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
START = "2026-11-11T22:56:34+00:00"


def save_json(path, value):
    path.write_text(json.dumps(value, indent=2, ensure_ascii=False) + "\n")


def mark(stage):
    path = BASE / "evidence/timings.json"
    times = json.loads(path.read_text()) if path.exists() else {"started": START}
    times[stage] = datetime.datetime.now(datetime.timezone.utc).isoformat()
    save_json(path, times)


def cell_value(cell):
    return "".join(t.text or "" for t in cell.findall(".//s:t", NS))


def main():
    (BASE / "output").mkdir(exist_ok=True)
    (BASE / "evidence").mkdir(exist_ok=True)
    with ZipFile(BASE / "input/skill-sample-4.xlsx") as source:
        sheet = source.read("xl/worksheets/sheet1.xml").decode()
        root = ET.fromstring(sheet)
        records = []
        for row in root.findall("s:sheetData/s:row", NS)[2:]:
            cells = {re.sub(r"\d", "", c.get("r")): cell_value(c) for c in row}
            if cells.get("B"):
                records.append({"$level": 1, "Department": cells["B"]})
            else:
                records.append({"$level": 2, "Name": cells["C"], "Position": cells["D"]})
        original = {"Data": records}
        save_json(BASE / "output/skill-sample-4.json", original)
        save_json(BASE / "evidence/source-inventory.json", {
            "parts": source.namelist(), "sheet": "Employees", "range": "A1:D112",
            "departments": sum(r["$level"] == 1 for r in records),
            "employees": sum(r["$level"] == 2 for r in records),
            "merges": [m.get("ref") for m in root.findall("s:mergeCells/s:mergeCell", NS)],
            "columns": [c.attrib for c in root.findall("s:cols/s:col", NS)],
        })
        mark("inspection_complete")
        # Edit only worksheet and PrintArea bytes; keep all other package parts intact.
        rows = re.search(r"<sheetData>(.*?)</sheetData>", sheet).group(1)
        rows = rows[:rows.index('<row r="5"')]
        rows = re.sub(r'(<c r=")([A-D])(\d+")',
                      lambda m: m[1] + chr(ord(m[2]) + 1) + m[3], rows)
        for old, new in [("Engineering", "{Department}"), ("Avery Chen", "{Name}"),
                         ("Software Engineer", "{Position}")]:
            rows = rows.replace(">" + old + "<", ">" + new + "<")
        for number, level in [(3, 1), (4, 2)]:
            rows = re.sub(rf'(<row r="{number}"[^>]*>)',
                          rf'\1<c r="A{number}" t="inlineStr"><is><t>$level({level}).Data</t></is></c>', rows)
        sheet = re.sub(r"<sheetData>.*?</sheetData>", "<sheetData>" + rows + "</sheetData>", sheet)
        sheet = sheet.replace('ref="A1:D112"', 'ref="A1:E4"')
        sheet = re.sub(r"<mergeCells.*?</mergeCells>",
                       '<mergeCells count="1"><mergeCell ref="C3:E3" /></mergeCells>', sheet)
        sheet = re.sub(r'(<col min=")(\d+)(" max=")(\d+)',
                       lambda m: m[1] + str(int(m[2]) + 1) + m[3] + str(int(m[4]) + 1), sheet)
        sheet = sheet.replace('<selection activeCell="A1" sqref="A1"',
                              '<selection activeCell="B1" sqref="B1"')
        workbook = source.read("xl/workbook.xml").decode().replace(
            "'Employees'!$A$1:$D$112", "'Employees'!$B$1:$E$4")
        with ZipFile(BASE / "output/skill-sample-4.template.xlsx", "w") as target:
            for item in source.infolist():
                content = sheet.encode() if item.filename == "xl/worksheets/sheet1.xml" else (
                    workbook.encode() if item.filename == "xl/workbook.xml" else source.read(item.filename))
                target.writestr(item, content)
    # Unequal sizes include an empty department and growth beyond ten employees.
    modified = {"Data": [
        {"$level": 1, "Department": "Platform Engineering"},
        {"$level": 2, "Name": "Alex Rivera", "Position": "Principal Engineer"},
        {"$level": 2, "Name": "Sam Bell", "Position": "QA Lead"},
        {"$level": 1, "Department": "Research"},
        {"$level": 1, "Department": "Customer Operations"},
        *[{"$level": 2, "Name": f"Support Agent {i:02}", "Position": "Support Specialist"}
          for i in range(1, 13)],
    ]}
    save_json(BASE / "evidence/modified-data.json", modified)
    save_json(BASE / "evidence/empty-data.json", {"Data": []})
    save_json(BASE / "evidence/empty-children-data.json", {"Data": [
        {"$level": 1, "Department": "Research"}, {"$level": 1, "Department": "Operations"}]})
    expanded = copy.deepcopy(original)
    expanded["Data"].extend([{"$level": 1, "Department": "New Department"},
                             {"$level": 2, "Name": "Robin Fox", "Position": "Coordinator"}])
    save_json(BASE / "evidence/expanded-data.json", expanded)
    mark("initial_template_authored")
    print("Authored template; extracted 10 departments and 100 employees.")


if __name__ == "__main__":
    main()

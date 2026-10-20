"""Small supplied-plan probes for revised D1/D2/D3/D5/D6. Not a template engine."""
import copy
import datetime as dt
import json
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as E
from fixtures import el, cell, col, xy, sheet, write_package, S

BASE = Path(__file__).resolve().parent
OUT = BASE / "nested-evidence"
OUT.mkdir(exist_ok=True)

def refs(root):
    return {c.get("r"): c for c in root.iter(f"{{{S}}}c")}

def groups(records, leaves):
    result, closed = [], set()
    for record in records:
        key = record["$path"]
        if not result or result[-1][0]["$path"] != key:
            assert key not in closed, "noncontiguous path"
            closed.add(key)
            result.append([])
        result[-1].append(record)
    assert all(len(g) == (leaves or 1) and len({r["$level"] for r in g}) == 1 for g in result)
    return result

def blocks(levels):
    # Minimal proposed table boundary: a blank vertical-control row terminates the marked run.
    runs, current = [], []
    for row, marker in enumerate(levels, 1):
        if marker is None:
            if current:
                runs.append(current)
                current = []
        else:
            current.append((row, marker))
    if current:
        runs.append(current)
    return runs

def prepare():
    original = {
        1: {4: "$h(1)", 5: "$h(1)"}, 2: {5: "$h(2)"},
        6: {2: "DISPOSABLE ABOVE"},
        7: {1: "$header(1).Bids", 4: "{Company}", 6: "Static right"},
        8: {1: "$header(2).Criteria", 5: "{Name}"},
        9: {1: "$level(10).Items", 2: "{Name}", 5: "=SUBTOTAL(9,E10:E11)"},
        10: {1: "$l(20).Items", 2: "{Name}", 5: "{Price}"},
        11: {1: "$level(20).Items", 2: "", 5: "{Delta}"},
        12: {1: "$level(10).Items", 2: "Footer", 5: "=SUBTOTAL(9,E10:E11)"},
        13: {2: "Grand", 5: "=SUBTOTAL(9,E9:E12)"}, 14: {2: "Separator"},
        15: {1: "$level(5).Other", 2: "{Name}"}, 16: {2: "After second table"},
        17: {2: "DISPOSABLE BELOW"}, 18: {7: "DISPOSABLE RIGHT"}}
    template = sheet(original, ["D7:E7", "B10:B11"])
    # Authored zero width is an independent column property, not loss of physical column identity.
    template.find(f"{{{S}}}cols")[3].set("width", "0")
    template.find(f"{{{S}}}cols")[3].set("hidden", "1")
    next(row for row in template.find(f"{{{S}}}sheetData") if row.get("r")=="14").set("ht","0")
    next(row for row in template.find(f"{{{S}}}sheetData") if row.get("r")=="14").set("hidden","1")
    write_package(OUT / "nested-template.xlsx", {"Report": template}, [("_xlnm.Print_Area", "Report!$B$7:$F$16", 0)])
    result_log = {}
    for case, companies, criteria in [("full", 2, 3), ("one-criterion",2,1), ("seven-criteria",2,7), ("no-criteria", 2, 0), ("no-companies", 0, 3)]:
        leaves = companies * criteria
        records = []
        for level, path, name, base in [(10,"001","Lot 1",100),(20,"001/001","A",10),(20,"001/002","B",20),(10,"002","Lot 2",200),(20,"002/001","C",30)]:
            for i in range(leaves or 1):
                records.append({"$level":level,"$path":path,"Name":name if i == 0 else "not ordinary name","Price":base+i,"Delta":1})
        names=["Price","Term","Quality"]+["Criterion "+str(i) for i in range(4,criteria+1)]
        data = {"Bids":[{"Company":name} for name in ["Alpha","Beta"][:companies]], "Criteria":[{"Name":name} for name in names[:criteria]], "Items":records,"Other":[{"Name":s} for s in ["X","Y","Z"]]}
        (OUT / (case + "-input.json")).write_text(json.dumps(data,indent=2))
        logical = groups(records, leaves)
        # Source identity, output instance and parent ownership are retained before any mutation.
        rowplan = [(7,None,None),(8,None,None)]
        for parent, children in [(logical[0],logical[1:3]),(logical[3],logical[4:5])]:
            p = parent[0]["$path"]
            rowplan.append((9,parent,p))
            for child in children:
                rowplan.extend([(10,child,p),(11,child,p)])
            rowplan.append((12,parent,p))
        rowplan.extend([(13,None,None),(14,None,None)])
        rowplan.extend((15,[record],None) for record in data["Other"])
        rowplan.append((16,None,None))
        rowmap = {}
        for dest,(src,_,_) in enumerate(rowplan,1):
            rowmap.setdefault(src,[]).append(dest)
        # Recursive geometry for D:E parent, E child; parent-only D appears once per company.
        columns = [(2,None,None),(3,None,None)]
        for company in range(companies):
            columns.append((4,company,None))
            columns.extend((5,company,criterion) for criterion in range(criteria))
        columns.append((6,None,None))
        root = el("worksheet")
        root.append(el("sheetFormatPr",{"defaultRowHeight":"15"}))
        root.append(el("cols"))
        for dest,(src,_,_) in enumerate(columns,1):
            root[-1].append(el("col",{"min":str(dest),"max":str(dest),"width":"0" if src==4 else str(8+src),"customWidth":"1",**({"hidden":"1"} if src==4 else {})}))
        root.append(el("sheetData"))
        requests,targets = [],[]
        for dest,(src,group,parent) in enumerate(rowplan,1):
            row = el("row",{"r":str(dest),"ht":"0" if src==14 else str(18+src%3),"customHeight":"1",**({"hidden":"1"} if src==14 else {})})
            for dc,(sc,company,criterion) in enumerate(columns,1):
                if sc not in original[src]:
                    continue
                value = original[src][sc]
                if value.startswith("{"):
                    name = value[1:-1]
                    if src==7:
                        value = data["Bids"][company][name]
                    elif src==8:
                        value = data["Criteria"][criterion][name]
                    else:
                        i = company*criteria+criterion if sc==5 else 0
                        value = group[i][name]
                address = col(dc)+str(dest)
                node = cell(address,value,1 if sc>=4 else 0)
                if isinstance(value,str) and value.startswith("="):
                    rmap = dict(rowmap)
                    if parent:
                        for key in [9,10,11,12]:
                            rmap[key] = [pos for pos in rowmap[key] if rowplan[pos-1][2]==parent]
                    cmap = {2:[1],3:[2],6:[len(columns)]}
                    if company is not None:
                        cmap[4] = [pos for pos,(s,p,c) in enumerate(columns,1) if s==4 and p==company]
                        cmap[5] = [dc] # the formula belongs to this child intersection
                    requests.append({"Formula":value[1:],"Rows":rmap,"Cols":cmap})
                    targets.append(address)
                row.append(node)
            root[-1].append(row)
        projected = []
        literal = []
        for company in range(companies):
            positions = [i for i,(_,p,_) in enumerate(columns,1) if p==company]
            if len(positions)>1:
                projected.append(f"{col(positions[0])}1:{col(positions[-1])}1")
                literal.append(f"{col(positions[0])}1:{col(positions[0]+1)}1")
        for pos in rowmap[10]:
            projected.append(f"A{pos}:A{pos+1}")
        root.append(el("mergeCells",{"count":str(len(projected))}))
        for merge in projected:
            root[-1].append(el("mergeCell",{"ref":merge}))
        # Printing is now a report-relative rectangle; external rows/cols/cells are absent.
        for tag in ["printOptions","pageMargins","pageSetup"]:
            root.append(copy.deepcopy(template.find(f"{{{S}}}{tag}")))
        other = sheet({2:{2:"=Report!"+col(4)+str(rowmap[13][0])}} if leaves else {2:{2:"No intersections"}})
        # Keep second sheet report content only, also crop its control margins in this fixture.
        other_root = el("worksheet")
        other_root.append(el("sheetData"))
        other_root[-1].append(el("row",{"r":"1"}))
        other_root[-1][-1].append(cell("A1","=Report!D13" if leaves else "No intersections"))
        write_package(OUT/(case+"-pending.xlsx"),{"Report":root,"Summary":other_root},[("_xlnm.Print_Area",f"Report!$A$1:${col(len(columns))}$18",0),("_xlnm.Print_Area","Summary!$A$1",1)])
        (OUT/(case+"-requests.json")).write_text(json.dumps(requests))
        (OUT/(case+"-targets.json")).write_text(json.dumps(targets))
        assert len(rowplan)==18
        assert [c for c in root.iter(f"{{{S}}}c") if c.find(f"{{{S}}}is/{{{S}}}t") is not None and "DISPOSABLE" in (c.find(f"{{{S}}}is/{{{S}}}t").text or "")] == []
        assert sum(1 for _,(src,_,_) in enumerate(rowplan) if src==10)==3
        assert len(root.find(f"{{{S}}}cols"))==len(columns)
        company_labels=[c.find(f"{{{S}}}is/{{{S}}}t").text for c in refs(root).values() if c.find(f"{{{S}}}is/{{{S}}}t") is not None and c.find(f"{{{S}}}is/{{{S}}}t").text in ["Alpha","Beta"]]
        assert company_labels==["Alpha","Beta"][:companies]
        assert len(projected)==(companies if criteria else 0)+3
        result_log[case] = {"logicalRows":len(logical),"records":len(records),"leaves":leaves,"rows":18,"columns":len(columns),"parentMerges":projected[:companies] if criteria else [],"literalCopyMerges":literal,"zeroWidthColumns":[col(i) for i,(src,_,_) in enumerate(columns,1) if src==4],"formulas":len(requests)}
    assert result_log["full"]["parentMerges"] == ["C1:F1","G1:J1"]
    assert result_log["full"]["literalCopyMerges"] == ["C1:D1","G1:H1"]
    # Same controls can mean one four-row item or two adjacent two-row tables.
    ambiguous = ["$l(10).Items"]*4
    assert len(blocks(ambiguous))==1
    assert len(blocks(ambiguous[:2]+[None]+ambiguous[2:]))==2
    result_log["boundaryCounterexample"] = {"markers":ambiguous,"interpretations":["one four-row item","two two-row tables"],"proposedSeparatorProducesTables":2}
    # Recursively nested dimensions, including parent-only columns, need no depth-specific engine.
    def width(node):
        own,children,count = node
        return count*(own+sum(width(child) for child in children))
    assert width((1,[(1,[(1,[],4)],3)],2))==32
    result_log["threeLevelGeometry"] = {"counts":[2,3,4],"parentOnlyColumnsPerLevel":1,"physicalColumns":32}
    (OUT/"geometry-results.json").write_text(json.dumps(result_log,indent=2))
    print("Nested fixtures prepared: 2 companies x 0/1/3/7 criteria, zero companies; D:E/E, zero-width parent columns, cropping, marker-only boundary counterexample, three-level width 32.")

def finish():
    for case in ["full","one-criterion","seven-criteria","no-criteria","no-companies"]:
        for backend in (["internal"] if "--internal-only" in sys.argv else ["internal","closed"]):
            formulas = json.loads((OUT/f"{case}-{backend}-formulas.json").read_text())
            targets = json.loads((OUT/f"{case}-targets.json").read_text())
            with zipfile.ZipFile(OUT/f"{case}-pending.xlsx") as z, zipfile.ZipFile(OUT/f"{case}-{backend}.xlsx","w",zipfile.ZIP_DEFLATED) as dest:
                for name in z.namelist():
                    payload = z.read(name)
                    if name=="xl/worksheets/sheet1.xml":
                        root = E.fromstring(payload)
                        cells = refs(root)
                        assert len(formulas)==len(targets)
                        for address,formula in zip(targets,formulas):
                            cells[address].find(f"{{{S}}}f").text=formula
                        assert cells["A4"].find(f"{{{S}}}is/{{{S}}}t").text=="A"
                        assert cells["A15"].find(f"{{{S}}}is/{{{S}}}t").text=="X"
                        assert cells["A17"].find(f"{{{S}}}is/{{{S}}}t").text=="Z"
                        assert next(row for row in root.find(f"{{{S}}}sheetData") if row.get("r")=="14").get("ht")=="0"
                        assert next(row for row in root.find(f"{{{S}}}sheetData") if row.get("r")=="14").get("hidden")=="1"
                        for column in root.find(f"{{{S}}}cols"):
                            if column.get("width")=="0": assert column.get("hidden")=="1"
                        if case=="full":
                            assert cells["D3"].find(f"{{{S}}}f").text=="SUBTOTAL(9,D4:D7)"
                            assert cells["J13"].find(f"{{{S}}}f").text=="SUBTOTAL(9,J3:J12)"
                        payload = E.tostring(root)
                    dest.writestr(name,payload)
        if "--internal-only" not in sys.argv:
            assert json.loads((OUT/f"{case}-internal-formulas.json").read_text())==json.loads((OUT/f"{case}-closed-formulas.json").read_text())
    print("Nested geometry/value/formula assertions passed for selected backends, including ordinary rows under both empty axes and independent second table.")

def dates():
    serials = {}
    for system in [1900,1904]:
        date=dt.date(2026,10,20)
        serial=(date-dt.date(1899,12,31)).days+(date>=dt.date(1900,3,1)) if system==1900 else (date-dt.date(1904,1,1)).days
        serials[str(system)]=serial
        root=sheet({1:{2:serial,3:"=YEAR(B1)",4:"=MONTH(B1)",5:"=DAY(B1)",6:True},2:{2:"2026-10-20",3:"2026-10-20 formatted text"}})
        # Assert actual cell types independently of the generic fixture value helper.
        cells=refs(root)
        cells["F1"].set("t","b")
        cells["F1"].find(f"{{{S}}}v").text="1"
        cells["B1"].set("s","1")
        path=OUT/f"date-{system}.xlsx"
        write_package(path,{"Dates":root})
        with zipfile.ZipFile(path) as z:
            parts={n:z.read(n) for n in z.namelist()}
        wb=E.fromstring(parts["xl/workbook.xml"])
        wb.insert(0,el("workbookPr",{"date1904":"1" if system==1904 else "0"}))
        parts["xl/workbook.xml"]=E.tostring(wb)
        style=E.fromstring(parts["xl/styles.xml"])
        style.find(f"{{{S}}}numFmts")[0].set("formatCode","yyyy-mm-dd")
        parts["xl/styles.xml"]=E.tostring(style)
        with zipfile.ZipFile(path,"w",zipfile.ZIP_DEFLATED) as z:
            for name,payload in parts.items(): z.writestr(name,payload)
    assert serials["1900"]-serials["1904"]==1462
    (OUT/"date-results.json").write_text(json.dumps(serials,indent=2))
    print("Native serial date and Boolean fixtures prepared for 1900/1904 systems; serial difference 1462.")

def recalc():
    def values(path,sheetid=1):
        with zipfile.ZipFile(BASE/"recalculated"/path) as z:
            root=E.fromstring(z.read(f"xl/worksheets/sheet{sheetid}.xml"))
        return {key:c.find(f"{{{S}}}v").text for key,c in refs(root).items() if c.find(f"{{{S}}}v") is not None}
    for case in ["full","one-criterion","seven-criteria"]:
        for backend in (["internal"] if "--internal-only" in sys.argv else ["internal","closed"]):
            vals=values(f"{case}-{backend}.xlsx")
            expected={"D3":"32","D9":"31","D13":"63"}
            expected.update({"J3":"42","J13":"78"} if case=="full" else ({"F13":"66"} if case=="one-criterion" else {"R13":"102"}))
            for addr,value in expected.items():
                assert vals[addr]==value,(addr,vals[addr],value)
            assert values(f"{case}-{backend}.xlsx",2)["A1"]=="63"
    for system in [1900,1904]:
        vals=values(f"date-{system}.xlsx")
        assert [vals[c+"1"] for c in ["C","D","E","F"]]==["2026","10","20","1"]
    # LibreOffice may normalize stored dimensions; authored hidden states preserve zero visible size.
    for case in ["full","one-criterion","seven-criteria","no-criteria","no-companies"]:
        with zipfile.ZipFile(BASE/"recalculated"/f"{case}-internal.xlsx") as z:
            root=E.fromstring(z.read("xl/worksheets/sheet1.xml"))
        separator=next(row for row in root.iter(f"{{{S}}}row") if row.get("r")=="14")
        assert separator.get("hidden") in ["1","true"]
        geometry=json.loads((OUT/"geometry-results.json").read_text())[case]
        for address in geometry["zeroWidthColumns"]:
            _,index=xy(address+"1")
            column=next(c for c in root.find(f"{{{S}}}cols") if int(c.get("min"))<=index<=int(c.get("max")))
            assert column.get("hidden") in ["1","true"]
    print("Selected backends pass nested group/grand/cross-sheet totals; both native date systems, Boolean values and authored hidden states verified.")

if __name__=="__main__":
    {"prepare":prepare,"finish":finish,"dates":dates,"recalc":recalc}[sys.argv[1]]()

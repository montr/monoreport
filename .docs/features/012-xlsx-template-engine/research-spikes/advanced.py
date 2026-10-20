"""Attach representative unchanged metadata/drawings to the isolated compatibility fixture."""
import struct
import zlib
import zipfile
from pathlib import Path
from xml.etree import ElementTree as E

BASE = Path(__file__).resolve().parent / "evidence"
S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P = "http://schemas.openxmlformats.org/package/2006/relationships"
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
X = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"

def anchor(c, r):
    return f"<xdr:col>{c}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{r}</xdr:row><xdr:rowOff>0</xdr:rowOff>"

pic = f'''<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="Pixel"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="image1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>'''
chartframe = f'''<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="{C}"><c:chart r:id="chart1"/></a:graphicData></a:graphic></xdr:graphicFrame>'''
drawing = f'<xdr:wsDr xmlns:xdr="{X}" xmlns:a="{A}" xmlns:r="{R}" xmlns:c="{C}">'
for start, end, payload in [(anchor(1, 18), anchor(2, 19), pic), (anchor(3, 18), anchor(6, 23), chartframe)]:
    drawing += f'<xdr:twoCellAnchor><xdr:from>{start}</xdr:from><xdr:to>{end}</xdr:to>{payload}<xdr:clientData/></xdr:twoCellAnchor>'
drawing += '</xdr:wsDr>'
chart = f'''<c:chartSpace xmlns:c="{C}" xmlns:a="{A}"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Values</c:v></c:tx><c:cat><c:strRef><c:f>Report!$C$15:$C$16</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>Report!$B$15:$B$16</c:f></c:numRef></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea><c:plotVisOnly val="1"/></c:chart></c:chartSpace>'''
def chunk(kind, payload):
    return struct.pack(">I", len(payload)) + kind + payload + struct.pack(">I", zlib.crc32(kind + payload) & 0xffffffff)

pixel = b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)) + chunk(b"IDAT", zlib.compress(b"\x00\xff\xff\xff")) + chunk(b"IEND", b"")
parts = {
    "xl/drawings/drawing1.xml": drawing.encode(),
    "xl/drawings/_rels/drawing1.xml.rels": f'<Relationships xmlns="{P}"><Relationship Id="image1" Type="{R}/image" Target="../media/image1.png"/><Relationship Id="chart1" Type="{R}/chart" Target="../charts/chart1.xml"/></Relationships>'.encode(),
    "xl/media/image1.png": pixel,
    "xl/charts/chart1.xml": chart.encode(),
    "xl/comments1.xml": f'<comments xmlns="{S}"><authors><author>Research</author></authors><commentList><comment ref="B15" authorId="0"><text><t>Static note</t></text></comment></commentList></comments>'.encode(),
}
types = {
    "/xl/drawings/drawing1.xml": "application/vnd.openxmlformats-officedocument.drawing+xml",
    "/xl/charts/chart1.xml": "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
    "/xl/comments1.xml": "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
}
for filename in ["compatibility.xlsx", "compatibility-output.xlsx"]:
    with zipfile.ZipFile(BASE / filename) as z:
        content = {n: z.read(n) for n in z.namelist()}
    root = E.fromstring(content["xl/worksheets/sheet1.xml"])
    root.insert(len(root) - 1, E.Element("{" + S + "}drawing", {"{" + R + "}id": "drawing1"}))
    content["xl/worksheets/sheet1.xml"] = E.tostring(root)
    root = E.fromstring(content["xl/worksheets/_rels/sheet1.xml.rels"])
    for rid, kind, target in [("drawing1", "drawing", "../drawings/drawing1.xml"), ("comments1", "comments", "../comments1.xml")]:
        root.append(E.Element("{" + P + "}Relationship", {"Id": rid, "Type": R + "/" + kind, "Target": target}))
    E.register_namespace("", P)
    content["xl/worksheets/_rels/sheet1.xml.rels"] = E.tostring(root)
    root = E.fromstring(content["[Content_Types].xml"])
    t = "http://schemas.openxmlformats.org/package/2006/content-types"
    root.insert(2, E.Element("{" + t + "}Default", {"Extension": "png", "ContentType": "image/png"}))
    for path, kind in types.items():
        root.append(E.Element("{" + t + "}Override", {"PartName": path, "ContentType": kind}))
    E.register_namespace("", t)
    content["[Content_Types].xml"] = E.tostring(root)
    content.update(parts)
    with zipfile.ZipFile(BASE / ("advanced-" + filename), "w", zipfile.ZIP_DEFLATED) as z:
        for name, payload in content.items():
            z.writestr(name, payload)
with zipfile.ZipFile(BASE / "advanced-compatibility.xlsx") as a, zipfile.ZipFile(BASE / "advanced-compatibility-output.xlsx") as b:
    for name in a.namelist():
        if name != "xl/worksheets/sheet1.xml":
            assert a.read(name) == b.read(name), name
print("Static image/chart/comment relationships, parts and anchors preserved byte-for-byte.")
with zipfile.ZipFile(BASE / "advanced-compatibility.xlsx") as z, zipfile.ZipFile(BASE / "unsafe-shift.xlsx", "w", zipfile.ZIP_DEFLATED) as result:
    for name in z.namelist():
        payload = z.read(name)
        if name == "xl/worksheets/sheet1.xml":
            root = E.fromstring(payload)
            for row in root.find("{" + S + "}sheetData"):
                number = int(row.get("r"))
                if number >= 14:
                    row.set("r", str(number + 3))
                    for cell in row:
                        old = cell.get("r")
                        cell.set("r", ''.join(ch for ch in old if ch.isalpha()) + str(number + 3))
            payload = E.tostring(root)
        result.writestr(name, payload)
print("Created intentionally unsafe-shift.xlsx: rows moved, table/name/CF/DV/chart/comment references deliberately stale.")

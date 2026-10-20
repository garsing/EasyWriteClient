# -*- coding: utf-8 -*-
import zipfile
from pathlib import Path

src = Path(r"C:\Users\garsing\AppData\Local\Temp\ew-xlchart-last.xlsx")
raw = zipfile.ZipFile(src).read("xl/charts/chart1.xml")
text = raw.decode("utf-8-sig")


def repack(new_text, dest):
    blob = new_text.encode("utf-8")
    dest = Path(dest)
    tmp = dest.with_suffix(".zip")
    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(tmp, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/charts/chart1.xml":
                data = blob
            zout.writestr(item.filename, data)
    if dest.exists():
        dest.unlink()
    tmp.replace(dest)


repack(text, r"C:\Users\garsing\AppData\Local\Temp\ew-a-nobom.xlsx")

NL = "\r\n"
old = (
    '<c:axId val="1833562384" />' + NL
    + '        <c:axId val="1833579184" />' + NL
    + "      </c:lineChart>"
)
new = (
    '<c:axId val="1833562384" />' + NL
    + '        <c:axId val="1833579185" />' + NL
    + "      </c:lineChart>"
)
t = text
if old not in t:
    print("AX PATTERN MISSING")
else:
    t = t.replace(old, new, 1)
sec = (
    "<c:valAx>" + NL
    + '        <c:axId val="1833579185" />' + NL
    + "        <c:scaling>" + NL
    + '          <c:orientation val="minMax" />' + NL
    + "        </c:scaling>" + NL
    + '        <c:axPos val="r" />' + NL
    + '        <c:crossAx val="1833562384" />' + NL
    + '        <c:crosses val="max" />' + NL
    + "      </c:valAx>"
)
sec2 = (
    "<c:valAx>" + NL
    + '        <c:axId val="1833579185" />' + NL
    + "        <c:scaling>" + NL
    + '          <c:orientation val="minMax" />' + NL
    + "        </c:scaling>" + NL
    + '        <c:delete val="0" />' + NL
    + '        <c:axPos val="r" />' + NL
    + '        <c:numFmt formatCode="General" sourceLinked="1" />' + NL
    + '        <c:majorTickMark val="out" />' + NL
    + '        <c:minorTickMark val="none" />' + NL
    + '        <c:tickLblPos val="nextTo" />' + NL
    + '        <c:crossAx val="1833562384" />' + NL
    + '        <c:crosses val="max" />' + NL
    + '        <c:crossBetween val="between" />' + NL
    + "      </c:valAx>"
)
if sec not in t:
    print("SEC PATTERN MISSING")
else:
    t = t.replace(sec, sec2, 1)
repack(t, r"C:\Users\garsing\AppData\Local\Temp\ew-b-axis.xlsx")

idx = text.find("00000004-B281")
Path(r"d:\project\EasyWrite\EasyWriteClient\Test\overnight-logs\xlg550\chart-slice.txt").write_text(
    text[idx - 500 : idx + 900], encoding="utf-8"
)
pos = text.find("00000004-B281")
ext_start = text.rfind("<c:extLst>", 0, pos)
ext_end = text.find("</c:extLst>", pos) + len("</c:extLst>")
sp_start = text.find("<c:spPr>", pos)
ser_end = text.find("</c:ser>", pos)
ext = text[ext_start:ext_end]
tail = text[sp_start:ser_end]
ordered = text[:ext_start] + tail + ext + text[ser_end:]
repack(ordered, r"C:\Users\garsing\AppData\Local\Temp\ew-c-order.xlsx")
repack(t[: t.find("00000004-B281")] + "unused", r"C:\Users\garsing\AppData\Local\Temp\ew-skip.xlsx") if False else None
# axis + order
pos = t.find("00000004-B281")
ext_start = t.rfind("<c:extLst>", 0, pos)
ext_end = t.find("</c:extLst>", pos) + len("</c:extLst>")
sp_start = t.find("<c:spPr>", pos)
ser_end = t.find("</c:ser>", pos)
both = t[:ext_start] + t[sp_start:ser_end] + t[ext_start:ext_end] + t[ser_end:]
repack(both, r"C:\Users\garsing\AppData\Local\Temp\ew-d-both.xlsx")
only_id = text.replace(old, new, 1)
repack(only_id, r"C:\Users\garsing\AppData\Local\Temp\ew-e-idonly.xlsx")
only_fill = text.replace(sec, sec2, 1)
repack(only_fill, r"C:\Users\garsing\AppData\Local\Temp\ew-f-fillonly.xlsx")
print("ax_ok", old in text, "sec_ok", sec in text, "idx", idx)

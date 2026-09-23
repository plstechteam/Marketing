"""
Replace one worksheet inside an existing .xlsx without touching anything else.

The workbook in SharePoint is shared: people add their own tabs next to
Deals (the first was "Maz"), with formulas, charts, pivots and formatting
this job knows nothing about. Rebuilding the file — or loading and saving it
with a library such as openpyxl, which drops charts and images it cannot
read — would destroy that work.

So the Deals sheet is swapped at the package level instead. An .xlsx is a zip
of XML parts; this module writes a fresh XML part for Deals and copies every
other part across byte for byte. Only these parts change:

  xl/worksheets/<Deals>.xml   the new data
  xl/styles.xml               two date formats appended, once (then reused)
  xl/workbook.xml             fullCalcOnLoad, so formulas that read Deals
                              recalculate when the file is next opened
  xl/calcChain.xml            removed (with its two references) — Excel
                              rebuilds it; a stale one can trigger "repair"

Text is written as inline strings, so the shared-strings table other sheets
use is never touched either.

Nothing here parses and re-serialises XML with a generic library: that would
rename namespace prefixes that attributes such as mc:Ignorable depend on, and
Excel would reject the file. Edits are small, targeted string operations.
"""

import io
import math
import re
import zipfile
from datetime import date, datetime
from xml.sax.saxutils import escape

import numpy as np

DATETIME_FORMAT = "yyyy-mm-dd hh:mm:ss"
DATE_FORMAT = "yyyy-mm-dd"

# Characters XML 1.0 forbids; a stray one in a HubSpot note would make the
# whole part unreadable.
_ILLEGAL_XML = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]")
_EXCEL_EPOCH = datetime(1899, 12, 30)
_MAX_CELL_TEXT = 32767


class SpliceError(Exception):
    """The workbook is not in a shape this module can edit safely."""


def column_letter(index):
    """0 -> A, 25 -> Z, 26 -> AA."""
    letters = ""
    index += 1
    while index:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _serial(value):
    """date / datetime -> Excel serial number."""
    if isinstance(value, datetime):
        delta = value - _EXCEL_EPOCH
    else:
        delta = datetime(value.year, value.month, value.day) - _EXCEL_EPOCH
    return delta.days + delta.seconds / 86400 + delta.microseconds / 86400e6


def _text(value):
    return escape(_ILLEGAL_XML.sub("", str(value))[:_MAX_CELL_TEXT])


def sheet_xml(df, s_datetime, s_date):
    """The worksheet part for df: a header row, then one row per record.

    s_datetime / s_date are the cellXfs indexes that carry the two date
    formats in the target workbook's styles.xml.
    """
    ncols = len(df.columns)
    last = f"{column_letter(max(ncols - 1, 0))}{len(df) + 1}"
    letters = [column_letter(i) for i in range(ncols)]
    out = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<dimension ref="A1:{last}"/>'
        '<sheetViews><sheetView workbookViewId="0">'
        '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        '</sheetView></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/><sheetData>'
    ]
    out.append('<row r="1">')
    for c, name in enumerate(df.columns):
        out.append(f'<c r="{letters[c]}1" t="inlineStr"><is><t>{_text(name)}</t></is></c>')
    out.append("</row>")

    values = df.astype(object).where(df.notna(), None)
    for r, row in enumerate(values.itertuples(index=False, name=None), start=2):
        out.append(f'<row r="{r}">')
        for c, v in enumerate(row):
            if v is None:
                continue
            ref = f"{letters[c]}{r}"
            if isinstance(v, datetime):
                out.append(f'<c r="{ref}" s="{s_datetime}"><v>{_serial(v)!r}</v></c>')
            elif isinstance(v, date):
                out.append(f'<c r="{ref}" s="{s_date}"><v>{_serial(v)!r}</v></c>')
            elif isinstance(v, (bool, np.bool_)):
                out.append(f'<c r="{ref}" t="b"><v>{int(bool(v))}</v></c>')
            elif isinstance(v, (int, np.integer)):
                out.append(f'<c r="{ref}"><v>{int(v)}</v></c>')
            elif isinstance(v, (float, np.floating)):
                if math.isnan(v) or math.isinf(v):
                    continue
                out.append(f'<c r="{ref}"><v>{float(v)!r}</v></c>')
            else:
                text = _text(v)
                space = ' xml:space="preserve"' if text != text.strip() else ""
                out.append(f'<c r="{ref}" t="inlineStr"><is><t{space}>{text}</t></is></c>')
        out.append("</row>")
    out.append("</sheetData>"
               '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
               "</worksheet>")
    return "".join(out).encode("utf-8")


def _attrs(tag):
    return dict(re.findall(r'([\w:]+)="([^"]*)"', tag))


def ensure_date_styles(styles):
    """Return (styles_xml, s_datetime, s_date), reusing matching formats and
    cell styles when present so repeated runs do not grow the file."""
    xf_indexes = []
    for code in (DATETIME_FORMAT, DATE_FORMAT):
        fmt_id = None
        for tag in re.findall(r"<numFmt\b[^>]*/>", styles):
            a = _attrs(tag)
            if a.get("formatCode") == code:
                fmt_id = int(a["numFmtId"])
                break
        if fmt_id is None:
            ids = [int(i) for i in re.findall(r'<numFmt\b[^>]*numFmtId="(\d+)"', styles)]
            fmt_id = max(ids + [163]) + 1
            tag = f'<numFmt numFmtId="{fmt_id}" formatCode="{code}"/>'
            m = re.search(r"<numFmts\b[^>]*>", styles)
            if m:
                close = styles.index("</numFmts>", m.end())
                styles = styles[:close] + tag + styles[close:]
                styles = _bump_count(styles, "numFmts")
            else:
                m = re.search(r"<styleSheet\b[^>]*>", styles)
                if not m:
                    raise SpliceError("styles.xml has no <styleSheet>")
                styles = styles[:m.end()] + f'<numFmts count="1">{tag}</numFmts>' + styles[m.end():]

        m = re.search(r"<cellXfs\b[^>]*>(.*?)</cellXfs>", styles, re.S)
        if not m:
            raise SpliceError("styles.xml has no <cellXfs>")
        xfs = re.findall(r"<xf\b[^>]*?(?:/>|>.*?</xf>)", m.group(1), re.S)
        found = None
        for i, xf in enumerate(xfs):
            a = _attrs(xf[: xf.index(">") + 1])
            if (a.get("numFmtId") == str(fmt_id) and a.get("fontId", "0") == "0"
                    and a.get("fillId", "0") == "0" and a.get("borderId", "0") == "0"
                    and a.get("applyNumberFormat") == "1" and "<alignment" not in xf
                    and "<protection" not in xf):
                found = i
                break
        if found is None:
            new = (f'<xf numFmtId="{fmt_id}" fontId="0" fillId="0" borderId="0" '
                   f'xfId="0" applyNumberFormat="1"/>')
            close = m.end(1)
            styles = styles[:close] + new + styles[close:]
            styles = _bump_count(styles, "cellXfs")
            found = len(xfs)
        xf_indexes.append(found)
    return styles, xf_indexes[0], xf_indexes[1]


def _bump_count(xml, element):
    def inc(m):
        return f'{m.group(1)}{int(m.group(2)) + 1}"'
    return re.sub(rf'(<{element}\b[^>]*\bcount=")(\d+)"', inc, xml, count=1)


def _sheet_part(workbook, rels, sheet_name):
    """Zip path of the worksheet part named sheet_name."""
    rid = None
    for tag in re.findall(r"<sheet\b[^>]*/?>", workbook):
        a = _attrs(tag)
        if a.get("name") == escape(sheet_name, {'"': "&quot;"}):
            rid = next((v for k, v in a.items() if k.endswith(":id")), None)
            break
    if rid is None:
        raise SpliceError(f"no sheet named '{sheet_name}' in the workbook")
    for tag in re.findall(r"<Relationship\b[^>]*/?>", rels):
        a = _attrs(tag)
        if a.get("Id") == rid:
            target = a["Target"]
            return target.lstrip("/") if target.startswith("/") else "xl/" + target
    raise SpliceError(f"sheet '{sheet_name}' points at a relationship ({rid}) that does not exist")


def replace_sheet(xlsx_bytes, sheet_name, df):
    """Return (new_xlsx_bytes, report) with sheet_name's data replaced by df.

    report lists the parts changed and the parts copied untouched, so a run
    can prove nothing else moved.
    """
    src = zipfile.ZipFile(io.BytesIO(xlsx_bytes))
    names = src.namelist()
    workbook = src.read("xl/workbook.xml").decode("utf-8")
    rels = src.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    types = src.read("[Content_Types].xml").decode("utf-8")
    styles = src.read("xl/styles.xml").decode("utf-8")

    part = _sheet_part(workbook, rels, sheet_name)
    if part not in names:
        raise SpliceError(f"sheet part {part} is missing from the package")
    sheet_rels = part.replace("worksheets/", "worksheets/_rels/") + ".rels"
    if sheet_rels in names:
        # A table, drawing or comment attached to Deals itself would be left
        # pointing at cells that moved. Deals is code-owned and has none; if
        # one appears, stop rather than guess.
        raise SpliceError(f"'{sheet_name}' has attached objects ({sheet_rels}); not editing it")

    styles, s_datetime, s_date = ensure_date_styles(styles)
    new_sheet = sheet_xml(df, s_datetime, s_date)

    changed = {part: new_sheet, "xl/styles.xml": styles.encode("utf-8")}
    dropped = set()

    if "xl/calcChain.xml" in names:
        dropped.add("xl/calcChain.xml")
        rels = re.sub(r'<Relationship\b[^>]*Target="[^"]*calcChain\.xml"[^>]*/>', "", rels)
        types = re.sub(r'<Override\b[^>]*PartName="/xl/calcChain\.xml"[^>]*/>', "", types)
        changed["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")
        changed["[Content_Types].xml"] = types.encode("utf-8")

    m = re.search(r"<calcPr\b[^>]*?/?>", workbook)
    if m:
        tag = m.group(0)
        if "fullCalcOnLoad=" not in tag:
            end = -2 if tag.endswith("/>") else -1
            workbook = workbook.replace(tag, tag[:end] + ' fullCalcOnLoad="1"' + tag[end:], 1)
    else:
        anchor = "</definedNames>" if "</definedNames>" in workbook else "</sheets>"
        workbook = workbook.replace(anchor, anchor + '<calcPr fullCalcOnLoad="1"/>', 1)
    changed["xl/workbook.xml"] = workbook.encode("utf-8")

    out = io.BytesIO()
    untouched = []
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            if info.filename in dropped:
                continue
            if info.filename in changed:
                new_info = zipfile.ZipInfo(info.filename, date_time=info.date_time)
                new_info.compress_type = zipfile.ZIP_DEFLATED
                dst.writestr(new_info, changed[info.filename])
            else:
                dst.writestr(info, src.read(info.filename))
                untouched.append(info.filename)
    report = {
        "sheet_part": part,
        "changed": sorted(changed),
        "dropped": sorted(dropped),
        "untouched": untouched,
        "sheets": [_attrs(t).get("name") for t in re.findall(r"<sheet\b[^>]*/?>", workbook)],
    }
    return out.getvalue(), report


def untouched_parts_identical(before, after, report):
    """True when every part the report calls untouched is byte-identical."""
    a = zipfile.ZipFile(io.BytesIO(before))
    b = zipfile.ZipFile(io.BytesIO(after))
    return all(a.read(n) == b.read(n) for n in report["untouched"])


def sheet_row_count(xlsx_bytes, part):
    """Rows (header included) in a worksheet part — a cheap post-check."""
    with zipfile.ZipFile(io.BytesIO(xlsx_bytes)) as z:
        return z.read(part).count(b"<row ")


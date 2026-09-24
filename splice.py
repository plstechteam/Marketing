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
    """The worksheet part for df, as one bytes object (tests, small sheets)."""
    return b"".join(sheet_xml_chunks(df, s_datetime, s_date))


def sheet_xml_chunks(df, s_datetime, s_date, rows_per_chunk=2000):
    """The worksheet part for df, yielded in pieces: a header row, then one
    row per record.

    Streamed rather than built whole: the full history is ~230,000 rows and
    the finished XML several hundred MB, which is written straight into the
    zip a chunk at a time instead of held in memory.

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

    yield "".join(out).encode("utf-8")
    out = []

    values = df.astype(object).where(df.notna(), None)
    for r, row in enumerate(values.itertuples(index=False, name=None), start=2):
        if (r - 2) % rows_per_chunk == 0 and out:
            yield "".join(out).encode("utf-8")
            out = []
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
    yield "".join(out).encode("utf-8")


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
    changed = {part: lambda: sheet_xml_chunks(df, s_datetime, s_date),
               "xl/styles.xml": styles.encode("utf-8")}
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
                body = changed[info.filename]
                if callable(body):
                    # No force_zip64: some Excel versions misread Zip64
                    # entries, and the sheet stays far below the 2 GB where
                    # zipfile would need it (it raises there, and the run
                    # fails without writing).
                    with dst.open(new_info, "w") as f:
                        for chunk in body():
                            f.write(chunk)
                else:
                    dst.writestr(new_info, body)
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



# ----------------------------------------------------------------------
# Formulas in other tabs that read the sheet this job owns
# ----------------------------------------------------------------------
# Formulas point at Deals by column letter (Deals!AE:AE, 'Deals'!$C$2...). A
# column inserted in the middle of Deals shifts every letter to its right, and
# those formulas silently start reading a different column. So before writing,
# a run finds every column letter another tab references and checks the header
# under it is unchanged.

_REF = re.compile(r"(?:'([^']+)'|([A-Za-z_][\w.]*))!(\$?[A-Z]{1,3}\$?\d*(?::\$?[A-Z]{1,3}\$?\d*)?)")


def column_index(letters):
    """A -> 0, Z -> 25, AA -> 26."""
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def _columns_in_ref(ref):
    """Column letters a reference covers: 'AE:AE' -> {AE}, 'A2:C9' -> {A,B,C}."""
    parts = [re.sub(r"[\$\d]", "", p) for p in ref.split(":")]
    parts = [p for p in parts if p]
    if not parts:
        return set()
    lo, hi = column_index(parts[0]), column_index(parts[-1])
    if hi - lo > 200:           # a whole-row-style reference; not a column pin
        return set()
    return {column_letter(i) for i in range(min(lo, hi), max(lo, hi) + 1)}


def referenced_columns(xlsx_bytes, sheet_name):
    """{column letter: sorted list of 'Tab!Cell' formulas that read it} for
    every formula in every other tab (and defined names) that points at
    sheet_name by column letter."""
    import html
    z = zipfile.ZipFile(io.BytesIO(xlsx_bytes))
    workbook = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    own = _sheet_part(workbook, rels, sheet_name)
    names = {}
    for tag in re.findall(r"<sheet\b[^>]*/?>", workbook):
        a = _attrs(tag)
        rid = next((v for k, v in a.items() if k.endswith(":id")), None)
        for rtag in re.findall(r"<Relationship\b[^>]*/?>", rels):
            ra = _attrs(rtag)
            if ra.get("Id") == rid:
                t = ra["Target"]
                names[t.lstrip("/") if t.startswith("/") else "xl/" + t] = html.unescape(a.get("name", ""))
    found = {}

    def scan(formula, where):
        for quoted, bare, ref in _REF.findall(html.unescape(formula)):
            if (quoted or bare) == sheet_name:
                for col in _columns_in_ref(ref):
                    found.setdefault(col, set()).add(where)

    for part, tab in names.items():
        if part == own or part not in z.namelist():
            continue
        xml = z.read(part).decode("utf-8", "replace")
        for m in re.finditer(r'<c\b[^>]*\br="([A-Z]+\d+)"[^>]*>(.*?)</c>', xml, re.S):
            for f in re.findall(r"<f\b[^>]*>(.*?)</f>", m.group(2), re.S):
                scan(f, f"{tab}!{m.group(1)}")
    for m in re.finditer(r'<definedName\b[^>]*name="([^"]*)"[^>]*>(.*?)</definedName>', workbook, re.S):
        scan(m.group(2), f"name {m.group(1)}")
    for part in z.namelist():
        if part.startswith("xl/charts/") and part.endswith(".xml"):
            xml = z.read(part).decode("utf-8", "replace")
            for f in re.findall(r"<c:f>(.*?)</c:f>", xml, re.S):
                scan(f, f"chart {part.rsplit('/', 1)[-1]}")
    return {col: sorted(w) for col, w in found.items()}


def header_row(xlsx_bytes, sheet_name):
    """The first row of sheet_name as {column letter: text}."""
    import html
    z = zipfile.ZipFile(io.BytesIO(xlsx_bytes))
    workbook = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    part = _sheet_part(workbook, rels, sheet_name)
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        sst = z.read("xl/sharedStrings.xml").decode("utf-8")
        shared = ["".join(re.findall(r"<t\b[^>]*>(.*?)</t>", si, re.S))
                  for si in re.findall(r"<si>(.*?)</si>", sst, re.S)]
    with z.open(part) as f:
        head = b""
        while b"</row>" not in head:
            chunk = f.read(65536)
            if not chunk:
                break
            head += chunk
    row = re.search(r"<row\b[^>]*>(.*?)</row>", head.decode("utf-8", "replace"), re.S)
    out = {}
    if not row:
        return out
    for m in re.finditer(r'<c\b([^>]*)>(.*?)</c>', row.group(1), re.S):
        a = _attrs("<c" + m.group(1) + ">")
        col = re.sub(r"\d", "", a.get("r", ""))
        body = m.group(2)
        if a.get("t") == "s":
            v = re.search(r"<v>(\d+)</v>", body)
            text = shared[int(v.group(1))] if v and int(v.group(1)) < len(shared) else ""
        else:
            text = "".join(re.findall(r"<t\b[^>]*>(.*?)</t>", body, re.S)) or \
                "".join(re.findall(r"<v>(.*?)</v>", body, re.S))
        out[col] = html.unescape(text)
    return out


def shifted_references(before, after, sheet_name):
    """[(column, tabs/cells, header before, header after)] for every column
    another tab reads whose header would change between the two files."""
    refs = referenced_columns(before, sheet_name)
    old, new = header_row(before, sheet_name), header_row(after, sheet_name)
    return [(col, where, old.get(col), new.get(col))
            for col, where in sorted(refs.items(), key=lambda kv: column_index(kv[0]))
            if old.get(col) != new.get(col)]

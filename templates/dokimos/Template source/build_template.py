"""Builds the Dokimos design-calculation template package (and an example filled with the
user's soil-stress calculation). Word then finishes it: logo, contents, .dotx."""
import os
import re
import sys
import uuid
import zipfile
from xml.sax.saxutils import escape

HERE = os.path.dirname(os.path.abspath(__file__))
SOURCE = os.path.join(HERE, "unpacked-1")
OUT = os.path.join(HERE, "out")
os.makedirs(OUT, exist_ok=True)

# ---- Brand (from dokimos.net) ------------------------------------------------------------
INK, INK_SOFT, INK_FAINT = "1C1710", "5B5648", "8A8574"
PAPER, PAPER_ALT, LINE = "F7F4EE", "EDE8DC", "D9D3C3"
BLUE, GREEN, BROWN, RED = "2E63C4", "2FA968", "3A2D00", "B3261E"
OK_TEXT = "1E7A47"  # the brand green, darkened for text on white
HEAD, BODY, LIGHT = "Libre Franklin SemiBold", "Libre Franklin", "Libre Franklin Light"  # the user's choice
COMPANY, WEBSITE = "Dokimos Consulting Engineers", "dokimos.net"

# A4, 0.6 in margins
PAGE_W, PAGE_H, MARGIN = 11906, 16838, 864
TEXT_W = PAGE_W - 2 * MARGIN  # 10178

DATA_NS = "urn:dokimos-net:calculation"
DATA_ID = "{5C7C1F2A-8E3B-4D2A-9B1E-6A0D2C4F7B11}"
CORE_ID = "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}"
FIELDS = ["project", "projectNo", "client", "location", "calcNo", "revision", "date",
          "codes", "preparedBy", "checkedBy", "approvedBy"]

_ids = iter(range(1001, 99999))
DOCPARTS = {}  # name -> (placeholder text, paragraph style or None)


def x(s):
    return escape(s, {'"': "&quot;"})


# ---- Runs and paragraphs ---------------------------------------------------------------------

def rpr(style=None, font=None, size=None, color=None, bold=False, italic=False, caps=False, spacing=None):
    parts = []
    if style: parts.append(f'<w:rStyle w:val="{style}"/>')
    if font: parts.append(f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>')
    if bold: parts.append("<w:b/><w:bCs/>")
    if italic: parts.append("<w:i/><w:iCs/>")
    if caps: parts.append("<w:caps/>")
    if color: parts.append(f'<w:color w:val="{color}"/>')
    if spacing is not None: parts.append(f'<w:spacing w:val="{spacing}"/>')
    if size: parts.append(f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>')
    return f"<w:rPr>{''.join(parts)}</w:rPr>" if parts else ""


def run(text, **fmt):
    return f'<w:r>{rpr(**fmt)}<w:t xml:space="preserve">{x(text)}</w:t></w:r>'


def tab(**fmt):
    return f"<w:r>{rpr(**fmt)}<w:tab/></w:r>"


def p(content="", style=None, ppr=""):
    props = (f'<w:pStyle w:val="{style}"/>' if style else "") + ppr
    return f"<w:p>{'<w:pPr>' + props + '</w:pPr>' if props else ''}{content}</w:p>"


def field(instr, result, **fmt):
    r = rpr(**fmt)
    return (f'<w:r>{r}<w:fldChar w:fldCharType="begin"/></w:r>'
            f'<w:r>{r}<w:instrText xml:space="preserve"> {x(instr)} </w:instrText></w:r>'
            f'<w:r>{r}<w:fldChar w:fldCharType="separate"/></w:r>'
            f'<w:r>{r}<w:t>{x(result)}</w:t></w:r>'
            f'<w:r>{r}<w:fldChar w:fldCharType="end"/></w:r>')


def placeholder_part(name, text, style=None):
    DOCPARTS[name] = (text, style)
    return f'<w:placeholder><w:docPart w:val="{name}"/></w:placeholder>'


def bound(field_name, text, value=None, date=False, **fmt):
    """An inline control bound to the calculation details (filled once, shown everywhere)."""
    name = "Dokimos_" + field_name
    placeholder = placeholder_part(name, text)
    binding = (f'<w:dataBinding w:prefixMappings="xmlns:ns0=\'{DATA_NS}\'" '
               f'w:xpath="/ns0:calculation[1]/ns0:{field_name}[1]" w:storeItemID="{DATA_ID}"/>')
    kind = ('<w:date><w:dateFormat w:val="d MMMM yyyy"/><w:lid w:val="en-GB"/>'
            '<w:storeMappedDataAs w:val="dateTime"/><w:calendar w:val="gregorian"/></w:date>') if date else "<w:text/>"
    shown = run(value, **fmt) if value else run(text, style="PlaceholderText", **{k: v for k, v in fmt.items() if k not in ("style", "color")})
    return (f'<w:sdt><w:sdtPr>{rpr(**fmt)}<w:alias w:val="{field_name}"/><w:tag w:val="dokimos:{field_name}"/>'
            f'<w:id w:val="{next(_ids)}"/>{placeholder}{"" if value else "<w:showingPlcHdr/>"}{binding}{kind}</w:sdtPr>'
            f'<w:sdtContent>{shown}</w:sdtContent></w:sdt>')


def core(prop, text, value=None, **fmt):
    """An inline control bound to the document's own Title or Subject property."""
    name = "Dokimos_core_" + prop
    placeholder = placeholder_part(name, text)
    binding = (f'<w:dataBinding w:prefixMappings="xmlns:ns0=\'http://purl.org/dc/elements/1.1/\' '
               f'xmlns:ns1=\'http://schemas.openxmlformats.org/package/2006/metadata/core-properties\'" '
               f'w:xpath="/ns1:coreProperties[1]/ns0:{prop}[1]" w:storeItemID="{CORE_ID}"/>')
    shown = run(value, **fmt) if value else run(text, style="PlaceholderText", **{k: v for k, v in fmt.items() if k not in ("style", "color")})
    return (f'<w:sdt><w:sdtPr>{rpr(**fmt)}<w:alias w:val="{prop.capitalize()}"/><w:tag w:val="dokimos:{prop}"/>'
            f'<w:id w:val="{next(_ids)}"/>{placeholder}{"" if value else "<w:showingPlcHdr/>"}{binding}<w:text/></w:sdtPr>'
            f'<w:sdtContent>{shown}</w:sdtContent></w:sdt>')


def prompt(name, text, style=None):
    """Guidance that disappears when you type (a temporary placeholder)."""
    placeholder = placeholder_part(name, text, style)
    return (f'<w:sdt><w:sdtPr><w:id w:val="{next(_ids)}"/>{placeholder}<w:temporary/><w:showingPlcHdr/></w:sdtPr>'
            f'<w:sdtContent>{p(run(text, style="PlaceholderText"), style)}</w:sdtContent></w:sdt>')


def inline_prompt(name, text, **fmt):
    placeholder = placeholder_part(name, text)
    return (f'<w:sdt><w:sdtPr>{rpr(**fmt)}<w:id w:val="{next(_ids)}"/>{placeholder}<w:temporary/><w:showingPlcHdr/><w:text/></w:sdtPr>'
            f'<w:sdtContent>{run(text, style="PlaceholderText")}</w:sdtContent></w:sdt>')


# ---- Tables ------------------------------------------------------------------------------------

def border(side, size, color, val="single"):
    return f'<w:{side} w:val="{val}" w:sz="{size}" w:space="0" w:color="{color}"/>'


def cell(content, width, shade=None, vmerge=None, valign=None, borders="", margins="", span=None):
    props = f'<w:tcW w:w="{width}" w:type="dxa"/>'
    if span: props += f'<w:gridSpan w:val="{span}"/>'
    if vmerge: props += f'<w:vMerge w:val="restart"/>' if vmerge == "restart" else "<w:vMerge/>"
    if borders: props += f"<w:tcBorders>{borders}</w:tcBorders>"
    if shade: props += f'<w:shd w:val="clear" w:color="auto" w:fill="{shade}"/>'
    if margins: props += f"<w:tcMar>{margins}</w:tcMar>"
    if valign: props += f'<w:vAlign w:val="{valign}"/>'
    return f"<w:tc><w:tcPr>{props}</w:tcPr>{content or p()}</w:tc>"


def table(rows, widths, borders="", style=None, margins=None, layout_fixed=True, look=None, indent=0, header_rows=0,
          at_page_bottom=False):
    props = f'<w:tblStyle w:val="{style}"/>' if style else ""
    if at_page_bottom:  # floats at the bottom of the text area, whatever is above it
        props += ('<w:tblpPr w:leftFromText="0" w:rightFromText="0" w:topFromText="0" w:bottomFromText="0" '
                  'w:vertAnchor="margin" w:horzAnchor="margin" w:tblpXSpec="left" w:tblpYSpec="bottom"/><w:tblOverlap w:val="never"/>')
    props += f'<w:tblW w:w="{sum(widths)}" w:type="dxa"/>'
    if indent: props += f'<w:tblInd w:w="{indent}" w:type="dxa"/>'
    if borders: props += f"<w:tblBorders>{borders}</w:tblBorders>"
    if layout_fixed: props += '<w:tblLayout w:type="fixed"/>'
    m = margins or (40, 90, 40, 90)
    props += (f'<w:tblCellMar><w:top w:w="{m[0]}" w:type="dxa"/><w:left w:w="{m[1]}" w:type="dxa"/>'
              f'<w:bottom w:w="{m[2]}" w:type="dxa"/><w:right w:w="{m[3]}" w:type="dxa"/></w:tblCellMar>')
    props += f'<w:tblLook w:val="{look or "04A0"}" w:firstRow="{1 if header_rows else 0}" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    body = ""
    for i, (cells, height) in enumerate(rows):
        trpr = ""
        if height: trpr += f'<w:trHeight w:val="{abs(height)}" w:hRule="{"exact" if height < 0 else "atLeast"}"/>'
        if i < header_rows: trpr += "<w:tblHeader/>"
        body += f"<w:tr>{'<w:trPr>' + trpr + '</w:trPr>' if trpr else ''}{''.join(cells)}</w:tr>"
    return f"<w:tbl><w:tblPr>{props}</w:tblPr><w:tblGrid>{grid}</w:tblGrid>{body}</w:tbl>"


NOBORDERS = "".join(border(s, 0, "auto", "nil") for s in ("top", "left", "bottom", "right", "insideH", "insideV"))


def label(text):
    return p(run(text, style="FieldLabel"), "TitleBlock")


def brand(logo_marker, word_size, sub_size, logo_w, text_w, sub_spacing):
    """Logo mark beside DOKIMOS / CONSULTING ENGINEERS, as on dokimos.net."""
    left = p(run(logo_marker), "TitleBlock", '<w:jc w:val="left"/>')
    right = (p(run("DOKIMOS", font=BODY, size=word_size, color=INK, spacing=6, bold=True), "TitleBlock")
             + p(run("CONSULTING ENGINEERS", font=HEAD, size=sub_size, color=INK_SOFT, caps=True, spacing=sub_spacing), "TitleBlock"))
    return table([([cell(left, logo_w, valign="center"), cell(right, text_w, valign="center")], None)],
                 [logo_w, text_w], NOBORDERS, margins=(0, 0, 0, 60))


# ---- Header and footer -------------------------------------------------------------------------

def sheet_number(numbered, **fmt):
    """Front matter (contents) shows just its roman page number; the body shows "1 of 12", counting
    only its own pages (SECTIONPAGES), not the cover and contents."""
    if not numbered:
        return field("PAGE", "i", **fmt)
    return field("PAGE", "1", **fmt) + run(" of ", **fmt) + field("SECTIONPAGES", "1", **fmt)


def header_sheet(numbered=True):
    w = [2640, 3738, 1900, 1900]
    outer = (border("top", 8, INK) + border("left", 8, INK) + border("bottom", 8, INK) + border("right", 8, INK)
             + border("insideH", 4, INK_FAINT) + border("insideV", 4, INK_FAINT))
    value = dict(font=BODY, size=17, color=INK)
    row1 = [
        cell(brand("{{LOGO_SHEET}}", 24, 8, 420, w[0] - 180 - 420, 10) + p(style="TitleBlock", ppr='<w:spacing w:after="0"/><w:rPr><w:sz w:val="2"/></w:rPr>'), w[0], vmerge="restart", valign="center"),
        cell(label("Project") + p(bound("project", "[Project name]", **value), "TitleBlock"), w[1]),
        cell(label("Project no.") + p(bound("projectNo", "[No.]", **value), "TitleBlock"), w[2]),
        cell(label("Sheet") + p(sheet_number(numbered, **value), "TitleBlock"), w[3]),
    ]
    row2 = [
        cell("", w[0], vmerge="continue"),
        cell(label("Calculation") + p(core("title", "[Calculation title]", **value), "TitleBlock"), w[1]),
        cell(label("Calc no. / Rev") + p(bound("calcNo", "[No.]", **value) + run(" / ", **value) + bound("revision", "[Rev]", **value), "TitleBlock"), w[2]),
        cell(label("Made / Checked") + p(bound("preparedBy", "[By]", **value) + run(" / ", **value) + bound("checkedBy", "[By]", **value), "TitleBlock"), w[3]),
    ]
    tbl = table([(row1, 360), (row2, 360)], w, outer, margins=(50, 90, 50, 90))
    return tbl + p(style="Header", ppr='<w:spacing w:after="200"/>')


def cover_logo():
    """The cover has no header: its logo is on the page itself, half an inch lower than a header's would be."""
    end = p(style="TitleBlock", ppr='<w:rPr><w:sz w:val="2"/></w:rPr>')  # a table cell must end with a paragraph
    drop = p(style="Spacer", ppr='<w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/></w:rPr>')
    logo = table([([cell(brand("{{LOGO_COVER}}", 40, 12, 680, 3600, 28) + end, 4280)], None)], [4280], NOBORDERS, margins=(0, 0, 0, 0))
    after = p(style="Spacer", ppr='<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:rPr><w:sz w:val="2"/></w:rPr>')
    return drop + logo + after  # a paragraph between tables, or Word joins them (and a floating one next would crash it)


def footer(numbered=True):
    v = dict(font=HEAD, size=14, color=INK_SOFT)
    left = run("Uncontrolled when printed.", font=LIGHT, size=13, color=INK_FAINT)  # no company name or website (user's request)
    right = (run("Calc no. ", **v) + bound("calcNo", "[No.]", **v) + run("  ·  Rev ", **v) + bound("revision", "[Rev]", **v)
             + run("  ·  Page ", **v) + sheet_number(numbered, **v))
    ppr = (f'<w:pBdr>{border("top", 4, LINE).replace("w:space=\"0\"", "w:space=\"4\"")}</w:pBdr>'
           f'<w:tabs><w:tab w:val="right" w:pos="{TEXT_W}"/></w:tabs>')
    return p(left + tab(**v) + right, "Footer", ppr)


# ---- Body --------------------------------------------------------------------------------------

def brand_bar():
    cells = [cell(p(style="TitleBlock", ppr='<w:rPr><w:sz w:val="2"/></w:rPr>'), 900, shade=GREEN),
             cell(p(style="TitleBlock", ppr='<w:rPr><w:sz w:val="2"/></w:rPr>'), 900, shade=BLUE)]
    return table([(cells, -80)], [900, 900], NOBORDERS, margins=(0, 0, 0, 0))


def details_table(values):
    w = [1500, 3589, 1500, 3589]
    rows = [("Project", "project", "[Project name]", "Project no.", "projectNo", "[Project number]"),
            ("Client", "client", "[Client]", "Calc no.", "calcNo", "[Calculation number]"),
            ("Location", "location", "[Site or location]", "Revision", "revision", "[Revision]"),
            ("Design codes", "codes", "[Codes, e.g. Eurocodes]", "Date", "date", "[Date]")]
    lines = border("top", 8, INK) + border("bottom", 8, INK) + border("insideH", 4, LINE)
    out = []
    for l1, f1, t1, l2, f2, t2 in rows:
        out.append(([cell(label(l1), w[0], valign="center"),
                     cell(p(bound(f1, t1, values.get(f1)), "TitleBlockValue"), w[1], valign="center"),
                     cell(label(l2), w[2], valign="center"),
                     cell(p(bound(f2, t2, values.get(f2), date=(f2 == "date")), "TitleBlockValue"), w[3], valign="center")], 400))
    return table(out, w, lines, margins=(60, 80, 60, 80))


def revision_table():
    w = [800, 1500, 3578, 1433, 1433, 1434]
    head = [cell(p(run(t), "TableHeading"), wd) for t, wd in zip(("Rev", "Date", "Description", "Prepared", "Checked", "Approved"), w)]
    first = [cell(p(run("0"), "TableText"), w[0]),
             cell(p(inline_prompt("Dokimos_rev_date", "[Date]"), "TableText"), w[1]),
             cell(p(run("First issue"), "TableText"), w[2]),
             cell(p(inline_prompt("Dokimos_rev_by", "[Name]"), "TableText"), w[3]),
             cell(p(inline_prompt("Dokimos_rev_checked", "[Name]"), "TableText"), w[4]),
             cell(p(inline_prompt("Dokimos_rev_approved", "[Name]"), "TableText"), w[5])]
    empty = [cell(p(style="TableText"), wd) for wd in w]
    return table([(head, 300), (first, 340), (empty, 340), (list(empty), 340)], w, style="DokimosTable", header_rows=1)


def cover(values, title=None, subject=None):
    # Cover, bottom-aligned: title, subject, project details and document control sit together at
    # the foot of the page; the space is above them, under the logo.
    tiny = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:rPr><w:sz w:val="2"/></w:rPr>'
    foot = (p(core("title", "[Calculation title]", title), "Title", '<w:spacing w:before="0"/>')
            + p(core("subject", "[Subject: the structure, element or question this calculation answers]", subject), "Subtitle")
            + brand_bar() + p(style="Spacer", ppr='<w:spacing w:after="480" w:line="240" w:lineRule="auto"/>')
            + details_table(values) + p(run("Document control"), "SectionLabel") + revision_table()
            + p(style="Spacer", ppr=tiny))  # a table cell must end with a paragraph
    # Three sections: the cover (no header or footer, no number), the contents (i, ii, ...) and the
    # body (1, 2, ...). A section ends with the paragraph that carries its properties.
    return "".join([
        cover_logo(),
        table([([cell(foot, TEXT_W)], None)], [TEXT_W], NOBORDERS, margins=(0, 0, 0, 0), at_page_bottom=True),
        p(style="Spacer", ppr=tiny + section(None, None, first=True)),  # holds the floating table on the cover; ends the cover
        p(run("Contents"), "SectionLabel"),
        p(f'<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-2" \\h \\z \\u </w:instrText></w:r>'
          f'<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
          + run("Right-click here and choose Update Field to list the sections.", style="PlaceholderText"), "TOC1")
        + p('<w:r><w:fldChar w:fldCharType="end"/></w:r>', "TOC1"),
        p(style="Spacer", ppr=tiny + section("rIdHeaderContents", None, "lowerRoman")),  # ends the contents; no footers (user's request)
    ])


GUIDE = {
    "scope": "State what this calculation covers: the structure or element, the design stage (concept, detailed design, construction) and what the result is used for.",
    "codes": "List the design codes with their editions and national annexes, e.g. EN 1992-1-1:2004 (Eurocode 2) with its National Annex.",
    "references": "List the drawings, reports and other calculations this one relies on, with their revisions.",
    "assumptions": "List the assumptions and simplifications, and anything still to be confirmed.",
    "materials": "Material grades and design strengths, e.g. concrete C30/37, reinforcement B500B. Type them as Calcpad lines, select them and press Evaluate (Alt+Shift+Enter).",
    "geometry": "Dimensions and section properties, as Calcpad lines, e.g. a line with only: L = 6m",
    "loads": "Loads and load combinations, with where they come from.",
    "calculations": "Work through the calculation step by step. Paste or type Calcpad here and press Evaluate: headings and text become Word paragraphs and the equations are calculated. Long calculations run in the background.",
    "results": "Summarise the governing results against their limits, e.g. bearing pressure 182 kPa ≤ 250 kPa: OK. Give the conclusions and any actions.",
}


def body(example=None):
    ex = example or {}
    g = lambda key, style=None: ex.get(key) or prompt("Dokimos_guide_" + key, GUIDE[key], style)
    return "".join([
        p(run("Scope"), "Heading1"), g("scope"),
        p(run("Design basis"), "Heading1"),
        p(run("Codes and standards"), "Heading2"), g("codes", "ListBullet"),
        p(run("References"), "Heading2"), g("references", "ListBullet"),
        p(run("Assumptions"), "Heading2"), g("assumptions", "ListBullet"),
        p(run("Design inputs"), "Heading1"),
        p(run("Materials"), "Heading2"), g("materials"),
        p(run("Geometry"), "Heading2"), g("geometry"),
        p(run("Loads"), "Heading2"), g("loads"),
        p(run("Calculations"), "Heading1"), g("calculations"),
        p(run("Results and conclusions"), "Heading1"),
        ex.get("results") or prompt("Dokimos_guide_results", GUIDE["results"], "ResultSummary"),
    ])


def section(header, footer_, numbering=None, first=False):
    """Section properties. The cover has no header or footer; the contents restart at i, the body at 1."""
    refs = ""
    if header: refs += f'<w:headerReference w:type="default" r:id="{header}"/>'
    if footer_: refs += f'<w:footerReference w:type="default" r:id="{footer_}"/>'
    start = "" if first else '<w:type w:val="nextPage"/>'
    pages = f'<w:pgNumType w:fmt="{numbering}" w:start="1"/>' if numbering else ""
    return (f"<w:sectPr>{refs}{start}<w:pgSz w:w=\"{PAGE_W}\" w:h=\"{PAGE_H}\"/>"
            f'<w:pgMar w:top="{MARGIN}" w:right="{MARGIN}" w:bottom="{MARGIN}" w:left="{MARGIN}" w:header="504" w:footer="432" w:gutter="0"/>'
            f'{pages}<w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>')


def section_properties():
    return section("rIdHeaderSheet", None, "decimal")  # no footer: the sheet header has the numbers


# ---- Styles, numbering, theme, settings ---------------------------------------------------------

def style(kind, sid, name, ppr="", rp="", based="Normal", next_=None, extra="", ui=None, q=True, default=False, link=None):
    attrs = f' w:type="{kind}" w:styleId="{sid}"' + (' w:default="1"' if default else "") + (' w:customStyle="1"' if sid not in BUILTIN else "")
    s = f'<w:style{attrs}><w:name w:val="{name}"/>'
    if based: s += f'<w:basedOn w:val="{based}"/>'
    if next_: s += f'<w:next w:val="{next_}"/>'
    if link: s += f'<w:link w:val="{link}"/>'
    if ui is not None: s += f'<w:uiPriority w:val="{ui}"/>'
    s += extra  # semiHidden / unhideWhenUsed come before qFormat
    if q: s += "<w:qFormat/>"
    if ppr: s += f"<w:pPr>{ppr}</w:pPr>"
    if rp: s += f"<w:rPr>{rp}</w:rPr>"
    return s + "</w:style>"


BUILTIN = {"Normal", "Heading1", "Heading2", "Heading3", "Title", "Subtitle", "Header", "Footer", "Caption",
           "ListBullet", "ListNumber", "TOC1", "TOC2", "TableNormal", "NoList", "PlaceholderText", "Hyperlink",
           "TableGrid", "DefaultParagraphFont"}


def fonts(theme_major=False, font=None):
    if font:
        return f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>'
    return ('<w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>' if theme_major
            else '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>')


def color(hexv, theme=None):
    return f'<w:color w:val="{hexv}"' + (f' w:themeColor="{theme}"' if theme else "") + "/>"


def sz(half_points):
    return f'<w:sz w:val="{half_points}"/><w:szCs w:val="{half_points}"/>'


def styles_xml():
    S = []
    S.append(style("paragraph", "Normal", "Normal", based=None, default=True, ui=0,
                   ppr='<w:spacing w:after="120" w:line="264" w:lineRule="auto"/>'))
    S.append(style("character", "DefaultParagraphFont", "Default Paragraph Font", based=None, default=True, ui=1, q=False,
                   extra="<w:semiHidden/><w:unhideWhenUsed/>"))
    S.append('<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/>'
             '<w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/>'
             '<w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>')
    S.append('<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style>')
    heading_ppr = lambda before, after, lvl: (f'<w:keepNext/><w:keepLines/><w:numPr><w:ilvl w:val="{lvl}"/><w:numId w:val="1"/></w:numPr>'
                                               f'<w:spacing w:before="{before}" w:after="{after}" w:line="240" w:lineRule="auto"/><w:outlineLvl w:val="{lvl}"/>')
    S.append(style("paragraph", "Heading1", "heading 1", next_="Normal", ui=9, link="Heading1Char",
                   ppr=heading_ppr(360, 140, 0).replace("<w:spacing", f'<w:pBdr>{border("bottom", 6, INK).replace("w:space=\"0\"", "w:space=\"3\"")}</w:pBdr><w:spacing'),
                   rp=fonts(True) + "<w:caps/>" + color(INK, "text1") + '<w:spacing w:val="10"/>' + sz(24)))
    S.append(style("paragraph", "Heading2", "heading 2", next_="Normal", ui=9, link="Heading2Char",
                   ppr=heading_ppr(240, 80, 1), rp=fonts(True) + color(BLUE, "accent1") + sz(22)))
    S.append(style("paragraph", "Heading3", "heading 3", next_="Normal", ui=9,
                   ppr=heading_ppr(200, 60, 2), rp=fonts(True) + color(INK_SOFT, "accent3") + sz(20)))
    for n in (1, 2):
        S.append(style("character", f"Heading{n}Char", f"Heading {n} Char", based="DefaultParagraphFont", ui=9, q=False, link=f"Heading{n}",
                       rp=fonts(True) + sz(26 if n == 1 else 23)))
    S.append(style("paragraph", "Title", "Title", next_="Subtitle", ui=10,
                   ppr='<w:spacing w:before="120" w:after="60" w:line="240" w:lineRule="auto"/><w:contextualSpacing/>',
                   rp=fonts(True) + color(INK, "text1") + '<w:kern w:val="28"/>' + sz(48)))
    S.append(style("paragraph", "Subtitle", "Subtitle", next_="Normal", ui=11,
                   ppr='<w:spacing w:after="240" w:line="264" w:lineRule="auto"/>', rp=fonts(font=LIGHT) + color(INK_SOFT, "accent3") + sz(26)))
    S.append(style("paragraph", "Kicker", "Kicker", next_="Title", ui=12,
                   ppr='<w:spacing w:before="480" w:after="40"/>', rp=fonts(True) + "<w:caps/>" + color(BLUE, "accent1") + '<w:spacing w:val="40"/>' + sz(16)))
    S.append(style("paragraph", "SectionLabel", "Section Label", next_="Normal", ui=12,
                   ppr='<w:keepNext/><w:spacing w:before="360" w:after="100"/>', rp=fonts(True) + "<w:caps/>" + color(INK, "text1") + '<w:spacing w:val="30"/>' + sz(16)))
    S.append(style("paragraph", "Spacer", "Spacer", ui=13, q=False, ppr='<w:spacing w:after="240" w:line="240" w:lineRule="auto"/>', rp=sz(8)))
    S.append(style("paragraph", "TitleBlock", "Title Block", ui=13, q=False, ppr='<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>', rp=fonts(True) + sz(17)))
    S.append(style("paragraph", "TitleBlockValue", "Title Block Value", based="TitleBlock", ui=13, q=False, rp=fonts(False) + sz(20)))
    S.append(style("character", "FieldLabel", "Field Label", based="DefaultParagraphFont", ui=13,
                   rp=fonts(True) + "<w:caps/>" + color(INK_FAINT, "accent4") + '<w:spacing w:val="16"/>' + sz(13)))
    S.append(style("paragraph", "TableHeading", "Table Heading", ui=14,
                   ppr='<w:keepNext/><w:spacing w:before="20" w:after="20" w:line="240" w:lineRule="auto"/>',
                   rp=fonts(True) + "<w:caps/>" + color("FFFFFF", "background1") + '<w:spacing w:val="10"/>' + sz(14)))
    S.append(style("paragraph", "TableText", "Table Text", ui=14, ppr='<w:spacing w:before="20" w:after="20" w:line="240" w:lineRule="auto"/>', rp=sz(18)))
    S.append(style("paragraph", "Header", "header", ui=99, q=False, link=None, extra="<w:unhideWhenUsed/>",
                   ppr=f'<w:tabs><w:tab w:val="center" w:pos="{TEXT_W // 2}"/><w:tab w:val="right" w:pos="{TEXT_W}"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>',
                   rp=sz(4)))
    S.append(style("paragraph", "Footer", "footer", ui=99, q=False, extra="<w:unhideWhenUsed/>",
                   ppr=f'<w:tabs><w:tab w:val="right" w:pos="{TEXT_W}"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>',
                   rp=fonts(True) + color(INK_SOFT, "accent3") + sz(14)))
    S.append(style("paragraph", "Caption", "caption", next_="Normal", ui=35, extra="<w:unhideWhenUsed/>",
                   ppr='<w:spacing w:before="60" w:after="200" w:line="240" w:lineRule="auto"/>', rp=fonts(True) + color(INK_SOFT, "accent3") + sz(16)))
    S.append(style("paragraph", "ListBullet", "List Bullet", ui=36, extra="<w:unhideWhenUsed/>",
                   ppr='<w:numPr><w:numId w:val="2"/></w:numPr><w:spacing w:after="60"/><w:contextualSpacing/>'))
    S.append(style("paragraph", "ListNumber", "List Number", ui=37, extra="<w:unhideWhenUsed/>",
                   ppr='<w:numPr><w:numId w:val="3"/></w:numPr><w:spacing w:after="60"/><w:contextualSpacing/>'))
    S.append(style("paragraph", "ResultSummary", "Result Summary", ui=15,
                   ppr=(f'<w:pBdr>{border("left", 24, GREEN).replace("w:space=\"0\"", "w:space=\"8\"")}</w:pBdr>'
                        f'<w:shd w:val="clear" w:color="auto" w:fill="{PAPER}"/><w:spacing w:before="120" w:after="120"/><w:ind w:left="170" w:right="113"/>')))
    S.append(style("paragraph", "Note", "Note", ui=16, ppr='<w:spacing w:after="120"/>', rp="<w:i/><w:iCs/>" + color(INK_SOFT, "accent3") + sz(18)))
    S.append(style("character", "CheckOK", "Check OK", based="DefaultParagraphFont", ui=17, rp="<w:b/><w:bCs/>" + color(OK_TEXT)))
    S.append(style("character", "CheckNotOK", "Check Not OK", based="DefaultParagraphFont", ui=17, rp="<w:b/><w:bCs/>" + color(RED, "accent5")))
    S.append(style("character", "Reference", "Reference", based="DefaultParagraphFont", ui=18, rp=fonts(True) + color(INK_SOFT, "accent3") + sz(16)))
    S.append(style("character", "PlaceholderText", "Placeholder Text", based="DefaultParagraphFont", ui=99, q=False,
                   extra="<w:semiHidden/>", rp=color(INK_FAINT, "accent4")))
    S.append(style("character", "Hyperlink", "Hyperlink", based="DefaultParagraphFont", ui=99, q=False, extra="<w:unhideWhenUsed/>",
                   rp=color(BLUE, "hyperlink") + '<w:u w:val="single"/>'))
    S.append(style("paragraph", "TOC1", "toc 1", next_="Normal", ui=39, q=False, extra="<w:unhideWhenUsed/>",
                   ppr=f'<w:tabs><w:tab w:val="left" w:pos="680"/><w:tab w:val="right" w:leader="dot" w:pos="{TEXT_W}"/></w:tabs><w:spacing w:before="80" w:after="20"/>',
                   rp=fonts(True) + sz(19)))
    S.append(style("paragraph", "TOC2", "toc 2", next_="Normal", ui=39, q=False, extra="<w:unhideWhenUsed/>",
                   ppr=f'<w:tabs><w:tab w:val="left" w:pos="1360"/><w:tab w:val="right" w:leader="dot" w:pos="{TEXT_W}"/></w:tabs><w:spacing w:after="20"/><w:ind w:left="680"/>',
                   rp=sz(19)))
    S.append(f'<w:style w:type="table" w:styleId="DokimosTable" w:customStyle="1"><w:name w:val="Dokimos Table"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="40"/>'
             f'<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr>{sz(19)}</w:rPr>'
             f'<w:tblPr><w:tblBorders>{border("top", 8, INK)}{border("bottom", 8, INK)}{border("insideH", 4, LINE)}</w:tblBorders>'
             f'<w:tblCellMar><w:top w:w="40" w:type="dxa"/><w:left w:w="90" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/><w:right w:w="90" w:type="dxa"/></w:tblCellMar></w:tblPr>'
             f'<w:tcPr><w:vAlign w:val="center"/></w:tcPr>'
             f'<w:tblStylePr w:type="firstRow"><w:rPr>{fonts(True)}<w:caps/>{color("FFFFFF", "background1")}{sz(15)}</w:rPr>'
             f'<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="{INK}" w:themeFill="text1"/></w:tcPr></w:tblStylePr>'
             f'<w:tblStylePr w:type="band2Horz"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="{PAPER}"/></w:tcPr></w:tblStylePr></w:style>')
    S.append(f'<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="39"/>'
             f'<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders>'
             + "".join(border(s, 4, INK_FAINT) for s in ("top", "left", "bottom", "right", "insideH", "insideV"))
             + "</w:tblBorders></w:tblPr></w:style>")
    defaults = (f'<w:docDefaults><w:rPrDefault><w:rPr>{fonts(False)}<w:sz w:val="20"/><w:szCs w:val="20"/>'
                f'<w:lang w:val="en-GB" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>'
                f'<w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="264" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>')
    latent = '<w:latentStyles w:defLockedState="0" w:defUIPriority="99" w:defSemiHidden="0" w:defUnhideWhenUsed="0" w:defQFormat="0" w:count="376"/>'
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'{defaults}{latent}{"".join(S)}</w:styles>')


def numbering_xml():
    def lvl(i, fmt, text, ind, hang, style=None, font=None, start=1):
        r = f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:hint="default"/></w:rPr>' if font else ""
        return (f'<w:lvl w:ilvl="{i}"><w:start w:val="{start}"/><w:numFmt w:val="{fmt}"/>'
                + (f'<w:pStyle w:val="{style}"/>' if style else "")
                + f'<w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="{ind}" w:hanging="{hang}"/></w:pPr>{r}</w:lvl>')
    headings = "".join([lvl(0, "decimal", "%1", 680, 680, "Heading1"), lvl(1, "decimal", "%1.%2", 680, 680, "Heading2"),
                        lvl(2, "decimal", "%1.%2.%3", 680, 680, "Heading3")]
                       + [lvl(i, "decimal", "%1." * 0 + "".join(f"%{k + 1}." for k in range(i + 1)).rstrip("."), 680, 680) for i in range(3, 9)])
    bullets = "".join([lvl(0, "bullet", "•", 340, 227, font=BODY), lvl(1, "bullet", "–", 680, 227, font=BODY)]
                      + [lvl(i, "bullet", "–", 340 * (i + 1), 227, font=BODY) for i in range(2, 9)])
    numbers = "".join([lvl(0, "decimal", "%1.", 340, 340), lvl(1, "lowerLetter", "%2)", 680, 340)]
                      + [lvl(i, "lowerRoman", f"%{i + 1}.", 340 * (i + 1), 340) for i in range(2, 9)])
    abstract = lambda nid, levels, kind: (f'<w:abstractNum w:abstractNumId="{nid}"><w:multiLevelType w:val="{kind}"/>{levels}</w:abstractNum>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            + abstract(10, headings, "multilevel") + abstract(11, bullets, "hybridMultilevel") + abstract(12, numbers, "hybridMultilevel")
            + '<w:num w:numId="1"><w:abstractNumId w:val="10"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="11"/></w:num>'
            + '<w:num w:numId="3"><w:abstractNumId w:val="12"/></w:num></w:numbering>')


def theme_xml():
    t = open(os.path.join(SOURCE, "word", "theme", "theme1.xml"), encoding="utf-8").read()
    clr = (f'<a:clrScheme name="Dokimos"><a:dk1><a:srgbClr val="{INK}"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>'
           f'<a:dk2><a:srgbClr val="{BROWN}"/></a:dk2><a:lt2><a:srgbClr val="{PAPER_ALT}"/></a:lt2>'
           f'<a:accent1><a:srgbClr val="{BLUE}"/></a:accent1><a:accent2><a:srgbClr val="{GREEN}"/></a:accent2>'
           f'<a:accent3><a:srgbClr val="{INK_SOFT}"/></a:accent3><a:accent4><a:srgbClr val="{INK_FAINT}"/></a:accent4>'
           f'<a:accent5><a:srgbClr val="{RED}"/></a:accent5><a:accent6><a:srgbClr val="{BROWN}"/></a:accent6>'
           f'<a:hlink><a:srgbClr val="{BLUE}"/></a:hlink><a:folHlink><a:srgbClr val="{INK_SOFT}"/></a:folHlink></a:clrScheme>')
    t = re.sub(r"<a:clrScheme.*?</a:clrScheme>", clr, t, flags=re.S)
    t = re.sub(r'(<a:majorFont>\s*<a:latin typeface=")[^"]*', r"\g<1>" + HEAD, t)
    t = re.sub(r'(<a:minorFont>\s*<a:latin typeface=")[^"]*', r"\g<1>" + BODY, t)
    t = re.sub(r'<a:fontScheme name="[^"]*"', '<a:fontScheme name="Dokimos"', t)
    t = re.sub(r'<a:theme ([^>]*)name="[^"]*"', r'<a:theme \1name="Dokimos"', t)
    return t


def settings_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
            '<w:zoom w:percent="100"/><w:proofState w:spelling="clean" w:grammar="clean"/><w:defaultTabStop w:val="680"/>'
            '<w:characterSpacingControl w:val="doNotCompress"/><w:compat>'
            '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
            '<m:mathPr><m:mathFont m:val="Cambria Math"/><m:brkBin m:val="before"/><m:brkBinSub m:val="--"/><m:smallFrac m:val="0"/>'
            '<m:dispDef/><m:lMargin m:val="0"/><m:rMargin m:val="0"/><m:defJc m:val="centerGroup"/><m:wrapIndent m:val="1440"/>'
            '<m:intLim m:val="subSup"/><m:naryLim m:val="undOvr"/></m:mathPr>'
            '<w:themeFontLang w:val="en-GB"/><w:decimalSymbol w:val="."/><w:listSeparator w:val=","/></w:settings>')


def glossary_xml():
    parts = ""
    for name, (text, pstyle) in DOCPARTS.items():
        guid = "{" + str(uuid.uuid5(uuid.NAMESPACE_URL, "dokimos/" + name)).upper() + "}"
        parts += (f'<w:docPart><w:docPartPr><w:name w:val="{name}"/><w:category><w:name w:val="General"/>'
                  f'<w:gallery w:val="placeholder"/></w:category><w:types><w:type w:val="bbPlcHdr"/></w:types>'
                  f'<w:behaviors><w:behavior w:val="content"/></w:behaviors><w:guid w:val="{guid}"/></w:docPartPr>'
                  f'<w:docPartBody>{p(run(text, style="PlaceholderText"), pstyle)}</w:docPartBody></w:docPart>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:docParts>{parts}</w:docParts></w:glossaryDocument>')


def font_table():
    entries = "".join(f'<w:font w:name="{f}"><w:charset w:val="00"/><w:family w:val="{fam}"/><w:pitch w:val="variable"/></w:font>'
                      for f, fam in ((BODY, "swiss"), (HEAD, "swiss"), (LIGHT, "swiss"), ("Cambria Math", "roman"),
                                     ("Consolas", "modern")))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f"{entries}</w:fonts>")


# ---- Package -----------------------------------------------------------------------------------

def root_start():
    src = open(os.path.join(SOURCE, "word", "document.xml"), encoding="utf-8").read()
    return re.search(r"<w:document[^>]*>", src).group(0)


def part(root, content):
    ns = re.findall(r'xmlns:\w+="[^"]*"', root_start())
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<{root} {" ".join(ns)}>{content}</{root.split()[0]}>'


def build(path, values=None, title=None, subject=None, example=None):
    values = values or {}
    DOCPARTS.clear()
    document = (root_start().replace("<w:document", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document', 1)
                + "<w:body>" + cover(values, title, subject) + body(example) + section_properties() + "</w:body></w:document>")
    headers = {"header1.xml": part("w:hdr", header_sheet()), "header2.xml": part("w:hdr", header_sheet(numbered=False))}
    data = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<calculation xmlns="{DATA_NS}">'
            + "".join(f"<{f}>{x(values.get(f, ''))}</{f}>" for f in FIELDS) + "</calculation>")
    item_props = (f'<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n<ds:datastoreItem ds:itemID="{DATA_ID}" '
                  f'xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs>'
                  f'<ds:schemaRef ds:uri="{DATA_NS}"/></ds:schemaRefs></ds:datastoreItem>')
    rel = lambda rid, kind, target: f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{kind}" Target="{target}"/>'
    doc_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                + rel("rIdStyles", "styles", "styles.xml") + rel("rIdSettings", "settings", "settings.xml")
                + rel("rIdNumbering", "numbering", "numbering.xml") + rel("rIdFonts", "fontTable", "fontTable.xml")
                + rel("rIdTheme", "theme", "theme/theme1.xml") + rel("rIdHeaderSheet", "header", "header1.xml")
                + rel("rIdHeaderContents", "header", "header2.xml")
                + rel("rIdGlossary", "glossaryDocument", "glossary/document.xml")
                + rel("rIdData", "customXml", "../customXml/item1.xml") + "</Relationships>")
    glossary_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                     + rel("rId1", "styles", "styles.xml") + rel("rId2", "settings", "settings.xml") + rel("rId3", "fontTable", "fontTable.xml")
                     + rel("rId4", "numbering", "numbering.xml") + "</Relationships>")
    ct = lambda name, kind: f'<Override PartName="{name}" ContentType="application/vnd.openxmlformats-officedocument.{kind}"/>'
    content_types = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                     '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                     '<Default Extension="xml" ContentType="application/xml"/>'
                     + ct("/word/document.xml", "wordprocessingml.document.main+xml")
                     + ct("/word/styles.xml", "wordprocessingml.styles+xml") + ct("/word/settings.xml", "wordprocessingml.settings+xml")
                     + ct("/word/numbering.xml", "wordprocessingml.numbering+xml") + ct("/word/fontTable.xml", "wordprocessingml.fontTable+xml")
                     + ct("/word/theme/theme1.xml", "theme+xml")
                     + "".join(ct(f"/word/{h}", "wordprocessingml." + ("header" if h.startswith("header") else "footer") + "+xml") for h in headers)
                     + ct("/word/glossary/document.xml", "wordprocessingml.document.glossary+xml")
                     + ct("/word/glossary/styles.xml", "wordprocessingml.styles+xml") + ct("/word/glossary/settings.xml", "wordprocessingml.settings+xml")
                     + ct("/word/glossary/fontTable.xml", "wordprocessingml.fontTable+xml")
                     + ct("/word/glossary/numbering.xml", "wordprocessingml.numbering+xml")
                     + ct("/customXml/itemProps1.xml", "customXmlProperties+xml")
                     + '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
                     + ct("/docProps/app.xml", "extended-properties+xml") + "</Types>")
    package_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
                    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
                    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
                    "</Relationships>")
    core_props = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
                  'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
                  'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
                  f"<dc:title>{x(title or '')}</dc:title><dc:subject>{x(subject or '')}</dc:subject><dc:creator>{COMPANY}</dc:creator>"
                  f"<cp:lastModifiedBy>{COMPANY}</cp:lastModifiedBy></cp:coreProperties>")
    app = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
           f"<Application>Microsoft Office Word</Application><Company>{COMPANY}</Company></Properties>")
    styles = styles_xml()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", package_rels)
        z.writestr("docProps/core.xml", core_props)
        z.writestr("docProps/app.xml", app)
        z.writestr("word/document.xml", document)
        z.writestr("word/_rels/document.xml.rels", doc_rels)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/numbering.xml", numbering_xml())
        z.writestr("word/settings.xml", settings_xml())
        z.writestr("word/fontTable.xml", font_table())
        z.writestr("word/theme/theme1.xml", theme_xml())
        for name, content in headers.items():
            z.writestr(f"word/{name}", content)
        z.writestr("word/glossary/document.xml", glossary_xml())
        z.writestr("word/glossary/_rels/document.xml.rels", glossary_rels)
        z.writestr("word/glossary/styles.xml", styles)
        z.writestr("word/glossary/settings.xml", settings_xml())
        z.writestr("word/glossary/fontTable.xml", font_table())
        z.writestr("word/glossary/numbering.xml", numbering_xml())
        z.writestr("customXml/item1.xml", data)
        z.writestr("customXml/itemProps1.xml", item_props)
        z.writestr("customXml/_rels/item1.xml.rels",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/></Relationships>')
    print("wrote", path, len(DOCPARTS), "placeholders")


def source_blocks():
    """The body elements of the user's document, as XML, in order."""
    src = open(os.path.join(SOURCE, "word", "document.xml"), encoding="utf-8").read()
    body_xml = re.search(r"<w:body>(.*)</w:body>", src, re.S).group(1)
    blocks, depth, start, i = [], 0, 0, 0
    for m in re.finditer(r"<(/?)w:(p|tbl|sdt|sectPr)\b[^>]*?(/?)>", body_xml):
        closing, selfclosing = m.group(1) == "/", m.group(3) == "/"
        if not closing and not selfclosing:
            if depth == 0: start = m.start()
            depth += 1
        elif closing:
            depth -= 1
            if depth == 0: blocks.append(body_xml[start:m.end()])
        elif selfclosing and depth == 0:
            blocks.append(m.group(0))
    return [b for b in blocks if not b.startswith("<w:sectPr")]


def write_cover_pattern(path, size=1000, spacing=62, reach=1000.0, line_alpha=0.11):
    """A quiet corner graphic for the cover: a triangulated grid (a truss mesh) that fades out from
    the top-right corner, over two very pale triangles in the brand green and blue."""
    import math
    h = spacing * math.sqrt(3) / 2
    corner = (size, 0.0)

    def alpha(x, y):
        d = math.hypot(x - corner[0], y - corner[1])
        return line_alpha * max(0.0, 1.0 - d / reach) ** 1.6

    lines = []
    rows = int(size / h) + 2
    cols = int(size / spacing) + 2
    point = lambda i, j: (i * spacing + (spacing / 2 if j % 2 else 0), j * h)
    for j in range(rows):
        for i in range(-1, cols):
            x1, y1 = point(i, j)
            ends = [point(i + 1, j)]
            if j % 2:
                ends += [point(i, j + 1), point(i + 1, j + 1)]
            else:
                ends += [point(i - 1, j + 1), point(i, j + 1)]
            for x2, y2 in ends:
                a = alpha((x1 + x2) / 2, (y1 + y2) / 2)
                if a > 0.004 and min(x1, x2) < size and min(y1, y2) < size:
                    lines.append(f'<line x1="{x1:.1f}" y1="{y1:.1f}" x2="{x2:.1f}" y2="{y2:.1f}" stroke-opacity="{a:.3f}"/>')
    # The two pale triangles are cut along mesh lines (60°), starting where a mesh line meets the top edge.
    def triangle(nodes_from_left, colour, opacity):
        x0 = nodes_from_left * spacing
        y1 = (size - x0) * math.sqrt(3)
        return f'<polygon points="{x0:.1f},0 {size},0 {size},{y1:.1f}" fill="#{colour}" fill-opacity="{opacity}"/>'

    svg = (f'<svg xmlns="http://www.w3.org/2000/svg" width="{size}" height="{size}" viewBox="0 0 {size} {size}">'
           + triangle(7, GREEN, 0.055) + triangle(12, BLUE, 0.06) +
           f'<g stroke="#{INK}" stroke-width="1.2" fill="none">{"".join(lines)}</g></svg>')
    with open(path, "w", encoding="utf-8") as f:
        f.write(svg)
    print("wrote", path, len(lines), "lines")


if __name__ == "__main__":
    write_cover_pattern(os.path.join(OUT, "cover-pattern.svg"))
    build(os.path.join(OUT, "dokimos-template.docx"))
    b = source_blocks()
    text = lambda s: "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", s))
    print("source blocks:", [text(s)[:30] for s in b])
    title = text(b[0])
    example = {
        "scope": b[1],
        "materials": b[5] + b[6],
        "geometry": b[2] + b[3],
        "loads": b[4],
        "calculations": p(run("Stress equations"), "Heading2") + "".join(b[8:12]),
    }
    build(os.path.join(OUT, "dokimos-example.docx"), title=title, example=example)

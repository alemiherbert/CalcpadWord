"""Template v3 from the user's own version (their logo and pattern positions kept):
- the cover is its own section: its header (logo, pattern) stays, no footer;
- the contents are a section numbered i, ii, ...; the body a section numbered from 1;
- no footers anywhere; the sheet header counts body pages only ("1 of N")."""
import re
import sys
import zipfile

src, dst = sys.argv[1:3]
with zipfile.ZipFile(src) as z:
    order = [i.filename for i in z.infolist()]
    parts = {name: z.read(name) for name in order}


def text(name):
    return parts[name].decode("utf-8")


doc = text("word/document.xml")
final = re.search(r"<w:sectPr\b[^>]*>.*?</w:sectPr>(?=</w:body>)", doc, re.S)
page_size = re.search(r"<w:pgSz\b[^>]*/>", final.group(0)).group(0)
page_margins = re.search(r"<w:pgMar\b[^>]*/>", final.group(0)).group(0)


def section(refs, numbering=None, first=False, title_page=False):
    return ("<w:sectPr>" + refs + ("" if first else '<w:type w:val="nextPage"/>') + page_size + page_margins
            + (f'<w:pgNumType w:fmt="{numbering}" w:start="1"/>' if numbering else "")
            + '<w:cols w:space="708"/>' + ("<w:titlePg/>" if title_page else "") + '<w:docGrid w:linePitch="360"/></w:sectPr>')


def paragraph_start(before):
    return [m.start() for m in re.finditer(r"<w:p[ >]", doc[:before])][-1]


# 1. The cover ends with the spacer paragraph just before "Contents".
contents = doc.index('<w:pStyle w:val="SectionLabel"/><w:pageBreakBefore/>')
holder = paragraph_start(paragraph_start(contents))
end_of_props = doc.index("</w:pPr>", holder)
doc = doc[:end_of_props] + section('<w:headerReference w:type="first" r:id="rId12"/>', first=True, title_page=True) + doc[end_of_props:]
doc = doc.replace('<w:pStyle w:val="SectionLabel"/><w:pageBreakBefore/>', '<w:pStyle w:val="SectionLabel"/>', 1)

# 2. The contents end just before the first heading (Scope).
scope = doc.index('<w:pStyle w:val="Heading1"/><w:pageBreakBefore/>')
start = paragraph_start(scope)
contents_end = ('<w:p><w:pPr><w:pStyle w:val="Spacer"/><w:spacing w:after="0"/><w:rPr><w:sz w:val="2"/></w:rPr>'
                + section('<w:headerReference w:type="default" r:id="rIdContentsHeader"/>', "lowerRoman") + "</w:pPr></w:p>")
doc = doc[:start] + contents_end + doc[start:]
doc = doc.replace('<w:pStyle w:val="Heading1"/><w:pageBreakBefore/>', '<w:pStyle w:val="Heading1"/>', 1)

# 3. The body: the sheet header, numbered from 1; no footer. (The last sectPr: the ones added above come first.)
last = doc.rindex("<w:sectPr")
last_end = doc.index("</w:sectPr>", last) + len("</w:sectPr>")
assert doc[last_end:].startswith("</w:body>")
doc = doc[:last] + section('<w:headerReference w:type="default" r:id="rId9"/>', "decimal") + doc[last_end:]
parts["word/document.xml"] = doc.encode("utf-8")

# Headers: the body's counts its own pages; the contents' shows only its page number (i).
sheet = text("word/header2.xml")
assert sheet.count(" NUMPAGES ") == 1
parts["word/header2.xml"] = sheet.replace(" NUMPAGES ", " SECTIONPAGES ").encode("utf-8")
contents_header, n = re.subn(r'<w:r\b[^>]*>(?:(?!</w:r>).)*?<w:t xml:space="preserve"> of </w:t></w:r>.*?NUMPAGES.*?'
                             r'<w:fldChar w:fldCharType="end"/></w:r>', "", sheet, count=1, flags=re.S)
assert n == 1, "sheet number not found"
parts["word/header4.xml"] = contents_header.encode("utf-8")
parts["word/_rels/header4.xml.rels"] = parts["word/_rels/header2.xml.rels"]

# No footers and no even-page header: their parts and relationships go.
gone = {"word/header1.xml", "word/footer1.xml", "word/footer2.xml", "word/footer3.xml"}
rels = text("word/_rels/document.xml.rels")
for rid in ("rId8", "rId10", "rId11", "rId13"):
    rels, n = re.subn(rf'<Relationship Id="{rid}"[^>]*/>', "", rels)
    assert n == 1, rid
rels = rels.replace("</Relationships>", '<Relationship Id="rIdContentsHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header4.xml"/></Relationships>')
parts["word/_rels/document.xml.rels"] = rels.encode("utf-8")
types = text("[Content_Types].xml")
for name in gone:
    types, n = re.subn(rf'<Override PartName="/{name}"[^>]*/>', "", types)
    assert n == 1, name
types = types.replace("</Types>", '<Override PartName="/word/header4.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>')
parts["[Content_Types].xml"] = types.encode("utf-8")

with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as out:
    for name in order + ["word/header4.xml", "word/_rels/header4.xml.rels"]:
        if name not in gone:
            out.writestr(name, parts[name])
print("wrote", dst, "sections:", doc.count("<w:sectPr>"))

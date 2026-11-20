"""Template v4 from v3: no guidance lines, since a calculation is assumed to be for submission
(the user, 2026-11-20).
- each guidance placeholder under a body heading becomes an empty paragraph in the same style
  (bullets where a list goes), so nothing prints until you type;
- the cover fields lose their hints: [Subject], [Design codes];
- the guidance's own entries in the glossary go.
Everything else stays byte for byte."""
import re
import sys
import zipfile

src, dst = sys.argv[1:3]
with zipfile.ZipFile(src) as z:
    order = [i.filename for i in z.infolist()]
    parts = {name: z.read(name) for name in order}

doc = parts["word/document.xml"].decode("utf-8")
glossary = parts["word/glossary/document.xml"].decode("utf-8")

# 1. Guidance: a temporary placeholder holding a whole paragraph.
guidance = re.compile(r"<w:sdt><w:sdtPr>(?:(?!</w:sdtPr>).)*?<w:docPart w:val=\"(\w+)\"/>(?:(?!</w:sdtPr>).)*?<w:temporary/>"
                      r"(?:(?!</w:sdtPr>).)*?</w:sdtPr>(?:<w:sdtEndPr/>)?<w:sdtContent>(<w:p\b[^>]*>)(<w:pPr>.*?</w:pPr>)?"
                      r"(?:(?!</w:p>).)*</w:p></w:sdtContent></w:sdt>", re.S)
removed = []


def empty_paragraph(m):
    removed.append(m.group(1))
    return m.group(2) + (m.group(3) or "") + "</w:p>"


doc, n = guidance.subn(empty_paragraph, doc)
assert n == 9, f"expected 9 guidance placeholders, found {n}"
for name in removed:
    glossary, k = re.subn(rf"<w:docPart><w:docPartPr><w:name w:val=\"{name}\"/>.*?</w:docPart>", "", glossary, flags=re.S)
    assert k == 1, name

# 2. The cover's hints.
for old, new in (("[Subject: the structure, element or question this calculation answers]", "[Subject]"),
                 ("[Codes, e.g. Eurocodes]", "[Design codes]")):
    assert doc.count(old) == 1 and glossary.count(old) == 1, old
    doc = doc.replace(old, new)
    glossary = glossary.replace(old, new)

parts["word/document.xml"] = doc.encode("utf-8")
parts["word/glossary/document.xml"] = glossary.encode("utf-8")
with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as out:
    for name in order:
        out.writestr(name, parts[name])
print("wrote", dst, "- guidance removed:", n)

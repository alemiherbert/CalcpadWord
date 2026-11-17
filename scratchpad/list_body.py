"""Lists a .docx/.dotx body: top-level paragraphs and content controls, with style and text."""
import re
import sys
import zipfile

doc = zipfile.ZipFile(sys.argv[1]).read("word/document.xml").decode("utf-8")
body = doc[doc.index("<w:body>"):]
depth = 0
for m in re.finditer(r"<w:p\b.*?</w:p>|<w:sdt>.*?</w:sdt>|<w:tbl>|</w:tbl>|<w:sectPr", body, re.S):
    x = m.group(0)
    if x == "<w:tbl>":
        print("  [table]")
        continue
    if x in ("</w:tbl>", "<w:sectPr"):
        if x == "<w:sectPr":
            print("  ---- section break ----")
        continue
    style = re.search(r'<w:pStyle w:val="([^"]+)"', x)
    tag = re.search(r'<w:tag w:val="([^"]+)"', x)
    text = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", x)).strip()
    if text or tag:
        print((style.group(1) if style else "-")[:16].ljust(16), ("[" + tag.group(1)[:40] + "] " if tag else "") + text[:100])

"""Makes the generic twin of a Dokimos template: the same file without the logo and the name.

    python make_generic.py "<Dokimos X vN.dotx>" "<X vN.dotx>"

Removes the logo (in the headers: a small table of the mark, an inline picture, and the name
in Word text) and its image files, and renames what
says Dokimos: field tags, the data namespace, the table style, the theme, and the file
properties. Everything else is copied byte for byte, so nothing moves. Fails if any "Dokimos"
is left.
"""
import re
import sys
import zipfile

RENAMES = [
    ("urn:dokimos-net:calculation", "urn:calcpad-word:calculation"),
    ('w:val="dokimos:', 'w:val="calc:'),
    ("DokimosTable", "CalculationTable"),
    ("Dokimos Table", "Calculation Table"),
    ('name="Dokimos"', 'name="Calculation"'),
    ("<Company>Dokimos Consulting Engineers</Company>", "<Company></Company>"),
    ("<Template>Dokimos ", "<Template>"),
]

# The innermost table holding an inline picture: the logo (the corner pattern is anchored).
# A table is always followed by a paragraph, so its cell or header stays valid without it.
LOGO_TABLE = re.compile(
    r"<w:tbl>(?:(?!<w:tbl>|</w:tbl>).)*<wp:inline\b(?:(?!<w:tbl>|</w:tbl>).)*</w:tbl>", re.S)
EMBED = re.compile(r'r:embed="([^"]+)"')
HEADER = re.compile(r"word/header\d+\.xml$")


def relationship(rels, rid):
    match = re.search(r'<Relationship [^>]*Id="%s"[^>]*/>' % re.escape(rid), rels)
    if not match:
        raise SystemExit("no relationship " + rid)
    return match.group(0)


def main(source, target):
    with zipfile.ZipFile(source) as z:
        infos = z.infolist()
        data = {i.filename: z.read(i.filename) for i in infos}

    logo_files = set()
    removed = 0
    for name in [n for n in data if HEADER.search(n)]:
        xml = data[name].decode("utf-8")
        logos = LOGO_TABLE.findall(xml)
        if not logos:
            continue
        rels_name = name.replace("word/", "word/_rels/") + ".rels"
        rels = data[rels_name].decode("utf-8")
        for logo in logos:
            for rid in EMBED.findall(logo):
                rel = relationship(rels, rid)
                logo_files.add("word/" + re.search(r'Target="([^"]+)"', rel).group(1))
                rels = rels.replace(rel, "")
            xml = xml.replace(logo, "")
            removed += 1
        data[name] = xml.encode("utf-8")
        data[rels_name] = rels.encode("utf-8")
    if removed == 0:
        raise SystemExit("no logo found in " + source)

    # Image files the logo used, unless something else still does.
    all_rels = "".join(v.decode("utf-8") for k, v in data.items() if k.endswith(".rels"))
    for f in logo_files:
        if 'Target="%s"' % f[len("word/"):] not in all_rels:
            del data[f]

    for name, content in data.items():
        if not name.endswith((".xml", ".rels")):
            continue
        text = content.decode("utf-8")
        for old, new in RENAMES:
            text = text.replace(old, new)
        data[name] = text.encode("utf-8")

    left = [n for n, c in data.items() if b"okimos" in c.lower() or b"okimos" in n.lower().encode()]
    if left:
        raise SystemExit("Dokimos is still in: " + ", ".join(left))

    with zipfile.ZipFile(target, "w") as out:
        for info in infos:
            if info.filename in data:
                out.writestr(info, data[info.filename], compress_type=zipfile.ZIP_DEFLATED)
    print("%s: %d logo(s) and %d image file(s) removed" % (target, removed, len([f for f in logo_files if f not in data])))


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit(__doc__)
    main(sys.argv[1], sys.argv[2])

import sys, zipfile, xml.dom.minidom
for path in sys.argv[1:]:
    z = zipfile.ZipFile(path)
    bad = 0
    for name in z.namelist():
        if name.endswith(".xml") or name.endswith(".rels"):
            try: xml.dom.minidom.parseString(z.read(name))
            except Exception as e: bad += 1; print("  BAD", name, e)
    print(path.split("\\")[-1], "parts:", len(z.namelist()), "malformed:", bad)

"""Replaces the cover pattern's pictures (SVG and its PNG fallback) in the user's document, keeping
every other part, and the pattern's position and size, exactly as the user left them."""
import sys
import zipfile

source, fallback, target = sys.argv[1:4]
with zipfile.ZipFile(fallback) as f:
    new_svg = f.read("word/media/image2.svg")
    new_png = f.read("word/media/image1.png")

replace = {"word/media/image4.svg": new_svg, "word/media/image3.png": new_png}
with zipfile.ZipFile(source) as src, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as out:
    for item in src.infolist():
        data = replace.pop(item.filename, None)
        out.writestr(item, data if data is not None else src.read(item.filename))
if replace:
    sys.exit("not found: " + ", ".join(replace))
print("replaced the pattern pictures:", target)

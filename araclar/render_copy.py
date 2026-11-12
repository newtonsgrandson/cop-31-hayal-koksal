"""Make a preview copy of a COP31-template deck whose EMF+ logo (COP31 Türkiye) is swapped for a PNG.

LibreOffice cannot draw the EMF+ logo embedded in the template, so PDFs/PNGs rendered from the
original file miss it. The .pptx that is delivered is left untouched; only this copy is rendered.

Usage: python render_copy.py <src.pptx> <dst.pptx> <cop31_logo.png>
"""
import sys
import zipfile

src, dst, logo = sys.argv[1:4]
zin = zipfile.ZipFile(src)
emfs = [n for n in zin.namelist() if n.startswith("ppt/media/") and n.lower().endswith(".emf")]
with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
    for info in zin.infolist():
        data = zin.read(info.filename)
        if info.filename.endswith(".rels"):
            x = data.decode("utf8")
            for e in emfs:
                x = x.replace("../media/" + e.split("/")[-1], "../media/cop31_render.png")
            data = x.encode("utf8")
        zout.writestr(info, data)
    zout.write(logo, "ppt/media/cop31_render.png")
print("render copy:", dst, "| swapped", emfs)

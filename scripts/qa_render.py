"""Contrôle visuel sans LibreOffice (macOS) : rend chaque slide d'un .pptx et chaque page d'un .pdf en PNG via Quick Look.
Quick Look ne rend que la première slide/page : on fait tourner chaque slide en première position (copie temporaire) et on découpe le PDF page par page.
Usage :
  qa_render.py pptx <deck.pptx> <dossier_sortie> [--size 3200] [--pdf <deck.pdf>]   # PNG par slide (+ assemblage PDF image optionnel)
  qa_render.py pdf  <doc.pdf>   <dossier_sortie> [--size 1400]                        # PNG par page
Limites : polices absentes du Mac substituées ; graphiques natifs PowerPoint NON rendus par Quick Look (préférer des graphiques dessinés en formes)."""
import sys, os, re, zipfile, subprocess, argparse


def render_pptx(src, out, size, pdf=None):
    os.makedirs(out, exist_ok=True)
    z = zipfile.ZipFile(src); entries = [(i.filename, z.read(i.filename)) for i in z.infolist()]
    pres = dict(entries)["ppt/presentation.xml"].decode(); ids = re.findall(r"<p:sldId [^>]*/>", pres)
    pngs = []
    for i in range(len(ids)):
        order = [ids[i]] + [x for j, x in enumerate(ids) if j != i]
        new = re.sub(r"<p:sldIdLst>.*?</p:sldIdLst>", "<p:sldIdLst>" + "".join(order) + "</p:sldIdLst>", pres, flags=re.S)
        tmp = os.path.join(out, f"_s{i + 1:02d}.pptx")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
            for fn, data in entries: zo.writestr(fn, new.encode() if fn == "ppt/presentation.xml" else data)
        subprocess.run(["qlmanage", "-t", "-s", str(size), "-o", out, tmp], capture_output=True)
        png = tmp + ".png"; final = os.path.join(out, f"slide-{i + 1:02d}.png")
        if os.path.exists(png): os.replace(png, final); pngs.append(final)
        os.remove(tmp)
    print(f"{len(pngs)} slides rendues dans {out}")
    if pdf and pngs:
        from PIL import Image
        imgs = [Image.open(p).convert("RGB") for p in pngs]
        imgs[0].save(pdf, save_all=True, append_images=imgs[1:], resolution=size / 10.0)
        print("PDF image :", pdf)


def render_pdf(src, out, size):
    from pypdf import PdfReader, PdfWriter
    os.makedirs(out, exist_ok=True)
    r = PdfReader(src); n = 0
    for i, p in enumerate(r.pages, 1):
        tmp = os.path.join(out, f"_p{i:02d}.pdf"); w = PdfWriter(); w.add_page(p); w.write(tmp)
        subprocess.run(["qlmanage", "-t", "-s", str(size), "-o", out, tmp], capture_output=True)
        png = tmp + ".png"
        if os.path.exists(png): os.replace(png, os.path.join(out, f"page-{i:02d}.png")); n += 1
        os.remove(tmp)
    print(f"{n}/{len(r.pages)} pages rendues dans {out}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(); ap.add_argument("kind", choices=["pptx", "pdf"]); ap.add_argument("src"); ap.add_argument("out")
    ap.add_argument("--size", type=int, default=None); ap.add_argument("--pdf", default=None); a = ap.parse_args()
    if a.kind == "pptx": render_pptx(a.src, a.out, a.size or 3200, a.pdf)
    else: render_pdf(a.src, a.out, a.size or 1400)

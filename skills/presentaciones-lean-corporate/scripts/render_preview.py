#!/usr/bin/env python3
"""
render_preview.py — Renderiza un .pptx a imágenes para revisarlo visualmente.

    python render_preview.py presentacion.pptx              # -> preview/slide-01.png ... + contact_sheet.png
    python render_preview.py presentacion.pptx --dpi 110 --out preview_v2
    python render_preview.py presentacion.pptx --slides 3,5  # solo esas diapositivas

Requiere LibreOffice (soffice) y poppler (pdftoppm). El render de LibreOffice es
muy fiel para formas, texto y gráficos; pequeñas diferencias de kerning respecto
a PowerPoint son normales (por eso se dejan márgenes de seguridad en los layouts).

SIEMPRE mira las imágenes resultantes antes de entregar: es la única forma de
detectar solapamientos, textos cortados o desequilibrios visuales.
"""
import argparse
import glob
import math
import os
import shutil
import subprocess
import sys
import tempfile


def to_pdf(pptx, workdir):
    profile = tempfile.mkdtemp(prefix="lo_profile_")
    cmd = ["soffice", f"-env:UserInstallation=file://{profile}", "--headless",
           "--convert-to", "pdf", "--outdir", workdir, pptx]
    subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                   timeout=240)
    shutil.rmtree(profile, ignore_errors=True)
    pdf = os.path.join(workdir, os.path.splitext(os.path.basename(pptx))[0] + ".pdf")
    if not os.path.exists(pdf):
        sys.exit("ERROR: LibreOffice no generó el PDF")
    return pdf


def contact_sheet(pngs, out, cols=4, thumb_w=480):
    from PIL import Image, ImageDraw, ImageFont
    thumbs = []
    for p in pngs:
        im = Image.open(p).convert("RGB")
        r = thumb_w / im.width
        thumbs.append(im.resize((thumb_w, int(im.height * r))))
    th = thumbs[0].height
    pad, label_h = 16, 26
    rows = math.ceil(len(thumbs) / cols)
    sheet = Image.new("RGB", (cols * (thumb_w + pad) + pad,
                              rows * (th + pad + label_h) + pad), (232, 234, 238))
    draw = ImageDraw.Draw(sheet)
    try:
        font = ImageFont.truetype("DejaVuSans-Bold.ttf", 16)
    except Exception:
        font = ImageFont.load_default()
    for i, t in enumerate(thumbs):
        r, c = divmod(i, cols)
        x = pad + c * (thumb_w + pad)
        y = pad + r * (th + pad + label_h)
        draw.text((x, y), f"{i + 1}", fill=(60, 60, 60), font=font)
        sheet.paste(t, (x, y + label_h))
        draw.rectangle([x - 1, y + label_h - 1, x + thumb_w, y + label_h + th],
                       outline=(190, 194, 200))
    sheet.save(out)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--out", default="preview")
    ap.add_argument("--dpi", type=int, default=80)
    ap.add_argument("--slides", help="Lista separada por comas (1-based)")
    ap.add_argument("--no-sheet", action="store_true")
    a = ap.parse_args()
    os.makedirs(a.out, exist_ok=True)
    for f in glob.glob(os.path.join(a.out, "slide-*.png")):
        os.remove(f)
    work = tempfile.mkdtemp(prefix="render_")
    pdf = to_pdf(os.path.abspath(a.pptx), work)
    subprocess.run(["pdftoppm", "-png", "-r", str(a.dpi), pdf,
                    os.path.join(work, "p")], check=True, stderr=subprocess.DEVNULL)
    pages = sorted(glob.glob(os.path.join(work, "p-*.png")))
    wanted = None
    if a.slides:
        wanted = {int(x) for x in a.slides.split(",")}
    out_files = []
    for i, p in enumerate(pages, 1):
        if wanted and i not in wanted:
            continue
        dst = os.path.join(a.out, f"slide-{i:02d}.png")
        shutil.move(p, dst)
        out_files.append(dst)
    shutil.copy(pdf, os.path.join(a.out, "preview.pdf"))
    shutil.rmtree(work, ignore_errors=True)
    print(f"{len(out_files)} imágenes en {a.out}/")
    if out_files and not a.no_sheet:
        sheet = contact_sheet(out_files, os.path.join(a.out, "contact_sheet.png"))
        print("Hoja de contactos:", sheet)


if __name__ == "__main__":
    main()

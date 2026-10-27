#!/usr/bin/env python3
"""
pptx_ops.py — Operaciones puntuales sobre un .pptx existente sin reconstruirlo.

    python pptx_ops.py in.pptx out.pptx replace "2025" "2026"            # en todo el deck, conserva formato
    python pptx_ops.py in.pptx out.pptx delete 4,7                      # borra diapositivas (1-based)
    python pptx_ops.py in.pptx out.pptx move 12 3                       # mueve la 12 a la posición 3
    python pptx_ops.py in.pptx out.pptx notes 5 "Texto de las notas"
    python pptx_ops.py in.pptx out.pptx chart 6 datos.json              # nuevos datos para el 1er gráfico de la diap. 6
          (datos.json: {"categories": [...], "series": [{"name": "...", "values": [...]}]})

También importable: from pptx_ops import replace_text, delete_slides, move_slide, set_notes, update_chart
"""
import json
import sys

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.oxml.ns import qn


def _iter_text_frames(shapes):
    for sh in shapes:
        if sh.shape_type == 6:  # grupo
            yield from _iter_text_frames(sh.shapes)
        if sh.has_text_frame:
            yield sh.text_frame
        if getattr(sh, "has_table", False) and sh.has_table:
            for row in sh.table.rows:
                for cell in row.cells:
                    yield cell.text_frame


def replace_text(prs, old, new):
    """Reemplaza texto run a run (conserva el formato). Si `old` está partido en
    varios runs del mismo párrafo, fusiona esos runs en el primero."""
    count = 0
    for slide in prs.slides:
        for tf in _iter_text_frames(slide.shapes):
            for p in tf.paragraphs:
                if old not in p.text:
                    continue
                hit = False
                for r in p.runs:
                    if old in r.text:
                        r.text = r.text.replace(old, new)
                        hit = True
                        count += 1
                if not hit and p.runs:  # texto partido entre runs
                    full = p.text.replace(old, new)
                    p.runs[0].text = full
                    for r in p.runs[1:]:
                        r._r.getparent().remove(r._r)
                    count += 1
    return count


def delete_slides(prs, numbers):
    sldIdLst = prs.slides._sldIdLst
    ids = list(sldIdLst)
    for n in sorted(set(numbers), reverse=True):
        sid = ids[n - 1]
        prs.part.drop_rel(sid.get(qn("r:id")))
        sldIdLst.remove(sid)


def move_slide(prs, src, dst):
    sldIdLst = prs.slides._sldIdLst
    el = list(sldIdLst)[src - 1]
    sldIdLst.remove(el)
    sldIdLst.insert(dst - 1, el)


def set_notes(prs, n, text):
    prs.slides[n - 1].notes_slide.notes_text_frame.text = text


def update_chart(prs, n, data, index=0):
    charts = [sh for sh in prs.slides[n - 1].shapes if getattr(sh, "has_chart", False) and sh.has_chart]
    if not charts:
        raise SystemExit(f"La diapositiva {n} no tiene gráficos")
    cd = CategoryChartData()
    cd.categories = data["categories"]
    for s in data["series"]:
        cd.add_series(s["name"], s["values"])
    charts[index].chart.replace_data(cd)  # mantiene formato y colores


def main():
    if len(sys.argv) < 4:
        sys.exit(__doc__)
    src, out, op, *args = sys.argv[1:]
    prs = Presentation(src)
    if op == "replace":
        print(f"{replace_text(prs, args[0], args[1])} reemplazos")
    elif op == "delete":
        delete_slides(prs, [int(x) for x in args[0].split(",")])
    elif op == "move":
        move_slide(prs, int(args[0]), int(args[1]))
    elif op == "notes":
        set_notes(prs, int(args[0]), args[1])
    elif op == "chart":
        with open(args[1], encoding="utf-8") as fh:
            update_chart(prs, int(args[0]), json.load(fh))
    else:
        sys.exit(f"Operación desconocida: {op}")
    prs.save(out)
    print(f"OK: {out}")


if __name__ == "__main__":
    main()

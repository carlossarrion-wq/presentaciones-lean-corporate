#!/usr/bin/env python3
"""Edición QUIRÚRGICA de celdas directamente en el XML del .xlsx/.xlsm.

Cuándo usarlo en lugar de openpyxl: el libro contiene elementos que openpyxl pierde o
degrada al guardar (segmentaciones/slicers, formas y cuadros de texto, controles de
formulario, minigráficos/sparklines, formato condicional avanzado, gráficos con formato
complejo, Power Query/conexiones, temas personalizados...). Este script reescribe SOLO las
celdas indicadas y deja el resto del paquete byte a byte idéntico.

Uso:
    python patch_cells.py libro.xlsx -o libro_mod.xlsx \
        --set "Ventas!B5=1250.5" \
        --set "Ventas!C5==SUM(C2:C4)" \
        --set "Ventas!A5=Total"            # texto
        --set "'Hoja con espacios'!D2=true"  # booleano
        --set "Ventas!E5="                 # vaciar la celda (conserva el formato)
    python patch_cells.py libro.xlsx -o out.xlsx --edits cambios.json

cambios.json: [{"sheet": "Ventas", "cell": "B5", "value": 1250.5},
               {"sheet": "Ventas", "cell": "C5", "formula": "SUM(C2:C4)"},
               {"sheet": "Ventas", "cell": "A5", "value": "Total"},
               {"sheet": "Ventas", "cell": "F5", "value": "2026-03-31", "type": "date"}]

Reglas de interpretación de --set (valor tras el primer '='):
    '=...'  → fórmula     'true'/'false' → booleano     número → número
    texto   → texto (usa el prefijo ' para forzar texto: "'00123")

Después: el libro se marca con fullCalcOnLoad para que Excel recalcule al abrir, y se
elimina calcChain.xml (Excel lo regenera). Valida con: python recalc.py out.xlsx --check-only
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import posixpath
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NSR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NSP = "http://schemas.openxmlformats.org/package/2006/relationships"
NSCT = "http://schemas.openxmlformats.org/package/2006/content-types"
q = lambda tag: f"{{{NS}}}{tag}"  # noqa: E731

CELL_RE = re.compile(r"^([A-Z]{1,3})(\d+)$")


def col_to_idx(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def parse_target(target: str) -> tuple[str, str]:
    sheet, cell = target.rsplit("!", 1)
    if sheet.startswith("'") and sheet.endswith("'"):
        sheet = sheet[1:-1].replace("''", "'")
    cell = cell.replace("$", "").upper()
    if not CELL_RE.match(cell):
        raise ValueError(f"Referencia de celda no válida: {cell}")
    return sheet, cell


def parse_set(expr: str) -> dict:
    # Separar en el primer '=' que sigue a la referencia de celda (Hoja!A1=valor)
    m = re.match(r"^(.*?![$A-Za-z]+\$?\d+)=(.*)$", expr, re.S)
    if not m:
        raise ValueError(f"Formato --set inválido: {expr!r} (esperado Hoja!A1=valor)")
    sheet, cell = parse_target(m.group(1))
    raw = m.group(2)
    edit = {"sheet": sheet, "cell": cell}
    if raw == "":
        edit["value"] = None
    elif raw.startswith("="):
        edit["formula"] = raw[1:]
    elif raw.startswith("'"):
        edit["value"] = raw[1:]
    elif raw.lower() in ("true", "false"):
        edit["value"] = raw.lower() == "true"
    else:
        try:
            edit["value"] = int(raw) if re.fullmatch(r"-?\d+", raw) else float(raw)
        except ValueError:
            edit["value"] = raw
    return edit


def excel_serial(value: str) -> float:
    d = dt.datetime.fromisoformat(value)
    base = dt.datetime(1899, 12, 30)
    delta = d - base
    return delta.days + delta.seconds / 86400


def sheet_paths(z: zipfile.ZipFile) -> dict[str, str]:
    wb = etree.fromstring(z.read("xl/workbook.xml"))
    rels = etree.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rid_to_target = {r.get("Id"): r.get("Target") for r in rels.findall(f"{{{NSP}}}Relationship")}
    out = {}
    for s in wb.find(q("sheets")).findall(q("sheet")):
        target = rid_to_target[s.get(f"{{{NSR}}}id")]
        path = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
        out[s.get("name")] = path
    return out


def get_or_create_cell(sheet_data, cell_ref: str):
    col, row = CELL_RE.match(cell_ref).groups()
    row_i = int(row)
    row_el = None
    for r in sheet_data.findall(q("row")):
        ri = int(r.get("r"))
        if ri == row_i:
            row_el = r
            break
        if ri > row_i:
            row_el = etree.Element(q("row"), r=str(row_i))
            r.addprevious(row_el)
            break
    if row_el is None:
        row_el = etree.SubElement(sheet_data, q("row"), r=str(row_i))
    row_el.attrib.pop("spans", None)
    target_idx = col_to_idx(col)
    for c in row_el.findall(q("c")):
        cref = c.get("r")
        cc, _ = CELL_RE.match(cref).groups()
        ci = col_to_idx(cc)
        if ci == target_idx:
            return c
        if ci > target_idx:
            new = etree.Element(q("c"), r=cell_ref)
            c.addprevious(new)
            return new
    return etree.SubElement(row_el, q("c"), r=cell_ref)


def apply_edit(cell, edit: dict):
    f_old = cell.find(q("f"))
    if f_old is not None and f_old.get("t") == "shared" and f_old.get("ref"):
        raise RuntimeError(
            f"{edit['sheet']}!{edit['cell']} es la celda maestra de una fórmula compartida "
            f"(rango {f_old.get('ref')}). Modificarla rompería las celdas dependientes: edita "
            "ese rango completo con openpyxl o escribe la fórmula en cada celda del rango.")
    if f_old is not None and f_old.get("t") == "array":
        raise RuntimeError(f"{edit['sheet']}!{edit['cell']} contiene una fórmula matricial.")
    style = cell.get("s")
    for child in list(cell):
        cell.remove(child)
    for attr in list(cell.attrib):
        if attr not in ("r", "s"):
            del cell.attrib[attr]
    if style is not None:
        cell.set("s", style)

    if "formula" in edit:
        f = etree.SubElement(cell, q("f"))
        f.text = edit["formula"].lstrip("=")
        return
    value = edit.get("value")
    if value is None:
        return
    if edit.get("type") == "date":
        etree.SubElement(cell, q("v")).text = repr(excel_serial(str(value)))
        if style is None:
            print(f"Aviso: {edit['cell']} no tiene estilo; la fecha se verá como número. "
                  "Aplica un formato de fecha (o elige una celda ya formateada).", file=sys.stderr)
        return
    if isinstance(value, bool):
        cell.set("t", "b")
        etree.SubElement(cell, q("v")).text = "1" if value else "0"
    elif isinstance(value, (int, float)):
        etree.SubElement(cell, q("v")).text = repr(value) if isinstance(value, float) else str(value)
    else:
        cell.set("t", "inlineStr")
        is_ = etree.SubElement(cell, q("is"))
        t = etree.SubElement(is_, q("t"))
        t.text = str(value)
        if t.text != t.text.strip():
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def patch(src: Path, dst: Path, edits: list[dict]):
    with zipfile.ZipFile(src) as z:
        paths = sheet_paths(z)
        by_part: dict[str, list[dict]] = {}
        for e in edits:
            if e["sheet"] not in paths:
                raise KeyError(f"Hoja '{e['sheet']}' no existe. Hojas: {list(paths)}")
            by_part.setdefault(paths[e["sheet"]], []).append(e)

        new_parts: dict[str, bytes] = {}
        for part, part_edits in by_part.items():
            root = etree.fromstring(z.read(part))
            sd = root.find(q("sheetData"))
            for e in part_edits:
                apply_edit(get_or_create_cell(sd, e["cell"]), e)
            new_parts[part] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

        # Forzar recálculo al abrir y eliminar calcChain
        wb = etree.fromstring(z.read("xl/workbook.xml"))
        calc = wb.find(q("calcPr"))
        if calc is None:
            calc = etree.SubElement(wb, q("calcPr"))
            # calcPr debe ir tras <definedNames>/<sheets>; reubicar antes de elementos posteriores
            for tag in ("oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes",
                        "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"):
                nxt = wb.find(q(tag))
                if nxt is not None:
                    nxt.addprevious(calc)
                    break
        calc.set("fullCalcOnLoad", "1")
        new_parts["xl/workbook.xml"] = etree.tostring(wb, xml_declaration=True, encoding="UTF-8", standalone=True)

        drop = {"xl/calcChain.xml"}
        rels = etree.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        for r in rels.findall(f"{{{NSP}}}Relationship"):
            if r.get("Target", "").endswith("calcChain.xml"):
                rels.remove(r)
        new_parts["xl/_rels/workbook.xml.rels"] = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)
        ct = etree.fromstring(z.read("[Content_Types].xml"))
        for o in ct.findall(f"{{{NSCT}}}Override"):
            if o.get("PartName") == "/xl/calcChain.xml":
                ct.remove(o)
        new_parts["[Content_Types].xml"] = etree.tostring(ct, xml_declaration=True, encoding="UTF-8", standalone=True)

        tmp = dst.with_suffix(dst.suffix + ".tmp")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
            for info in z.infolist():
                if info.filename in drop:
                    continue
                data = new_parts.get(info.filename, None)
                if data is None:
                    out.writestr(info, z.read(info.filename))
                else:
                    out.writestr(info, data)
    tmp.replace(dst)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file")
    ap.add_argument("-o", "--output", required=True, help="Fichero de salida (puede ser el mismo)")
    ap.add_argument("--set", action="append", default=[], help="Hoja!A1=valor (repetible)")
    ap.add_argument("--edits", help="JSON con la lista de cambios")
    a = ap.parse_args()
    edits = [parse_set(s) for s in a.set]
    if a.edits:
        for e in json.loads(Path(a.edits).read_text(encoding="utf-8")):
            e["cell"] = e["cell"].replace("$", "").upper()
            edits.append(e)
    if not edits:
        sys.exit("No hay cambios: usa --set o --edits")
    patch(Path(a.file), Path(a.output), edits)
    print(json.dumps({"status": "ok", "edits": len(edits), "output": a.output}, ensure_ascii=False))


if __name__ == "__main__":
    main()

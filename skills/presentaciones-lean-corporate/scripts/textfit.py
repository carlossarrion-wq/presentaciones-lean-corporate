"""
textfit.py — Medición aproximada de texto para evitar desbordamientos.

PowerPoint no expone un motor de maquetación a python-pptx, así que estimamos
cuántas líneas ocupa un texto con las métricas reales de la fuente (vía PIL).
Si la fuente exacta no está instalada se usa su equivalente métrico:
  Calibri -> Carlito, Arial -> Liberation Sans, Cambria -> Caladea,
  Times New Roman -> Liberation Serif.
Si tampoco existe, se usa una aproximación por anchura media de carácter.

Uso típico:
    from textfit import fit_size, text_height
    size = fit_size(paragraphs, width_in, height_in, "Calibri", bold=False,
                    max_pt=20, min_pt=12)
"""
from __future__ import annotations

import functools
import re
import shutil
import subprocess

try:
    from PIL import ImageFont
except ImportError:  # pragma: no cover
    ImageFont = None

METRIC_EQUIVALENTS = {
    "calibri": "Carlito",
    "calibri light": "Carlito",
    "arial": "Liberation Sans",
    "helvetica": "Liberation Sans",
    "cambria": "Caladea",
    "times new roman": "Liberation Serif",
    "courier new": "Liberation Mono",
}

# Altura de línea relativa al tamaño de fuente (aprox. PowerPoint, interlineado 1.0)
LINE_HEIGHT = {
    "ibm plex sans": 1.3, "ibm plex sans light": 1.3,
    "calibri": 1.22, "carlito": 1.22,
    "arial": 1.15, "liberation sans": 1.15,
    "cambria": 1.17, "caladea": 1.17,
}
DEFAULT_LINE_HEIGHT = 1.2


@functools.lru_cache(maxsize=64)
def _font_file(family: str, bold: bool) -> str | None:
    if not shutil.which("fc-match"):
        return None
    candidates = [family]
    eq = METRIC_EQUIVALENTS.get(family.lower())
    if eq:
        candidates.insert(0, eq)  # preferimos el equivalente métrico conocido
    style = ":bold" if bold else ":regular"
    for fam in candidates:
        try:
            out = subprocess.run(
                ["fc-match", "-f", "%{file}|%{family}", f"{fam}{style}"],
                capture_output=True, text=True, timeout=5,
            ).stdout
        except Exception:
            continue
        if "|" in out:
            path, matched = out.split("|", 1)
            # "IBM Plex Sans Light" en negrita resuelve a la familia base "IBM Plex Sans"
            base = re.sub(r"\s+(light|medium|medm|semibold|smbld|text|thin|extralight|extlt)$",
                          "", fam.lower())
            if fam.lower() in matched.lower() or base in matched.lower():
                return path
    return None


@functools.lru_cache(maxsize=128)
def _pil_font(family: str, bold: bool):
    if ImageFont is None:
        return None
    path = _font_file(family, bold)
    if not path:
        return None
    try:
        return ImageFont.truetype(path, 100)  # medimos a 100 pt y escalamos
    except Exception:
        return None


def text_width_pt(text: str, family: str, size_pt: float, bold: bool = False) -> float:
    font = _pil_font(family, bold)
    if font is None:
        factor = 0.55 if bold else 0.5  # anchura media de carácter
        return len(text) * size_pt * factor
    return font.getlength(text) * size_pt / 100.0


def wrap_lines(text: str, width_pt: float, family: str, size_pt: float, bold=False) -> int:
    """Número de líneas que ocupa `text` en una caja de `width_pt` puntos."""
    total = 0
    for hard_line in str(text).split("\n"):
        words = hard_line.split()
        if not words:
            total += 1
            continue
        lines, current = 1, ""
        for w in words:
            trial = (current + " " + w).strip()
            if text_width_pt(trial, family, size_pt, bold) <= width_pt:
                current = trial
            else:
                if current:
                    lines += 1
                current = w
                # palabra más larga que la caja: se parte
                while text_width_pt(current, family, size_pt, bold) > width_pt and len(current) > 1:
                    lines += 1
                    cut = max(1, int(len(current) * width_pt / text_width_pt(current, family, size_pt, bold)))
                    current = current[cut:]
        total += lines
    return total


def line_height_factor(family: str) -> float:
    return LINE_HEIGHT.get(family.lower(), DEFAULT_LINE_HEIGHT)


def text_height(paragraphs, width_in, family, size_pt, bold=False,
                line_spacing=1.0, space_after_pt=0.0, indent_in=0.0) -> float:
    """Altura (pulgadas) que ocupan los párrafos. `paragraphs`: lista de str o (str, level)."""
    width_pt = width_in * 72
    lh = size_pt * line_height_factor(family) * line_spacing
    h = 0.0
    for p in paragraphs:
        level = 0
        if isinstance(p, (tuple, list)):
            p, level = p[0], p[1]
        w = width_pt - (indent_in * 72) * (1 + level)
        h += wrap_lines(p, max(w, 10), family, size_pt, bold) * lh + space_after_pt
    if paragraphs:
        h -= space_after_pt
    return h / 72.0


def fit_size(paragraphs, width_in, height_in, family, bold=False, max_pt=24, min_pt=10,
             line_spacing=1.0, space_after_ratio=0.0, indent_in=0.0, step=0.5) -> float:
    """Mayor tamaño (pt) entre min_pt y max_pt con el que el texto cabe en la caja.
    space_after_ratio: espacio tras párrafo como fracción del tamaño de fuente."""
    size = max_pt
    while size >= min_pt:
        h = text_height(paragraphs, width_in, family, size, bold, line_spacing,
                        space_after_pt=size * space_after_ratio, indent_in=indent_in)
        if h <= height_in:
            return size
        size -= step
    return min_pt


def fits(paragraphs, width_in, height_in, family, size_pt, bold=False, line_spacing=1.0,
         space_after_pt=0.0, indent_in=0.0, tolerance=1.02) -> bool:
    h = text_height(paragraphs, width_in, family, size_pt, bold, line_spacing,
                    space_after_pt, indent_in)
    return h <= height_in * tolerance

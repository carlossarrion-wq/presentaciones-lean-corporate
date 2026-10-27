#!/usr/bin/env python3
"""
build_deck.py — Genera un .pptx profesional a partir de una especificación JSON.

    python build_deck.py deck.json -o salida.pptx
    python build_deck.py deck.json -o salida.pptx --theme consultoria
    python build_deck.py deck.json -o salida.pptx --theme-file mi_marca.json
    python build_deck.py deck.json -o salida.pptx --template plantilla_corporativa.pptx
    python build_deck.py --list-layouts

La especificación describe QUÉ dice cada diapositiva (layout + contenido); este
script decide CÓMO se dibuja: rejilla, márgenes, jerarquía tipográfica, colores,
ajuste de texto para que nada se desborde, gráficos nativos editables, tablas
estilizadas, notas del orador, numeración y pie.

Ver references/spec-format.md para el formato completo y references/layouts.md
para cuándo usar cada layout.
"""
from __future__ import annotations

import argparse
import copy
import json
import math
import os
import re
import sys
import tempfile
import zipfile

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from textfit import fit_size, line_height_factor, text_height, wrap_lines  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
THEMES_PATH = os.path.join(HERE, "..", "themes", "themes.json")
TEMPLATES_DIR = os.path.join(HERE, "..", "templates")
DEFAULT_PROFILE = "lean-corporate"   # esta versión del skill usa la plantilla Lean por defecto

# ----------------------------------------------------------------------------
# Rejilla (pulgadas, lienzo 16:9 = 13.333 x 7.5)
# ----------------------------------------------------------------------------
W, H = 13.333, 7.5
M = 0.6                 # margen lateral
TOP = 0.45              # inicio del título
FOOTER_Y = 7.02         # línea base del pie
BOTTOM = 6.78           # límite inferior del contenido (sin fuente ni conclusión)
GAP = 0.3               # separación estándar entre bloques

SCALE = 1.0            # pulgadas reales por pulgada de diseño (2.0 en la plantilla Lean)


def IN(v):
    """Pulgadas de diseño -> EMU reales."""
    return Inches(v * SCALE)


def PT(v):
    """Puntos de diseño -> puntos reales."""
    return Pt(v * SCALE)


ALIGN = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}
ANCHOR = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM}


# ----------------------------------------------------------------------------
# Utilidades básicas
# ----------------------------------------------------------------------------
def rgb(hex_str: str) -> RGBColor:
    return RGBColor.from_string(hex_str.replace("#", "").upper())


def mix(c1: str, c2: str, t: float) -> str:
    """Interpola dos colores hex (t=0 -> c1, t=1 -> c2)."""
    a = [int(c1[i:i + 2], 16) for i in (0, 2, 4)]
    b = [int(c2[i:i + 2], 16) for i in (0, 2, 4)]
    return "".join(f"{round(x + (y - x) * t):02X}" for x, y in zip(a, b))


EMPH = re.compile(r"\*\*(.+?)\*\*")


def plain(text) -> str:
    return EMPH.sub(r"\1", str(text or ""))


def runs_of(text):
    """'Texto con **énfasis**' -> [(texto, False), (énfasis, True)]."""
    out, pos = [], 0
    text = str(text or "")
    for m in EMPH.finditer(text):
        if m.start() > pos:
            out.append((text[pos:m.start()], False))
        out.append((m.group(1), True))
        pos = m.end()
    if pos < len(text):
        out.append((text[pos:], False))
    return out or [("", False)]


def norm_items(items):
    """Acepta ["a", {"text": "b", "level": 1}, {"text": "c", "sub": ["d"]}] -> [(txt, lvl)]."""
    out = []
    for it in items or []:
        if isinstance(it, str):
            out.append((it, 0))
        elif isinstance(it, dict):
            out.append((it.get("text", ""), int(it.get("level", 0))))
            for sub in it.get("sub", []) or []:
                out.append((sub if isinstance(sub, str) else sub.get("text", ""), 1))
        elif isinstance(it, (list, tuple)):
            out.append((it[0], int(it[1])))
    return out


class SpecError(Exception):
    pass


# ----------------------------------------------------------------------------
# Contexto de dibujo
# ----------------------------------------------------------------------------
class Deck:
    def __init__(self, theme: dict, meta: dict, template: str | None = None,
                 profile: dict | None = None):
        global W, H, M, BOTTOM, SCALE
        self.t = theme
        self.meta = meta
        self.profile = profile
        self.template = template
        if template:
            self.prs = Presentation(_load_template(template))
            _remove_all_slides(self.prs)
        else:
            self.prs = Presentation()
            self.prs.slide_width = Inches(13.333)  # noqa
            self.prs.slide_height = Inches(7.5)  # noqa
        real_w = self.prs.slide_width / 914400
        real_h = self.prs.slide_height / 914400
        # Trabajamos siempre en "pulgadas de diseño" (ancho 13,333) y escalamos al lienzo real
        SCALE = real_w / 13.333
        W, H = 13.333, round(real_h / SCALE, 3)
        if abs(W / H - 16 / 9) > 0.02:
            print(f"AVISO: el lienzo es {real_w:.2f}x{real_h:.2f} in (no 16:9). Los layouts "
                  "están optimizados para 16:9.", file=sys.stderr)
        if profile:
            g = profile.get("grid", {})
            M = g.get("margin", M)
            BOTTOM = g.get("bottom", BOTTOM)
            self.sizes = profile.get("sizes", {})
            self.roles = profile.get("roles", {})
            names = {l.name: l for l in self.prs.slide_layouts}
            for role, r in self.roles.items():
                if r["layout"] not in names:
                    raise SpecError(f"El perfil pide el layout '{r['layout']}' para '{role}', "
                                    f"pero la plantilla no lo tiene: {', '.join(names)}")
            self.layout_by_name = names
        self.blank = _pick_blank_layout(self.prs)
        self.page = 0
        self.hf = self.t["font_heading"]
        self.bf = self.t["font_body"]
        self.ef = self.t.get("font_emph", self.hf)
        self.acc_txt = self.t.get("accent_text", self.t["accent"])
        self.current_section = None   # sección vigente (se arrastra entre diapositivas)
        self.warnings = []

    # -- diapositiva -------------------------------------------------------
    def new_slide(self, dark=False, bg=None, role="content"):
        """Nueva diapositiva. Con perfil de plantilla usa el layout del rol y conserva
        solo los marcadores que el rol declara (title/subtitle/body)."""
        if self.profile:
            r = self.roles.get(role) or self.roles["content"]
            slide = self.prs.slides.add_slide(self.layout_by_name[r["layout"]])
            keep = {v for k, v in r.items() if k in ("title", "subtitle", "body")}
            for ph in list(slide.placeholders):
                if ph.placeholder_format.idx not in keep:
                    ph._element.getparent().remove(ph._element)
            self.page += 1
            return slide
        slide = self.prs.slides.add_slide(self.blank)
        for ph in list(slide.placeholders):
            ph._element.getparent().remove(ph._element)
        self.page += 1
        keep_template_bg = self.template and self.meta.get("use_template_background", True)
        if dark:
            self.rect(slide, 0, 0, W, H, self.t["bg_dark"])
        elif bg:
            self.rect(slide, 0, 0, W, H, bg)
        elif not keep_template_bg:
            self.rect(slide, 0, 0, W, H, self.t["bg"])
        return slide

    @staticmethod
    def materialize(ph):
        """Fija en la diapositiva la geometría heredada del layout (antes de moverla)."""
        l, t, w, h = ph.left, ph.top, ph.width, ph.height
        ph.left, ph.top, ph.width, ph.height = l, t, w, h
        return ph

    def ph(self, slide, idx):
        for p in slide.placeholders:
            if p.placeholder_format.idx == idx:
                return p
        return None

    def drop_ph(self, slide, idx):
        p = self.ph(slide, idx)
        if p is not None:
            p._element.getparent().remove(p._element)

    def fill_ph(self, slide, idx, paragraphs, size=None, bold=None, color=None, font=None,
                emph_color=None, align=None, autonum=False, para_sizes=None, para_bold=None,
                para_colors=None, space_before=None):
        """Escribe en un marcador de la plantilla conservando su formato heredado.
        Solo se fija explícitamente lo que se pasa (tamaño, negrita, color, fuente)."""
        p_ = self.ph(slide, idx)
        if p_ is None:
            return None
        tf = p_.text_frame
        tf.clear()
        if isinstance(paragraphs, str):
            paragraphs = [paragraphs]
        for i, ptxt in enumerate(paragraphs):
            para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            if align:
                para.alignment = ALIGN[align]
            if space_before is not None and i > 0:
                para.space_before = PT(space_before)
            if autonum:
                pPr = para._p.get_or_add_pPr()
                pPr.set("marL", str(int(IN(0.46))))
                pPr.set("indent", str(-int(IN(0.46))))
                an = etree.SubElement(pPr, qn("a:buAutoNum"))
                an.set("type", "arabicPeriod")
            psize = (para_sizes[i] if para_sizes and i < len(para_sizes) else size)
            pbold = (para_bold[i] if para_bold and i < len(para_bold) else bold)
            pcol = (para_colors[i] if para_colors and i < len(para_colors) else color)
            for txt, em in runs_of(ptxt):
                r = para.add_run()
                r.text = txt
                f = r.font
                if psize:
                    f.size = PT(psize)
                if pbold or em:
                    f.bold = True
                    f.name = font if font == self.hf else self.ef
                elif font:
                    f.name = font
                if em and emph_color:
                    f.color.rgb = rgb(emph_color)
                elif pcol:
                    f.color.rgb = rgb(pcol)
        return p_

    def lead(self, slide, s, top, x=None, w=None):
        """Párrafo introductorio bajo el título (estilo documento-oferta)."""
        if not s.get("lead"):
            return top
        x = M if x is None else x
        w = (W - 2 * M) if w is None else w
        paras = s["lead"] if isinstance(s["lead"], list) else [s["lead"]]
        sz = self.sizes.get("lead", 15) if self.profile else 15
        mn = self.sizes.get("lead_min", 12) if self.profile else 12
        avail = (BOTTOM - top) * 0.4
        size = fit_size([plain(p) for p in paras], w, avail, self.bf, max_pt=sz, min_pt=mn,
                        line_spacing=1.05, space_after_ratio=0.35)
        h = text_height([plain(p) for p in paras], w, self.bf, size, line_spacing=1.05,
                        space_after_pt=size * 0.35) + 0.04
        self.text(slide, x, top, w, h, paras, size=size, line_spacing=1.05,
                  space_after=size * 0.35, color=self.t["text"])
        return top + h + 0.25

    # -- formas ------------------------------------------------------------
    def rect(self, slide, x, y, w, h, fill=None, line=None, line_w=0.75,
             shape=MSO_SHAPE.RECTANGLE, radius=None):
        shp = slide.shapes.add_shape(shape, IN(x), IN(y), IN(w), IN(h))
        if fill:
            shp.fill.solid()
            shp.fill.fore_color.rgb = rgb(fill)
        else:
            shp.fill.background()
        if line:
            shp.line.color.rgb = rgb(line)
            shp.line.width = PT(line_w)
        else:
            shp.line.fill.background()
        if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
            shp.adjustments[0] = radius
        _strip_style(shp)
        tf = shp.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        return shp

    def line(self, slide, x1, y1, x2, y2, color, width=1.0):
        ln = slide.shapes.add_connector(1, IN(x1), IN(y1), IN(x2), IN(y2))
        ln.line.color.rgb = rgb(color)
        ln.line.width = PT(width)
        _strip_style(ln)
        return ln

    # -- texto -------------------------------------------------------------
    def text(self, slide, x, y, w, h, paragraphs, size=16, color=None, font=None,
             bold=False, italic=False, align="left", anchor="top", line_spacing=1.0,
             space_after=0.0, emph_color=None, fit=None, caps=False, letter_spacing=None,
             bullets=None, shape=None):
        """Caja de texto. `paragraphs`: str | [str] | [(str, level)].
        fit=(min_pt,) -> reduce el tamaño hasta que quepa (máx = size).
        bullets=dict(char, color, indent) -> viñetas por nivel."""
        if isinstance(paragraphs, str):
            paragraphs = [paragraphs]
        if shape is not None:  # texto dentro de una forma: medimos la forma
            tfm = shape.text_frame
            w = (shape.width - (tfm.margin_left or 0) - (tfm.margin_right or 0)) / 914400 / SCALE
            h = (shape.height - (tfm.margin_top or 0) - (tfm.margin_bottom or 0)) / 914400 / SCALE
            if shape.auto_shape_type in (MSO_SHAPE.CHEVRON, MSO_SHAPE.PENTAGON):
                w -= 0.25  # la punta de la flecha no es área útil
        paras = [(p, 0) if isinstance(p, str) else (p[0], p[1]) for p in paragraphs]
        font = font or self.bf
        color = color or self.t["text"]
        indent = bullets.get("indent", 0.28) if bullets else 0.0
        if caps:
            paras = [(str(p).upper(), lvl) for p, lvl in paras]
        mparas = [(plain(p), lvl if lvl >= 0 else -1) for p, lvl in paras]
        if fit is not None:
            size = fit_size(mparas, w, h, font, bold,
                            max_pt=size, min_pt=fit, line_spacing=line_spacing,
                            space_after_ratio=(space_after / size if size else 0),
                            indent_in=indent)
            space_after = space_after * (size / max(size, 1))
        # control de desbordamiento (se registra para el QA)
        needed = text_height(mparas, w, font, size, bold,
                             line_spacing, space_after_pt=space_after, indent_in=indent)
        if needed > h * 1.05 and any(plain(p).strip() for p, _ in paras):
            self.warnings.append(
                f"Diapositiva {self.page}: el texto '{plain(paras[0][0])[:40]}…' necesita "
                f"{needed:.2f}in y la caja mide {h:.2f}in (posible desbordamiento).")

        box = shape or slide.shapes.add_textbox(IN(x), IN(y), IN(w), IN(h))
        tf = box.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.NONE
        tf.margin_left = tf.margin_right = IN(0.0)
        tf.margin_top = tf.margin_bottom = IN(0.0)
        tf.vertical_anchor = ANCHOR[anchor]
        for i, (ptxt, lvl) in enumerate(paras):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = ALIGN[align]
            p.line_spacing = line_spacing
            if space_after and i < len(paras) - 1:
                p.space_after = PT(space_after)
            lvl_size = size if lvl <= 0 else max(size - 2, 9)
            heading = lvl == -2
            for txt, em in runs_of(ptxt):
                r = p.add_run()
                r.text = txt
                f = r.font
                f.name = self.ef if ((em or heading or bold) and font == self.bf) else font
                f.size = PT(lvl_size)
                f.bold = bold or em or heading
                f.italic = italic
                f.color.rgb = rgb(emph_color if (em and emph_color) else
                                  self.t["primary"] if heading else
                                  (self.t["text_muted"] if lvl > 0 and bullets else color))
                if letter_spacing:
                    r._r.get_or_add_rPr().set("spc", str(int(letter_spacing * 100 * SCALE)))
            if bullets and lvl >= 0:
                _set_bullet(p, bullets.get("char", self.t.get("bullet_char", "•"))
                            if lvl == 0 else "–",
                            bullets.get("color", self.t["accent"]), indent, lvl,
                            bullets.get("font"))
            elif bullets:
                pPr = p._p.get_or_add_pPr()
                pPr.set("marL", "0")
                pPr.set("indent", "0")
                etree.SubElement(pPr, qn("a:buNone"))
        return box

    # -- piezas comunes ----------------------------------------------------
    def title(self, slide, s, x=None, w=None, dark=False):
        """Título de la diapositiva. Devuelve la Y donde empieza el contenido.
        - Modo genérico: antetítulo opcional + título-acción dibujados por el motor.
        - Modo plantilla: título grande = `section` (o `title`) en el marcador de título;
          subtítulo = `title` (o `subtitle`) en el marcador de subtítulo. Luego `lead`."""
        if self.profile:
            return self._title_profile(slide, s)
        x = M if x is None else x
        w = w or (W - 2 * M)
        y = TOP
        kicker = s.get("kicker")
        if kicker:
            self.text(slide, x, y - 0.05, w, 0.3, kicker, size=11, bold=True, caps=True,
                      color=self.acc_txt, letter_spacing=1.5)
            y += 0.3
        title = s.get("title", "")
        size = fit_size([plain(title)], w, 1.05, self.hf, self.t["heading_bold"],
                        max_pt=s.get("title_size", 28), min_pt=20)
        lines = wrap_lines(plain(title), w * 72, self.hf, size, self.t["heading_bold"])
        h = lines * size * line_height_factor(self.hf) / 72 + 0.05
        self.text(slide, x, y, w, h, title, size=size, font=self.hf,
                  bold=self.t["heading_bold"],
                  color=self.t["text_on_dark"] if dark else self.t["text"],
                  emph_color=self.t["accent"])
        top = max(y + h + 0.32, 1.35)
        return self.lead(slide, s, top, x, w)

    def _title_profile(self, slide, s):
        r = self.roles["content"]
        g = self.profile.get("grid", {})
        if s.get("section"):
            self.current_section = s["section"]
        sec = s.get("section") or (self.current_section if s.get("title") else None)
        big = sec or s.get("title", "")
        sub = s.get("title") if sec else s.get("subtitle")
        tw = (W - 2 * M) * 0.95   # margen interno del marcador + seguridad de render
        tmax = s.get("title_size", self.sizes.get("title", 40))
        tmin = self.sizes.get("title_min", 28)
        tsz = fit_size([plain(big)], tw, 0.95, self.hf, True, max_pt=tmax, min_pt=tmin)
        tlines = wrap_lines(plain(big), tw * 72, self.hf, tsz, True)
        ph_t = self.fill_ph(slide, r["title"], big, size=tsz, font=self.hf, bold=True)
        shift = 0.0
        if tlines > 1:
            if sec:
                self.warnings.append(f"Diapositiva {self.page}: la sección '{plain(big)[:40]}…' "
                                     "no cabe en una línea: acórtala.")
            extra = (tlines - 1) * tsz * 0.85 / 72   # interlineado del título = 80 %
            self.materialize(ph_t).height = ph_t.height + IN(extra)
            shift = extra
        top = g.get("content_top_no_subtitle", 1.5) + shift
        if sub:
            ssz = fit_size([plain(sub)], tw, 0.5, self.bf, max_pt=self.sizes.get("subtitle", 27),
                           min_pt=self.sizes.get("subtitle_min", 20))
            lines = wrap_lines(plain(sub), tw * 72, self.bf, ssz)
            phs = self.fill_ph(slide, r["subtitle"], sub, size=ssz, emph_color=self.acc_txt,
                               font=self.bf)
            self.materialize(phs)
            if shift:
                phs.top = phs.top + IN(shift)
            top = g.get("content_top", 1.95) + shift
            if lines > 1:  # subtítulo largo: 2 líneas, el contenido baja
                extra = (lines - 1) * ssz * 1.2 / 72
                phs.height = phs.height + IN(extra)
                top += extra
                if lines > 2:
                    self.warnings.append(f"Diapositiva {self.page}: el subtítulo ocupa "
                                         f"{lines} líneas; resúmelo (≤ 60 caracteres).")
        else:
            self.drop_ph(slide, r["subtitle"])
        return self.lead(slide, s, top)

    def footer(self, slide, s, dark=False):
        muted = self.t["text_on_dark_muted"] if dark else self.t["text_muted"]
        src = s.get("source")
        if src and not re.match(r"^(fuente|source)", src, re.I):
            src = "Fuente: " + src
        if self.profile:  # logo y nº de página los pone la plantilla
            g = self.profile.get("grid", {})
            if src:
                self.text(slide, g.get("source_x", 2.2), g.get("source_y", 6.88),
                          g.get("source_w", 9.5), 0.22, src, size=9, color=muted, italic=True)
            if s.get("note"):
                ny = g.get("source_y", 6.88) - (0.3 if src else 0)
                self.text(slide, g.get("source_x", 2.2), ny,
                          g.get("source_w", 9.5), 0.25, s["note"], size=9, color=muted, fit=7)
            return
        foot = self.meta.get("footer")
        if foot:
            self.text(slide, M, FOOTER_Y, 8, 0.25, foot, size=9, color=muted)
        if self.meta.get("slide_numbers", True):
            self.text(slide, W - M - 1, FOOTER_Y, 1, 0.25, str(self.page), size=9,
                      color=muted, align="right")
        if src:
            self.text(slide, M, FOOTER_Y - 0.3, W - 2 * M, 0.25, src, size=9,
                      color=muted, italic=True)

    def _src_gap(self, s):
        return 0.3 if (s.get("source") and not self.profile) else 0.0

    def body_box(self, s, top):
        bottom = BOTTOM - self._src_gap(s)
        if self.profile and s.get("note") and s.get("source"):
            bottom -= 0.3
        if s.get("takeaway"):
            bottom -= 0.85
        return M, top, W - 2 * M, bottom - top

    def takeaway(self, slide, s):
        if not s.get("takeaway"):
            return
        bottom = BOTTOM - self._src_gap(s) - (0.3 if self.profile and s.get("note")
                                               and s.get("source") else 0)
        y, h = bottom - 0.66, 0.66
        self.rect(slide, M, y, W - 2 * M, h, self.t["surface"])
        self.rect(slide, M, y, 0.07, h, self.t["accent"])
        self.text(slide, M + 0.3, y + 0.08, W - 2 * M - 0.5, h - 0.16, s["takeaway"],
                  size=16, bold=True, anchor="middle", fit=12, emph_color=self.acc_txt,
                  font=self.ef if self.profile else None)

    def band(self, slide, x, y, w, text, h=0.4, size=15):
        """Banda de cabecera de bloque (estilo plantilla: color de marca + texto blanco)."""
        self.rect(slide, x, y, w, h, self.t.get("band", self.t["primary"]))
        self.text(slide, x + 0.12, y, w - 0.24, h, text, size=size, fit=11, bold=True,
                  font=self.ef, color="FFFFFF", anchor="middle")
        return y + h

    def notes(self, slide, s):
        if s.get("notes"):
            slide.notes_slide.notes_text_frame.text = s["notes"]

    def circle_label(self, slide, cx, cy, d, label, fill, color="FFFFFF", size=14):
        c = self.rect(slide, cx - d / 2, cy - d / 2, d, d, fill, shape=MSO_SHAPE.OVAL)
        self.text(slide, 0, 0, 0, 0, str(label), size=size, bold=True, color=color,
                  align="center", anchor="middle", shape=c, font=self.hf)
        return c


# ----------------------------------------------------------------------------
# XML helpers
# ----------------------------------------------------------------------------
def _strip_style(shape):
    """Elimina <p:style> (evita sombras/efectos del tema por defecto)."""
    st = shape._element.find(qn("p:style"))
    if st is not None:
        shape._element.remove(st)


def _set_bullet(p, char, color, indent_in, level=0, font=None):
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(int(IN(indent_in * (level + 1)))))
    pPr.set("indent", str(-int(IN(indent_in))))
    for tag in ("a:buClr", "a:buSzPct", "a:buFont", "a:buChar", "a:buNone"):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)
    buClr = etree.SubElement(pPr, qn("a:buClr"))
    clr = etree.SubElement(buClr, qn("a:srgbClr"))
    clr.set("val", color)
    sz = etree.SubElement(pPr, qn("a:buSzPct"))
    sz.set("val", "100000")
    bf = etree.SubElement(pPr, qn("a:buFont"))
    bf.set("typeface", font or "Arial")
    bc = etree.SubElement(pPr, qn("a:buChar"))
    bc.set("char", char)


def _cell_style(cell, fill=None, border_bottom=None, border_w=0.75, border_top=None):
    tcPr = cell._tc.get_or_add_tcPr()
    for child in list(tcPr):
        tcPr.remove(child)

    def ln(tag, color, w):
        el = etree.SubElement(tcPr, qn(tag))
        if color:
            el.set("w", str(int(PT(w))))
            sf = etree.SubElement(el, qn("a:solidFill"))
            c = etree.SubElement(sf, qn("a:srgbClr"))
            c.set("val", color)
        else:
            el.set("w", "0")
            etree.SubElement(el, qn("a:noFill"))

    ln("a:lnL", None, 0)
    ln("a:lnR", None, 0)
    ln("a:lnT", border_top, border_w)
    ln("a:lnB", border_bottom, border_w)
    if fill:
        sf = etree.SubElement(tcPr, qn("a:solidFill"))
        c = etree.SubElement(sf, qn("a:srgbClr"))
        c.set("val", fill)
    else:
        etree.SubElement(tcPr, qn("a:noFill"))


def _load_template(path):
    """python-pptx no abre .potx directamente: cambiamos el content-type en una copia."""
    if not path.lower().endswith((".potx", ".potm")):
        return path
    tmp = tempfile.NamedTemporaryFile(suffix=".pptx", delete=False).name
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(
                    b"presentationml.template.main+xml",
                    b"presentationml.presentation.main+xml")
            zout.writestr(item, data)
    return tmp


def _remove_all_slides(prs):
    sldIdLst = prs.slides._sldIdLst
    for sldId in list(sldIdLst):
        rId = sldId.get(qn("r:id"))
        prs.part.drop_rel(rId)
        sldIdLst.remove(sldId)


def _pick_blank_layout(prs):
    names = ("blank", "en blanco", "vacío", "vacio", "leer", "vide")
    for layout in prs.slide_layouts:
        if layout.name.strip().lower() in names:
            return layout
    return min(prs.slide_layouts, key=lambda l: len(l.placeholders))


# ----------------------------------------------------------------------------
# Layouts
# ----------------------------------------------------------------------------
def l_title(d: Deck, s):
    if d.profile:
        return _p_cover(d, s)
    sl = d.new_slide(dark=True)
    t = d.t
    img = s.get("image")
    text_w = (W * 0.55 - M) if img else (W - 2 * M - 1.5)
    if img:
        _picture_fill(d, sl, img, W * 0.58, 0, W * 0.42, H)
    d.rect(sl, M, 2.05, 0.9, 0.07, t["accent"])
    y = 2.35
    if s.get("kicker"):
        d.text(sl, M, y, text_w, 0.35, s["kicker"], size=13, bold=True, caps=True,
               color=t["accent"], letter_spacing=2)
        y += 0.45
    title = s.get("title", "")
    size = fit_size([plain(title)], text_w, 2.0, d.hf, True, max_pt=44, min_pt=28)
    lines = wrap_lines(plain(title), text_w * 72, d.hf, size, True)
    th = lines * size * line_height_factor(d.hf) / 72 + 0.1
    d.text(sl, M, y, text_w, th, title, size=size, font=d.hf, bold=True,
           color=t["text_on_dark"], emph_color=t["accent"])
    y += th + 0.2
    if s.get("subtitle"):
        d.text(sl, M, y, text_w, 1.0, s["subtitle"], size=20, fit=14,
               color=t["text_on_dark_muted"])
    meta_line = " · ".join(x for x in [s.get("author") or d.meta.get("author"),
                                       s.get("date") or d.meta.get("date")] if x)
    if meta_line:
        d.text(sl, M, 6.45, text_w, 0.35, meta_line, size=13, color=t["text_on_dark_muted"])
    d.notes(sl, s)


def l_section(d: Deck, s):
    if d.profile:
        return _p_section(d, s)
    sl = d.new_slide(dark=True)
    t = d.t
    num = s.get("number")
    y = 2.3
    if num:
        d.text(sl, M, y - 0.4, 4, 1.4, str(num), size=80, bold=True, font=d.hf,
               color=t["accent"])
        y += 1.2
    d.text(sl, M, y, W - 2 * M - 2, 1.3, s.get("title", ""), size=40, fit=28, bold=True,
           font=d.hf, color=t["text_on_dark"], emph_color=t["accent"])
    if s.get("subtitle"):
        d.text(sl, M, y + 1.35, W - 2 * M - 3, 1.0, s["subtitle"], size=18, fit=13,
               color=t["text_on_dark_muted"])
    d.notes(sl, s)


def l_agenda(d: Deck, s):
    if d.profile:
        return _p_agenda(d, s)
    """Panel oscuro a la izquierda con el título; lista numerada a la derecha."""
    sl = d.new_slide()
    t = d.t
    pw = W * 0.34
    d.rect(sl, 0, 0, pw, H, t["bg_dark"])
    d.rect(sl, M, 2.55, 0.7, 0.07, t["accent"])
    if s.get("kicker"):
        d.text(sl, M, 2.0, pw - 2 * M, 0.35, s["kicker"], size=12, bold=True, caps=True,
               color=t["accent"], letter_spacing=1.5)
    d.text(sl, M, 2.8, pw - M - 0.4, 2.2, s.get("title", "Agenda"), size=40, fit=26,
           bold=True, font=d.hf, color=t["text_on_dark"], emph_color=t["accent"])
    items = s.get("items", [])
    if not 1 <= len(items) <= 10:
        raise SpecError("agenda: entre 1 y 10 puntos")
    current = s.get("current")
    x = pw + 0.8
    w = W - x - M
    cols = 2 if len(items) > 6 else 1
    per_col = math.ceil(len(items) / cols)
    col_w = (w - 0.5 * (cols - 1)) / cols
    avail = H - 1.6
    row_h = min(avail / per_col, 1.15)
    y0 = (H - row_h * per_col) / 2
    for i, it in enumerate(items):
        c, r = divmod(i, per_col)
        cx = x + c * (col_w + 0.5)
        cy = y0 + r * row_h
        active = current is None or current == i
        num_color = t["accent"] if active else t["border"]
        txt_color = t["text"] if active else t["text_muted"]
        label = it if isinstance(it, str) else it.get("title", "")
        desc = None if isinstance(it, str) else it.get("text")
        d.text(sl, cx, cy, 0.9, row_h, f"{i + 1:02d}", size=26, bold=True,
               font=d.hf, color=num_color, anchor="middle")
        box = d.text(sl, cx + 1.0, cy + 0.08, col_w - 1.0, row_h - 0.16, label, size=22,
                     fit=14, bold=current == i, color=txt_color, anchor="middle")
        if desc:
            p = box.text_frame.add_paragraph()
            p.space_before = PT(3)
            run = p.add_run()
            run.text = desc
            run.font.size = PT(14)
            run.font.name = d.bf
            run.font.color.rgb = rgb(t["text_muted"])
        if r < per_col - 1 and i < len(items) - 1:
            d.line(sl, cx, cy + row_h, cx + col_w, cy + row_h, t["border"], 0.75)
    if d.meta.get("slide_numbers", True):
        d.text(sl, W - M - 1, FOOTER_Y, 1, 0.25, str(d.page), size=9,
               color=t["text_muted"], align="right")
    d.notes(sl, s)


def l_bullets(d: Deck, s):
    sl = d.new_slide()
    top = d.title(sl, s)
    x, y, w, h = d.body_box(s, top)
    callout = s.get("callout")
    text_w = w * 0.6 if callout else min(w, 10.5)
    items = norm_items(s.get("bullets", []))
    d.text(sl, x, y, text_w, h, items, size=s.get("size", 22), fit=13, space_after=10,
           line_spacing=1.05, bullets={"color": d.t["accent"], "indent": 0.32},
           emph_color=None)
    if callout:
        cx = x + text_w + 0.5
        cw = w - text_w - 0.5
        ch = min(h, 3.6)
        d.rect(sl, cx, y, cw, ch, d.t["surface"])
        d.rect(sl, cx, y, cw, 0.07, d.t["accent"])
        if callout.get("value"):
            d.text(sl, cx + 0.35, y + 0.4, cw - 0.7, 1.3, callout["value"], size=60,
                   fit=32, bold=True, font=d.hf, color=d.t["primary"]
                   if not d.t["dark_content"] else d.t["accent"])
            d.text(sl, cx + 0.35, y + 1.75, cw - 0.7, ch - 2.0, callout.get("label", ""),
                   size=16, fit=11, color=d.t["text_muted"])
        else:
            d.text(sl, cx + 0.35, y + 0.4, cw - 0.7, ch - 0.7, callout.get("text", ""),
                   size=18, fit=12, color=d.t["text"], emph_color=d.t["accent"])
    d.takeaway(sl, s)
    d.footer(sl, s)
    d.notes(sl, s)


def l_two_column(d: Deck, s):
    sl = d.new_slide()
    top = d.title(sl, s)
    x, y, w, h = d.body_box(s, top)
    cols = [s.get("left", {}), s.get("right", {})]
    contrast = s.get("style") == "contrast"
    cw = (w - 0.5) / 2
    sizes = []
    for c in cols:
        items = norm_items(c.get("bullets", [])) or [c.get("text", "")]
        sizes.append(fit_size([(plain(t), l) for t, l in items], cw - (0.6 if contrast else 0),
                              h - 0.9 - (0.4 if contrast else 0), d.bf, max_pt=20, min_pt=12,
                              space_after_ratio=0.4, indent_in=0.3))
    size = min(sizes)
    if contrast:
        need = max(text_height([(plain(t), l) for t, l in (norm_items(c.get("bullets", []))
                                or [(c.get("text", ""), 0)])], cw - 0.6, d.bf, size,
                               space_after_pt=size * 0.4, indent_in=0.3) for c in cols)
        h = min(h, max(h * 0.55, need + 0.8 + 0.7))
    for i, c in enumerate(cols):
        cx = x + i * (cw + 0.5)
        pad = 0.0
        if contrast:
            fill = d.t["surface"] if i == 0 else mix(d.t["primary"], d.t["bg"], 0.9)
            d.rect(sl, cx, y, cw, h, fill)
            pad = 0.3
        head_color = d.t["text_muted"] if (contrast and i == 0) else d.t["accent"]
        d.rect(sl, cx + pad, y + pad, 0.6, 0.06, head_color)
        d.text(sl, cx + pad, y + pad + 0.15, cw - 2 * pad, 0.5, c.get("heading", ""), size=20,
               fit=14, bold=True, font=d.hf, color=d.t["text"])
        items = norm_items(c.get("bullets", []))
        body = items if items else [c.get("text", "")]
        d.text(sl, cx + pad, y + pad + 0.8, cw - 2 * pad, h - 0.9 - 2 * pad, body, size=size,
               space_after=size * 0.4,
               bullets={"color": head_color, "indent": 0.3} if items else None)
    d.takeaway(sl, s)
    d.footer(sl, s)
    d.notes(sl, s)


def l_cards(d: Deck, s):
    sl = d.new_slide()
    top = d.title(sl, s)
    x, y, w, h = d.body_box(s, top)
    cards = s.get("cards", [])
    n = len(cards)
    if not 1 <= n <= 8:
        raise SpecError("cards: entre 1 y 8 tarjetas")
    cols = n if n <= 4 else (3 if n in (5, 6) else 4)
    rows = math.ceil(n / cols)
    gap = 0.28
    cw = (w - gap * (cols - 1)) / cols
    ch = (h - gap * (rows - 1)) / rows
    pad = 0.28
    head_y = 1.0 if s.get("numbered", True) or any(c.get("icon") for c in cards) else pad
    title_size = min(fit_size([plain(c.get("title", ""))], cw - 2 * pad, 0.75, d.hf, True,
                              max_pt=18, min_pt=13) for c in cards)
    title_h = max(wrap_lines(plain(c.get("title", "")), (cw - 2 * pad) * 72, d.hf, title_size, True)
                  for c in cards) * title_size * 1.25 / 72
    body_h = ch - head_y - title_h - 0.15 - pad
    body_size = min(fit_size([plain(c.get("text", ""))], cw - 2 * pad, body_h, d.bf,
                             max_pt=17, min_pt=10.5, line_spacing=1.1) for c in cards)
    for i, c in enumerate(cards):
        r, k = divmod(i, cols)
        cx, cy = x + k * (cw + gap), y + r * (ch + gap)
        hl = c.get("highlight")
        d.rect(sl, cx, cy, cw, ch, d.t["primary"] if hl else d.t["surface"])
        on_dark = hl
        if head_y > pad:
            label = c.get("icon") or f"{i + 1:02d}"
            d.circle_label(sl, cx + pad + 0.28, cy + pad + 0.28, 0.56, label,
                           d.t["accent"], size=14 if len(str(label)) <= 2 else 10)
        d.text(sl, cx + pad, cy + head_y, cw - 2 * pad, title_h, c.get("title", ""),
               size=title_size, bold=True, font=d.hf,
               color=d.t["text_on_dark"] if on_dark else d.t["text"])
        d.text(sl, cx + pad, cy + head_y + title_h + 0.15, cw - 2 * pad, body_h,
               c.get("text", ""), size=body_size, line_spacing=1.1,
               color=mix("FFFFFF", d.t["primary"], 0.18) if on_dark else d.t["text_muted"],
               emph_color=d.t["text_on_dark"] if on_dark else d.t["text"])
    d.takeaway(sl, s)
    d.footer(sl, s)
    d.notes(sl, s)


def _delta_style(d, k):
    trend = k.get("trend")
    delta = str(k.get("delta", ""))
    if not trend and delta:
        trend = "down" if delta.strip().startswith(("-", "−")) else "up"
    good = k.get("good", "up")
    if trend == "flat" or not trend:
        return "■", d.t["text_muted"]
    arrow = "▲" if trend == "up" else "▼"
    color = d.t["positive"] if trend == good else d.t["negative"]
    return arrow, color


def l_kpis(d: Deck, s):
    sl = d.new_slide()
    top = d.title(sl, s)
    x, y, w, h = d.body_box(s, top)
    kpis = s.get("kpis", [])
    n = len(kpis)
    if not 1 <= n <= 6:
        raise SpecError("kpis: entre 1 y 6 indicadores")
    extra = s.get("bullets")
    tile_h = min(2.6, h * 0.58) if extra else min(h, 3.4)
    if not extra:
        y = y + (h - tile_h) / 2 * 0.5
    gap = 0.3
    tw = (w - gap * (n - 1)) / n
    vh = min(1.1, tile_h * 0.4)
    vsize = min(fit_size([str(k.get("value", ""))], tw - 0.6, vh, d.hf, True,
                         max_pt=60, min_pt=24) for k in kpis)
    pad_t = min(0.35, tile_h * 0.12)
    ly = pad_t + vh + 0.08
    lh = tile_h - ly - (0.5 if any(k.get("delta") for k in kpis) else 0.2)
    for i, k in enumerate(kpis):
        tx = x + i * (tw + gap)
        d.rect(sl, tx, y, tw, tile_h, d.t["surface"])
        d.rect(sl, tx, y, tw, 0.07, d.t["accent"] if i == s.get("highlight", -1) else d.t["primary"])
        d.text(sl, tx + 0.3, y + pad_t, tw - 0.6, vh, k.get("value", ""), size=vsize,
               bold=True, font=d.hf,
               color=d.t["accent"] if d.t["dark_content"] else d.t["primary"])
        d.text(sl, tx + 0.3, y + ly, tw - 0.6, lh, k.get("label", ""),
               size=15, fit=11, color=d.t["text"])
        if k.get("delta"):
            arrow, color = _delta_style(d, k)
            d.text(sl, tx + 0.3, y + tile_h - 0.5, tw - 0.6, 0.3,
                   f"{arrow} {k['delta']}" + (f"  {k['delta_label']}" if k.get("delta_label") else ""),
                   size=12, bold=True, color=color)
    if extra:
        by = y + tile_h + 0.4
        d.text(sl, x, by, w, (y + h) - by, norm_items(extra), size=18, fit=12,
               space_after=6, bullets={"color": d.t["accent"]})
    d.takeaway(sl, s)
    d.footer(sl, s)
    d.notes(sl, s)


def l_big_number(d: Deck, s):
    sl = d.new_slide()
    top = d.title(sl, s) if s.get("title") else 1.0
    x, y, w, h = d.body_box(s, top)
    lw = w * 0.48
    vs = fit_size([str(s.get("value", ""))], lw, 2.4, d.hf, True, max_pt=150, min_pt=54)
    vh = vs * line_height_factor(d.hf) / 72
    cy = y + max(0, (h - vh - 1.0) / 2)
    d.text(sl, x, cy, lw, vh, s.get("value", ""), size=vs, bold=True, font=d.hf,
           color=d.t["accent"])
    d.text(sl, x, cy + vh + 0.05, lw, 1.0, s.get("label", ""), size=20, fit=13,
           bold=True, color=d.t["text"])
    rx = x + lw + 0.6
    d.rect(sl, rx, y + 0.2, 0.05, h - 0.4, d.t["border"])
    ctx = s.get("context")
    if ctx:
        items = norm_items(ctx) if isinstance(ctx, list) else [ctx]
        d.text(sl, rx + 0.5, y + 0.2, w - lw - 1.1, h - 0.4, items, size=20, fit=12,
               anchor="middle", space_after=10, emph_color=d.t["accent"],
               bullets={"color": d.t["accent"]} if isinstance(ctx, list) else None)
    d.takeaway(sl, s)
    d.footer(sl, s)
    d.notes(sl, s)


def _bare_content_slide(d: Deck, s):
    """Diapositiva de contenido de la plantilla sin título (o solo con la sección)."""
    sl = d.new_slide()
    r = d.roles["content"]
    if s.get("section"):
        d.fill_ph(sl, r["title"], s["section"], size=d.sizes.get("title", 40), font=d.hf, bold=True)
    else:
        d.drop_ph(sl, r["title"])
    d.drop_ph(sl, r["subtitle"])
    return sl


def l_statement(d: Deck, s):
    dark = s.get("dark", False) and not d.profile
    if d.profile:
        sl = _bare_content_slide(d, s)
    else:
        sl = d.new_slide(dark=dark, bg=None if dark else d.t["surface"])
    col = d.t["text_on_dark"] if dark else (d.t["primary"] if d.profile else d.t["text"])
    if s.get("kicker"):
        d.text(sl, 1.4, 1.9, W - 2.8, 0.4, s["kicker"], size=13, bold=True, caps=True,
               color=d.acc_txt, letter_spacing=2, align="center")
    d.text(sl, 1.4, 2.4, W - 2.8, 2.6, s.get("text", s.get("title", "")), size=40, fit=24,
           font=d.hf, bold=True, color=col,
           emph_color=d.t["accent"] if not d.profile else d.acc_txt, align="center",
           anchor="middle", line_spacing=1.05)
    if s.get("subtitle"):
        d.text(sl, 2, 5.2, W - 4, 0.9, s["subtitle"], size=18, fit=13, align="center",
               color=d.t["text_on_dark_muted"] if dark else d.t["text_muted"])
    d.footer(sl, s, dark=dark)
    d.notes(sl, s)


# --- variantes de plantilla (perfil) ----------------------------------------
def _p_cover(d: Deck, s):
    r = d.roles["cover"]
    sl = d.new_slide(role="cover")
    title = s.get("title", "")
    cw = d.profile.get("grid", {}).get("cover_title_w", 10.2)
    sw = d.profile.get("grid", {}).get("cover_subtitle_w", 8.0)
    tsz = fit_size([plain(title)], cw * 0.95, 2.4, d.hf, True, max_pt=58, min_pt=36)
    pt_ = d.fill_ph(sl, r["title"], title, size=tsz, font=d.hf, bold=True)
    d.materialize(pt_).width = IN(cw)
    sub = s.get("subtitle", "")
    meta_line = " / ".join(x for x in [s.get("version") or d.meta.get("version"),
                                       s.get("date") or d.meta.get("date")] if x)
    paras, bolds, sizes = [], [], []
    if sub:
        paras.append(sub); bolds.append(True); sizes.append(27.5)
    if s.get("client") or d.meta.get("client"):
        paras.append(s.get("client") or d.meta.get("client")); bolds.append(False); sizes.append(22)
    if meta_line:
        paras.append(meta_line); bolds.append(False); sizes.append(24)
    if paras:
        bw = sw * 0.95
        def need(k):
            return sum(text_height([plain(p_)], bw, d.ef if b else d.bf, z * k, b)
                       for p_, b, z in zip(paras, bolds, sizes)) + 0.1 * len(paras)
        k = 1.0
        while k > 0.7 and need(k) > 2.3:
            k -= 0.05
        ph = d.fill_ph(sl, r["subtitle"], paras, para_bold=bolds,
                       para_sizes=[z * k for z in sizes], space_before=6, font=d.bf)
        d.materialize(ph).width = IN(sw)
        if need(k) > ph.height / 914400 / SCALE:
            d.materialize(ph).height = IN(min(2.4, need(k) + 0.1))
    else:
        d.drop_ph(sl, r["subtitle"])
    if s.get("image"):
        d.warnings.append("La portada de la plantilla no admite imagen: se ignora 'image'.")
    d.notes(sl, s)


def _p_section(d: Deck, s):
    r = d.roles["section"]
    sl = d.new_slide(role="section")
    num = s.get("number")
    if num is not None and str(num).strip().rstrip(".").isdigit():
        num = str(int(str(num).strip().rstrip(".")))
    title = f"{str(num).rstrip('.')}. {s.get('title', '')}" if num else s.get("title", "")
    d.current_section = title
    cw = d.profile.get("grid", {}).get("cover_title_w", 10.2)
    sw = d.profile.get("grid", {}).get("cover_subtitle_w", 8.0)
    tsz = fit_size([plain(title)], cw * 0.95, 2.4, d.hf, True, max_pt=52, min_pt=34)
    pt_ = d.fill_ph(sl, r["title"], title, size=tsz, font=d.hf, bold=True)
    d.materialize(pt_).width = IN(cw)
    if s.get("subtitle"):
        ssz = fit_size([plain(s["subtitle"])], sw * 0.95, 1.0, d.bf, max_pt=24, min_pt=16)
        ps_ = d.fill_ph(sl, r["subtitle"], s["subtitle"], size=ssz, font=d.bf)
        d.materialize(ps_).width = IN(sw)
    else:
        d.drop_ph(sl, r["subtitle"])
    d.notes(sl, s)


def _p_agenda(d: Deck, s):
    r = d.roles["agenda"]
    sl = d.new_slide(role="agenda")
    d.fill_ph(sl, r["title"], s.get("title", "Contenido"), size=30, font=d.hf, bold=True)
    items = s.get("items", [])
    labels = [it if isinstance(it, str) else it.get("title", "") for it in items]
    current = s.get("current")
    ph = d.ph(sl, r["body"])
    bw = ph.width / 914400 / SCALE
    bh = ph.height / 914400 / SCALE
    size = fit_size([plain(x) for x in labels], bw - 0.5, bh, d.bf, max_pt=17, min_pt=12,
                    line_spacing=1.0, space_after_ratio=0.45)
    cols = [None if current is None else (d.t["primary"] if i == current else d.t["text_muted"])
            for i in range(len(labels))]
    bolds = [current == i for i in range(len(labels))]
    d.fill_ph(sl, r["body"], labels, size=size, autonum=True, para_colors=cols,
              para_bold=bolds, space_before=size * 0.45, font=d.bf)
    d.notes(sl, s)


def _p_closing(d: Deck, s):
    r = d.roles["closing"]
    sl = d.new_slide(role="closing")
    paras, sizes, bolds = [s.get("title", "Gracias")], [28], [True]
    if s.get("subtitle"):
        paras.append(s["subtitle"]); sizes.append(19); bolds.append(False)
    for st in s.get("next_steps", []) or []:
        paras.append("▪  " + st); sizes.append(18); bolds.append(False)
    contact = s.get("contact")
    if contact:
        for c in (contact if isinstance(contact, list) else [contact]):
            paras.append(c); sizes.append(15); bolds.append(False)
    ph = d.ph(sl, r["body"])
    bw, bh = ph.width / 914400 / SCALE, ph.height / 914400 / SCALE
    k = 1.0
    while k > 0.6 and text_height([plain(p) for p in paras], bw, d.bf, 15 * k,
                                  space_after_pt=6) > bh * 0.9:
        k -= 0.05
    d.fill_ph(sl, r["body"], paras, para_sizes=[x * k for x in sizes], para_bold=bolds,
              space_before=8, color="FFFFFF", font=d.bf)
    d.notes(sl, s)


CHART_TYPES = {
    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "bar": XL_CHART_TYPE.BAR_CLUSTERED,
    "stacked-column": XL_CHART_TYPE.COLUMN_STACKED,
    "stacked-bar": XL_CHART_TYPE.BAR_STACKED,
    "stacked-column-100": XL_CHART_TYPE.COLUMN_STACKED_100,
    "line": XL_CHART_TYPE.LINE_MARKERS,
    "area": XL_CHART_TYPE.AREA,
    "pie": XL_CHART_TYPE.PIE,
    "doughnut": XL_CHART_TYPE.DOUGHNUT,
    "scatter": XL_CHART_TYPE.XY_SCATTER,
}


def draw_chart(d: Deck, sl, spec, x, y, w, h):
    ctype = spec.get("type", "column")
    if ctype not in CHART_TYPES:
        raise SpecError(f"chart.type '{ctype}' no válido. Usa: {', '.join(CHART_TYPES)}")
    t = d.t
    cats = spec.get("categories", [])
    series = spec.get("series", [])
    if not series:
        raise SpecError("chart.series vacío")
    cd = CategoryChartData()
    cd.categories = cats
    nf = spec.get("number_format", "General")
    for se in series:
        cd.add_series(se.get("name", ""), se.get("values", []), number_format=nf)
    gf = sl.shapes.add_chart(CHART_TYPES[ctype], IN(x), IN(y), IN(w), IN(h), cd)
    ch = gf.chart
    ch.has_title = False
    txt = t["text_muted"]
    ch.font.name = d.bf
    ch.font.size = PT(spec.get("font_size", 12))
    ch.font.color.rgb = rgb(txt)
    pie_like = ctype in ("pie", "doughnut")
    multi = len(series) > 1
    ch.has_legend = multi or pie_like
    if ch.has_legend:
        ch.legend.position = XL_LEGEND_POSITION.BOTTOM if not pie_like else XL_LEGEND_POSITION.RIGHT
        ch.legend.include_in_layout = False
        ch.legend.font.size = PT(12)
        ch.legend.font.color.rgb = rgb(t["text"])
    palette = t["chart"]
    highlight = spec.get("highlight")  # índice de categoría a resaltar
    labels = spec.get("data_labels", not multi or ctype.startswith("stacked") or pie_like)

    plot = ch.plots[0]
    if not pie_like:
        if hasattr(plot, "gap_width"):
            plot.gap_width = spec.get("gap_width", 60)
        if ctype.startswith("stacked") and hasattr(plot, "overlap"):
            plot.overlap = 100
        elif hasattr(plot, "overlap") and multi:
            plot.overlap = -10
        va = ch.value_axis
        ca = ch.category_axis
        va.has_major_gridlines = not labels
        if va.has_major_gridlines:
            va.major_gridlines.format.line.color.rgb = rgb(t["border"])
            va.major_gridlines.format.line.width = PT(0.5)
        va.format.line.fill.background()
        va.tick_labels.font.size = PT(11)
        va.tick_labels.number_format = nf
        va.tick_labels.number_format_is_linked = False
        if labels and ctype in ("column", "bar", "stacked-column", "stacked-bar",
                                "stacked-column-100"):
            va.visible = False
        if spec.get("y_min") is not None:
            va.minimum_scale = spec["y_min"]
        if spec.get("y_max") is not None:
            va.maximum_scale = spec["y_max"]
        ca.format.line.color.rgb = rgb(t["border"])
        ca.tick_labels.font.size = PT(12)
        ca.tick_labels.font.color.rgb = rgb(t["text"])
        from pptx.enum.chart import XL_TICK_MARK
        ca.major_tick_mark = XL_TICK_MARK.NONE
        va.major_tick_mark = XL_TICK_MARK.NONE
        if ctype == "bar" or ctype == "stacked-bar":
            ca.reverse_order = True  # primera categoría arriba
            from pptx.enum.chart import XL_TICK_LABEL_POSITION  # noqa
            va.crosses  # mantener por defecto

    for si, se in enumerate(plot.series):
        color = series[si].get("color") or palette[si % len(palette)]
        if pie_like:
            for pi in range(len(cats)):
                pt = se.points[pi]
                pt.format.fill.solid()
                pt.format.fill.fore_color.rgb = rgb(palette[pi % len(palette)])
                pt.format.line.color.rgb = rgb(t["bg"])
                pt.format.line.width = PT(1.5)
            continue
        if ctype == "line":
            se.format.line.color.rgb = rgb(color)
            se.format.line.width = PT(2.75)
            se.smooth = False
            from pptx.enum.chart import XL_MARKER_STYLE
            se.marker.style = XL_MARKER_STYLE.CIRCLE
            se.marker.size = max(2, min(72, int(7 * SCALE)))
            se.marker.format.fill.solid()
            se.marker.format.fill.fore_color.rgb = rgb(color)
            se.marker.format.line.color.rgb = rgb(t["bg"])
        else:
            se.format.fill.solid()
            se.format.fill.fore_color.rgb = rgb(color)
            se.format.line.fill.background()
            if highlight is not None and not multi:
                muted = mix(color, t["bg"], 0.55)
                for pi in range(len(cats)):
                    pt = se.points[pi]
                    pt.format.fill.solid()
                    pt.format.fill.fore_color.rgb = rgb(
                        t["accent"] if pi == highlight else muted)

    if labels:
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.font.size = PT(spec.get("label_size", 12))
        dl.font.bold = True
        dl.number_format = spec.get("label_format", nf if not pie_like else "0%")
        dl.number_format_is_linked = False
        if pie_like:
            dl.show_percentage = spec.get("show_percent", True)
            dl.show_value = not spec.get("show_percent", True)
            dl.font.color.rgb = rgb("FFFFFF")
        elif ctype.startswith("stacked"):
            dl.font.color.rgb = rgb("FFFFFF")
            dl.position = XL_LABEL_POSITION.CENTER
        elif ctype == "line":
            dl.position = XL_LABEL_POSITION.ABOVE
            dl.font.color.rgb = rgb(t["text"])
        elif ctype != "area":
            dl.position = XL_LABEL_POSITION.OUTSIDE_END
            dl.font.color.rgb = rgb(t["text"])
    if ctype == "doughnut":
        hole = plot._element.find(qn("c:holeSize"))
        if hole is None:
            hole = etree.SubElement(plot._element, qn("c:holeSize"))
        hole.set("val", "62")
    return gf


def l_chart(d: Deck, s):
    sl = d.new_slide()
    top = d.title(sl, s)
    x, y, w, h = d.body_box(s, top)
    insights = s.get("insights")
    cw = w * 0.63 if insights else w
    spec = s.get("chart", {})
    if spec.get("subtitle"):
        d.text(sl, x, y - 0.05, cw, 0.35, spec["subtitle"], size=13, bold=True,
               color=d.t["text_muted"])
        y += 0.35
        h -= 0.35
    draw_chart(d, sl, spec, x, y, cw, h)
    if insights:
        ix = x + cw + 0.45
        iw = w - cw - 0.45
        d.rect(sl, ix, y, iw, h, d.t["surface"])
        d.text(sl, ix + 0.3, y + 0.3, iw - 0.6, 0.35, s.get("insights_title", "Claves"),
               size=13, bold=True, caps=True, color=d.acc_txt, letter_spacing=1.2)
        d.text(sl, ix + 0.3, y + 0.8, iw - 0.6, h - 1.1, norm_items(insights), size=16,
               fit=11, space_after=9, bullets={"color": d.t["accent"], "indent": 0.25},
               emph_color=d.t["text"])
    d.takeaway(sl, s)
    d.footer(sl, s)
    d.notes(sl, s)


def _is_spacer(row):
    return all(not plain(c).strip() for c in row)


def _col_widths(cols, rows, w, given=None):
    nc = len(cols)
    if given:
        tot = sum(given)
        return [w * c / tot for c in given]
    lens = []
    for ci in range(nc):
        cells = [plain(cols[ci])] + [plain(r[ci]) if ci < len(r) else "" for r in rows]
        lens.append(max(4, min(40, max(len(c) for c in cells))))
    return [w * L / sum(lens) for L in lens]


def _row_heights(d, cols, rows, widths, size, pad_x=0.1, pad_y=0.06):
    out = []
    for ri, r in enumerate([cols] + rows):
        if ri > 0 and _is_spacer(r):
            out.append(0.12)
            continue
        mx = max(text_height([plain(r[ci]) if ci < len(r) else ""], widths[ci] - 2 * pad_x,
                             d.bf, size, ri == 0) for ci in range(len(cols)))
        out.append(mx + 2 * pad_y + 0.03)
    return out


def draw_table(d: Deck, sl, cols, rows, x, y, w, size, widths=None, col_align=None,
               highlight_row=None, zebra=True, header_style="dark", total_row=False,
               min_row_h=None):
    """Dibuja una tabla estilizada y devuelve su altura (pulgadas de diseño)."""
    nc, nr = len(cols), len(rows)
    widths = widths or _col_widths(cols, rows, w)
    pad_x, pad_y = 0.1, 0.06
    heights = _row_heights(d, cols, rows, widths, size, pad_x, pad_y)
    if min_row_h:
        heights = [max(hh, min_row_h) if not (i and _is_spacer(([cols] + rows)[i])) else hh
                   for i, hh in enumerate(heights)]
    gf = sl.shapes.add_table(nr + 1, nc, IN(x), IN(y), IN(w), IN(sum(heights)))
    tbl = gf.table
    tblPr = tbl._tbl.tblPr
    tblPr.set("firstRow", "0")
    tblPr.set("bandRow", "0")
    style = tblPr.find(qn("a:tableStyleId"))
    if style is None:
        style = etree.SubElement(tblPr, qn("a:tableStyleId"))
    style.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"  # "Sin estilo, sin cuadrícula"
    for ci, cw in enumerate(widths):
        tbl.columns[ci].width = IN(cw)
    for ri, hh in enumerate(heights):
        tbl.rows[ri].height = IN(hh)
    aligns = col_align or ["left"] * nc
    dark_head = header_style == "dark"
    head_fill = d.t.get("table_header", d.t["primary"]) if dark_head else \
        d.t.get("table_subheader", d.t["surface"])
    zi = 0
    for ri in range(nr + 1):
        data = cols if ri == 0 else rows[ri - 1]
        header = ri == 0
        spacer = (not header) and _is_spacer(data)
        is_total = total_row and ri == nr
        is_hl = (not header) and highlight_row is not None and (ri - 1) == highlight_row
        if not header and not spacer:
            zi += 1
        for ci in range(nc):
            cell = tbl.cell(ri, ci)
            val = data[ci] if ci < len(data) else ""
            fill = head_fill if header else (
                mix(d.t["accent"], d.t["bg"], 0.82) if is_hl else
                (d.t["surface"] if zebra and not spacer and zi % 2 == 0 and not is_total
                 else None))
            _cell_style(cell, fill=fill,
                        border_bottom=None if (header and dark_head) or spacer else
                        (d.t["primary"] if header or is_total else d.t["border"]),
                        border_top=d.t["primary"] if is_total else None,
                        border_w=1.25 if is_total or header else 0.75)
            cell.margin_left = cell.margin_right = IN(pad_x)
            cell.margin_top = cell.margin_bottom = IN(0 if spacer else pad_y)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = ALIGN.get(aligns[ci] if ci < len(aligns) else "left")
            if spacer:
                r = p.add_run()
                r.text = ""
                r.font.size = PT(4)
                continue
            for txt, em in runs_of(val):
                r = p.add_run()
                r.text = txt
                bold = header or em or is_total or (is_hl and ci == 0)
                r.font.name = d.ef if bold else d.bf
                r.font.size = PT(size)
                r.font.bold = bold
                r.font.color.rgb = rgb(
                    ("FFFFFF" if not d.t["dark_content"] else d.t["text_on_dark"])
                    if header and dark_head else
                    (d.t["primary"] if header else d.t["text"]))
    return sum(heights)


def l_table(d: Deck, s):
    """Tabla; si no cabe ni a tamaño mínimo, se pagina en varias diapositivas
    (mismo título + '(cont.)'), repitiendo la cabecera."""
    cols = s.get("columns", [])
    rows = s.get("rows", [])
    if not cols:
        raise SpecError("table: falta 'columns'")
    sl = d.new_slide()
    top = d.title(sl, s)
    x, y, w, h = d.body_box(s, top)
    if s.get("band"):
        y = d.band(sl, x, y, w, s["band"]) + 0.12
        h -= 0.52
    nr = len(rows)
    widths = _col_widths(cols, rows, w, s.get("col_widths"))
    aligns = s.get("col_align") or (["left"] + ["right"] * (len(cols) - 1)
                                    if s.get("numeric", False) else ["left"] * len(cols))
    size = s.get("size", 16 if nr <= 6 else 13 if nr <= 10 else 12)
    min_size = s.get("min_size", 10.5)
    while size > min_size and sum(_row_heights(d, cols, rows, widths, size)) > h:
        size -= 0.5
    opts = dict(widths=widths, col_align=aligns, zebra=s.get("zebra", True),
                header_style=s.get("header_style", "dark"), total_row=s.get("total_row", False))
    if sum(_row_heights(d, cols, rows, widths, size)) <= h or not s.get("paginate", True):
        if sum(_row_heights(d, cols, rows, widths, size)) > h:
            d.warnings.append(f"Diapositiva {d.page}: la tabla no cabe; usa 'paginate'.")
        small = nr <= 8 and sum(_row_heights(d, cols, rows, widths, size)) < h * 0.6
        draw_table(d, sl, cols, rows, x, y, w, size, highlight_row=s.get("highlight_row"),
                   min_row_h=0.5 if small else None, **opts)
        d.takeaway(sl, s)
        d.footer(sl, s)
        d.notes(sl, s)
        return
    # paginación: trozos que caben con la cabecera repetida
    heights = _row_heights(d, cols, rows, widths, size)
    head_h = heights[0]
    if d.profile:
        h_cont = BOTTOM - d.profile.get("grid", {}).get("content_top", 1.95)
    else:
        h_cont = BOTTOM - 1.4 - d._src_gap(s)
    if s.get("band"):
        h_cont -= 0.52
    chunks, cur, acc = [], [], 0.0
    for r, hh in zip(rows, heights[1:]):
        avail = h if not chunks else h_cont
        if cur and head_h + acc + hh > avail:
            chunks.append(cur)
            cur, acc = [], 0.0
        if not cur and _is_spacer(r):
            continue  # no empezar página con un separador
        cur.append(r)
        acc += hh
    if cur:
        chunks.append(cur)
    for ci, chunk in enumerate(chunks):
        if ci > 0:
            cont = dict(s)
            cont.pop("lead", None)
            cont["title"] = f"{s.get('title', '')} (cont.)"
            sl = d.new_slide()
            top = d.title(sl, cont)
            x, y, w, h = d.body_box(cont, top)
            if s.get("band"):
                y = d.band(sl, x, y, w, s["band"] + " (cont.)") + 0.12
        while chunk and _is_spacer(chunk[-1]):
            chunk = chunk[:-1]
        draw_table(d, sl, cols, chunk, x, y, w, size, **opts)
        if ci == len(chunks) - 1:
            d.takeaway(sl, s)
        d.footer(sl, s)
        d.notes(sl, s)


def l_text(d: Deck, s):
    """Texto largo estilo documento (oferta, términos, perfiles): párrafos, listas y
    encabezados; opcionalmente en 2 columnas y con bloque de firmas al pie."""
    sl = d.new_slide()
    top = d.title(sl, s)
    x, y, w, h = d.body_box(s, top)
    sigs = s.get("signatures") or []
    if sigs:
        h -= 1.45
    paras = []
    for b in s.get("blocks", []):
        if isinstance(b, str):
            paras.append((b, -1))
        elif "heading" in b:
            paras.append((b["heading"], -2))
        elif "bullets" in b:
            paras.extend(norm_items(b["bullets"]))
        elif "text" in b:
            paras.append((b["text"], -1))
    ncol = int(s.get("columns", 1))
    base = d.sizes.get("body", 15) if d.profile else 16
    mn = d.sizes.get("body_min", 11) if d.profile else 11.5
    if ncol == 2:
        # partir por bloques en el punto más cercano a la mitad del texto
        tot = sum(len(plain(t)) for t, _ in paras)
        acc, cut = 0, len(paras)
        for i, (t, _) in enumerate(paras):
            acc += len(plain(t))
            if acc >= tot / 2:
                cut = i + 1
                break
        while cut < len(paras) and paras[cut][1] > 0:
            cut += 1  # no separar sub-viñetas de su viñeta
        groups = [paras[:cut], paras[cut:]]
        cw = (w - 0.5) / 2
    else:
        groups, cw = [paras], w
    size = min(fit_size([(plain(t), max(l, -1)) for t, l in g], cw, h, d.bf, max_pt=base,
                        min_pt=mn, line_spacing=1.05, space_after_ratio=0.45,
                        indent_in=0.28) for g in groups if g)
    for gi, g in enumerate(groups):
        if g:
            d.text(sl, x + gi * (cw + 0.5), y, cw, h, g, size=size, line_spacing=1.05,
                   space_after=size * 0.45, bullets={"color": d.t["accent"], "indent": 0.28},
                   emph_color=None)
    if sigs:
        n = len(sigs)
        gap = 0.3
        sw = min((w - gap * (n - 1)) / n, 4.2)
        sy = y + h + 0.15
        for i, sg in enumerate(sigs):
            sx = x + i * (sw + gap)
            d.rect(sl, sx, sy, sw, 1.3, None, line=d.t["border"], line_w=1)
            d.text(sl, sx + 0.15, sy + 0.1, sw - 0.3, 0.6,
                   [sg.get("name", ""), sg.get("role", ""), sg.get("company", "")],
                   size=12, fit=9, color=d.t["text"])
            box = sl.shapes[-1].text_frame.paragraphs[0]
            for r in box.runs:
                r.font.bold = True
                r.font.name = d.ef
            d.text(sl, sx + 0.15, sy + 1.02, sw - 0.3, 0.22, sg.get("label", "Firma y fecha"),
                   size=9, color=d.t["text_muted"], italic=True)
    d.takeaway(sl, s)
    d.footer(sl, s)
    d.notes(sl, s)


def l_pricing(d: Deck, s):
    """Modelo económico: banda + resumen (concepto: importe) + tablas de desglose."""
    sl = d.new_slide()
    top = d.title(sl, s)
    x, y, w, h = d.body_box(s, top)
    summary = s.get("summary", [])
    tables = s.get("tables", [])
    def need(k):
        n = (0.42 * k + 0.15 if s.get("band") else 0) + len(summary) * 0.38 * k + 0.15
        for t in tables:
            widths = _col_widths(t["columns"], t["rows"], w, t.get("col_widths"))
            n += (0.36 * k + 0.1 if t.get("band") else 0) + sum(
                _row_heights(d, t["columns"], t["rows"], widths, 14 * k)) + 0.2
        return n
    k = 1.0
    while k > 0.72 and need(k) > h:
        k -= 0.04
    if need(k) > h:
        d.warnings.append(f"Diapositiva {d.page}: el modelo económico no cabe; "
                          "reparte las tablas en dos diapositivas.")
    if s.get("band"):
        y = d.band(sl, x, y, w, s["band"], h=0.42 * k, size=17 * k) + 0.15
    for it in summary:
        strong = it.get("strong")
        txt = f"**{it.get('label', '')}:** {it.get('value', '')}"
        if strong:
            txt = f"**{it.get('label', '')}: {it.get('value', '')}**"
        d.text(sl, x, y, w, 0.36 * k, txt, size=(16 if strong else 15) * k,
               color=d.t["primary"] if strong else d.t["text"], anchor="middle",
               emph_color=d.t["primary"])
        y += 0.38 * k
    y += 0.15
    for t in tables:
        if t.get("band"):
            y = d.band(sl, x, y, w, t["band"], h=0.36 * k, size=14 * k) + 0.1
        aligns = t.get("col_align") or (["left"] + ["right"] * (len(t["columns"]) - 1))
        hh = draw_table(d, sl, t["columns"], t["rows"], x, y, w, 14 * k,
                        widths=_col_widths(t["columns"], t["rows"], w, t.get("col_widths")),
                        col_align=aligns, zebra=False, header_style="light",
                        total_row=t.get("total_row", False))
        y += hh + 0.2
    d.takeaway(sl, s)
    d.footer(sl, s)
    d.notes(sl, s)


def l_timeline(d: Deck, s):
    sl = d.new_slide()
    top = d.title(sl, s)
    x, y, w, h = d.body_box(s, top)
    steps = s.get("steps", [])
    n = len(steps)
    if not 2 <= n <= 7:
        raise SpecError("timeline: entre 2 y 7 hitos")
    sw = w / n
    line_y = y + 1.0
    current = s.get("current")
    tsize = min(fit_size([plain(st.get("title", ""))], sw - 0.25, 0.7, d.hf, True, 17, 12)
                for st in steps)
    bsize = min(fit_size([plain(st.get("text", ""))], sw - 0.25, h - 2.0, d.bf, False, 14, 10,
                         line_spacing=1.1) for st in steps)
    need = max(text_height([plain(st.get("text", ""))], sw - 0.24, d.bf, bsize,
                           line_spacing=1.1) for st in steps)
    block = 2.15 + need
    shift = max(0, (h - block) * 0.4)
    y += shift
    line_y += shift
    h = min(h, block + 0.1)
    d.line(sl, x, line_y, x + w, line_y, d.t["border"], 2)
    for i, st in enumerate(steps):
        cx = x + sw * i + sw / 2
        done = current is not None and i < current
        active = current is None or i == current
        col = d.t["accent"] if active else (d.t["primary"] if done else d.t["border"])
        label_col = d.acc_txt if active else d.t["text_muted"]
        if current is not None and i < current:
            d.line(sl, x + sw * i + (sw / 2 if i == 0 else 0), line_y,
                   x + sw * (i + 1), line_y, d.t["primary"], 2)
        d.text(sl, cx - sw / 2 + 0.1, y, sw - 0.2, 0.55, st.get("label", ""), size=15,
               fit=11, bold=True, color=label_col, align="center", anchor="bottom")
        d.circle_label(sl, cx, line_y, 0.34, "", col)
        d.text(sl, cx - sw / 2 + 0.12, line_y + 0.4, sw - 0.24, 0.7, st.get("title", ""),
               size=tsize, bold=True, font=d.hf, align="center", color=d.t["text"])
        d.text(sl, cx - sw / 2 + 0.12, line_y + 1.15, sw - 0.24, h - 2.15,
               st.get("text", ""), size=bsize, line_spacing=1.1, align="center",
               color=d.t["text_muted"])
    d.takeaway(sl, s)
    d.footer(sl, s)
    d.notes(sl, s)


def l_process(d: Deck, s):
    sl = d.new_slide()
    top = d.title(sl, s)
    x, y, w, h = d.body_box(s, top)
    steps = s.get("steps", [])
    n = len(steps)
    if not 2 <= n <= 6:
        raise SpecError("process: entre 2 y 6 pasos")
    overlap = 0.12
    sw = (w + overlap * (n - 1)) / n
    ch = 0.95
    hl = s.get("highlight")
    tsize = min(fit_size([plain(st.get("title", ""))], sw - 0.75, ch - 0.2, d.hf, True,
                         17, 11) for st in steps)
    bsize = min(fit_size([plain(st.get("text", ""))], sw - 0.3, h - ch - 0.4, d.bf, False,
                         15, 10.5, line_spacing=1.1) for st in steps)
    need = max(text_height([plain(st.get("text", ""))], sw - 0.45, d.bf, bsize,
                           line_spacing=1.1) for st in steps)
    block = ch + 0.35 + need
    y = y + max(0, (h - block) * 0.4)
    h = min(h, block + 0.1)
    for i, st in enumerate(steps):
        sx = x + i * (sw - overlap)
        shape = MSO_SHAPE.PENTAGON if i == 0 else MSO_SHAPE.CHEVRON
        fill = d.t["accent"] if hl == i else mix(d.t["primary"], d.t["bg"], 0.0 + 0.12 * i
                                                  if not d.t["dark_content"] else 0)
        shp = d.rect(sl, sx, y, sw, ch, fill, shape=shape)
        try:
            shp.adjustments[0] = 0.32
        except Exception:
            pass
        shp.text_frame.margin_left = IN(0.35 if i else 0.2)
        shp.text_frame.margin_right = IN(0.3)
        d.text(sl, 0, 0, 0, 0, st.get("title", ""), size=tsize, bold=True, font=d.hf,
               color="FFFFFF", align="center", anchor="middle", shape=shp)
        d.text(sl, sx + 0.15, y + ch + 0.35, sw - 0.45, h - ch - 0.4, st.get("text", ""),
               size=bsize, line_spacing=1.1, color=d.t["text_muted"],
               emph_color=d.t["text"])
    d.takeaway(sl, s)
    d.footer(sl, s)
    d.notes(sl, s)


def l_comparison(d: Deck, s):
    sl = d.new_slide()
    top = d.title(sl, s)
    x, y, w, h = d.body_box(s, top)
    opts = s.get("options", [])
    n = len(opts)
    if not 2 <= n <= 4:
        raise SpecError("comparison: entre 2 y 4 opciones")
    gap = 0.3
    ow = (w - gap * (n - 1)) / n
    head_h = 0.85
    if any(o.get("recommended") for o in opts):  # hueco para la etiqueta "Recomendada"
        y += 0.25
        h -= 0.25
    body = []
    for o in opts:
        body.append([("✓ " + p, 0) for p in o.get("pros", [])] +
                    [("✕ " + c, 0) for c in o.get("cons", [])])
    size = min(fit_size([(plain(t), 0) for t, _ in b], ow - 0.6, h - head_h - 0.9, d.bf,
                        max_pt=18, min_pt=10.5, space_after_ratio=0.45) for b in body if b)
    for i, o in enumerate(opts):
        ox = x + i * (ow + gap)
        rec = o.get("recommended")
        d.rect(sl, ox, y, ow, h, d.t["bg"], line=d.t["accent"] if rec else d.t["border"],
               line_w=2.25 if rec else 0.75)
        d.rect(sl, ox, y, ow, head_h, d.t["primary"] if rec else d.t["surface"])
        d.text(sl, ox + 0.3, y, ow - 0.6, head_h, o.get("name", ""), size=19, fit=13,
               bold=True, font=d.hf, anchor="middle",
               color="FFFFFF" if rec else d.t["text"])
        if rec:
            badge = d.rect(sl, ox + ow - 1.75, y - 0.2, 1.55, 0.36, d.t["accent"],
                           shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
            d.text(sl, 0, 0, 0, 0, o.get("badge", "Recomendada"), size=10.5, bold=True,
                   color="FFFFFF", align="center", anchor="middle", shape=badge)
        by = y + head_h + 0.3
        bh = h - head_h - 0.6 - (0.6 if o.get("verdict") else 0)
        box = sl.shapes.add_textbox(IN(ox + 0.3), IN(by), IN(ow - 0.6), IN(bh))
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        first = True
        for kind, items in (("pro", o.get("pros", [])), ("con", o.get("cons", []))):
            for it in items:
                p = tf.paragraphs[0] if first else tf.add_paragraph()
                first = False
                p.space_after = PT(size * 0.45)
                for txt, em in runs_of(it):
                    r = p.add_run()
                    r.text = txt
                    r.font.size = PT(size)
                    r.font.name = d.bf
                    r.font.bold = em
                    r.font.color.rgb = rgb(d.t["text"])
                _set_bullet(p, "✓" if kind == "pro" else "✕",
                            d.t["positive"] if kind == "pro" else d.t["negative"], 0.3,
                            font="Segoe UI Symbol")
        if o.get("verdict"):
            d.line(sl, ox + 0.3, y + h - 0.75, ox + ow - 0.3, y + h - 0.75, d.t["border"])
            d.text(sl, ox + 0.3, y + h - 0.65, ow - 0.6, 0.5, o["verdict"], size=14, fit=10,
                   bold=True, anchor="middle", color=d.acc_txt if rec else d.t["text"])
    d.takeaway(sl, s)
    d.footer(sl, s)
    d.notes(sl, s)


def l_matrix(d: Deck, s):
    sl = d.new_slide()
    top = d.title(sl, s)
    x, y, w, h = d.body_box(s, top)
    q = s.get("quadrants", [])
    if len(q) != 4:
        raise SpecError("matrix: exactamente 4 cuadrantes (sup-izq, sup-dcha, inf-izq, inf-dcha)")
    side_text = s.get("bullets")
    mw = w * 0.62 if side_text else w
    ax = x + 0.45
    aw, ah = mw - 0.45, h - 0.4
    gap = 0.12
    qw, qh = (aw - gap) / 2, (ah - gap) / 2
    hl = s.get("highlight")
    size = min(fit_size([plain(qq.get("text", ""))], qw - 0.5, qh - 0.95, d.bf, False, 14, 10)
               for qq in q)
    for i, qq in enumerate(q):
        r, c = divmod(i, 2)
        qx, qy = ax + c * (qw + gap), y + r * (qh + gap)
        on = hl == i
        d.rect(sl, qx, qy, qw, qh, d.t["primary"] if on else d.t["surface"])
        d.text(sl, qx + 0.25, qy + 0.22, qw - 0.5, 0.5, qq.get("title", ""), size=17, fit=12,
               bold=True, font=d.hf, color="FFFFFF" if on else d.t["text"])
        d.text(sl, qx + 0.25, qy + 0.75, qw - 0.5, qh - 0.95, qq.get("text", ""), size=size,
               color=mix("FFFFFF", d.t["primary"], 0.18) if on else d.t["text_muted"])
    # ejes
    d.text(sl, ax, y + ah + 0.08, aw, 0.3, (s.get("x_label", "") + "  →").strip(" →") + "  →",
           size=12, bold=True, color=d.t["text_muted"], align="center")
    ylab = d.text(sl, x - 1.2, y + ah / 2 - 0.15, ah, 0.3,
                  (s.get("y_label", "")) + "  →", size=12, bold=True,
                  color=d.t["text_muted"], align="center")
    ylab.rotation = -90
    ylab.left = IN(x + 0.22 - ah / 2)
    ylab.top = IN(y + ah / 2 - 0.15)
    if side_text:
        sx = x + mw + 0.5
        d.text(sl, sx, y, w - mw - 0.5, h, norm_items(side_text), size=16, fit=11,
               space_after=9, bullets={"color": d.t["accent"]}, emph_color=d.t["text"])
    d.takeaway(sl, s)
    d.footer(sl, s)
    d.notes(sl, s)


def l_quote(d: Deck, s):
    dark = s.get("dark", False) and not d.profile
    sl = _bare_content_slide(d, s) if d.profile else d.new_slide(dark=dark)
    col = d.t["text_on_dark"] if dark else d.t["text"]
    d.text(sl, M + 0.4, 0.3, 2, 3.0, "“", size=160, bold=True, font=d.hf,
           color=d.t["accent"])
    d.text(sl, M + 1.2, 2.2, W - 2 * M - 2.2, 2.8, s.get("quote", ""), size=32, fit=20,
           font=d.hf, italic=True, color=col, anchor="middle", line_spacing=1.1,
           emph_color=d.t["accent"])
    who = s.get("author", "")
    role = s.get("role", "")
    d.rect(sl, M + 1.2, 5.4, 0.6, 0.05, d.t["accent"])
    d.text(sl, M + 1.2, 5.6, 9, 0.4, who, size=16, bold=True, color=col)
    if role:
        d.text(sl, M + 1.2, 6.0, 9, 0.4, role, size=14,
               color=d.t["text_on_dark_muted"] if dark else d.t["text_muted"])
    d.footer(sl, s, dark=dark)
    d.notes(sl, s)


def _picture_fill(d: Deck, sl, path, x, y, w, h):
    """Inserta una imagen recortada para llenar la caja (sin deformar)."""
    if not path or not os.path.exists(path):
        ph = d.rect(sl, x, y, w, h, mix(d.t["border"], d.t["bg"], 0.3))
        d.text(sl, 0, 0, 0, 0, f"[Imagen: {os.path.basename(str(path)) or 'pendiente'}]",
               size=14, color=d.t["text_muted"], align="center", anchor="middle", shape=ph)
        d.warnings.append(f"Diapositiva {d.page}: imagen no encontrada '{path}' "
                          "(se ha dejado un marcador).")
        return ph
    from PIL import Image
    with Image.open(path) as im:
        iw, ih = im.size
    pic = sl.shapes.add_picture(path, IN(x), IN(y), IN(w), IN(h))
    box_r, img_r = w / h, iw / ih
    if img_r > box_r:
        crop = (1 - box_r / img_r) / 2
        pic.crop_left = pic.crop_right = crop
    else:
        crop = (1 - img_r / box_r) / 2
        pic.crop_top = pic.crop_bottom = crop
    return pic


def l_image(d: Deck, s):
    if d.profile:
        return _p_image(d, s)
    sl = d.new_slide()
    side = s.get("side", "right")
    iw = W * 0.45
    ix = W - iw if side == "right" else 0
    _picture_fill(d, sl, s.get("image"), ix, 0, iw, H)
    tx = M if side == "right" else iw + 0.6
    tw = W - iw - M - 0.6
    top = d.title(sl, s, x=tx, w=tw)
    bottom = BOTTOM - (0.3 if s.get("source") else 0)
    items = norm_items(s.get("bullets", []))
    if items:
        d.text(sl, tx, top, tw, bottom - top, items, size=20, fit=12, space_after=9,
               bullets={"color": d.t["accent"]}, emph_color=d.t["text"])
    elif s.get("text"):
        d.text(sl, tx, top, tw, bottom - top, s["text"], size=20, fit=12, line_spacing=1.1)
    if s.get("caption"):
        cap = d.rect(sl, ix, H - 0.5, iw, 0.5, "000000")
        cap.fill.fore_color.rgb = rgb("000000")
        _set_alpha(cap, 55)
        d.text(sl, ix + 0.25, H - 0.45, iw - 0.5, 0.4, s["caption"], size=11,
               color="FFFFFF", anchor="middle")
    # pie solo en el lado del texto
    muted = d.t["text_muted"]
    if d.meta.get("footer"):
        d.text(sl, tx, FOOTER_Y, tw - 1, 0.25, d.meta["footer"], size=9, color=muted)
    if d.meta.get("slide_numbers", True):
        d.text(sl, tx + tw - 1, FOOTER_Y, 1, 0.25, str(d.page), size=9, color=muted,
               align="right")
    if s.get("source"):
        d.text(sl, tx, FOOTER_Y - 0.3, tw, 0.25, "Fuente: " + s["source"], size=9,
               italic=True, color=d.t["text_muted"])
    d.notes(sl, s)


def _p_image(d: Deck, s):
    """Con plantilla: la imagen va dentro del área de contenido (no a sangre)."""
    sl = d.new_slide()
    top = d.title(sl, s)
    x, y, w, h = d.body_box(s, top)
    side = s.get("side", "right")
    iw = w * 0.5
    tw = w - iw - 0.5
    ix = x + w - iw if side == "right" else x
    tx = x if side == "right" else x + iw + 0.5
    _picture_fill(d, sl, s.get("image"), ix, y, iw, h - (0.35 if s.get("caption") else 0))
    if s.get("caption"):
        d.text(sl, ix, y + h - 0.3, iw, 0.3, s["caption"], size=11, italic=True,
               color=d.t["text_muted"])
    items = norm_items(s.get("bullets", []))
    if items:
        d.text(sl, tx, y, tw, h, items, size=18, fit=11, space_after=8,
               bullets={"color": d.t["accent"]}, emph_color=d.t["text"])
    elif s.get("text"):
        d.text(sl, tx, y, tw, h, s["text"], size=18, fit=11, line_spacing=1.1)
    d.footer(sl, s)
    d.notes(sl, s)


def _set_alpha(shape, pct):
    sf = shape.fill._xPr.find(qn("a:solidFill"))
    if sf is not None and len(sf):
        a = etree.SubElement(sf[0], qn("a:alpha"))
        a.set("val", str(int(pct * 1000)))


def l_closing(d: Deck, s):
    if d.profile:
        return _p_closing(d, s)
    sl = d.new_slide(dark=True)
    t = d.t
    d.rect(sl, M, 2.0, 0.9, 0.07, t["accent"])
    d.text(sl, M, 2.3, W - 2 * M - 2, 1.4, s.get("title", "Gracias"), size=48, fit=30,
           bold=True, font=d.hf, color=t["text_on_dark"], emph_color=t["accent"])
    y = 3.75
    if s.get("subtitle"):
        d.text(sl, M, y, W - 2 * M - 3, 0.9, s["subtitle"], size=20, fit=14,
               color=t["text_on_dark_muted"])
        y += 1.0
    if s.get("next_steps"):
        d.text(sl, M, y, W - 2 * M - 3, 6.3 - y, norm_items(s["next_steps"]), size=18, fit=12,
               space_after=6, color=t["text_on_dark"], bullets={"color": t["accent"]})
    contact = s.get("contact")
    if contact:
        lines = contact if isinstance(contact, list) else [contact]
        d.text(sl, M, 6.35, W - 2 * M, 0.5, "   ·   ".join(lines), size=13,
               color=t["text_on_dark_muted"])
    d.notes(sl, s)


LAYOUTS = {
    "title": (l_title, "Portada: título, subtítulo, autor/fecha; imagen opcional a la derecha."),
    "section": (l_section, "Separador de sección con número grande."),
    "agenda": (l_agenda, "Índice numerado (1–10 puntos); 'current' resalta la sección actual."),
    "bullets": (l_bullets, "Título + viñetas (máx. 6) con 'callout' opcional (cifra o frase)."),
    "two-column": (l_two_column, "Dos columnas con encabezado; style='contrast' para antes/después."),
    "cards": (l_cards, "2–8 tarjetas (pilares, capacidades, beneficios)."),
    "kpis": (l_kpis, "2–6 indicadores con valor, etiqueta y variación."),
    "big-number": (l_big_number, "Una cifra protagonista + contexto."),
    "statement": (l_statement, "Mensaje clave a pantalla completa (transición o conclusión)."),
    "chart": (l_chart, "Gráfico nativo editable + panel de 'insights' opcional."),
    "table": (l_table, "Tabla estilizada (máx. ~12 filas × 6 columnas)."),
    "timeline": (l_timeline, "Cronología horizontal de 2–7 hitos; 'current' marca el actual."),
    "process": (l_process, "Proceso en flechas de 2–6 pasos con descripción."),
    "comparison": (l_comparison, "2–4 opciones con pros/contras y recomendada."),
    "matrix": (l_matrix, "Matriz 2×2 con ejes y cuadrante destacado."),
    "quote": (l_quote, "Cita o testimonio."),
    "image": (l_image, "Imagen a sangre en media pantalla + texto."),
    "closing": (l_closing, "Cierre: mensaje final, próximos pasos y contacto."),
    "text": (l_text, "Texto largo tipo documento: párrafos, listas, encabezados; 1–2 columnas; firmas."),
    "pricing": (l_pricing, "Modelo económico: banda, resumen de importes y tablas de desglose."),
}


# ----------------------------------------------------------------------------
# Principal
# ----------------------------------------------------------------------------
def load_theme(name=None, theme_file=None, overrides=None):
    with open(THEMES_PATH, encoding="utf-8") as fh:
        themes = json.load(fh)
    if theme_file:
        with open(theme_file, encoding="utf-8") as fh:
            custom = json.load(fh)
        base = copy.deepcopy(themes.get(custom.get("base", "ejecutivo")))
        base.update(custom)
        theme = base
    else:
        name = name or "ejecutivo"
        if name not in themes or name.startswith("_"):
            raise SpecError(f"Tema '{name}' desconocido. Disponibles: "
                            f"{', '.join(k for k in themes if not k.startswith('_'))}")
        theme = copy.deepcopy(themes[name])
    if overrides:
        theme.update(overrides)
    base = themes["ejecutivo"]
    hexre = re.compile(r"^[0-9A-Fa-f]{6}$")
    for k, v in list(theme.items()):
        if isinstance(v, str) and k in base and hexre.match(str(base[k])):
            v = v.lstrip("#")
            if not hexre.match(v):
                print(f"AVISO: color '{k}'='{v}' no es hex válido; se usa {base[k]}",
                      file=sys.stderr)
                v = base[k]
            theme[k] = v.upper()
    theme["chart"] = [c.lstrip("#").upper() for c in theme.get("chart", [])
                      if hexre.match(c.lstrip("#"))] or base["chart"]
    missing = [k for k in base if k not in theme]
    for k in missing:
        theme[k] = base[k]
    return theme


def load_profile(name_or_path):
    """Carga un perfil de plantilla: nombre de carpeta en templates/ o ruta a profile.json."""
    if not name_or_path:
        return None, None
    path = name_or_path
    if not path.endswith(".json"):
        path = os.path.join(TEMPLATES_DIR, name_or_path, "profile.json")
    if not os.path.exists(path):
        avail = [d for d in os.listdir(TEMPLATES_DIR)] if os.path.isdir(TEMPLATES_DIR) else []
        raise SpecError(f"Perfil '{name_or_path}' no encontrado. Disponibles: {', '.join(avail)}")
    with open(path, encoding="utf-8") as fh:
        prof = json.load(fh)
    tpl = os.path.join(os.path.dirname(os.path.abspath(path)), prof["template"])
    if not os.path.exists(tpl):
        raise SpecError(f"El perfil apunta a una plantilla que no existe: {tpl}")
    return prof, tpl


def build(spec: dict, out: str, theme_name=None, theme_file=None, template=None,
          profile=None):
    meta = spec.get("meta", {})
    if profile is None:
        profile = meta.get("profile", DEFAULT_PROFILE)
        if theme_name or theme_file or template or meta.get("theme") or meta.get("template"):
            profile = meta.get("profile")  # tema/plantilla explícitos: sin perfil salvo que se pida
    prof, prof_tpl = load_profile(profile) if profile and profile != "none" else (None, None)
    if prof:
        theme = load_theme("ejecutivo", None, {**prof.get("theme", {}),
                                               **(spec.get("theme_overrides") or {})})
        deck = Deck(theme, meta, prof_tpl, profile=prof)
    else:
        theme = load_theme(theme_name or meta.get("theme"), theme_file or meta.get("theme_file"),
                           spec.get("theme_overrides"))
        deck = Deck(theme, meta, template or meta.get("template"))
    base_dir = spec.get("_base_dir", ".")
    for i, s in enumerate(spec.get("slides", []), 1):
        lay = s.get("layout", "bullets")
        if lay not in LAYOUTS:
            raise SpecError(f"Diapositiva {i}: layout '{lay}' no existe. "
                            f"Usa uno de: {', '.join(LAYOUTS)}")
        if s.get("image") and not os.path.isabs(s["image"]):
            cand = os.path.join(base_dir, s["image"])
            if os.path.exists(cand):
                s["image"] = cand
        try:
            LAYOUTS[lay][0](deck, s)
        except SpecError as e:
            raise SpecError(f"Diapositiva {i} ({lay}): {e}")
    cp = deck.prs.core_properties
    cp.title = meta.get("title", "")
    cp.author = meta.get("author", "")
    cp.subject = meta.get("subject", "")
    deck.prs.save(out)
    return deck


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("spec", nargs="?", help="Especificación JSON del deck")
    ap.add_argument("-o", "--out", default="presentacion.pptx")
    ap.add_argument("--theme", help="Nombre del tema (ver themes/themes.json)")
    ap.add_argument("--theme-file", help="JSON con un tema propio (puede incluir 'base')")
    ap.add_argument("--template", help=".pptx/.potx corporativo cuyo patrón se reutiliza")
    ap.add_argument("--profile", help="Perfil de plantilla (carpeta en templates/ o ruta a "
                    f"profile.json). Por defecto: {DEFAULT_PROFILE}")
    ap.add_argument("--no-profile", action="store_true",
                    help="Usar los temas genéricos en lugar de la plantilla corporativa")
    ap.add_argument("--list-layouts", action="store_true")
    ap.add_argument("--list-themes", action="store_true")
    a = ap.parse_args()
    if a.list_layouts:
        for k, (_, desc) in LAYOUTS.items():
            print(f"{k:12s} {desc}")
        return
    if a.list_themes:
        with open(THEMES_PATH, encoding="utf-8") as fh:
            for k, v in json.load(fh).items():
                if not k.startswith("_"):
                    print(f"{k:14s} {v.get('description', '')}")
        return
    if not a.spec:
        ap.error("falta la especificación JSON")
    with open(a.spec, encoding="utf-8") as fh:
        spec = json.load(fh)
    spec["_base_dir"] = os.path.dirname(os.path.abspath(a.spec))
    try:
        prof = "none" if a.no_profile else a.profile
        deck = build(spec, a.out, a.theme, a.theme_file, a.template, prof)
    except SpecError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(2)
    print(f"OK: {a.out} ({deck.page} diapositivas)")
    for w in deck.warnings:
        print("AVISO:", w)


if __name__ == "__main__":
    main()

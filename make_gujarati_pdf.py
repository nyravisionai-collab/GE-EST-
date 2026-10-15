#!/usr/bin/env python3
"""
ગુજરાતી કોટેશન (DOCX) → PDF
============================
'GE EST માતૃશ્રી બંગ્લોઝ - ગુજરાતી.docx' ફાઈલમાંથી બરાબર એ જ લખાણ, એ જ ક્રમ,
એ જ રંગો અને એ જ પેજ-બ્રેક સાથે ગુજરાતી PDF બનાવે છે.

ગુજરાતી + Latin લખાણ 'Noto Sans Gujarati' અને 'Noto Sans' ફોન્ટને જોડીને બનાવેલા
એક જ embedded ફોન્ટમાં લખાય છે.  ગુજરાતીના જોડાક્ષર (ligature) બરાબર બેસે તે માટે
fpdf2 નું text-shaping (uharfbuzz) ચાલુ કરેલું છે.
"""

import os
import re
import sys

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from fpdf import FPDF
from fpdf.enums import Align, MethodReturnValue

HERE = os.path.dirname(os.path.abspath(__file__))
DOCX = os.path.join(HERE, "GE EST માતૃશ્રી બંગ્લોઝ - ગુજરાતી.docx")
PDF_OUT = os.path.join(HERE, "GE EST માતૃશ્રી બંગ્લોઝ - ગુજરાતી.pdf")

# ફોન્ટ: Noto Sans Gujarati (ગુજરાતી) + Noto Sans (Latin/ચિહ્નો) – બંનેને એક ફોન્ટમાં
# જોડી દેવાય છે, જેથી એક જ ફોન્ટમાં બધા અક્ષરો મળે અને PDF માં લખાણ ચોક્કસ રીતે બેસે.
FONT_DIR = os.path.join(HERE, "fonts")
CACHE_DIR = "/tmp/gujarati_fonts"
CS_REG = os.path.join(FONT_DIR, "NotoSansGujarati-Regular.ttf")
CS_BOLD = os.path.join(FONT_DIR, "NotoSansGujarati-Bold.ttf")
LAT_REG = os.path.join(FONT_DIR, "NotoSans-Regular.ttf")
LAT_BOLD = os.path.join(FONT_DIR, "NotoSans-Bold.ttf")


def prepare_fonts():
    """ગુજરાતી + Latin ફોન્ટને જોડીને એક ફોન્ટ બનાવો (cache સાથે)."""
    os.makedirs(CACHE_DIR, exist_ok=True)
    out = {}
    for style, cs, lat in (("Regular", CS_REG, LAT_REG), ("Bold", CS_BOLD, LAT_BOLD)):
        target = os.path.join(CACHE_DIR, "Combined-%s.ttf" % style)
        if not os.path.exists(target) or os.path.getmtime(target) < max(
            os.path.getmtime(cs), os.path.getmtime(lat)
        ):
            from fontTools.merge import Merger

            Merger().merge([cs, lat]).save(target)
        out[style] = target
    return out["Regular"], out["Bold"]

NAVY = (31, 78, 121)      # 1F4E79
BLUE = (46, 117, 182)     # 2E75B6
DARK = (26, 26, 26)       # 1A1A1A
GRAY = (64, 64, 64)       # 404040
LGRAY = (120, 120, 120)
BAND = (222, 234, 246)    # DEEAF6
LINE = (190, 200, 210)

MARGIN = 15
USABLE = 210 - 2 * MARGIN


# ==========================================================================
#  દસ્તાવેજમાંથી બ્લોક્સ વાંચો
# ==========================================================================
def para_kind(par):
    """paragraph ના ફોર્મેટિંગ પરથી તેનો પ્રકાર નક્કી કરો."""
    runs = [r for r in par.runs if r.text.strip()]
    if not runs:
        return ("empty", {})
    r = runs[0]
    size = r.font.size.pt if r.font.size else 10
    bold = bool(r.font.bold)
    color = str(r.font.color.rgb) if (r.font.color and r.font.color.type is not None) else "1A1A1A"
    align = par.paragraph_format.alignment
    centered = str(align).startswith("CENTER")
    bullet = par._p.pPr is not None and par._p.pPr.find(qn("w:numPr")) is not None
    info = dict(size=size, bold=bold, color=color, centered=centered, bullet=bullet)
    if bool(bullet):
        return ("bullet", info)
    if bold and size >= 19:
        return ("company", info)
    if bold and size >= 17:
        return ("doctitle", info)
    if color == "1F4E79" and bold and size >= 12.5:
        return ("section", info)
    if color == "2E75B6" and bold and size >= 10.5:
        return ("subhead", info)
    if color == "2E75B6" and bold:
        return ("subsub", info)
    if color == "2E75B6" and centered:
        return ("subtitle", info)
    if color == "2E75B6" and bold:
        return ("toc", info)
    if color in ("404040",) and centered:
        return ("address", info)
    if size <= 8.5 and color == "1A1A1A":
        return ("small", info)
    if color == "666666":
        return ("note", info)
    return ("body", info)


def table_rows(tbl):
    """(rows, grid_widths) – દરેક cell: {text, span}"""
    grid = tbl.find(qn("w:tblGrid"))
    widths = [int(float(c.get(qn("w:w")))) for c in grid.findall(qn("w:gridCol"))]
    rows = []
    for tr in tbl.findall(qn("w:tr")):
        row = []
        for tc in tr.findall(qn("w:tc")):
            span = 1
            tcPr = tc.find(qn("w:tcPr"))
            if tcPr is not None:
                gs = tcPr.find(qn("w:gridSpan"))
                if gs is not None:
                    span = int(gs.get(qn("w:val")))
            paras = [p for p in tc.findall(qn("w:p"))]
            text = "\n".join(
                "".join(t.text or "" for t in p.iter(qn("w:t"))) for p in paras
            ).strip("\n")
            row.append(dict(text=text.strip(), span=span))
        rows.append(row)
    return rows, widths


def collect_blocks(document):
    """બોડી + ટેબલ – દસ્તાવેજના ક્રમમાં."""
    blocks = []
    tidx = 0
    for child in document.element.body.iterchildren():
        if child.tag == qn("w:p"):
            blocks.append(("p", Paragraph(child, document)))
        elif child.tag == qn("w:tbl"):
            rows, widths = table_rows(child)
            blocks.append(("tbl", (tidx, rows, widths)))
            tidx += 1
    return blocks


def para_page_break(par):
    for br in par._p.iter(qn("w:br")):
        if br.get(qn("w:type")) == "page":
            return True
    return False


# ==========================================================================
#  PDF
# ==========================================================================
class QuotePDF(FPDF):
    plain = False          # True હોય તો હેડર/ફૂટર વગર (QA માટે)

    def header(self):
        if self.plain or self.page_no() == 1:
            return
        self.set_font("NotoGuj", "", 7.5)
        self.set_text_color(*LGRAY)
        self.set_xy(MARGIN, 8)
        self.cell(USABLE, 5, "ઘનશ્યામ ઇલેક્ટ્રિકલ્સ  —  ઇલેક્ટ્રિકલ કામનું કોટેશન", align="L")
        self.set_draw_color(*LINE)
        self.line(MARGIN, 13.5, 210 - MARGIN, 13.5)
        self.set_y(18)

    def footer(self):
        if self.plain:
            return
        self.set_y(-13)
        self.set_font("NotoGuj", "", 7.5)
        self.set_text_color(*LGRAY)
        self.cell(USABLE, 5, f"પાનું {self.page_no()}", align="C")


def wrap_height(pdf, text, width, line_h, font=None, style="", size=None):
    if font:
        pdf.set_font(font, style, size)
    return pdf.multi_cell(
        w=width, h=line_h, text=text, dry_run=True, output=MethodReturnValue.HEIGHT
    )


def draw_paragraph(pdf, kind, info, text):
    """એક paragraph દોરો (પ્રકાર પ્રમાણે ફોન્ટ/રંગ/જગ્યા)."""
    style = dict(
        company=("NotoGuj", "B", 17, NAVY, 4.0, 0, "C"),
        doctitle=("NotoGuj", "B", 15, NAVY, 3.0, 2, "C"),
        subtitle=("NotoGuj", "", 10.5, BLUE, 2.0, 0, "C"),
        address=("NotoGuj", "", 9.5, GRAY, 0.5, 0, "C"),
        section=("NotoGuj", "B", 12.5, NAVY, 4.0, 0, "L"),
        subhead=("NotoGuj", "B", 10.5, BLUE, 1.5, 0, "L"),
        subsub=("NotoGuj", "B", 9.5, BLUE, 1.0, 0, "L"),
        toc=("NotoGuj", "B", 10, BLUE, 0.2, 0, "L"),
        body=("NotoGuj", "", 10, DARK, 1.0, 0, "L"),
        note=("NotoGuj", "", 9, (102, 102, 102), 1.5, 0, "L"),
        small=("NotoGuj", "", 8.5, DARK, 0.6, 0, "L"),
        bullet=("NotoGuj", "", info.get("size", 10), DARK, 0.3, 0, "L"),
    )
    font, fstyle, size, color, space_before, indent, align = style[kind]
    if kind == "bullet":
        size = min(size, 10)
        indent = 6
        text = "•   " + text
    pdf.set_font(font, fstyle, size)
    pdf.set_text_color(*color)
    left = MARGIN + indent
    width = USABLE - indent
    line_h = size * 0.56
    if pdf.get_y() + space_before > pdf.page_break_trigger:
        pdf.add_page()
    pdf.set_y(pdf.get_y() + space_before)
    pdf.set_x(left)
    pdf.multi_cell(w=width, h=line_h, text=text, align=align)
    # section headings નીચે પાતળી લાઇન
    if kind == "section":
        pdf.set_draw_color(*LINE)
        pdf.line(MARGIN, pdf.get_y() + 1.2, 210 - MARGIN, pdf.get_y() + 1.2)
        pdf.set_y(pdf.get_y() + 2.5)


# ---- ટેબલ -----------------------------------------------------------------
def draw_table_header(pdf, cols, widths, x):
    line_h = 5
    y0 = pdf.get_y()
    heights = [
        wrap_height(pdf, cols[i], widths[i] - 3, line_h, "NotoGuj", "B", 9)
        for i in range(len(cols))
    ]
    h = max(heights) + 2
    pdf.set_fill_color(*NAVY)
    pdf.set_text_color(255, 255, 255)
    pdf.set_draw_color(*NAVY)
    cx = x
    for i, col in enumerate(cols):
        pdf.rect(cx, y0, widths[i], h, style="DF")
        pdf.set_font("NotoGuj", "B", 9)
        pdf.set_xy(cx + 1.5, y0 + 1)
        pdf.multi_cell(w=widths[i] - 3, h=line_h, text=col, align="C")
        cx += widths[i]
    pdf.set_y(y0 + h)
    pdf.set_text_color(*DARK)


def draw_pricing_row(pdf, row, widths, x, banded=False):
    """એક row દોરો.  દરેક cell પોતાના span સાથે – કોઈ પણ જોડાયેલા cells હોય તો પણ સાચી જગ્યાએ."""
    line_h = 4.6
    # cell → (કઈ કોલમથી શરૂ, કેટલી કોલમ, લખાણ, alignment)
    cells, col = [], 0
    for cell in row:
        span = cell.get("span", 1)
        cells.append((col, span, cell["text"], cell.get("align")))
        col += span
    heights = []
    for c0, span, text, _ in cells:
        w = sum(widths[c0 : c0 + span]) - 3
        heights.append(wrap_height(pdf, text, w, line_h, "NotoGuj", "", 9))
    h = max(heights) + 1.6
    if pdf.get_y() + h > pdf.page_break_trigger:
        pdf.add_page()
        draw_table_header(pdf, HEADER_COLS, widths, x)
    y0 = pdf.get_y()
    pdf.set_draw_color(*LINE)
    pdf.set_fill_color(*(BAND if banded else (255, 255, 255)))
    pdf.rect(x, y0, sum(widths), h, style="DF")
    for c0, span, text, align in cells:
        cx = x + sum(widths[:c0])
        w = sum(widths[c0 : c0 + span])
        for k in range(c0, c0 + span):        # દરેક કોલમ વચ્ચેની ઊભી લાઇન
            gx = x + sum(widths[:k])
            if gx > x and not (c0 < k < c0 + span and text):
                pdf.line(gx, y0, gx, y0 + h)
        pdf.set_font("NotoGuj", "B" if banded else "", 9)
        pdf.set_text_color(*(NAVY if banded else DARK))
        pdf.set_xy(cx + 1.5, y0 + 0.9)
        if not align:
            align = "L" if span > 1 else "C"
        pdf.multi_cell(w=w - 3, h=line_h, text=text, align=align)
    pdf.set_y(y0 + h)
    pdf.set_text_color(*DARK)


def draw_kv_table(pdf, rows, x, key_w=55, val_w=None, small=False):
    """2-કોલમ ટેબલ (સ્પષ્ટીકરણ / પ્રમાણપત્ર)."""
    line_h = 4.8
    val_w = val_w or (USABLE - key_w)
    pdf.set_font("NotoGuj", "", 9.5)
    for key, val in rows:
        hk = wrap_height(pdf, key, key_w - 3, line_h, "NotoGuj", "B", 9.5)
        hv = wrap_height(pdf, val, val_w - 3, line_h, "NotoGuj", "", 9.5)
        h = max(hk, hv, line_h) + 1.6
        if pdf.get_y() + h > pdf.page_break_trigger:
            pdf.add_page()
        y0 = pdf.get_y()
        pdf.set_draw_color(*LINE)
        pdf.rect(x, y0, key_w, h, style="D")
        pdf.rect(x + key_w, y0, val_w, h, style="D")
        pdf.set_font("NotoGuj", "B", 9.5)
        pdf.set_xy(x + 1.5, y0 + 0.8)
        pdf.multi_cell(w=key_w - 3, h=line_h, text=key)
        pdf.set_font("NotoGuj", "", 9.5)
        pdf.set_xy(x + key_w + 1.5, y0 + 0.8)
        pdf.multi_cell(w=val_w - 3, h=line_h, text=val)
        pdf.set_y(y0 + h)


def draw_signature_block(pdf, rows, x):
    """સહી બ્લોક: ડાબે ક્લાયન્ટ, જમણે ઘનશ્યામ ઇલેક્ટ્રિકલ્સ (સહી માટે જગ્યા સાથે)."""
    if not rows:
        return
    row = rows[0]
    left_txt = row[0]["text"] if len(row) > 0 else ""
    right_txt = row[-1]["text"] if len(row) > 2 else ""
    col_w = (USABLE - 25) / 2
    y0 = pdf.get_y() + 5
    if y0 + 45 > pdf.page_break_trigger:
        pdf.add_page()
        y0 = pdf.get_y() + 5
    pdf.set_draw_color(*LGRAY)
    for i, txt in enumerate((left_txt, right_txt)):
        cx = x + i * (col_w + 25)
        cy = y0
        for line in [l for l in txt.split("\n") if l.strip()]:
            pdf.set_xy(cx, cy)
            pdf.set_font("NotoGuj", "B", 9.5)
            pdf.set_text_color(*DARK)
            pdf.multi_cell(w=col_w, h=5, text=line, align="L")
            cy = pdf.get_y() + 1.5
        pdf.line(cx, cy + 11, cx + col_w * 0.72, cy + 11)
    pdf.set_y(y0 + 34)


def draw_glossary(pdf, rows, x):
    draw_kv_table(pdf, [(r[0]["text"], r[1]["text"]) for r in rows], x, key_w=42)
    pdf.set_y(pdf.get_y() + 2)


HEADER_COLS = ["ક્રમ", "કામનું વર્ણન", "એકમ", "જથ્થો", "સામગ્રી સહિતનો દર (₹)", "કુલ"]

# કિંમત ટેબલના સ્તંભની પહોળાઈ (mm માં) – કુલ 190 mm
PRICE_WIDTHS = [12, 74, 26, 16, 40, 22]


# ==========================================================================
def build_pdf():
    document = Document(DOCX)
    blocks = collect_blocks(document)

    reg, bold = prepare_fonts()
    pdf = QuotePDF(orientation="P", unit="mm", format="A4")
    pdf.set_margins(MARGIN, 15, MARGIN)
    pdf.set_auto_page_break(True, margin=18)
    pdf.add_font("NotoGuj", "", reg)
    pdf.add_font("NotoGuj", "B", bold)
    pdf.set_text_shaping(True)
    pdf.add_page()

    for kind, payload in blocks:
        if kind == "p":
            par = payload
            text = par.text.strip()
            if para_page_break(par):
                pdf.add_page()
                continue
            if not text:
                pdf.set_y(pdf.get_y() + 1.6)
                continue
            pkind, info = para_kind(par)
            draw_paragraph(pdf, pkind, info, text)
        else:
            tidx, rows, grid = payload
            x = MARGIN
            total = sum(grid)
            widths = [g / total * USABLE for g in grid]
            if tidx == 1:  # કિંમત ટેબલ
                widths = PRICE_WIDTHS
                pdf.set_y(pdf.get_y() + 1.5)
                draw_table_header(pdf, [r["text"] for r in rows[0]], widths, x)
                for r in rows[1:]:
                    if len(r) == 3 and r[0]["span"] >= 3:
                        title = r[0]["text"].strip()
                        if title:  # વિભાગ-શીર્ષક (A. / B. / C. ...)
                            draw_pricing_row(pdf, [dict(text=title, span=6)], widths, x, banded=True)
                        else:      # 'અંદાજિત કુલ' ની લાઈન
                            total_txt = next((c["text"].strip() for c in r if c["text"].strip()), "")
                            draw_pricing_row(
                                pdf,
                                [dict(text="", span=4),
                                 dict(text=total_txt, span=1, align="R"),
                                 dict(text="", span=1)],
                                widths, x, banded=True,
                            )
                    else:
                        draw_pricing_row(pdf, r, widths, x)
                pdf.set_y(pdf.get_y() + 2)
            elif tidx == 0:  # સાઇટ માહિતી
                left = [t for t in rows[0][0]["text"].split("\n") if t.strip()]
                right = [t for t in rows[0][-1]["text"].split("\n") if t.strip()]
                y0 = pdf.get_y() + 1
                for i in range(max(len(left), len(right))):
                    for j, col in enumerate((left, right)):
                        if i < len(col):
                            pdf.set_xy(x + j * (USABLE / 2), y0 + i * 5)
                            pdf.set_font("NotoGuj", "B", 9.5)
                            pdf.set_text_color(*DARK)
                            pdf.multi_cell(w=USABLE / 2 - 4, h=5, text=col[i])
                pdf.set_y(y0 + max(len(left), len(right)) * 5 + 2)
            elif tidx == 3:  # પ્રમાણપત્ર માહિતી
                draw_kv_table(pdf, [(r[0]["text"], r[1]["text"]) for r in rows], x, key_w=62)
                pdf.set_y(pdf.get_y() + 3)
            elif tidx == 5:  # શબ્દાવલી
                draw_glossary(pdf, rows, x)
            else:  # સહી બ્લોક
                draw_signature_block(pdf, rows, x)

    pdf.output(PDF_OUT)
    print("બનાવેલી PDF:", os.path.basename(PDF_OUT), "| પાનાં:", pdf.page_no())


if __name__ == "__main__":
    build_pdf()

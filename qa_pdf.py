#!/usr/bin/env python3
"""
QA – ગુજરાતી PDF ના લખાણની ચકાસણી
==================================
દરેક paragraph ને અલગ પાના પર એ જ પહોળાઈમાં દોરે છે અને PDF માંથી પાછું લખાણ
કાઢીને મૂળ લખાણ સાથે સરખાવે છે.  ખૂટતા/બદલાયેલા અક્ષરો (shaping કે font
fallback ની ભૂલ) તરત પકડાય છે.

    python qa_pdf.py            # merged-font mode
    python qa_pdf.py fallback   # fpdf2 fallback-font mode
"""

import os
import re
import sys

import pymupdf
from docx import Document
from fpdf import FPDF

import make_gujarati_pdf as gen

HERE = os.path.dirname(os.path.abspath(__file__))
MODE = sys.argv[1] if len(sys.argv) > 1 else "merged"
TEST_PDF = "/tmp/qa_%s.pdf" % MODE


def build_test_pdf():
    document = Document(gen.DOCX)
    pdf = gen.QuotePDF(orientation="P", unit="mm", format="A4")
    pdf.plain = True
    pdf.set_margins(gen.MARGIN, 15, gen.MARGIN)
    pdf.set_auto_page_break(False)
    reg, bold = gen.prepare_fonts()
    pdf.add_font("NotoGuj", "", reg)
    pdf.add_font("NotoGuj", "B", bold)
    if MODE == "fallback":
        pdf.add_font("DejaVu", "", gen.LAT_REG)
        pdf.add_font("DejaVu", "B", gen.LAT_BOLD)
        pdf.set_fallback_fonts(["DejaVu"])
    pdf.set_text_shaping(True)

    items = []
    for kind, payload in gen.collect_blocks(document):
        if kind != "p":
            continue
        par = payload
        text = par.text.strip()
        if not text or gen.para_page_break(par):
            continue
        pkind, info = gen.para_kind(par)
        items.append((pkind, info, text))

    for pkind, info, text in items:
        pdf.add_page()
        gen.draw_paragraph(pdf, pkind, info, text)
    pdf.output(TEST_PDF)
    return items


def normalize(s):
    s = s.replace("\u00ad", "")
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def gujarati_skeleton(s):
    """ફક્ત ગુજરાતી અક્ષરો અને અંક – PDF માંથી કાઢેલા લખાણમાં Latin/ચિહ્નોનું
    બદલાવું એ extraction ની મર્યાદા છે, પણ ગુજરાતી અક્ષર ખોટા પડે તો એ વાસ્તવિક
    રેન્ડરિંગ ખામી છે – એ જ અહીં ચકાસાય છે."""
    return "".join(re.findall(r"[\u0A80-\u0AFF0-9]+", s))


def main():
    items = build_test_pdf()
    doc = pymupdf.open(TEST_PDF)
    assert doc.page_count == len(items), (doc.page_count, len(items))
    bad = []
    for i, (pkind, info, text) in enumerate(items):
        got = gujarati_skeleton(doc[i].get_text())
        want = gujarati_skeleton(normalize(text))
        if got != want:
            bad.append((i, pkind, want, got))
    print("મોડ: %s" % MODE)
    print("કુલ paragraphs ચકાસાયા: %d | રેન્ડરિંગ ખામીવાળા: %d" % (len(items), len(bad)))
    for i, pkind, want, got in bad[:15]:
        k = 0
        while k < min(len(want), len(got)) and want[k] == got[k]:
            k += 1
        print("-" * 70)
        print("પાનું %d [%s]" % (i + 1, pkind))
        print("  જોઈએ : ...%s" % want[max(0, k - 25) : k + 35])
        print("  મળ્યું : ...%s" % got[max(0, k - 25) : k + 35])
    return bad


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""
માતૃશ્રી બંગ્લોઝ કોટેશન  →  સંપૂર્ણ ગુજરાતી આવૃત્તિ
=================================================
In   : 'GE EST માતૃશ્રી બંગ્લોઝ.docx'   (English quotation, original)
Out  : 'GE EST માતૃશ્રી બંગ્લોઝ - ગુજરાતી.docx'

* અનુવાદ માટેનો શબ્દકોશ તે જ ફાઈલમાંથી લેવામાં આવે છે જે પહેલાં બનાવેલી છે
  ('GE EST - Bilingual English-Gujarati.docx'), એટલે અહીંનું દરેક ગુજરાતી વાક્ય
  બરાબર એ જ છે જે ધંધાદારી રીતે પહેલેથી મંજૂર થયેલું છે.
* લેઆઉટ, ફોન્ટ, રંગ, ટેબલ, બોર્ડર, પેજ-બ્રેક – બધું જૂનું જ રહે છે; ફક્ત લખાણ બદલાય છે.
* ગુજરાતી લખાણના runs પર Word નો complex-script ફોન્ટ (Nirmala UI) અને
  ભાષા gu-IN સેટ થાય છે, જેથી Word/Google Docs માં સાચું રેન્ડરિંગ થાય.
"""

import json
import os
import re
import shutil
import sys

from docx import Document
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "GE EST માતૃશ્રી બંગ્લોઝ.docx")
BILINGUAL = os.path.join(HERE, "GE EST - Bilingual English-Gujarati.docx")
DST = os.path.join(HERE, "GE EST માતૃશ્રી બંગ્લોઝ - ગુજરાતી.docx")
REPORT = os.path.join(HERE, "translation-report.txt")

CS_FONT = "Nirmala UI"
LANG = "gu-IN"

# જે શબ્દો ગુજરાતીમાં પણ અંગ્રેજીમાં જ રહેવા જોઈએ (ટેકનિકલ સંક્ષેપ / કાનૂની નામ)
KEEP_LATIN = [
    "MDB", "SDB", "MCB", "RCCB", "ELCB", "SPD", "CAT 6", "CCTV", "DVR", "LED",
    "ISI", "NDA", "AMC", "PPE", "GST", "EV", "DTH", "UPI", "NEFT", "RTGS", "IMPS",
    "IS / IEC", "IEC", "Indian Electricity Act", "Indian Contract Act",
    "Arbitration and Conciliation Act", "as-built", "WhatsApp", "INR", "QR",
    "HARDIK HIRPARA", "GE-", "Email", "Mobile", "gmail.com", "ghanshyamelectricals101",
]


# ==========================================================================
# 1) XML સ્તરે દસ્તાવેજમાં ફરવાનું (દરેક paragraph / cell બરાબર એક વાર)
# ==========================================================================
def iter_body_children(elm):
    """w:p અને w:tbl – ક્રમમાં."""
    for child in elm.iterchildren():
        if child.tag == qn("w:p"):
            yield ("p", child)
        elif child.tag == qn("w:tbl"):
            yield ("tbl", child)


def iter_paragraphs(root_elm, part):
    """બોડી/ટેબલ/સેલમાંના તમામ paragraphs – એક પણ વાર પુનરાવર્તન વગર."""
    for kind, elm in iter_body_children(root_elm):
        if kind == "p":
            yield Paragraph(elm, part)
        else:
            for tr in elm.iterchildren(qn("w:tr")):
                for tc in tr.iterchildren(qn("w:tc")):
                    yield from iter_paragraphs(tc, part)


def cell_texts(root_elm, part):
    """દરેક table-cell નું લખાણ (paragraph વચ્ચે newline) – શબ્દકોશ બનાવવા માટે."""
    for kind, elm in iter_body_children(root_elm):
        if kind == "tbl":
            for tr in elm.iterchildren(qn("w:tr")):
                for tc in tr.iterchildren(qn("w:tc")):
                    paras = list(iter_paragraphs(tc, part))
                    if paras:
                        yield "\n".join(p.text for p in paras)
                    yield from cell_texts(tc, part)


def document_texts(document):
    for p in iter_paragraphs(document.element.body, document):
        yield p.text
    yield from cell_texts(document.element.body, document)


# ==========================================================================
# 2) બે-ભાષી ફાઈલમાંથી શબ્દકોશ
# ==========================================================================
def build_dictionary():
    doc = Document(BILINGUAL)
    table = {}
    for text in document_texts(doc):
        if "\n" not in text:
            continue
        parts = text.split("\n")
        en = parts[0].strip()
        gu = "\n".join(parts[1:]).strip()
        if en and gu and en not in table:
            table[en] = gu
    # ટેબલ કોષોમાં ક્યાંક અંગ્રેજી+ગુજરાતી એક જ લાઈનમાં ' / ' થી છૂટા પડેલા હોય તે પણ લેવા
    for text in document_texts(doc):
        for line in text.split("\n"):
            if " / " in line and re.search(r"[\u0A80-\u0AFF]", line):
                continue
    return table


# ==========================================================================
# 3) ખાસ કિસ્સા – એક જ paragraph માં જુદા ફોર્મેટિંગવાળા runs
# ==========================================================================
SPECIAL_RUNS = {
    "Mobile: +91 92655-60630   |   Email: ghanshyamelectricals101@gmail.com": [
        "મોબાઇલ: +91 92655-6063",
        "0",
        "   |   ઈમેઇલ: ghanshyamelectricals101@gmail.com",
    ],
    "Including Complete Terms, Conditions & Non-Disclosure Agreement": [
        "સંપૂર્ણ શરતો, નિયમો અને ગોપનીયતા કરાર",
        " સહિત",
    ],
    "Quotation No.: GE-_________": ["કોટેશન નં.: ", "GE-_________"],
    "Date: ____________________": ["તારીખ: ", "____________________"],
    "Valid Until: 10 Days from Date": ["માન્યતા: ", "તારીખથી 10 દિવસ સુધી"],
    "3.1 Work Process": ["3.1 ", "કામ કરવાની પ્રક્રિયા"],
    "Party A: Ghanshyam Electricals, 441-The Galleria, Yogi Chowk, Surat, Gujarat - 395006": [
        "પક્ષ A: ",
        "ઘનશ્યામ ઇલેક્ટ્રિકલ્સ, 441-ધ ગેલેરિયા, યોગી ચોક, સુરત, ગુજરાત - 395006",
    ],
    "Party B: The Client (as named in the Quotation)": [
        "પક્ષ B: ",
        "ક્લાયન્ટ (કોટેશનમાં દર્શાવ્યા મુજબ)",
    ],
    "Warranty claims shall be registered via WhatsApp or phone call to +91 92655-60630.": [
        "વોરંટીના દાવા WhatsApp અથવા +91 92655-6063",
        "0",
        " પર ફોન કરીને નોંધાવવાના રહેશે.",
    ],
}

TOC_TERMS = {
    "Scope of Work": "કામનો વ્યાપ",
    "Pricing": "કિંમત",
    "Terms & Conditions": "શરતો અને નિયમો",
    "Payment Conditions": "ચુકવણીની શરતો",
    "Warranty & After Service": "વોરંટી અને વેચાણ પછીની સેવા",
    "Limitation of Liability": "જવાબદારીની મર્યાદા",
    "Health & Safety": "આરોગ્ય અને સલામતી",
    "Waste & Environment": "કચરો અને પર્યાવરણ",
    "Quotation Acceptance": "કોટેશનની સ્વીકૃતિ",
    "Project Completion Certificate": "પ્રોજેક્ટ પૂર્ણતા પ્રમાણપત્ર",
    "Glossary of Terms": "શબ્દાવલી",
    "Non-Disclosure Agreement": "ગોપનીયતા કરાર",
}


# ==========================================================================
# 4) અનુવાદ
# ==========================================================================
def set_cs_font(run):
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = rPr.makeelement(qn("w:rFonts"), {})
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:cs"), CS_FONT)
    lang = rPr.find(qn("w:lang"))
    if lang is None:
        lang = rPr.makeelement(qn("w:lang"), {})
        rPr.append(lang)
    lang.set(qn("w:bidi"), LANG)


def rewrite_paragraph(par, dictionary):
    src = par.text
    key = src.strip()
    if not key:
        return False

    non_empty = [r for r in par.runs if r.text.strip()]
    if not non_empty:
        return False

    # વિષયસૂચિ : '1.  Scope of Work' → '1.  કામનો વ્યાપ'
    m = re.match(r"^(\d{1,2}\.\s+)(.+)$", key)
    if m and m.group(2).strip() in TOC_TERMS and len(non_empty) > 1:
        non_empty[0].text = m.group(1)
        set_cs_font(non_empty[0])
        non_empty[1].text = TOC_TERMS[m.group(2).strip()]
        set_cs_font(non_empty[1])
        for r in non_empty[2:]:
            r.text = ""
        return True

    if key in SPECIAL_RUNS:
        replacement = SPECIAL_RUNS[key]
        if len(replacement) == len(non_empty):
            for run, txt in zip(non_empty, replacement):
                run.text = txt
                set_cs_font(run)
            return True
        non_empty[0].text = dictionary.get(key, " ".join(replacement))
        set_cs_font(non_empty[0])
        for r in non_empty[1:]:
            r.text = ""
        return True

    translation = dictionary.get(key)
    if translation is None:
        return False

    non_empty[0].text = translation
    set_cs_font(non_empty[0])
    for r in non_empty[1:]:
        r.text = ""
    return True


def translate(document, dictionary):
    changed, untranslated = 0, []
    for par in iter_paragraphs(document.element.body, document):
        if rewrite_paragraph(par, dictionary):
            changed += 1
        elif par.text.strip():
            untranslated.append(par.text.strip())
    return changed, untranslated


# ==========================================================================
# 5) ચકાસણી (QA)
# ==========================================================================
def qa(document):
    problems = []
    for par in iter_paragraphs(document.element.body, document):
        text = par.text.strip()
        if not text:
            continue
        cleaned = text
        for allowed in KEEP_LATIN:
            cleaned = cleaned.replace(allowed, " ")
        leftover = re.findall(r"[A-Za-z]{2,}", cleaned)
        if leftover:
            problems.append((text, leftover))
    return problems


def polish(document):
    """Word માં ગુજરાતીને ગુજરાતી તરીકે ઓળખવા માટે દસ્તાવેજની ડિફોલ્ટ ભાષા સેટ કરો."""
    styles = document.styles.element
    doc_defaults = styles.find(qn("w:docDefaults"))
    if doc_defaults is not None:
        rpr_default = doc_defaults.find(qn("w:rPrDefault"))
        if rpr_default is not None:
            rpr = rpr_default.find(qn("w:rPr"))
            if rpr is None:
                rpr = rpr_default.makeelement(qn("w:rPr"), {})
                rpr_default.append(rpr)
            lang = rpr.find(qn("w:lang"))
            if lang is None:
                lang = rpr.makeelement(qn("w:lang"), {})
                rpr.append(lang)
            lang.set(qn("w:val"), "en-IN")
            lang.set(qn("w:bidi"), LANG)

    props = document.core_properties
    props.title = "ઇલેક્ટ્રિકલ કામનું કોટેશન – ઘનશ્યામ ઇલેક્ટ્રિકલ્સ (માતૃશ્રી બંગ્લોઝ)"
    props.subject = "ઇલેક્ટ્રિકલ કામનું કોટેશન – ગુજરાતી"
    props.language = LANG
    props.comments = "મૂળ અંગ્રેજી કોટેશનનો સંપૂર્ણ ગુજરાતી અનુવાદ."


def main():
    dictionary = build_dictionary()
    dictionary.setdefault("GE-", "GE-")

    shutil.copyfile(SRC, DST)
    document = Document(DST)
    changed, untranslated = translate(document, dictionary)
    polish(document)
    document.save(DST)

    problems = qa(Document(DST))

    lines = []
    lines.append("માતૃશ્રી બંગ્લોઝ કોટેશન – ગુજરાતી અનુવાદ રિપોર્ટ")
    lines.append("=" * 46)
    lines.append("સ્રોત ફાઈલ : %s" % os.path.basename(SRC))
    lines.append("નવી ફાઈલ  : %s" % os.path.basename(DST))
    lines.append("શબ્દકોશ    : %d વાક્ય/શબ્દસમૂહ" % len(dictionary))
    lines.append("ગુજરાતીમાં અનુવાદિત paragraphs : %d" % changed)
    lines.append("અનુવાદિત ન થયેલા (સંખ્યા/સંક્ષેપ) : %d" % len(untranslated))
    for t in sorted(set(untranslated)):
        lines.append("    – %s" % t)
    lines.append("")
    lines.append("ચકાસણી (QA) – ગુજરાતીમાં ન થયેલા અંગ્રેજી શબ્દો: %d" % len(problems))
    for text, words in problems:
        lines.append("    – %s   → %s" % (words, text[:90]))
    report = "\n".join(lines)
    open(REPORT, "w", encoding="utf-8").write(report)
    print(report)

    # સાથે PDF પણ બનાવો (જો fpdf2 / fontTools ઉપલબ્ધ હોય તો)
    try:
        import make_gujarati_pdf

        make_gujarati_pdf.build_pdf()
    except Exception as exc:  # pragma: no cover
        print("PDF બનાવવામાં અડચણ:", exc)


if __name__ == "__main__":
    main()

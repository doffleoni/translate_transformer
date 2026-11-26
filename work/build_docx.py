# -*- coding: utf-8 -*-
"""Сборка русского перевода статьи «Attention Is All You Need» в DOCX."""
import re
import os
import sys

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_LINE_SPACING
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from content1 import BLOCKS          # noqa: E402
from content2 import BLOCKS2         # noqa: E402
from content3 import BLOCKS3         # noqa: E402
from tables import TABLES            # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
FIG = os.path.join(HERE, 'fig')
OUT = os.path.abspath(os.path.join(HERE, '..', 'Attention_Is_All_You_Need_RU.docx'))

BODY_FONT = 'Liberation Serif'
MATH_FONT = 'DejaVu Serif'          # надёжный рендеринг ∈, ℝ, −∞
TEXT_WIDTH_IN = 6.5

TOKEN_RE = re.compile(r'\[\[(/?)(b|i|sup|sub|m|fn)(?::([^\]]*))?\]\]')


class FootnoteStore:
    """Сноски внизу страницы: ссылка в тексте + w:footnotes.xml."""

    CT = ('application/vnd.openxmlformats-officedocument'
          '.wordprocessingml.footnotes+xml')

    def __init__(self):
        self.items = []

    def add(self, text):
        self.items.append(text)
        return len(self.items)

    def reference_run(self, paragraph, number):
        run = paragraph.add_run()
        run.style = paragraph.part.document.styles['FootnoteReference']
        run._r.append(parse_xml(
            '<w:footnoteReference %s w:id="%d"/>' % (nsdecls('w'), number)))
        return run

    def attach(self, document):
        if not self.items:
            return
        head = (
            '<w:footnotes %s>'
            '<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr>'
            '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
            '<w:r><w:separator/></w:r></w:p></w:footnote>'
            '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr>'
            '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
            '<w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
        ) % nsdecls('w')
        body = []
        for i, text in enumerate(self.items, start=1):
            safe = (text.replace('&', '&amp;').replace('<', '&lt;')
                        .replace('>', '&gt;'))
            body.append(
                '<w:footnote w:id="%d"><w:p><w:pPr>'
                '<w:pStyle w:val="FootnoteText"/>'
                '<w:jc w:val="both"/></w:pPr>'
                '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>'
                '<w:footnoteRef/></w:r>'
                '<w:r><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s"/>'
                '</w:rPr><w:t xml:space="preserve"> %s</w:t></w:r>'
                '</w:p></w:footnote>' % (i, BODY_FONT, BODY_FONT, BODY_FONT, safe))
        blob = (head + ''.join(body) + '</w:footnotes>').encode('utf-8')
        part = Part(PackURI('/word/footnotes.xml'), self.CT, blob,
                    document.part.package)
        document.part.relate_to(part, RT.FOOTNOTES)


def set_run_font(run, name=BODY_FONT, size=None, bold=None, italic=None,
                 sup=False, sub=False, math=False):
    run.font.name = name
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.insert(0, rfonts)
    for attr in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
        rfonts.set(qn(attr), name)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    if sup:
        run.font.superscript = True
    if sub:
        run.font.subscript = True
    return run


def render(paragraph, text, footnotes=None, size=None, base_italic=False):
    """Разбирает разметку и добавляет в абзац соответствующие runs."""
    pos = 0
    bold = italic = sup = sub = math = False
    while True:
        m = TOKEN_RE.search(text, pos)
        chunk = text[pos:m.start()] if m else text[pos:]
        if chunk:
            run = paragraph.add_run(chunk)
            set_run_font(run, MATH_FONT if (math or base_italic) else BODY_FONT,
                         size=size, bold=bold, italic=italic or math or base_italic,
                         sup=sup, sub=sub)
        if not m:
            break
        closing, kind, arg = m.group(1), m.group(2), m.group(3)
        on = (closing != '/')
        if kind == 'b':
            bold = on
        elif kind == 'i':
            italic = on
        elif kind == 'sup':
            sup = on
        elif kind == 'sub':
            sub = on
        elif kind == 'm':
            math = on
        elif kind == 'fn' and on and footnotes is not None:
            number = footnotes.add(FOOTNOTES[arg])
            footnotes.reference_run(paragraph, number)
        pos = m.end()


def add_para(doc, text, style=None, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
             size=None, space_after=6, space_before=0, first_line=None,
             left_indent=None, hanging=None, footnotes=None, keep_with_next=False,
             line_spacing=1.12):
    p = doc.add_paragraph(style=style)
    pf = p.paragraph_format
    pf.alignment = align
    pf.space_after = Pt(space_after)
    pf.space_before = Pt(space_before)
    pf.line_spacing = line_spacing
    if first_line is not None:
        pf.first_line_indent = Inches(first_line)
    if left_indent is not None:
        pf.left_indent = Inches(left_indent)
    if hanging is not None:
        pf.first_line_indent = Inches(-hanging)
    pf.keep_with_next = keep_with_next
    render(p, text, footnotes=footnotes, size=size)
    return p


# ------------------------------------------------------------------ таблицы
def set_cell_border(cell, **kwargs):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = tcPr.find(qn('w:tcBorders'))
    if borders is None:
        borders = OxmlElement('w:tcBorders')
        tcPr.append(borders)
    for edge in ('top', 'left', 'bottom', 'right'):
        if edge in kwargs:
            tag = 'w:{}'.format(edge)
            el = borders.find(qn(tag))
            if el is None:
                el = OxmlElement(tag)
                borders.append(el)
            for key, val in kwargs[edge].items():
                el.set(qn('w:{}'.format(key)), str(val))


def booktabs_rule(cell, edge='bottom', sz=6):
    set_cell_border(cell, **{edge: {'val': 'single', 'sz': sz, 'space': 0,
                                    'color': '000000'}})


def clear_cell_borders(cell):
    for edge in ('top', 'left', 'bottom', 'right'):
        set_cell_border(cell, **{edge: {'val': 'nil'}})


def fill_cell(cell, text, size, align=WD_ALIGN_PARAGRAPH.CENTER, bold=False,
              footnotes=None):
    cell.text = ''
    p = cell.paragraphs[0]
    p.paragraph_format.alignment = align
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.line_spacing = 1.05
    render(p, text, footnotes=footnotes, size=size)
    if bold:
        for r in p.runs:
            r.font.bold = True


def add_table(doc, spec, footnotes):
    caption, header, rows, widths, size = (spec['caption'], spec['header'],
                                           spec['rows'], spec['widths'],
                                           spec.get('size', 8.5))
    add_para(doc, caption, align=WD_ALIGN_PARAGRAPH.JUSTIFY, size=8.5,
             space_after=4, space_before=8, keep_with_next=True)
    table = doc.add_table(rows=1 + len(rows), cols=len(header))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    tblPr = table._tbl.tblPr
    layout = OxmlElement('w:tblLayout')
    layout.set(qn('w:type'), 'fixed')
    tblPr.append(layout)

    for j, w in enumerate(widths):
        for i in range(len(table.rows)):
            table.cell(i, j).width = Inches(w)

    for j, h in enumerate(header):
        fill_cell(table.cell(0, j), h, size, bold=True)
    for i, row in enumerate(rows, start=1):
        for j, val in enumerate(row):
            align = (WD_ALIGN_PARAGRAPH.LEFT if j == 0
                     else WD_ALIGN_PARAGRAPH.CENTER)
            fill_cell(table.cell(i, j), val, size, align=align)

    for r_i, c_i, span in spec.get('merges', []):
        a = table.cell(r_i, c_i)
        b = table.cell(r_i, c_i + span - 1)
        merged = a.merge(b)
        merged.text = ''
        mp = merged.paragraphs[0]
        mp.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        mp.paragraph_format.space_after = Pt(1)
        mp.paragraph_format.space_before = Pt(1)
        render(mp, rows[r_i - 1][c_i], footnotes=footnotes, size=size)

    for i in range(len(table.rows)):
        for j in range(len(header)):
            clear_cell_borders(table.cell(i, j))
    for j in range(len(header)):
        booktabs_rule(table.cell(0, j), 'top', sz=10)
        booktabs_rule(table.cell(0, j), 'bottom', sz=5)
        booktabs_rule(table.cell(len(table.rows) - 1, j), 'bottom', sz=10)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(10)
    p.paragraph_format.space_before = Pt(0)
    set_run_font(p.add_run(''), size=2)
    return table


def add_figure(doc, filename, width_in, caption, size=8.5):
    p = doc.add_paragraph()
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(os.path.join(FIG, filename), width=Inches(width_in))
    cap = add_para(doc, caption, align=WD_ALIGN_PARAGRAPH.JUSTIFY, size=size,
                   space_after=10, space_before=0)
    return cap


def add_equation(doc, filename, width_in, number):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(6)
    pf.space_after = Pt(6)
    pf.line_spacing = 1.0
    pf.keep_together = True
    pf.keep_with_next = True
    tabs = pf.tab_stops
    tabs.add_tab_stop(Inches(TEXT_WIDTH_IN / 2.0), WD_TAB_ALIGNMENT.CENTER)
    tabs.add_tab_stop(Inches(TEXT_WIDTH_IN), WD_TAB_ALIGNMENT.RIGHT)
    set_run_font(p.add_run('\t'), size=10)
    p.add_run().add_picture(os.path.join(FIG, filename), width=Inches(width_in))
    set_run_font(p.add_run('\t' + ('(%s)' % number if number else '')), size=10)
    return p


def add_bullets(doc, items, footnotes):
    for item in items:
        p = add_para(doc, '\u2022\t' + item, size=10, space_after=4,
                     left_indent=0.45, hanging=0.22, footnotes=footnotes)
    return p


def add_refs(doc, refs):
    for i, ref in enumerate(refs, start=1):
        add_para(doc, '[%d]\t%s' % (i, ref), size=8.5, space_after=2,
                 left_indent=0.35, hanging=0.35,
                 align=WD_ALIGN_PARAGRAPH.JUSTIFY)


# --------------------------------------------------------------- документ
def build():
    doc = Document()

    normal = doc.styles['Normal']
    normal.font.name = BODY_FONT
    normal.font.size = Pt(10)
    normal._element.rPr.rFonts.set(qn('w:eastAsia'), BODY_FONT)

    for name, size, bold, italic, before, after in [
        ('Heading 1', 13, True, False, 14, 6),
        ('Heading 2', 11.5, True, False, 11, 4),
        ('Heading 3', 10.5, True, True, 9, 3),
    ]:
        st = doc.styles[name]
        st.font.name = BODY_FONT
        st.font.size = Pt(size)
        st.font.bold = bold
        st.font.italic = italic
        st.font.color.rgb = RGBColor(0, 0, 0)
        st.paragraph_format.space_before = Pt(before)
        st.paragraph_format.space_after = Pt(after)
        st.paragraph_format.keep_with_next = True
        st.paragraph_format.line_spacing = 1.0

    for name, size in [('FootnoteText', 8)]:
        st = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        st.font.name = BODY_FONT
        st.font.size = Pt(size)
        st.paragraph_format.space_after = Pt(0)
        st.paragraph_format.line_spacing = 1.0
    ref = doc.styles.add_style('FootnoteReference', WD_STYLE_TYPE.CHARACTER)
    ref.font.superscript = True
    ref.font.size = Pt(7)

    sec = doc.sections[0]
    # сохранить исходную нумерацию сносок оригинала (4 и 5)
    sectPr = sec._sectPr
    fpr = OxmlElement('w:footnotePr')
    num_start = OxmlElement('w:numStart')
    num_start.set(qn('w:val'), '4')
    fpr.append(num_start)
    sectPr.insert(0, fpr)
    sec.page_width = Inches(8.5)
    sec.page_height = Inches(11)
    for attr in ('left_margin', 'right_margin'):
        setattr(sec, attr, Inches(1.0))
    sec.top_margin = Inches(0.9)
    sec.bottom_margin = Inches(0.9)

    # номер страницы в колонтитуле
    footer_p = sec.footer.paragraphs[0]
    footer_p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer_p.add_run()
    set_run_font(run, size=9)
    fld = OxmlElement('w:fldSimple')
    fld.set(qn('w:instr'), 'PAGE')
    run._r.addnext(fld)

    footnotes = FootnoteStore()

    seq = BLOCKS + BLOCKS2 + BLOCKS3
    for idx, block in enumerate(seq):
        kind = block[0]
        nxt = seq[idx + 1][0] if idx + 1 < len(seq) else None
        if kind == 'title':
            add_para(doc, block[1], align=WD_ALIGN_PARAGRAPH.CENTER, size=20,
                     space_after=2, space_before=0)
            p = add_para(doc, 'Attention Is All You Need',
                         align=WD_ALIGN_PARAGRAPH.CENTER, size=13, space_after=2)
            for r in p.runs:
                r.font.italic = True
        elif kind == 'title_notice':
            add_para(doc, block[1], align=WD_ALIGN_PARAGRAPH.CENTER, size=8.5,
                     space_after=14)
        elif kind == 'title_sub':
            pass
        elif kind == 'authors':
            for name, org, mail in block[1]:
                line = '[[b]]%s[[/b]]' % name
                p = add_para(doc, line, align=WD_ALIGN_PARAGRAPH.CENTER,
                             size=10.5, space_after=0)
                for r in p.runs:
                    r.font.bold = True
                add_para(doc, org, align=WD_ALIGN_PARAGRAPH.CENTER, size=9,
                         space_after=0)
                if mail:
                    add_para(doc, mail, align=WD_ALIGN_PARAGRAPH.CENTER, size=8.5,
                             space_after=10)
                else:
                    doc.paragraphs[-1].paragraph_format.space_after = Pt(10)
        elif kind == 'title_note':
            add_para(doc, block[1], size=8, space_after=4)
        elif kind == 'title_conf':
            add_para(doc, block[1], align=WD_ALIGN_PARAGRAPH.CENTER, size=8.5,
                     space_after=6, space_before=8)
        elif kind == 'pagebreak':
            doc.add_page_break()
        elif kind in ('h1', 'h1_notoc'):
            doc.add_paragraph(block[1], style='Heading 1')
        elif kind == 'h2':
            doc.add_paragraph(block[1], style='Heading 2')
        elif kind == 'h3':
            doc.add_paragraph(block[1], style='Heading 3')
        elif kind == 'p':
            p = add_para(doc, block[1], footnotes=footnotes, size=10)
            if nxt in ('eq', 'fig', 'tab1', 'tab2', 'tab3', 'tab4'):
                p.paragraph_format.keep_with_next = True
        elif kind == 'p_abstract':
            add_para(doc, block[1], footnotes=footnotes, size=9.5,
                     left_indent=0.4, space_after=10)
        elif kind == 'bullets':
            add_bullets(doc, block[1], footnotes)
        elif kind == 'fig':
            add_figure(doc, block[1], block[2], block[3])
        elif kind == 'eq':
            add_equation(doc, block[1], block[2], block[3])
        elif kind in ('tab1', 'tab2', 'tab3', 'tab4'):
            add_table(doc, TABLES[kind], footnotes)
        elif kind == 'refs':
            add_refs(doc, block[1])
        else:
            raise ValueError('unknown block: %r' % (kind,))

    footnotes.attach(doc)
    doc.save(OUT)
    print('saved', OUT)
    print('paragraphs:', len(doc.paragraphs), 'footnotes:', len(footnotes.items))


FOOTNOTES = {
    '4': 'Чтобы пояснить, почему скалярные произведения становятся большими, '
         'предположим, что компоненты q и k — независимые случайные величины со '
         'средним 0 и дисперсией 1. Тогда их скалярное произведение '
         'q · k = Σ (i=1..d_k) q_i k_i имеет среднее 0 и дисперсию d_k.',
    '5': 'Мы использовали значения 2,8; 3,7; 6,0 и 9,5 TFLOPS для K80, K40, M40 и '
         'P100 соответственно.',
}

if __name__ == '__main__':
    build()

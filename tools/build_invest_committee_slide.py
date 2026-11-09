"""Генерация invest_committee_slide.pptx (python-pptx): проект развития площадки АО «Изотоп»."""
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt
from lxml import etree

OUT = Path(__file__).resolve().parents[1] / "invest_committee_slide.pptx"

NAVY = RGBColor(0x0B, 0x2A, 0x4A)
BLUE = RGBColor(0x00, 0x5B, 0xAA)
MID = RGBColor(0x2E, 0x7D, 0xC4)
LIGHT = RGBColor(0xEA, 0xF0, 0xF7)
LINE = RGBColor(0xC5, 0xD3, 0xE3)
GREY = RGBColor(0x55, 0x60, 0x6E)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
FONT = "Arial"

SLIDE_W, SLIDE_H = Inches(13.333), Inches(7.5)
PROJECT = "Проект развития площадки АО «Изотоп»"


def _font(run, size, bold=False, color=NAVY, italic=False):
    f = run.font
    f.name = FONT
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.color.rgb = color


def _bullet(paragraph, char="•", indent=Pt(12)):
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set("marL", str(int(indent)))
    pPr.set("indent", str(-int(indent)))
    for tag in ("a:buNone", "a:buChar", "a:buFont"):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)
    bu_font = etree.SubElement(pPr, qn("a:buFont"))
    bu_font.set("typeface", FONT)
    bu = etree.SubElement(pPr, qn("a:buChar"))
    bu.set("char", char)


def text(slide, x, y, w, h, paras, *, anchor=MSO_ANCHOR.TOP, margin=0.05, autofit=False):
    """paras: list of dicts {runs:[(text, size, bold, color)], align, bullet, space_after}."""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    m = Inches(margin)
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = m
    fill_paras(tf, paras)
    return tb


def fill_paras(tf, paras):
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = p.get("align", PP_ALIGN.LEFT)
        if "space_after" in p:
            para.space_after = Pt(p["space_after"])
        if "space_before" in p:
            para.space_before = Pt(p["space_before"])
        para.line_spacing = p.get("line_spacing", 1.0)
        if p.get("bullet"):
            _bullet(para, p.get("bullet_char", "•"), Pt(p.get("bullet_indent", 12)))
        for r in p["runs"]:
            t, size, bold, color = (list(r) + [False, NAVY])[:4] if len(r) < 4 else r
            run = para.add_run()
            run.text = t
            _font(run, size, bold, color)


def rect(slide, x, y, w, h, fill, line=None, shape=MSO_SHAPE.RECTANGLE, radius=None):
    s = slide.shapes.add_shape(shape, x, y, w, h)
    s.shadow.inherit = False
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(1)
    if radius is not None:
        s.adjustments[0] = radius
    return s


def p(runs, **kw):
    d = {"runs": runs}
    d.update(kw)
    return d


def header(slide, title, subtitle=None, num=None, title_size=26, title_h=Inches(0.6)):
    rect(slide, 0, 0, SLIDE_W, Inches(0.12), NAVY)
    rect(slide, 0, Inches(0.12), Inches(2.2), Inches(0.05), MID)
    text(slide, Inches(10.6), Inches(0.3), Inches(2.35), Inches(0.5),
         [p([("РОСАТОМ", 20, True, NAVY)], align=PP_ALIGN.RIGHT)], anchor=MSO_ANCHOR.MIDDLE)
    text(slide, Inches(10.6), Inches(0.75), Inches(2.35), Inches(0.3),
         [p([("АО «Изотоп»", 11, False, GREY)], align=PP_ALIGN.RIGHT)])
    text(slide, Inches(0.45), Inches(0.28), Inches(10.1), title_h,
         [p([(title, title_size, True, NAVY)], line_spacing=0.95)], anchor=MSO_ANCHOR.TOP)
    if subtitle:
        text(slide, Inches(0.45), Inches(0.3) + title_h, Inches(10.1), Inches(0.35),
             [p([(subtitle, 13, False, BLUE)])])
    # footer
    rect(slide, Inches(0.45), Inches(7.05), Inches(12.43), Emu(9525), LINE)
    text(slide, Inches(0.45), Inches(7.08), Inches(10), Inches(0.3),
         [p([("Госкорпорация «Росатом»  ·  " + PROJECT, 9, False, GREY)])])
    if num is not None:
        text(slide, Inches(12.2), Inches(7.08), Inches(0.68), Inches(0.3),
             [p([(str(num), 9, True, NAVY)], align=PP_ALIGN.RIGHT)])


def section_label(slide, x, y, w, label):
    text(slide, x, y, w, Inches(0.35), [p([(label.upper(), 12, True, BLUE)])], anchor=MSO_ANCHOR.MIDDLE)
    rect(slide, x, y + Inches(0.36), w, Pt(1.5), BLUE)


# ---------------------------------------------------------------- slide
def card(slide, x, y, w, h, title, bullets, size=11):
    rect(slide, x, y, w, h, WHITE, LINE)
    rect(slide, x, y, w, Inches(0.42), NAVY)
    text(slide, x + Inches(0.15), y, w - Inches(0.3), Inches(0.42),
         [p([(title, 13, True, WHITE)])], anchor=MSO_ANCHOR.MIDDLE)
    text(slide, x + Inches(0.12), y + Inches(0.5), w - Inches(0.24), h - Inches(0.55),
         [p([(b, size, False, NAVY)], bullet=True, space_after=3) for b in bullets])


def flow_box(slide, x, y, w, h, kicker, title, body, dark=False, footer=None):
    fill, fg, sub = (NAVY, WHITE, RGBColor(0x8F, 0xB8, 0xE6)) if dark else (LIGHT, NAVY, BLUE)
    rect(slide, x, y, w, h, fill, None if dark else LINE)
    paras = [
        p([(kicker.upper(), 9.5, True, sub)], space_after=1),
        p([(title, 14, True, fg)], space_after=3),
        p([(body, 11, False, fg)], line_spacing=1.0),
    ]
    if footer:
        paras.append(p([(footer, 9, True, sub)], space_before=4))
    text(slide, x + Inches(0.15), y + Inches(0.05), w - Inches(0.3), h - Inches(0.1), paras,
         anchor=MSO_ANCHOR.MIDDLE)


def slide_development(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header(
        s,
        "Проект развития: площадка АО «Изотоп» — производство и рынок "
        "для Научного дивизиона Росатома",
        "Концепция. Количественные параметры определяются на этапе ТЭО",
        title_size=22, title_h=Inches(0.8),
    )
    L, W = Inches(0.45), Inches(12.43)

    # идея проекта
    iy, ih = Inches(1.42), Inches(0.6)
    rect(s, L, iy, W, ih, LIGHT)
    rect(s, L, iy, Inches(0.09), ih, BLUE)
    text(s, L + Inches(0.25), iy, W - Inches(0.4), ih, [
        p([("Идея: ", 12, True, BLUE),
           ("разместить на площадке АО «Изотоп» производства, дополняющие предприятия Научного "
            "дивизиона: институты концентрируются на НИОКР и наработке радионуклидов, "
            "«Изотоп» снимает с них серийную и непрофильную нагрузку и берёт на себя рыночную функцию.",
            12, False, NAVY)]),
    ], anchor=MSO_ANCHOR.MIDDLE)

    # модель разделения ролей
    fy, fh, aw = Inches(2.13), Inches(1.15), Inches(0.45)
    w1, w2 = Inches(4.35), Inches(4.1)
    w3 = W - w1 - w2 - 2 * aw
    x1 = L
    x2 = x1 + w1 + aw
    x3 = x2 + w2 + aw
    flow_box(s, x1, fy, w1, fh, "Научный дивизион Росатома", "Наука и сырьё",
             "НИОКР, наработка радионуклидов на реакторах и циклотронах, разработка технологий",
             footer="НИФХИ · НИИТФА · ФЭИ · ИРМ · НИИАР · Радиевый институт")
    flow_box(s, x2, fy, w2, fh, "Площадка АО «Изотоп»", "Производство и рынок",
             "Серийный выпуск, качество, логистика, регистрация, продажи и сервис",
             dark=True,
             footer="GMP · ISO 13485 · GDP · GLP · ОСПОРБ / НРБ")
    flow_box(s, x3, fy, w3, fh, "Потребители", "Рынок",
             "Госзаказ Минздрава, онкоцентры и центры ядерной медицины, экспорт")
    for ax in (x1 + w1, x2 + w2):
        rect(s, ax + Inches(0.06), fy + fh // 2 - Inches(0.2), aw - Inches(0.12), Inches(0.4),
             MID, shape=MSO_SHAPE.RIGHT_ARROW)
    # обратная связь
    by = fy + fh + Inches(0.07)
    fb = rect(s, x1 + Inches(0.6), by, x3 + w3 - x1 - Inches(1.2), Inches(0.32),
              RGBColor(0xD6, 0xE3, 0xF2), shape=MSO_SHAPE.LEFT_ARROW)
    fb.adjustments[0] = 0.62
    fb.adjustments[1] = 0.9
    text(s, x1 + Inches(1.2), by, x3 + w3 - x1 - Inches(2.4), Inches(0.32), [
        p([("Обратная связь: рыночный запрос и real-world evidence → техзадания на НИОКР "
            "(цикл PDSA: Plan — Do — Study — Act)", 10, True, NAVY)], align=PP_ALIGN.CENTER),
    ], anchor=MSO_ANCHOR.MIDDLE, margin=0)

    # три колонки
    cy, ch, gap = Inches(3.72), Inches(2.38), Inches(0.22)
    cw = int((W - 2 * gap) / 3)
    card(s, L, cy, cw, ch, "Предпосылки", [
        "Институты загружены серийным выпуском и непрофильными задачами",
        "Научная инфраструктура не рассчитана на GMP-серию и рыночные объёмы",
        "Нет единого рыночного окна: продукт, регистрация, логистика, продажи",
        "Растущий спрос на РФЛП и задача технологического суверенитета",
    ])
    card(s, L + cw + gap, cy, cw, ch, "Что размещается на площадке", [
        "GMP-производство РФЛП по технологиям институтов",
        "Выпуск нерадиоактивных частей и наборов",
        "Упаковочные комплекты типа А для РФЛП",
        "Спецтранспорт (РВ и темп. режим), авто/авиа доставка",
        "Контроль качества, регистрация, фармаконадзор",
    ])
    card(s, L + 2 * (cw + gap), cy, cw, ch, "Эффекты", [
        "Институты: снятие доп. нагрузки, фокус на науке и наработке изотопов",
        "Росатом: единое рыночное окно и быстрый вывод продуктов на рынок",
        "Государство и пациенты: суверенитет и доступность ядерной медицины",
        "Экспорт через канал зарубежных проектов Росатома",
    ])

    # предлагаемое решение
    dy, dh = Inches(6.2), Inches(0.75)
    rect(s, L, dy, W, dh, NAVY)
    rect(s, L, dy, Inches(0.12), dh, MID)
    text(s, L + Inches(0.3), dy, Inches(2.3), dh, [
        p([("ПРЕДЛАГАЕМОЕ", 11, True, RGBColor(0x8F, 0xB8, 0xE6))]),
        p([("РЕШЕНИЕ", 11, True, RGBColor(0x8F, 0xB8, 0xE6))]),
    ], anchor=MSO_ANCHOR.MIDDLE)
    text(s, L + Inches(2.35), dy, W - Inches(2.5), dh, [
        p([("Одобрить концепцию и поручить: ", 12, True, WHITE),
           ("1) сформировать с институтами перечень производств для размещения на площадке; "
            "2) проработать модель взаимодействия и передачи технологий; "
            "3) подготовить ТЭО и финансовую модель", 12, False, WHITE)]),
    ], anchor=MSO_ANCHOR.MIDDLE)
    return s


def main():
    prs = Presentation()
    prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
    prs.core_properties.title = PROJECT
    prs.core_properties.subject = "Проект развития Госкорпорации «Росатом»: площадка АО «Изотоп»"
    slide_development(prs)
    prs.save(OUT)
    print(OUT)


if __name__ == "__main__":
    main()

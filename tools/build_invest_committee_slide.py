"""Генерация invest_committee_slide.pptx (python-pptx): АО «Изотоп», ИК Росатома."""
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
PROJECT = "АО «Изотоп» — национальный центр компетенций ядерной медицины"


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


def header(slide, title, subtitle=None, num=None):
    rect(slide, 0, 0, SLIDE_W, Inches(0.12), NAVY)
    rect(slide, 0, Inches(0.12), Inches(2.2), Inches(0.05), MID)
    text(slide, Inches(10.6), Inches(0.3), Inches(2.35), Inches(0.5),
         [p([("РОСАТОМ", 20, True, NAVY)], align=PP_ALIGN.RIGHT)], anchor=MSO_ANCHOR.MIDDLE)
    text(slide, Inches(10.6), Inches(0.75), Inches(2.35), Inches(0.3),
         [p([("АО «Изотоп»", 11, False, GREY)], align=PP_ALIGN.RIGHT)])
    long_title = len(title) > 70
    title_h = Inches(1.12) if long_title else Inches(0.6)
    text(slide, Inches(0.45), Inches(0.3), Inches(10.1), title_h,
         [p([(title, 22 if long_title else 26, True, NAVY)], line_spacing=0.95)],
         anchor=MSO_ANCHOR.TOP)
    if subtitle:
        text(slide, Inches(0.45), Inches(0.35) + title_h, Inches(10.1), Inches(0.4),
             [p([(subtitle, 13 if long_title else 14, False, BLUE)])])
    # footer
    rect(slide, Inches(0.45), Inches(7.05), Inches(12.43), Emu(9525), LINE)
    text(slide, Inches(0.45), Inches(7.08), Inches(10), Inches(0.3),
         [p([("Инвестиционный комитет Госкорпорации «Росатом»  ·  " + PROJECT, 9, False, GREY)])])
    if num is not None:
        text(slide, Inches(12.2), Inches(7.08), Inches(0.68), Inches(0.3),
             [p([(str(num), 9, True, NAVY)], align=PP_ALIGN.RIGHT)])


def section_label(slide, x, y, w, label):
    text(slide, x, y, w, Inches(0.35), [p([(label.upper(), 12, True, BLUE)])], anchor=MSO_ANCHOR.MIDDLE)
    rect(slide, x, y + Inches(0.36), w, Pt(1.5), BLUE)


# ---------------------------------------------------------------- slide 1
def slide_strategy(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header(
        s,
        "АО «Изотоп» — национальный центр компетенций ядерной медицины: "
        "100% суверенитет РФП к 20XX и выход на экспортные рынки $X млрд",
        "Полный жизненный цикл радиофармпрепаратов и медизделий с РВ — "
        "от НИОКР до клинического применения и экспорта",
        num=1,
    )
    section_label(s, Inches(0.45), Inches(2.05), Inches(12.43), "Стратегическая значимость")

    cards = [
        ("01", "Суверенитет", "X% → 100%",
         "доля отечественных РФП в онкологии РФ — с X% до 100% к 20XX"),
        ("02", "Здоровье нации", "+N тыс. пациентов / год",
         "ПЭТ/КТ-диагностики и радионуклидной терапии; вклад в цель "
         "ожидаемой продолжительности жизни 78+ лет к 2030"),
        ("03", "Позиция Росатома", "Топ-3 мирового рынка",
         "изотопной продукции; экспорт в N стран ЕАЭС, БРИКС, Азии, Африки"),
    ]
    x0, y0, gap = Inches(0.45), Inches(2.65), Inches(0.28)
    w = int((Inches(12.43) - 2 * gap) / 3)
    h = Inches(3.2)
    for i, (n, ttl, metric, body) in enumerate(cards):
        x = x0 + i * (w + gap)
        rect(s, x, y0, w, h, WHITE, LINE)
        rect(s, x, y0, w, Inches(0.75), NAVY)
        text(s, x + Inches(0.2), y0, Inches(0.7), Inches(0.75),
             [p([(n, 22, True, RGBColor(0x8F, 0xB8, 0xE6))])], anchor=MSO_ANCHOR.MIDDLE)
        text(s, x + Inches(0.85), y0, w - Inches(1.0), Inches(0.75),
             [p([(ttl, 17, True, WHITE)])], anchor=MSO_ANCHOR.MIDDLE)
        text(s, x + Inches(0.25), y0 + Inches(0.95), w - Inches(0.5), Inches(0.75),
             [p([(metric, 22, True, BLUE)])], anchor=MSO_ANCHOR.MIDDLE)
        rect(s, x + Inches(0.25), y0 + Inches(1.78), Inches(0.6), Pt(3), MID)
        text(s, x + Inches(0.25), y0 + Inches(1.95), w - Inches(0.5), Inches(1.2),
             [p([(body, 13.5, False, NAVY)], line_spacing=1.05)])

    # nat. projects footnote
    fy = Inches(6.12)
    rect(s, Inches(0.45), fy, Inches(12.43), Inches(0.78), LIGHT)
    rect(s, Inches(0.45), fy, Inches(0.09), Inches(0.78), BLUE)
    text(s, Inches(0.7), fy, Inches(12.1), Inches(0.78), [
        p([("Соответствие госприоритетам: ", 12, True, NAVY),
           ("национальные проекты «Здравоохранение», «Наука и университеты»; "
            "Стратегия развития медицинской промышленности РФ до 2030 года", 12, False, NAVY)]),
    ], anchor=MSO_ANCHOR.MIDDLE)
    return s


# ---------------------------------------------------------------- slide 2
def _set_block_arc(shape, start_deg, end_deg, thickness=0.22):
    gd = shape._element.spPr.find(qn("a:prstGeom")).find(qn("a:avLst"))
    for el in list(gd):
        gd.remove(el)
    for name, val in (("adj1", int(start_deg * 60000)), ("adj2", int(end_deg * 60000)),
                      ("adj3", int(thickness * 100000))):
        g = etree.SubElement(gd, qn("a:gd"))
        g.set("name", name)
        g.set("fmla", f"val {val}")


def slide_pdsa(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header(s, "Контур жизненного цикла РФП: цикл PDSA",
           "От госзаказа и НИОКР до клинического применения, непрерывного улучшения и экспорта", num=2)

    quads = [
        ("P", "PLAN", "Госзаказ и R&D", [
            "Техзадания на НИОКР по приоритетам Минздрава и Росатома",
            "Pipeline N молекул: Lu-177, Ac-225, Ga-68, F-18, Tc-99m, I-131",
            "Научные площадки: НИФХИ, НИИТФА, ФЭИ, ИРМ, НИИАР, Радиевый институт",
        ]),
        ("D", "DO", "Производство и логистика", [
            "GMP-завод полного цикла: N тыс. доз/год, запуск 20XX",
            "Собственная изотопная база: реакторы + циклотроны",
            "Логистика с учётом периода полураспада; покрытие N% территории РФ",
        ]),
        ("A", "ACT", "Оптимизация и экспорт", [
            "Непрерывное улучшение по данным real-world evidence",
            "Экспорт через канал зарубежных проектов Росатома",
            "Цель — $X млн/год к 20XX",
        ]),
        ("S", "STUDY", "Клиническое внедрение", [
            "Регистрация N препаратов до 20XX",
            "Сеть из N региональных центров ядерной медицины",
            "Клинические протоколы с НМИЦ онкологии им. Блохина / Герцена",
        ]),
    ]
    top, row_h, row_gap = Inches(1.6), Inches(2.6), Inches(0.2)
    card_w = Inches(4.95)
    xs = [Inches(0.45), SLIDE_W - Inches(0.45) - card_w]
    for idx, (letter, en, ru, bullets) in enumerate(quads):
        col, row = idx % 2, idx // 2
        x = xs[col]
        y = top + row * (row_h + row_gap)
        rect(s, x, y, card_w, row_h, WHITE, LINE)
        badge_x = x if col == 0 else x + card_w - Inches(0.75)
        rect(s, badge_x, y, Inches(0.75), row_h, NAVY)
        text(s, badge_x, y + Inches(0.1), Inches(0.75), Inches(0.8),
             [p([(letter, 32, True, WHITE)], align=PP_ALIGN.CENTER)])
        tx = x + Inches(0.95) if col == 0 else x + Inches(0.2)
        tw = card_w - Inches(1.15)
        text(s, tx, y + Inches(0.12), tw, Inches(0.45),
             [p([(en + " — ", 15, True, BLUE), (ru, 15, True, NAVY)])], anchor=MSO_ANCHOR.MIDDLE)
        rect(s, tx + Inches(0.05), y + Inches(0.62), tw - Inches(0.1), Pt(1), LINE)
        text(s, tx, y + Inches(0.72), tw, row_h - Inches(0.8),
             [p([(b, 12.5, False, NAVY)], bullet=True, space_after=5, line_spacing=1.0) for b in bullets])

    # central PDSA wheel
    cx, cy = SLIDE_W // 2, top + row_h + row_gap // 2
    d = Inches(2.4)
    wheel = [("P", 180, 270, NAVY), ("D", 270, 360, BLUE), ("S", 0, 90, MID), ("A", 90, 180, BLUE)]
    import math
    for letter, a0, a1, color in wheel:
        arc = rect(s, cx - d // 2, cy - d // 2, d, d, color, shape=MSO_SHAPE.BLOCK_ARC)
        _set_block_arc(arc, a0 + 5, a1 - 5, 0.2)
        mid = math.radians((a0 + a1) / 2)
        r = d / 2 * 0.8
        lx, ly = int(cx + r * math.cos(mid)), int(cy + r * math.sin(mid))
        text(s, lx - Inches(0.25), ly - Inches(0.25), Inches(0.5), Inches(0.5),
             [p([(letter, 18, True, WHITE)], align=PP_ALIGN.CENTER)], anchor=MSO_ANCHOR.MIDDLE, margin=0)
    # clockwise arrowheads at segment ends
    for ang in (270, 0, 90, 180):
        a = math.radians(ang)
        r = d / 2 * 0.8
        ax, ay = int(cx + r * math.cos(a)), int(cy + r * math.sin(a))
        tri = rect(s, ax - Inches(0.13), ay - Inches(0.13), Inches(0.26), Inches(0.26), MID, WHITE,
                   shape=MSO_SHAPE.ISOSCELES_TRIANGLE)
        tri.rotation = (ang + 180) % 360  # tip points clockwise
    inner = Inches(1.3)
    c = rect(s, cx - inner // 2, cy - inner // 2, inner, inner, LIGHT, shape=MSO_SHAPE.OVAL)
    text(s, cx - inner // 2, cy - inner // 2, inner, inner, [
        p([("PDSA", 18, True, NAVY)], align=PP_ALIGN.CENTER),
        p([("GMP-завод", 9, True, BLUE)], align=PP_ALIGN.CENTER),
        p([("полного цикла", 9, False, NAVY)], align=PP_ALIGN.CENTER),
    ], anchor=MSO_ANCHOR.MIDDLE, margin=0)
    text(s, cx - Inches(1.4), top - Inches(0.05), Inches(2.8), Inches(0.5), [
        p([("Сквозной цикл", 11, True, NAVY)], align=PP_ALIGN.CENTER),
        p([("ЖЦ РФП и медизделий с РВ", 10, False, GREY)], align=PP_ALIGN.CENTER),
    ], margin=0)
    text(s, cx - Inches(1.4), top + 2 * row_h + row_gap - Inches(0.5), Inches(2.8), Inches(0.5), [
        p([("Непрерывное улучшение", 11, True, NAVY)], align=PP_ALIGN.CENTER),
        p([("GLP · GMP · GDP · ISO 13485 · ОСПОРБ/НРБ", 9, False, GREY)], align=PP_ALIGN.CENTER),
    ], margin=0)
    return s


# ---------------------------------------------------------------- slide 3
def _cell(cell, runs, fill, align=PP_ALIGN.LEFT):
    cell.fill.solid()
    cell.fill.fore_color.rgb = fill
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.margin_left = cell.margin_right = Inches(0.12)
    cell.margin_top = cell.margin_bottom = Inches(0.04)
    tf = cell.text_frame
    tf.word_wrap = True
    para = tf.paragraphs[0]
    para.alignment = align
    for t, size, bold, color in runs:
        r = para.add_run()
        r.text = t
        _font(r, size, bold, color)


def slide_economics(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header(s, "Экономика проекта и решение ИК",
           "Ключевые параметры инвестиций, источники спроса и господдержки, риски и их митигирование", num=3)

    lx, ly, lw = Inches(0.45), Inches(1.6), Inches(7.35)
    section_label(s, lx, ly, lw, "Экономика и запрашиваемое решение")
    rows = [
        ("Инвестиции", [("X млрд руб.", 13, True, BLUE), ("  (капекс / НИОКР)", 12, False, NAVY)]),
        ("Выручка к 2030", [("X млрд руб.", 13, True, BLUE), ("  (экспорт X%)", 12, False, NAVY)]),
        ("IRR / окупаемость", [("IRR X%", 13, True, BLUE), ("  /  окупаемость X лет", 12, False, NAVY)]),
        ("Оффтейк", [("Госзаказ Минздрава + региональные онкоцентры", 12, False, NAVY)]),
        ("Господдержка", [("СЗПК, субсидии Минпромторга, ФРП", 12, False, NAVY)]),
    ]
    tbl_y = ly + Inches(0.55)
    row_h = Inches(0.56)
    gt = s.shapes.add_table(len(rows) + 1, 2, lx, tbl_y, lw, row_h * (len(rows) + 1))
    tbl = gt.table
    tbl.first_row = True
    tbl.horz_banding = False
    tbl.columns[0].width = Inches(2.3)
    tbl.columns[1].width = lw - Inches(2.3)
    _cell(tbl.cell(0, 0), [("Параметр", 12, True, WHITE)], NAVY)
    _cell(tbl.cell(0, 1), [("Значение", 12, True, WHITE)], NAVY)
    for i, (k, v) in enumerate(rows, start=1):
        fill = LIGHT if i % 2 else WHITE
        _cell(tbl.cell(i, 0), [(k, 12, True, NAVY)], fill)
        _cell(tbl.cell(i, 1), v, fill)
    for r in tbl.rows:
        r.height = row_h

    # risks
    rx, rw = Inches(8.1), Inches(4.78)
    section_label(s, rx, ly, rw, "Риски и митигирование")
    risks = [
        ("Санкционные", "Собственная реакторная и циклотронная база"),
        ("Регуляторные", "Опережающий R&D и ускоренные процедуры Минздрава"),
        ("Рыночные", "Ценовое преимущество и гарантированный госзаказ"),
    ]
    ry = ly + Inches(0.55)
    rh, rg = Inches(1.0), Inches(0.12)
    for i, (risk, mit) in enumerate(risks):
        y = ry + i * (rh + rg)
        rect(s, rx, y, rw, rh, WHITE, LINE)
        rect(s, rx, y, Inches(0.09), rh, BLUE)
        text(s, rx + Inches(0.22), y + Inches(0.08), rw - Inches(0.35), Inches(0.36),
             [p([(risk + " риски", 13, True, NAVY)])], anchor=MSO_ANCHOR.MIDDLE)
        text(s, rx + Inches(0.22), y + Inches(0.44), rw - Inches(0.35), Inches(0.5),
             [p([(mit, 12.5, False, NAVY)], bullet=True, bullet_char="→", bullet_indent=18)], anchor=MSO_ANCHOR.TOP)

    # decision bar
    dy = Inches(5.75)
    rect(s, Inches(0.45), dy, Inches(12.43), Inches(1.05), NAVY)
    rect(s, Inches(0.45), dy, Inches(0.12), Inches(1.05), MID)
    text(s, Inches(0.8), dy, Inches(11.9), Inches(1.05), [
        p([("ЗАПРАШИВАЕМОЕ РЕШЕНИЕ ИК", 12, True, RGBColor(0x8F, 0xB8, 0xE6))], space_after=3),
        p([("Одобрить инвестиции X млрд руб.;  контрольная точка — выход на проектную мощность к 20XX",
            17, True, WHITE)]),
    ], anchor=MSO_ANCHOR.MIDDLE)
    return s


def main():
    prs = Presentation()
    prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
    prs.core_properties.title = PROJECT
    prs.core_properties.subject = "Материалы к инвестиционному комитету Госкорпорации «Росатом»"
    slide_strategy(prs)
    slide_pdsa(prs)
    slide_economics(prs)
    prs.save(OUT)
    print(OUT)


if __name__ == "__main__":
    main()

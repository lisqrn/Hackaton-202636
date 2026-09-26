import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

# ==============================================================================
# MINIMALIST DEEPTECH DECK: Dark Graphite + Neon Emerald + Centered Visual Schemas
# ==============================================================================
COLOR_BG = RGBColor(0x0A, 0x0E, 0x11)              # #0A0E11 Deepest Dark Slate
COLOR_CARD = RGBColor(0x12, 0x18, 0x1D)            # #12181D Card Surface
COLOR_CARD_BORDER = RGBColor(0x1E, 0x29, 0x32)     # #1E2932 Clean border
COLOR_EMERALD = RGBColor(0x10, 0xB9, 0x81)         # #10B981 Deeptech Emerald
COLOR_NEON = RGBColor(0x34, 0xD3, 0x99)            # #34D399 Neon Emerald
COLOR_TEXT_WHITE = RGBColor(0xFF, 0xFF, 0xFF)      # #FFFFFF Pure White
COLOR_TEXT_MUTED = RGBColor(0x94, 0xA3, 0xB8)      # #94A3B8 Cool Gray
COLOR_CARD_GLOW = RGBColor(0x10, 0x2E, 0x22)       # #102E22 Glow Card

FONT_TITLE = "Plus Jakarta Sans"
FONT_BODY = "Plus Jakarta Sans"
FONT_MONO = "JetBrains Mono"

def create_deck():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    def set_bg(slide):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = COLOR_BG
        bg.line.color.rgb = COLOR_BG
        return bg

    def add_centered_header(slide, tag_text, title_text, slide_num):
        # Category Tag (Centered)
        tb_tag = slide.shapes.add_textbox(Inches(1.0), Inches(0.4), Inches(11.333), Inches(0.35))
        tf_t = tb_tag.text_frame
        p_t = tf_t.paragraphs[0]
        p_t.alignment = PP_ALIGN.CENTER
        p_t.text = tag_text.upper()
        p_t.font.name = FONT_MONO
        p_t.font.size = Pt(11)
        p_t.font.bold = True
        p_t.font.color.rgb = COLOR_NEON

        # Main Title (Centered)
        tb_tit = slide.shapes.add_textbox(Inches(1.0), Inches(0.75), Inches(11.333), Inches(0.75))
        tf_tit = tb_tit.text_frame
        p_tit = tf_tit.paragraphs[0]
        p_tit.alignment = PP_ALIGN.CENTER
        p_tit.text = title_text
        p_tit.font.name = FONT_TITLE
        p_tit.font.size = Pt(24)
        p_tit.font.bold = True
        p_tit.font.color.rgb = COLOR_TEXT_WHITE

        # Slide Number (Top Right)
        tb_num = slide.shapes.add_textbox(Inches(11.3), Inches(0.4), Inches(1.2), Inches(0.35))
        tf_num = tb_num.text_frame
        p_num = tf_num.paragraphs[0]
        p_num.alignment = PP_ALIGN.RIGHT
        p_num.text = f"0{slide_num} / 08"
        p_num.font.name = FONT_MONO
        p_num.font.size = Pt(11)
        p_num.font.bold = True
        p_num.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 1: Cover (Centered, Clean, Big Impact)
    # =========================================================================
    slide1 = prs.slides.add_slide(blank_layout)
    set_bg(slide1)

    tb_h = slide1.shapes.add_textbox(Inches(1.0), Inches(1.2), Inches(11.333), Inches(3.2))
    tf_h = tb_h.text_frame
    tf_h.word_wrap = True

    p_b = tf_h.paragraphs[0]
    p_b.alignment = PP_ALIGN.CENTER
    p_b.text = "DEEPTECH GIGAHACK 2026 • BIOMENTORHUB × PROFI TRACK"
    p_b.font.name = FONT_MONO
    p_b.font.size = Pt(12)
    p_b.font.bold = True
    p_b.font.color.rgb = COLOR_NEON

    p_main = tf_h.add_paragraph()
    p_main.alignment = PP_ALIGN.CENTER
    p_main.text = "\nPackShift"
    p_main.font.name = FONT_TITLE
    p_main.font.size = Pt(48)
    p_main.font.bold = True
    p_main.font.color.rgb = COLOR_TEXT_WHITE

    p_sub = tf_h.add_paragraph()
    p_sub.alignment = PP_ALIGN.CENTER
    p_sub.text = "0% ПЕРВИЧНОГО ПЛАСТИКА • 100% ЦИРКУЛЯРНОСТЬ"
    p_sub.font.name = FONT_MONO
    p_sub.font.size = Pt(16)
    p_sub.font.bold = True
    p_sub.font.color.rgb = COLOR_NEON

    p_desc = tf_h.add_paragraph()
    p_desc.alignment = PP_ALIGN.CENTER
    p_desc.text = "\nИнженерная система циркулярной пищевой упаковки для сети супермаркетов Profi"
    p_desc.font.name = FONT_BODY
    p_desc.font.size = Pt(15)
    p_desc.font.color.rgb = COLOR_TEXT_MUTED

    # 4 Centered Proof Badges
    m_data = [
        ("-88%", "Первичный пластик"),
        ("250°C", "Вызов гриля (Annex)"),
        ("+4 дня", "Свежесть продуктов"),
        ("Class A", "RecyClass & PPWR"),
    ]
    card_w = Inches(2.6)
    card_h = Inches(1.8)
    card_y = Inches(4.8)
    gap = Inches(0.3)
    start_x = Inches(1.1)

    for i, (val, label) in enumerate(m_data):
        cx = start_x + i * (card_w + gap)
        card = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, card_y, card_w, card_h)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD_GLOW if i == 1 else COLOR_CARD
        card.line.color.rgb = COLOR_EMERALD if i == 1 else COLOR_CARD_BORDER
        card.line.width = Pt(1.5 if i == 1 else 1)

        tb_c = slide1.shapes.add_textbox(cx + Inches(0.1), card_y + Inches(0.2), card_w - Inches(0.2), card_h - Inches(0.4))
        tf_c = tb_c.text_frame
        tf_c.word_wrap = True

        p_v = tf_c.paragraphs[0]
        p_v.alignment = PP_ALIGN.CENTER
        p_v.text = val
        p_v.font.name = FONT_MONO
        p_v.font.size = Pt(28)
        p_v.font.bold = True
        p_v.font.color.rgb = COLOR_NEON

        p_l = tf_c.add_paragraph()
        p_l.alignment = PP_ALIGN.CENTER
        p_l.text = label
        p_l.font.name = FONT_BODY
        p_l.font.size = Pt(12)
        p_l.font.bold = True
        p_l.font.color.rgb = COLOR_TEXT_WHITE

    # =========================================================================
    # SLIDE 2: Schema Split (Problem vs Solution)
    # =========================================================================
    slide2 = prs.slides.add_slide(blank_layout)
    set_bg(slide2)
    add_centered_header(slide2, "ТЕХНОЛОГИЧЕСКИЙ РАЗРЫВ", "Традиционная упаковка vs Решение PackShift", 2)

    # Left: The Bad Traditional Chain
    c_bad = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(1.8), Inches(5.2), Inches(5.0))
    c_bad.fill.solid()
    c_bad.fill.fore_color.rgb = COLOR_CARD
    c_bad.line.color.rgb = RGBColor(0x7F, 0x1D, 0x1D)
    c_bad.line.width = Pt(1.5)

    tb_b = slide2.shapes.add_textbox(Inches(1.2), Inches(2.0), Inches(4.8), Inches(4.5))
    tf_b = tb_b.text_frame
    tf_b.word_wrap = True

    p = tf_b.paragraphs[0]
    p.text = "❌ ТРАДИЦИОННАЯ УПАКОВКА (ТУПИК)"
    p.font.name = FONT_MONO
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0xEF, 0x44, 0x44)

    bad_nodes = [
        ("Многослойный ламинат (ПЭ+ALU)", "Не перерабатывается $\rightarrow$ штрафы EPR до +45%"),
        ("Плавление при 95°C–120°C", "Течь горячего жира в витрине и запах синтетики"),
        ("Токсичные добавки PFAS", "Миграция опасных фторорганик в пищу при нагреве"),
        ("«Парниковый эффект»", "Пар блокирован: корочка гриля размокает за 40 мин"),
    ]
    for b_title, b_desc in bad_nodes:
        p_it = tf_b.add_paragraph()
        p_it.text = f"\n▪ {b_title}\n  "
        p_it.font.bold = True
        p_it.font.size = Pt(12)
        p_it.font.color.rgb = COLOR_TEXT_WHITE
        r = p_it.add_run()
        r.text = b_desc
        r.font.bold = False
        r.font.size = Pt(10.5)
        r.font.color.rgb = COLOR_TEXT_MUTED

    # Right: The PackShift Deeptech Solution
    c_good = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.1), Inches(1.8), Inches(5.2), Inches(5.0))
    c_good.fill.solid()
    c_good.fill.fore_color.rgb = COLOR_CARD_GLOW
    c_good.line.color.rgb = COLOR_EMERALD
    c_good.line.width = Pt(1.5)

    tb_g = slide2.shapes.add_textbox(Inches(7.3), Inches(2.0), Inches(4.8), Inches(4.5))
    tf_g = tb_g.text_frame
    tf_g.word_wrap = True

    p = tf_g.paragraphs[0]
    p.text = "✔ ИННОВАЦИЯ PACKSHIFT DEEPTECH"
    p.font.name = FONT_MONO
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = COLOR_NEON

    good_nodes = [
        ("100% моно-материал rPET", "RecyClass Class A $\rightarrow$ прямое освобождение от EPR-налога"),
        ("Стойкость до 250°C (Dual-Oven)", "Более 6 часов в витрине при 95°C без усадки и дефектов"),
        ("100% PFAS-Free барьер Kit 12", "Водная нано-дисперсия альгината водорослей и хитозана"),
        ("Дышащая микропористая мембрана", "Отводит лишний пар, сохраняя корочку гриля хрустящей"),
    ]
    for g_title, g_desc in good_nodes:
        p_it = tf_g.add_paragraph()
        p_it.text = f"\n✔ {g_title}\n  "
        p_it.font.bold = True
        p_it.font.size = Pt(12)
        p_it.font.color.rgb = COLOR_NEON
        r = p_it.add_run()
        r.text = g_desc
        r.font.bold = False
        r.font.size = Pt(10.5)
        r.font.color.rgb = COLOR_TEXT_WHITE

    # =========================================================================
    # SLIDE 3: Material Schemas for Hot Grill (180°–250°C)
    # =========================================================================
    slide3 = prs.slides.add_slide(blank_layout)
    set_bg(slide3)
    add_centered_header(slide3, "СПЕЦИАЛЬНЫЙ ВЫЗОВ ХАКАТОНА", "Материаловедение: 2 решения для витрин гриля (250°C)", 3)

    # Box A
    c_a = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(1.8), Inches(5.3), Inches(5.0))
    c_a.fill.solid()
    c_a.fill.fore_color.rgb = COLOR_CARD_GLOW
    c_a.line.color.rgb = COLOR_EMERALD
    c_a.line.width = Pt(1.5)

    tb_a = slide3.shapes.add_textbox(Inches(1.2), Inches(2.0), Inches(4.9), Inches(4.5))
    tf_a = tb_a.text_frame
    tf_a.word_wrap = True

    p = tf_a.paragraphs[0]
    p.text = "РЕШЕНИЕ А: МОНО-ПАКЕТ"
    p.font.name = FONT_MONO
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_NEON

    p_t = tf_a.add_paragraph()
    p_t.text = "CPET Ultra-Flex"
    p_t.font.name = FONT_TITLE
    p_t.font.size = Pt(22)
    p_t.font.bold = True
    p_t.font.color.rgb = COLOR_TEXT_WHITE

    pts_a = [
        "Диапазон от -40°C до +250°C (Dual-Ovenable)",
        "Кристалличность rPET > 38% держит форму 6+ часов",
        "100% жиробарьер Kit Test 12 без PFAS",
        "100% вторичная переработка в поток ПЭТ-бутылок",
        "Себестоимость: -14% vs импортный ламинат с фольгой",
    ]
    for pt in pts_a:
        p_pt = tf_a.add_paragraph()
        p_pt.text = f"✔ {pt}"
        p_pt.font.size = Pt(11.5)
        p_pt.font.color.rgb = COLOR_TEXT_WHITE

    # Box B
    c_b = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.0), Inches(1.8), Inches(5.3), Inches(5.0))
    c_b.fill.solid()
    c_b.fill.fore_color.rgb = COLOR_CARD
    c_b.line.color.rgb = COLOR_CARD_BORDER
    c_b.line.width = Pt(1.5)

    tb_b = slide3.shapes.add_textbox(Inches(7.2), Inches(2.0), Inches(4.9), Inches(4.5))
    tf_b = tb_b.text_frame
    tf_b.word_wrap = True

    p = tf_b.paragraphs[0]
    p.text = "РЕШЕНИЕ Б: БИО-КОМПОЗИТ"
    p.font.name = FONT_MONO
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_NEON

    p_t = tf_b.add_paragraph()
    p_t.text = "LignoShield™ BioBox"
    p_t.font.name = FONT_TITLE
    p_t.font.size = Pt(22)
    p_t.font.bold = True
    p_t.font.color.rgb = COLOR_TEXT_WHITE

    pts_b = [
        "100% формованная багасса сахарного тростника",
        "Водный нано-барьер из морских водорослей и хитозана",
        "Микропористая мембрана отводит пар (хруст корочки)",
        "Термостойкость до 220°C, 6 часов в витрине гриля",
        "Домашний компост за 60 дней (OK Compost HOME)",
    ]
    for pt in pts_b:
        p_pt = tf_b.add_paragraph()
        p_pt.text = f"✔ {pt}"
        p_pt.font.size = Pt(11.5)
        p_pt.font.color.rgb = COLOR_TEXT_WHITE

    # =========================================================================
    # SLIDE 4: 4 Departments Quadrants Schema
    # =========================================================================
    slide4 = prs.slides.add_slide(blank_layout)
    set_bg(slide4)
    add_centered_header(slide4, "АРХИТЕКТУРА РЕШЕНИЙ", "Оптимизация 4 зон супермаркета сети Profi", 4)

    quads = [
        ("🍗 Кулинария (Гриль)", "CPET Ultra-Flex", "250°C", "Выдержка 6ч в витрине, 100% rPET, Kit 12 без PFAS"),
        ("🥬 Свежие Овощи", "EMAP Laser™", "+4 Дня", "Лазерная микроперфорация O₂/CO₂, -32% списаний"),
        ("🥩 Мясной Отдел", "Capillary Tray", "0 Салфеток", "Капиллярный замок экссудата на 45 мл, чистый рециклинг"),
        ("🥐 Пекарня", "NatureVent™", "Pulpable", "Крафт FSC + целлюлозное био-окно NatureFlex™, тест CEPI"),
    ]

    q_w = Inches(2.65)
    q_h = Inches(4.9)
    q_y = Inches(1.8)
    q_gap = Inches(0.24)

    for i, (title, tech, stat, desc) in enumerate(quads):
        cx = Inches(1.0) + i * (q_w + q_gap)
        card = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, q_y, q_w, q_h)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD
        card.line.color.rgb = COLOR_CARD_BORDER
        card.line.width = Pt(1)

        tb = slide4.shapes.add_textbox(cx + Inches(0.12), q_y + Inches(0.25), q_w - Inches(0.24), q_h - Inches(0.5))
        tf = tb.text_frame
        tf.word_wrap = True

        p_t = tf.paragraphs[0]
        p_t.alignment = PP_ALIGN.CENTER
        p_t.text = title
        p_t.font.name = FONT_BODY
        p_t.font.size = Pt(13)
        p_t.font.bold = True
        p_t.font.color.rgb = COLOR_TEXT_WHITE

        p_tc = tf.add_paragraph()
        p_tc.alignment = PP_ALIGN.CENTER
        p_tc.text = tech
        p_tc.font.name = FONT_MONO
        p_tc.font.size = Pt(14)
        p_tc.font.bold = True
        p_tc.font.color.rgb = COLOR_NEON

        p_s = tf.add_paragraph()
        p_s.alignment = PP_ALIGN.CENTER
        p_s.text = f"\n{stat}\n"
        p_s.font.name = FONT_MONO
        p_s.font.size = Pt(22)
        p_s.font.bold = True
        p_s.font.color.rgb = COLOR_NEON

        p_d = tf.add_paragraph()
        p_d.alignment = PP_ALIGN.CENTER
        p_d.text = desc
        p_d.font.name = FONT_BODY
        p_d.font.size = Pt(11)
        p_d.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 5: Economics on 100 Stores (Big Numbers)
    # =========================================================================
    slide5 = prs.slides.add_slide(blank_layout)
    set_bg(slide5)
    add_centered_header(slide5, "БИЗНЕС-КЕЙС ДЛЯ PROFI", "Экономический эффект на масштабе 100 магазинов", 5)

    econ_stats = [
        ("1 840 т", "Первичный пластик", "Тонн первичных полимеров замещено в год"),
        ("4 250 т", "Выбросы CO₂e", "Предотвращено парниковых газов Scope 3"),
        ("€1.28M", "Экономия на EPR", "Прямая годовая экономия на эко-сборах"),
        ("€700K", "Снижение списаний", "Сбереженная продукция (+4 дня свежести)"),
    ]

    es_w = Inches(2.65)
    es_h = Inches(3.2)
    es_y = Inches(1.8)
    es_gap = Inches(0.24)

    for i, (val, label, sub) in enumerate(econ_stats):
        cx = Inches(1.0) + i * (es_w + es_gap)
        card = slide5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, es_y, es_w, es_h)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD_GLOW if i == 2 else COLOR_CARD
        card.line.color.rgb = COLOR_EMERALD if i == 2 else COLOR_CARD_BORDER
        card.line.width = Pt(1.5 if i == 2 else 1)

        tb = slide5.shapes.add_textbox(cx + Inches(0.12), es_y + Inches(0.2), es_w - Inches(0.24), es_h - Inches(0.4))
        tf = tb.text_frame
        tf.word_wrap = True

        p_v = tf.paragraphs[0]
        p_v.alignment = PP_ALIGN.CENTER
        p_v.text = val
        p_v.font.name = FONT_MONO
        p_v.font.size = Pt(28)
        p_v.font.bold = True
        p_v.font.color.rgb = COLOR_NEON

        p_l = tf.add_paragraph()
        p_l.alignment = PP_ALIGN.CENTER
        p_l.text = label
        p_l.font.name = FONT_BODY
        p_l.font.size = Pt(13.5)
        p_l.font.bold = True
        p_l.font.color.rgb = COLOR_TEXT_WHITE

        p_s = tf.add_paragraph()
        p_s.alignment = PP_ALIGN.CENTER
        p_s.text = f"\n{sub}"
        p_s.font.name = FONT_BODY
        p_s.font.size = Pt(10.5)
        p_s.font.color.rgb = COLOR_TEXT_MUTED

    # Bottom CapEx Strip
    c_cx = slide5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(5.3), Inches(11.333), Inches(1.4))
    c_cx.fill.solid()
    c_cx.fill.fore_color.rgb = COLOR_CARD_GLOW
    c_cx.line.color.rgb = COLOR_EMERALD
    c_cx.line.width = Pt(1.5)

    tb_cx = slide5.shapes.add_textbox(Inches(1.2), Inches(5.45), Inches(10.9), Inches(1.1))
    tf_cx = tb_cx.text_frame
    tf_cx.word_wrap = True

    p_cx1 = tf_cx.paragraphs[0]
    p_cx1.alignment = PP_ALIGN.CENTER
    p_cx1.text = "⚡ НУЛЕВОЙ CAPEX (DROP-IN COMPATIBILITY):"
    p_cx1.font.name = FONT_MONO
    p_cx1.font.size = Pt(13)
    p_cx1.font.bold = True
    p_cx1.font.color.rgb = COLOR_NEON

    p_cx2 = tf_cx.add_paragraph()
    p_cx2.alignment = PP_ALIGN.CENTER
    p_cx2.text = (
        "100% совместимость со стандартными запайщиками и тепловыми витринами Profi. "
        "Себестоимость моно-пакетов CPET на 14% ниже импортных многослойных ламинатов с фольгой."
    )
    p_cx2.font.name = FONT_BODY
    p_cx2.font.size = Pt(12)
    p_cx2.font.color.rgb = COLOR_TEXT_WHITE

    # =========================================================================
    # SLIDE 6: Standards & Compliance (4 Shields)
    # =========================================================================
    slide6 = prs.slides.add_slide(blank_layout)
    set_bg(slide6)
    add_centered_header(slide6, "РЕГУЛЯТОРНЫЙ ЩИТ", "Международные стандарты безопасности и рециклинга", 6)

    st_quads = [
        ("RecyClass A", "RECYCLING", "Высший класс сортировки. Совместимость с NIR-сепараторами и механическим рециклингом в ЕС."),
        ("PPWR 2030", "REGULATION", "Упреждение европейских норм: >35% пищевого rPET, 0% первичного пластика в фасовке."),
        ("Zero PFAS", "CHEMICAL SAFETY", "100% PFAS-Free. Лабораторный тест на отсутствие фторорганики в контакте до 250°C."),
        ("EFSA / FDA", "FOOD CONTACT", "Сертификация Regulation EU No 10/2011 для прямого контакта с горячим жиром."),
    ]

    for i, (title, tag, desc) in enumerate(st_quads):
        cx = Inches(1.0) + i * (q_w + q_gap)
        card = slide6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, q_y, q_w, q_h)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD
        card.line.color.rgb = COLOR_CARD_BORDER
        card.line.width = Pt(1)

        tb = slide6.shapes.add_textbox(cx + Inches(0.12), q_y + Inches(0.3), q_w - Inches(0.24), q_h - Inches(0.6))
        tf = tb.text_frame
        tf.word_wrap = True

        p_tag = tf.paragraphs[0]
        p_tag.alignment = PP_ALIGN.CENTER
        p_tag.text = tag
        p_tag.font.name = FONT_MONO
        p_tag.font.size = Pt(11)
        p_tag.font.bold = True
        p_tag.font.color.rgb = COLOR_NEON

        p_tit = tf.add_paragraph()
        p_tit.alignment = PP_ALIGN.CENTER
        p_tit.text = f"\n{title}\n"
        p_tit.font.name = FONT_TITLE
        p_tit.font.size = Pt(20)
        p_tit.font.bold = True
        p_tit.font.color.rgb = COLOR_TEXT_WHITE

        p_d = tf.add_paragraph()
        p_d.alignment = PP_ALIGN.CENTER
        p_d.text = desc
        p_d.font.name = FONT_BODY
        p_d.font.size = Pt(11.5)
        p_d.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 7: Pipeline Flow (6 Weeks Roadmap)
    # =========================================================================
    slide7 = prs.slides.add_slide(blank_layout)
    set_bg(slide7)
    add_centered_header(slide7, "ДОРОЖНАЯ КАРТА ВНЕДРЕНИЯ", "План пилотного запуска в Profi: 6 недель", 7)

    pipe_steps = [
        ("ЭТАП 1 • 2 НЕДЕЛИ", "Лабораторный тест", "Тест пакетов CPET и боксов при 95°C–120°C (6 часов). Проверка удержания жира Kit 12 и протокол EFSA."),
        ("ЭТАП 2 • 4 НЕДЕЛИ", "Пилот в 5 магазинах", "Полевое тестирование в отделах Гриль и Овощи. Сбор отзывов поваров кулинарии и покупателей сети."),
        ("ЭТАП 3 • КВАРТАЛ 2", "Масштабирование", "Развертывание в сети Profi, подключение автоматического дашборда учета сокращения Scope 3 выбросов."),
    ]

    p_w = Inches(3.6)
    p_h = Inches(3.2)
    p_y = Inches(1.8)
    p_gap = Inches(0.26)

    for i, (tag, title, desc) in enumerate(pipe_steps):
        cx = Inches(1.0) + i * (p_w + p_gap)
        card = slide7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, p_y, p_w, p_h)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD_GLOW if i == 0 else COLOR_CARD
        card.line.color.rgb = COLOR_EMERALD if i == 0 else COLOR_CARD_BORDER
        card.line.width = Pt(1.5 if i == 0 else 1)

        tb = slide7.shapes.add_textbox(cx + Inches(0.15), p_y + Inches(0.25), p_w - Inches(0.3), p_h - Inches(0.5))
        tf = tb.text_frame
        tf.word_wrap = True

        p_t = tf.paragraphs[0]
        p_t.alignment = PP_ALIGN.CENTER
        p_t.text = tag
        p_t.font.name = FONT_MONO
        p_t.font.size = Pt(11)
        p_t.font.bold = True
        p_t.font.color.rgb = COLOR_NEON

        p_tit = tf.add_paragraph()
        p_tit.alignment = PP_ALIGN.CENTER
        p_tit.text = f"\n{title}\n"
        p_tit.font.name = FONT_TITLE
        p_tit.font.size = Pt(18)
        p_tit.font.bold = True
        p_tit.font.color.rgb = COLOR_TEXT_WHITE

        p_d = tf.add_paragraph()
        p_d.alignment = PP_ALIGN.CENTER
        p_d.text = desc
        p_d.font.name = FONT_BODY
        p_d.font.size = Pt(11.5)
        p_d.font.color.rgb = COLOR_TEXT_MUTED

    # Bottom Sample Ready Banner
    c_smp = slide7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(5.3), Inches(11.333), Inches(1.4))
    c_smp.fill.solid()
    c_smp.fill.fore_color.rgb = COLOR_CARD_GLOW
    c_smp.line.color.rgb = COLOR_NEON
    c_smp.line.width = Pt(1.5)

    tb_smp = slide7.shapes.add_textbox(Inches(1.2), Inches(5.45), Inches(10.9), Inches(1.1))
    tf_smp = tb_smp.text_frame
    tf_smp.word_wrap = True

    p_s1 = tf_smp.paragraphs[0]
    p_s1.alignment = PP_ALIGN.CENTER
    p_s1.text = "📦 ПАРТИЯ ИЗ 200 ОБРАЗЦОВ ГОТОВА К ПЕРЕДАЧЕ PROFI"
    p_s1.font.name = FONT_MONO
    p_s1.font.size = Pt(13)
    p_s1.font.bold = True
    p_s1.font.color.rgb = COLOR_NEON

    p_s2 = tf_smp.add_paragraph()
    p_s2.alignment = PP_ALIGN.CENTER
    p_s2.text = "Комплект пакетов CPET Ultra-Flex и термобоксов LignoShield подготовлен к испытаниям на тепловых витринах сети Profi."
    p_s2.font.name = FONT_BODY
    p_s2.font.size = Pt(12)
    p_s2.font.color.rgb = COLOR_TEXT_WHITE

    # =========================================================================
    # SLIDE 8: Conclusion & Call to Action
    # =========================================================================
    slide8 = prs.slides.add_slide(blank_layout)
    set_bg(slide8)

    c_f = slide8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.5), Inches(1.2), Inches(10.333), Inches(5.1))
    c_f.fill.solid()
    c_f.fill.fore_color.rgb = COLOR_CARD
    c_f.line.color.rgb = COLOR_EMERALD
    c_f.line.width = Pt(2)

    tb_f = slide8.shapes.add_textbox(Inches(1.8), Inches(1.5), Inches(9.733), Inches(4.5))
    tf_f = tb_f.text_frame
    tf_f.word_wrap = True

    p_f0 = tf_f.paragraphs[0]
    p_f0.alignment = PP_ALIGN.CENTER
    p_f0.text = "DEEPTECH GIGAHACK 2026 • BIOMENTORHUB × PROFI"
    p_f0.font.name = FONT_MONO
    p_f0.font.size = Pt(12)
    p_f0.font.bold = True
    p_f0.font.color.rgb = COLOR_NEON

    p_f1 = tf_f.add_paragraph()
    p_f1.alignment = PP_ALIGN.CENTER
    p_f1.text = "\nДелаем циркулярность прибыльной для Profi уже сегодня"
    p_f1.font.name = FONT_TITLE
    p_f1.font.size = Pt(32)
    p_f1.font.bold = True
    p_f1.font.color.rgb = COLOR_TEXT_WHITE

    p_f2 = tf_f.add_paragraph()
    p_f2.alignment = PP_ALIGN.CENTER
    p_f2.text = "0% первичного пластика • -14% себестоимость • Нулевой CapEx"
    p_f2.font.name = FONT_MONO
    p_f2.font.size = Pt(15)
    p_f2.font.bold = True
    p_f2.font.color.rgb = COLOR_NEON

    p_f3 = tf_f.add_paragraph()
    p_f3.alignment = PP_ALIGN.CENTER
    p_f3.text = "\nPackShift позволяет сети Profi закрыть норматив EU PPWR 2030, защитить маржинальность кулинарии и гарантировать пищевую безопасность покупателей."
    p_f3.font.name = FONT_BODY
    p_f3.font.size = Pt(13.5)
    p_f3.font.color.rgb = COLOR_TEXT_MUTED

    p_f4 = tf_f.add_paragraph()
    p_f4.alignment = PP_ALIGN.CENTER
    p_f4.text = "\n\nБлагодарим за внимание! Готовы ответить на ваши вопросы."
    p_f4.font.name = FONT_TITLE
    p_f4.font.size = Pt(18)
    p_f4.font.bold = True
    p_f4.font.color.rgb = COLOR_TEXT_WHITE

    output_path = os.path.join(os.path.dirname(__file__), "PackShift_Profi_Gigahack2026.pptx")
    prs.save(output_path)
    print(f"Minimalist centered schema deck saved to: {output_path}")

if __name__ == "__main__":
    create_deck()

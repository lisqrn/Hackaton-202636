import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

# ==============================================================================
# HACKATHON DEEPTECH DESIGN SYSTEM: Dark Slate Charcoal + Vibrant Emerald Neon
# (Никакой привязки к сайту: чистый технологичный питч-дек хакатона)
# ==============================================================================
COLOR_BG_DARK = RGBColor(0x0F, 0x14, 0x18)          # #0F1418 Deep Dark Charcoal
COLOR_CARD_BG = RGBColor(0x18, 0x20, 0x26)          # #182026 Dark Slate Card
COLOR_CARD_BORDER = RGBColor(0x24, 0x32, 0x3B)      # #24323B Card Border
COLOR_EMERALD = RGBColor(0x10, 0xB9, 0x81)          # #10B981 Deeptech Emerald
COLOR_EMERALD_NEON = RGBColor(0x34, 0xD3, 0x99)     # #34D399 Vibrant Glowing Emerald
COLOR_TEXT_WHITE = RGBColor(0xFF, 0xFF, 0xFF)       # #FFFFFF Pure White
COLOR_TEXT_MUTED = RGBColor(0x94, 0xA3, 0xB8)       # #94A3B8 Cool Gray
COLOR_ACCENT_BG = RGBColor(0x12, 0x2E, 0x24)        # #122E24 Emerald Glow Card

FONT_TITLE = "Plus Jakarta Sans"
FONT_BODY = "Plus Jakarta Sans"
FONT_MONO = "JetBrains Mono"

def create_deck():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    logo_path = os.path.join(os.path.dirname(__file__), "img", "logo.png")
    if not os.path.exists(logo_path):
        logo_path = os.path.join(os.path.dirname(__file__), "Проэкт Хакатон", "img", "logo.png")

    def set_slide_bg(slide):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = COLOR_BG_DARK
        bg.line.color.rgb = COLOR_BG_DARK
        return bg

    def add_top_bar(slide, category, title, slide_idx):
        # Category Tag
        tb_cat = slide.shapes.add_textbox(Inches(0.9), Inches(0.45), Inches(9.5), Inches(0.35))
        tf_cat = tb_cat.text_frame
        p_cat = tf_cat.paragraphs[0]
        p_cat.text = f"{category.upper()} • GIGAHACK 2026"
        p_cat.font.name = FONT_MONO
        p_cat.font.size = Pt(11)
        p_cat.font.bold = True
        p_cat.font.color.rgb = COLOR_EMERALD_NEON

        # Slide Title
        tb_tit = slide.shapes.add_textbox(Inches(0.9), Inches(0.75), Inches(9.5), Inches(0.8))
        tf_tit = tb_tit.text_frame
        p_tit = tf_tit.paragraphs[0]
        p_tit.text = title
        p_tit.font.name = FONT_TITLE
        p_tit.font.size = Pt(22)
        p_tit.font.bold = True
        p_tit.font.color.rgb = COLOR_TEXT_WHITE

        # Slide Number Badge on top right
        tb_num = slide.shapes.add_textbox(Inches(11.2), Inches(0.5), Inches(1.3), Inches(0.4))
        tf_num = tb_num.text_frame
        p_num = tf_num.paragraphs[0]
        p_num.alignment = PP_ALIGN.RIGHT
        p_num.text = f"0{slide_idx} / 08"
        p_num.font.name = FONT_MONO
        p_num.font.size = Pt(12)
        p_num.font.bold = True
        p_num.font.color.rgb = COLOR_TEXT_MUTED

        # Thin Emerald Line
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.9), Inches(1.5), Inches(11.533), Inches(0.02))
        line.fill.solid()
        line.fill.fore_color.rgb = COLOR_CARD_BORDER
        line.line.color.rgb = COLOR_CARD_BORDER

    # =========================================================================
    # SLIDE 1: Title Slide (Dark Slate + Emerald Neon)
    # =========================================================================
    slide1 = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide1)

    # Track Badge
    badge = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.9), Inches(0.9), Inches(6.2), Inches(0.45))
    badge.fill.solid()
    badge.fill.fore_color.rgb = COLOR_ACCENT_BG
    badge.line.color.rgb = COLOR_EMERALD
    badge.line.width = Pt(1)
    tf_b = badge.text_frame
    p_b = tf_b.paragraphs[0]
    p_b.text = "⚡ DEEPTECH GIGAHACK 2026 • BIOMENTORHUB × PROFI TRACK"
    p_b.font.name = FONT_MONO
    p_b.font.size = Pt(10)
    p_b.font.bold = True
    p_b.font.color.rgb = COLOR_EMERALD_NEON

    # Main Headline
    tb_h = slide1.shapes.add_textbox(Inches(0.9), Inches(1.6), Inches(11.5), Inches(2.2))
    tf_h = tb_h.text_frame
    tf_h.word_wrap = True
    p1 = tf_h.paragraphs[0]
    p1.text = "PackShift"
    p1.font.name = FONT_TITLE
    p1.font.size = Pt(46)
    p1.font.bold = True
    p1.font.color.rgb = COLOR_EMERALD_NEON

    p2 = tf_h.add_paragraph()
    p2.text = "Циркулярная эко-упаковка для пищевого ритейла"
    p2.font.name = FONT_TITLE
    p2.font.size = Pt(28)
    p2.font.bold = True
    p2.font.color.rgb = COLOR_TEXT_WHITE

    p3 = tf_h.add_paragraph()
    p3.text = "0% первичного пластика • Термостойкость 250°C • Zero PFAS • EU PPWR 2030 Ready"
    p3.font.name = FONT_MONO
    p3.font.size = Pt(13)
    p3.font.color.rgb = COLOR_EMERALD

    # Subtitle Paragraph
    tb_sub = slide1.shapes.add_textbox(Inches(0.9), Inches(3.9), Inches(11.2), Inches(0.85))
    tf_sub = tb_sub.text_frame
    tf_sub.word_wrap = True
    p_s = tf_sub.paragraphs[0]
    p_s.text = (
        "Инженерная система замещения первичного пластика на базе модифицированных моно-материалов rPET "
        "и биомиметических растительных барьеров. Разработано под операционные стандарты сети супермаркетов Profi."
    )
    p_s.font.name = FONT_BODY
    p_s.font.size = Pt(14)
    p_s.font.color.rgb = COLOR_TEXT_MUTED

    # 4 Key Proof Metric Cards
    metrics1 = [
        ("-88%", "Первичный пластик", "Замена на rPET и растительное сырье"),
        ("250°C", "Вызов гриля (Annex)", "6ч в тепловой витрине без размягчения"),
        ("+4 дня", "Срок годности", "Лазерная микроперфорация EMAP (-32% списаний)"),
        ("Class A", "RecyClass & PPWR", "100% моно-материалы в сортировочные потоки"),
    ]
    card_w = Inches(2.7)
    card_h = Inches(1.8)
    card_y = Inches(4.9)
    gap = Inches(0.24)

    for i, (m_val, m_lbl, m_dsc) in enumerate(metrics1):
        cx = Inches(0.9) + i * (card_w + gap)
        card = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, card_y, card_w, card_h)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD_BG
        card.line.color.rgb = COLOR_EMERALD if i == 1 else COLOR_CARD_BORDER
        card.line.width = Pt(1.5 if i == 1 else 1)

        tb_c = slide1.shapes.add_textbox(cx + Inches(0.15), card_y + Inches(0.15), card_w - Inches(0.3), card_h - Inches(0.3))
        tf_c = tb_c.text_frame
        tf_c.word_wrap = True

        p_v = tf_c.paragraphs[0]
        p_v.text = m_val
        p_v.font.name = FONT_MONO
        p_v.font.size = Pt(26)
        p_v.font.bold = True
        p_v.font.color.rgb = COLOR_EMERALD_NEON

        p_l = tf_c.add_paragraph()
        p_l.text = m_lbl
        p_l.font.name = FONT_BODY
        p_l.font.size = Pt(13)
        p_l.font.bold = True
        p_l.font.color.rgb = COLOR_TEXT_WHITE

        p_d = tf_c.add_paragraph()
        p_d.text = m_dsc
        p_d.font.name = FONT_BODY
        p_d.font.size = Pt(10.5)
        p_d.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 2: Problem & Regulatory Cliff
    # =========================================================================
    slide2 = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide2)
    add_top_bar(slide2, "Контекст и Вызов", "Регуляторный капкан ритейла и острая боль витрин", 2)

    # Left Column: Regulatory Cliff
    c_left = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.9), Inches(1.8), Inches(5.6), Inches(5.0))
    c_left.fill.solid()
    c_left.fill.fore_color.rgb = COLOR_CARD_BG
    c_left.line.color.rgb = COLOR_CARD_BORDER
    c_left.line.width = Pt(1)

    tb_l = slide2.shapes.add_textbox(Inches(1.15), Inches(2.0), Inches(5.1), Inches(4.6))
    tf_l = tb_l.text_frame
    tf_l.word_wrap = True
    p = tf_l.paragraphs[0]
    p.text = "⚖️ 1. Регуляторное давление и налоги"
    p.font.name = FONT_TITLE
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = COLOR_EMERALD_NEON

    bullets_l = [
        ("Директива EU PPWR 2030:", "Запрет многослойных неперерабатываемых полимеров. Штрафы ритейлерам за превышение квот первичного пластика."),
        ("Экологический сбор (EPR):", "Ставка налога на комбинированные ламинаты (ПЭ+алюминий) выросла на 45%. Неперерабатываемая упаковка становится прямым убытком."),
        ("Тотальный бан на PFAS:", "Европейский запрет пер- и полифторалкильных гидрофобизаторов в прямом пищевом контакте."),
        ("Аудит Scope 3 (CSRD):", "Сеть Profi обязана публично сокращать углеродный след упаковки по всей цепочке."),
    ]
    for k, v in bullets_l:
        p_item = tf_l.add_paragraph()
        p_item.text = f"• {k} "
        p_item.font.bold = True
        p_item.font.size = Pt(11.5)
        p_item.font.color.rgb = COLOR_TEXT_WHITE
        r = p_item.add_run()
        r.text = v
        r.font.bold = False
        r.font.color.rgb = COLOR_TEXT_MUTED

    # Right Column: The Hot Grill Pain
    c_right = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.8), Inches(5.6), Inches(5.0))
    c_right.fill.solid()
    c_right.fill.fore_color.rgb = COLOR_CARD_BG
    c_right.line.color.rgb = COLOR_EMERALD
    c_right.line.width = Pt(1.5)

    tb_r = slide2.shapes.add_textbox(Inches(7.05), Inches(2.0), Inches(5.1), Inches(4.6))
    tf_r = tb_r.text_frame
    tf_r.word_wrap = True
    p = tf_r.paragraphs[0]
    p.text = "🔥 2. Острая боль: Отдел гриля и кулинарии"
    p.font.name = FONT_TITLE
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = COLOR_EMERALD_NEON

    bullets_r = [
        ("Экстремальный режим:", "Выдержка в витрине при 85°C – 120°C до 6 часов. Пиковые температуры фасовки достигают 180°C – 250°C."),
        ("Дефекты текущих пакетов:", "Полиэтилен с фольгой плавится, дает течь жира и выделяет резкий запах нагретого пластика."),
        ("«Парниковый эффект»:", "Пар блокируется, аппетитная корочка гриля размокает в кашу за 40 минут — потеря лояльности покупателей."),
        ("Скрытые токсины:", "Жиростойкость текущих пакетов держится на PFAS, мигрирующих в горячую пищу."),
    ]
    for k, v in bullets_r:
        p_item = tf_r.add_paragraph()
        p_item.text = f"• {k} "
        p_item.font.bold = True
        p_item.font.size = Pt(11.5)
        p_item.font.color.rgb = COLOR_TEXT_WHITE
        r = p_item.add_run()
        r.text = v
        r.font.bold = False
        r.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 3: The Special Annex Challenge (Grill 180-250°C)
    # =========================================================================
    slide3 = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide3)
    add_top_bar(slide3, "Специальный вызов", "Инженерные решения для Горячего Гриля (180°C – 250°C)", 3)

    # Solution A Card
    c_sa = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.9), Inches(1.8), Inches(5.6), Inches(5.0))
    c_sa.fill.solid()
    c_sa.fill.fore_color.rgb = COLOR_CARD_BG
    c_sa.line.color.rgb = COLOR_EMERALD
    c_sa.line.width = Pt(1.5)

    tb_sa = slide3.shapes.add_textbox(Inches(1.15), Inches(2.0), Inches(5.1), Inches(4.6))
    tf_sa = tb_sa.text_frame
    tf_sa.word_wrap = True

    p = tf_sa.paragraphs[0]
    p.text = "РЕШЕНИЕ А: МОНО-ПАКЕТ"
    p.font.name = FONT_MONO
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_EMERALD

    p_t = tf_sa.add_paragraph()
    p_t.text = "PackShift CPET Ultra-Flex"
    p_t.font.name = FONT_TITLE
    p_t.font.size = Pt(21)
    p_t.font.bold = True
    p_t.font.color.rgb = COLOR_TEXT_WHITE

    specs_a = [
        ("Материал:", "Кристаллический модифицированный rPET (степень кристалличности > 38%)."),
        ("Термостойкость:", "от -40°C до +250°C (Dual-Ovenable: разогрев прямо в упаковке)."),
        ("Тепловая витрина:", "6+ часов при 95°C–120°C без усадки, протечек и запаха."),
        ("Жиростойкость:", "Высший класс Kit Test 12. 100% PFAS-Free."),
        ("Рециклинг:", "100% моно-материал. Поток ПЭТ-бутылок (RecyClass Class A)."),
        ("Себестоимость:", "-14% vs импортный барьерный ламинат с алюминием."),
    ]
    for k, v in specs_a:
        p_item = tf_sa.add_paragraph()
        p_item.text = f"{k} "
        p_item.font.bold = True
        p_item.font.size = Pt(11.5)
        p_item.font.color.rgb = COLOR_EMERALD_NEON
        r = p_item.add_run()
        r.text = v
        r.font.bold = False
        r.font.color.rgb = COLOR_TEXT_MUTED

    # Solution B Card
    c_sb = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.8), Inches(5.6), Inches(5.0))
    c_sb.fill.solid()
    c_sb.fill.fore_color.rgb = COLOR_CARD_BG
    c_sb.line.color.rgb = COLOR_CARD_BORDER
    c_sb.line.width = Pt(1)

    tb_sb = slide3.shapes.add_textbox(Inches(7.05), Inches(2.0), Inches(5.1), Inches(4.6))
    tf_sb = tb_sb.text_frame
    tf_sb.word_wrap = True

    p = tf_sb.paragraphs[0]
    p.text = "РЕШЕНИЕ Б: БИО-КОМПОЗИТ (PROFI BIO)"
    p.font.name = FONT_MONO
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_EMERALD

    p_t = tf_sb.add_paragraph()
    p_t.text = "LignoShield™ ThermoBox"
    p_t.font.name = FONT_TITLE
    p_t.font.size = Pt(21)
    p_t.font.bold = True
    p_t.font.color.rgb = COLOR_TEXT_WHITE

    specs_b = [
        ("Материал:", "100% формованная багасса сахарного тростника + альгинат/хитозан из водорослей."),
        ("Дышащая мембрана:", "Микропористая структура стравливает пар. Корочка гриля остается хрустящей."),
        ("Эко-жиробарьер:", "Водный нано-слой. 0% пластика, 0% фтора, 0% полиэтилена."),
        ("Термостойкость:", "До +220°C кратковременно, 6 часов в тепловой витрине."),
        ("Утилизация:", "Домашний компост за 60 дней (OK Compost HOME) или макулатура."),
        ("Позиционирование:", "Премиальная эко-линейка готовой кулинарии сети Profi."),
    ]
    for k, v in specs_b:
        p_item = tf_sb.add_paragraph()
        p_item.text = f"{k} "
        p_item.font.bold = True
        p_item.font.size = Pt(11.5)
        p_item.font.color.rgb = COLOR_EMERALD_NEON
        r = p_item.add_run()
        r.text = v
        r.font.bold = False
        r.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 4: Store Departments Portfolio (Vegetables, Meat, Bakery)
    # =========================================================================
    slide4 = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide4)
    add_top_bar(slide4, "Масштабирование", "Комплексная оптимизация отделов супермаркета", 4)

    dept_cols = [
        ("🥬 Свежие овощи и зелень", "EMAP Laser-Breathe™", [
            ("Проблема:", "Конденсат, плесень, 28-35% списаний зелени и ягод."),
            ("Инновация:", "Моно-пленка rBOPP с лазерной микроперфорацией. Баланс O₂ (3-5%) и CO₂ (5-8%)."),
            ("Эффект:", "+4 дня к сроку годности на полке."),
            ("Списания:", "Сокращение списаний в супермаркетах на 32%."),
            ("Материал:", "Моно-rBOPP • 100% Recyclable."),
        ]),
        ("🥩 Мясной и рыбный отдел", "Capillary Tray (Pad-Less)", [
            ("Проблема:", "Грязная синтетическая салфетка исключает лоток из рециклинга."),
            ("Инновация:", "Глубокоформованный моно-rPET лоток с капиллярным гидро-замком на 45 мл экссудата."),
            ("Эффект:", "0 синтетических салфеток-вкладышей."),
            ("Сортировка:", "100% чистый поток вторичной переработки."),
            ("Материал:", "85% rPET • RecyClass Class A."),
        ]),
        ("🥐 Пекарня и выпечка", "NatureVent Crisp™", [
            ("Проблема:", "Пакеты со смотровыми окнами из полиэтилена сжигаются."),
            ("Инновация:", "Крафт FSC + прозрачное био-окно NatureFlex™ из древесной массы + водный био-воск."),
            ("Эффект:", "Сохранение хрустящей корочки круассанов."),
            ("Утилизация:", "100% Pulpable (переработка вместе с бумагой)."),
            ("Материал:", "Крафт FSC + целлюлоза • 0% ПЭ."),
        ]),
    ]

    col_w = Inches(3.64)
    col_h = Inches(5.0)
    col_y = Inches(1.8)
    col_gap = Inches(0.3)

    for i, (dept_tag, tech_name, points) in enumerate(dept_cols):
        cx = Inches(0.9) + i * (col_w + col_gap)
        card = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, col_y, col_w, col_h)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD_BG
        card.line.color.rgb = COLOR_CARD_BORDER
        card.line.width = Pt(1)

        tb = slide4.shapes.add_textbox(cx + Inches(0.18), col_y + Inches(0.18), col_w - Inches(0.36), col_h - Inches(0.36))
        tf = tb.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        p.text = dept_tag.upper()
        p.font.name = FONT_MONO
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = COLOR_EMERALD

        p_tn = tf.add_paragraph()
        p_tn.text = tech_name
        p_tn.font.name = FONT_TITLE
        p_tn.font.size = Pt(18)
        p_tn.font.bold = True
        p_tn.font.color.rgb = COLOR_TEXT_WHITE

        for pk, pv in points:
            p_pt = tf.add_paragraph()
            p_pt.text = f"• {pk} "
            p_pt.font.bold = True
            p_pt.font.size = Pt(11)
            p_pt.font.color.rgb = COLOR_EMERALD_NEON
            r = p_pt.add_run()
            r.text = pv
            r.font.bold = False
            r.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 5: Economics & Business Case for Profi (100 Stores Scale)
    # =========================================================================
    slide5 = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide5)
    add_top_bar(slide5, "Бизнес-модель", "Экономика и ESG-эффект: Расчет на 100 супермаркетов Profi", 5)

    calc_metrics = [
        ("1 840 т", "Первичный пластик", "Тонн первичных полимеров предотвращено в год благодаря rPET"),
        ("4 250 т", "Сокращение CO₂e", "Предотвращено выбросов парниковых газов (Scope 3 Profi)"),
        ("€1 280 000", "Экономия на EPR", "Прямая годовая экономия на налоге за счет моно-материалов"),
        ("€700 000", "Снижение списаний", "Сбережение продукции за счет продления свежести (+4 дня)"),
    ]

    cm_w = Inches(2.7)
    cm_h = Inches(3.2)
    cm_y = Inches(1.8)
    cm_gap = Inches(0.24)

    for i, (val, title, sub) in enumerate(calc_metrics):
        cx = Inches(0.9) + i * (cm_w + cm_gap)
        card = slide5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, cm_y, cm_w, cm_h)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD_BG
        card.line.color.rgb = COLOR_EMERALD if i == 2 else COLOR_CARD_BORDER
        card.line.width = Pt(1.5 if i == 2 else 1)

        tb = slide5.shapes.add_textbox(cx + Inches(0.15), cm_y + Inches(0.2), cm_w - Inches(0.3), cm_h - Inches(0.4))
        tf = tb.text_frame
        tf.word_wrap = True

        p_v = tf.paragraphs[0]
        p_v.text = val
        p_v.font.name = FONT_MONO
        p_v.font.size = Pt(26)
        p_v.font.bold = True
        p_v.font.color.rgb = COLOR_EMERALD_NEON

        p_t = tf.add_paragraph()
        p_t.text = title
        p_t.font.name = FONT_BODY
        p_t.font.size = Pt(14)
        p_t.font.bold = True
        p_t.font.color.rgb = COLOR_TEXT_WHITE

        p_s = tf.add_paragraph()
        p_s.text = sub
        p_s.font.name = FONT_BODY
        p_s.font.size = Pt(11)
        p_s.font.color.rgb = COLOR_TEXT_MUTED

    # Bottom CapEx & Compatibility Strip
    c_capex = slide5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.9), Inches(5.3), Inches(11.533), Inches(1.5))
    c_capex.fill.solid()
    c_capex.fill.fore_color.rgb = COLOR_ACCENT_BG
    c_capex.line.color.rgb = COLOR_EMERALD
    c_capex.line.width = Pt(1)

    tb_cx = slide5.shapes.add_textbox(Inches(1.15), Inches(5.45), Inches(11.0), Inches(1.2))
    tf_cx = tb_cx.text_frame
    tf_cx.word_wrap = True

    p_cx1 = tf_cx.paragraphs[0]
    p_cx1.text = "💡 Нулевой CapEx для ритейлера (Drop-In Compatibility):"
    p_cx1.font.name = FONT_TITLE
    p_cx1.font.bold = True
    p_cx1.font.size = Pt(14)
    p_cx1.font.color.rgb = COLOR_EMERALD_NEON

    p_cx2 = tf_cx.add_paragraph()
    p_cx2.text = (
        "Упаковка PackShift полностью адаптирована под стандартные запайщики лотков, фасовочные линии "
        "и тепловые витрины супермаркетов Profi. Закупка нового оборудования не требуется. "
        "Себестоимость моно-пакетов CPET на 14% ниже импортных многослойных фольгированных ламинатов."
    )
    p_cx2.font.name = FONT_BODY
    p_cx2.font.size = Pt(12)
    p_cx2.font.color.rgb = COLOR_TEXT_WHITE

    # =========================================================================
    # SLIDE 6: Standards, Regulatory Compliance & IP
    # =========================================================================
    slide6 = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide6)
    add_top_bar(slide6, "Стандарты и Сертификация", "Регуляторный щит и соответствие нормативам ЕС", 6)

    standards = [
        ("RecyClass", "Grade A Recyclability", "Подтвержденная совместимость с оптической сортировкой (NIR) и механическим рециклингом в ЕС."),
        ("EU PPWR", "PPWR 2030 Ready", "Упреждение европейских норм: >35% пищевого rPET, 0% первичного пластика в фасовочных пакетах."),
        ("Zero Toxic", "100% PFAS-Free", "Сертификация отсутствия пер- и полифторалкильных веществ в контакте с горячим маслом до 250°C."),
        ("Food Safe", "EFSA / FDA Approved", "Соответствие Regulation EU No 10/2011 для прямого контакта с горячими и жирными продуктами."),
    ]

    st_w = Inches(2.7)
    st_h = Inches(5.0)
    st_y = Inches(1.8)
    st_gap = Inches(0.24)

    for i, (badge_txt, st_title, st_desc) in enumerate(standards):
        cx = Inches(0.9) + i * (st_w + st_gap)
        card = slide6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, st_y, st_w, st_h)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD_BG
        card.line.color.rgb = COLOR_CARD_BORDER
        card.line.width = Pt(1)

        tb = slide6.shapes.add_textbox(cx + Inches(0.18), st_y + Inches(0.2), st_w - Inches(0.36), st_h - Inches(0.4))
        tf = tb.text_frame
        tf.word_wrap = True

        p_b = tf.paragraphs[0]
        p_b.text = badge_txt.upper()
        p_b.font.name = FONT_MONO
        p_b.font.size = Pt(11)
        p_b.font.bold = True
        p_b.font.color.rgb = COLOR_EMERALD

        p_t = tf.add_paragraph()
        p_t.text = st_title
        p_t.font.name = FONT_TITLE
        p_t.font.size = Pt(17)
        p_t.font.bold = True
        p_t.font.color.rgb = COLOR_TEXT_WHITE

        p_d = tf.add_paragraph()
        p_d.text = st_desc
        p_d.font.name = FONT_BODY
        p_d.font.size = Pt(12)
        p_d.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 7: Pilot Implementation Roadmap (6 Weeks)
    # =========================================================================
    slide7 = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide7)
    add_top_bar(slide7, "Дорожная карта", "План пилотного запуска в Profi: 6 недель до результата", 7)

    steps = [
        ("ЭТАП 1", "Лабораторный тест", "2 недели", [
            "Тестирование пакетов CPET и термобоксов на витринах гриля Profi при 95°C–120°C.",
            "Проверка герметичности, жиростойкости (Kit 12) и отсутствия посторонних запахов.",
            "Протокол миграции EFSA для контакта с горячим куриным жиром.",
        ]),
        ("ЭТАП 2", "Пилот в 5 магазинах", "4 недели", [
            "Запуск в 5 флагманских супермаркетах Profi в отделах Гриль и Овощи.",
            "Сбор обратной связи от поваров кулинарии и покупателей.",
            "Замер динамики списаний и сохранности хрустящей корочки гриля.",
        ]),
        ("ЭТАП 3", "Масштабирование", "Квартал 2", [
            "Полноформатное развертывание в сети супермаркетов Profi.",
            "Подключение автоматизированного дашборда учета сокращения выбросов Scope 3.",
            "Оптимизация логистики оптовых поставок rPET моно-упаковки.",
        ]),
    ]

    st_w = Inches(3.64)
    st_h = Inches(3.6)
    st_y = Inches(1.8)
    st_gap = Inches(0.3)

    for i, (badge_txt, st_title, st_time, points) in enumerate(steps):
        cx = Inches(0.9) + i * (st_w + st_gap)
        card = slide7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, st_y, st_w, st_h)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD_BG
        card.line.color.rgb = COLOR_EMERALD if i == 0 else COLOR_CARD_BORDER
        card.line.width = Pt(1.5 if i == 0 else 1)

        tb = slide7.shapes.add_textbox(cx + Inches(0.18), st_y + Inches(0.18), st_w - Inches(0.36), st_h - Inches(0.36))
        tf = tb.text_frame
        tf.word_wrap = True

        p_b = tf.paragraphs[0]
        p_b.text = f"{badge_txt} • {st_time}"
        p_b.font.name = FONT_MONO
        p_b.font.size = Pt(11)
        p_b.font.bold = True
        p_b.font.color.rgb = COLOR_EMERALD

        p_t = tf.add_paragraph()
        p_t.text = st_title
        p_t.font.name = FONT_TITLE
        p_t.font.size = Pt(18)
        p_t.font.bold = True
        p_t.font.color.rgb = COLOR_TEXT_WHITE

        for pt in points:
            p_pt = tf.add_paragraph()
            p_pt.text = f"• {pt}"
            p_pt.font.name = FONT_BODY
            p_pt.font.size = Pt(11.5)
            p_pt.font.color.rgb = COLOR_TEXT_MUTED

    # Action Banner (200 samples ready)
    c_smpl = slide7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.9), Inches(5.7), Inches(11.533), Inches(1.2))
    c_smpl.fill.solid()
    c_smpl.fill.fore_color.rgb = COLOR_ACCENT_BG
    c_smpl.line.color.rgb = COLOR_EMERALD_NEON
    c_smpl.line.width = Pt(1.5)

    tb_sm = slide7.shapes.add_textbox(Inches(1.15), Inches(5.8), Inches(11.0), Inches(1.0))
    tf_sm = tb_sm.text_frame
    tf_sm.word_wrap = True

    p_sm1 = tf_sm.paragraphs[0]
    p_sm1.text = "📦 Комплект из 200 образцов готов к передаче технологической службе Profi:"
    p_sm1.font.name = FONT_TITLE
    p_sm1.font.bold = True
    p_sm1.font.size = Pt(14)
    p_sm1.font.color.rgb = COLOR_EMERALD_NEON

    p_sm2 = tf_sm.add_paragraph()
    p_sm2.text = (
        "Партия пакетов CPET Ultra-Flex и термобоксов LignoShield подготовлена для старта "
        "испытаний на тепловых витринах гриля Profi уже на следующей неделе."
    )
    p_sm2.font.name = FONT_BODY
    p_sm2.font.size = Pt(12)
    p_sm2.font.color.rgb = COLOR_TEXT_WHITE

    # =========================================================================
    # SLIDE 8: Conclusion & Call to Action (Pure Deeptech Dark Deck)
    # =========================================================================
    slide8 = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide8)

    c_box = slide8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.5), Inches(1.2), Inches(10.333), Inches(5.1))
    c_box.fill.solid()
    c_box.fill.fore_color.rgb = COLOR_CARD_BG
    c_box.line.color.rgb = COLOR_EMERALD
    c_box.line.width = Pt(2)

    tb_f = slide8.shapes.add_textbox(Inches(1.8), Inches(1.6), Inches(9.733), Inches(4.3))
    tf_f = tb_f.text_frame
    tf_f.word_wrap = True

    p_f0 = tf_f.paragraphs[0]
    p_f0.alignment = PP_ALIGN.CENTER
    p_f0.text = "DEEPTECH GIGAHACK 2026 • BIOMENTORHUB × PROFI"
    p_f0.font.name = FONT_MONO
    p_f0.font.size = Pt(12)
    p_f0.font.bold = True
    p_f0.font.color.rgb = COLOR_EMERALD_NEON

    p_f1 = tf_f.add_paragraph()
    p_f1.alignment = PP_ALIGN.CENTER
    p_f1.text = "PackShift"
    p_f1.font.name = FONT_TITLE
    p_f1.font.size = Pt(38)
    p_f1.font.bold = True
    p_f1.font.color.rgb = COLOR_TEXT_WHITE

    p_f2 = tf_f.add_paragraph()
    p_f2.alignment = PP_ALIGN.CENTER
    p_f2.text = "Делаем циркулярность прибыльной для Profi уже сегодня"
    p_f2.font.name = FONT_TITLE
    p_f2.font.size = Pt(20)
    p_f2.font.bold = True
    p_f2.font.color.rgb = COLOR_EMERALD_NEON

    p_f3 = tf_f.add_paragraph()
    p_f3.alignment = PP_ALIGN.CENTER
    p_f3.text = "\n0% первичного пластика • -14% себестоимость • Нулевой CapEx • Защита маржинальности Profi"
    p_f3.font.name = FONT_BODY
    p_f3.font.size = Pt(13)
    p_f3.font.color.rgb = COLOR_TEXT_MUTED

    p_f4 = tf_f.add_paragraph()
    p_f4.alignment = PP_ALIGN.CENTER
    p_f4.text = "\nБлагодарим за внимание! Готовы ответить на ваши вопросы."
    p_f4.font.name = FONT_TITLE
    p_f4.font.size = Pt(17)
    p_f4.font.bold = True
    p_f4.font.color.rgb = COLOR_TEXT_WHITE

    output_path = os.path.join(os.path.dirname(__file__), "PackShift_Profi_Gigahack2026.pptx")
    prs.save(output_path)
    print(f"Deeptech dark deck saved to: {output_path}")

if __name__ == "__main__":
    create_deck()

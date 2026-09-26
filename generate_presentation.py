import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

# 1. Colors & Design Palette (Exact PackShift Website Styling: Butter + Forest Pine + Emerald)
COLOR_BG_BUTTER = RGBColor(0xFA, 0xF4, 0xE4)       # #FAF4E4 (Butter Yellow Canvas)
COLOR_FOREST_DEEP = RGBColor(0x12, 0x35, 0x24)     # #123524 (Deep Forest Pine)
COLOR_FOREST_CARD = RGBColor(0x15, 0x3E, 0x2B)     # #153E2B (Forest Card Surface)
COLOR_LEAF_EMERALD = RGBColor(0x15, 0x80, 0x3D)    # #15803D (Leaf Emerald Accent)
COLOR_LEAF_VIBRANT = RGBColor(0x22, 0xC5, 0x5E)    # #22C55E (Vibrant Leaf Neon)
COLOR_TEXT_WHITE = RGBColor(0xFF, 0xFF, 0xFF)      # #FFFFFF (White)
COLOR_TEXT_MUTED = RGBColor(0x52, 0x6B, 0x5C)      # #526B5C (Forest Muted)
COLOR_CARD_LIGHT = RGBColor(0xFF, 0xFF, 0xFF)      # White Card on Butter Canvas
COLOR_BADGE_BG = RGBColor(0xFA, 0xF4, 0xE4)        # Translucent Badge Color
COLOR_BORDER_LIGHT = RGBColor(0xD8, 0xE2, 0xDC)    # Soft Card Border

FONT_TITLE = "Plus Jakarta Sans"
FONT_BODY = "Plus Jakarta Sans"
FONT_MONO = "JetBrains Mono"

def create_deck():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    logo_path = os.path.join(os.path.dirname(__file__), "Проэкт Хакатон", "img", "logo.png")

    def set_slide_background(slide, color=COLOR_BG_BUTTER):
        bg_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg_shape.fill.solid()
        bg_shape.fill.fore_color.rgb = color
        bg_shape.line.color.rgb = color
        return bg_shape

    def add_site_header(slide, title_text, tag_text="DEEPTECH GIGAHACK 2026 • BIOMENTORHUB × PROFI"):
        # Top Floating Navbar Bar (matching website floating nav)
        nav_bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.35), Inches(11.733), Inches(0.95))
        nav_bar.fill.solid()
        nav_bar.fill.fore_color.rgb = COLOR_FOREST_DEEP
        nav_bar.line.color.rgb = COLOR_LEAF_VIBRANT
        nav_bar.line.width = Pt(1)

        # Frosted glass logo plate inside navbar
        if os.path.exists(logo_path):
            plate = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(0.47), Inches(2.2), Inches(0.7))
            plate.fill.solid()
            plate.fill.fore_color.rgb = COLOR_BADGE_BG
            plate.line.color.rgb = COLOR_LEAF_VIBRANT
            plate.line.width = Pt(1)
            slide.shapes.add_picture(logo_path, Inches(1.15), Inches(0.53), width=Inches(1.9))

        # Header Title inside Navbar
        tb_nav = slide.shapes.add_textbox(Inches(3.5), Inches(0.4), Inches(8.8), Inches(0.85))
        tf_nav = tb_nav.text_frame
        tf_nav.word_wrap = True

        p_tag = tf_nav.paragraphs[0]
        p_tag.text = tag_text.upper()
        p_tag.font.name = FONT_BODY
        p_tag.font.size = Pt(10)
        p_tag.font.bold = True
        p_tag.font.color.rgb = COLOR_LEAF_VIBRANT

        p_tit = tf_nav.add_paragraph()
        p_tit.text = title_text
        p_tit.font.name = FONT_TITLE
        p_tit.font.size = Pt(18)
        p_tit.font.bold = True
        p_tit.font.color.rgb = COLOR_TEXT_WHITE

    # =========================================================================
    # SLIDE 1: Cover Slide (Exact Website Hero Section)
    # =========================================================================
    slide1 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide1, COLOR_BG_BUTTER)

    # Top floating navbar with logo
    add_site_header(slide1, "Революция пищевой упаковки ритейла: 0% первичного пластика", "DEEPTECH GIGAHACK 2026 • BIOMENTORHUB × PROFI")

    # Main Headline & Subtitle
    tb_hero = slide1.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(11.733), Inches(2.7))
    tf_h = tb_hero.text_frame
    tf_h.word_wrap = True

    p_badge = tf_h.paragraphs[0]
    p_badge.text = "🌱 ТРЕК BIOMENTORHUB × PROFI • ХАКАТОН 2026"
    p_badge.font.name = FONT_BODY
    p_badge.font.size = Pt(12)
    p_badge.font.bold = True
    p_badge.font.color.rgb = COLOR_LEAF_EMERALD

    p_h1 = tf_h.add_paragraph()
    p_h1.text = "PackShift: 0% первичного пластика, 100% циркулярность"
    p_h1.font.name = FONT_TITLE
    p_h1.font.size = Pt(32)
    p_h1.font.bold = True
    p_h1.font.color.rgb = COLOR_FOREST_DEEP

    p_h2 = tf_h.add_paragraph()
    p_h2.text = (
        "Инженерные решения для сети супермаркетов Profi: биомиметические барьеры против жира, "
        "термостойкость до 250°C без токсичных PFAS, продление свежести продуктов и прямое снижение эко-сбора."
    )
    p_h2.font.name = FONT_BODY
    p_h2.font.size = Pt(14)
    p_h2.font.color.rgb = COLOR_TEXT_MUTED

    # 4 Proof Cards (Identical to website hero-proof-grid)
    metrics = [
        ("-88%", "Первичный пластик", "Замена на rPET и биополимеры"),
        ("250°C", "Термостойкость гриля", "До 6ч в витрине без размягчения"),
        ("+4 дня", "Свежесть продуктов", "Лазерная микроперфорация EMAP"),
        ("Class A", "RecyClass & PPWR", "100% моно-материалы в сортировку"),
    ]
    card_w = Inches(2.75)
    card_h = Inches(2.5)
    card_y = Inches(4.5)
    card_gap = Inches(0.24)

    for i, (val, label, desc) in enumerate(metrics):
        cx = Inches(0.8) + i * (card_w + card_gap)
        card = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, card_y, card_w, card_h)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_FOREST_DEEP if i == 1 else COLOR_CARD_LIGHT
        card.line.color.rgb = COLOR_LEAF_VIBRANT if i == 1 else COLOR_BORDER_LIGHT
        card.line.width = Pt(1.5 if i == 1 else 1)

        tb_m = slide1.shapes.add_textbox(cx + Inches(0.15), card_y + Inches(0.2), card_w - Inches(0.3), card_h - Inches(0.4))
        tf_m = tb_m.text_frame
        tf_m.word_wrap = True

        p_v = tf_m.paragraphs[0]
        p_v.text = val
        p_v.font.name = FONT_MONO
        p_v.font.size = Pt(30)
        p_v.font.bold = True
        p_v.font.color.rgb = COLOR_LEAF_VIBRANT if i == 1 else COLOR_LEAF_EMERALD

        p_l = tf_m.add_paragraph()
        p_l.text = label
        p_l.font.name = FONT_BODY
        p_l.font.size = Pt(14)
        p_l.font.bold = True
        p_l.font.color.rgb = COLOR_TEXT_WHITE if i == 1 else COLOR_FOREST_DEEP

        p_d = tf_m.add_paragraph()
        p_d.text = desc
        p_d.font.name = FONT_BODY
        p_d.font.size = Pt(11)
        p_d.font.color.rgb = COLOR_BG_BUTTER if i == 1 else COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 2: Problem & Regulatory Challenge (Site styling)
    # =========================================================================
    slide2 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide2, COLOR_BG_BUTTER)
    add_site_header(slide2, "Двойной вызов ритейла: Регуляторный капкан и боль витрины", "ПРОБЛЕМА И КОНТЕКСТ")

    # Left Card
    left_card = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.6), Inches(5.7), Inches(5.4))
    left_card.fill.solid()
    left_card.fill.fore_color.rgb = COLOR_CARD_LIGHT
    left_card.line.color.rgb = COLOR_BORDER_LIGHT
    left_card.line.width = Pt(1)

    tb_lc = slide2.shapes.add_textbox(Inches(1.1), Inches(1.8), Inches(5.1), Inches(5.0))
    tf_lc = tb_lc.text_frame
    tf_lc.word_wrap = True
    p = tf_lc.paragraphs[0]
    p.text = "⚖️ 1. Регуляторное давление на Profi"
    p.font.name = FONT_TITLE
    p.font.size = Pt(19)
    p.font.bold = True
    p.font.color.rgb = COLOR_FOREST_DEEP

    bullets_l = [
        ("Директива EU PPWR 2030:", "Запрет многослойных неперерабатываемых полимеров. Штрафы ритейлерам за превышение квот первичного пластика."),
        ("Экологический сбор (EPR Tax):", "Рост ставки налога на комбинированные ламинаты (ПЭ+фольга) на 45%. Трудная упаковка становится прямым убытком."),
        ("Запрет PFAS (вечные токсины):", "Европейские регуляторы вводят полный бан на фторированные гидрофобизаторы в пищевом контакте."),
        ("CSRD аудит (Scope 3):", "Сеть Profi обязана публично декларировать и сокращать углеродный след упаковки по всей цепочке."),
    ]
    for b_title, b_desc in bullets_l:
        p_b = tf_lc.add_paragraph()
        p_b.text = f"• {b_title} "
        p_b.font.bold = True
        p_b.font.size = Pt(12)
        p_b.font.color.rgb = COLOR_FOREST_DEEP
        r = p_b.add_run()
        r.text = b_desc
        r.font.bold = False
        r.font.color.rgb = COLOR_TEXT_MUTED

    # Right Card (Forest Deep Card)
    right_card = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.6), Inches(5.7), Inches(5.4))
    right_card.fill.solid()
    right_card.fill.fore_color.rgb = COLOR_FOREST_DEEP
    right_card.line.color.rgb = COLOR_LEAF_VIBRANT
    right_card.line.width = Pt(1.5)

    tb_rc = slide2.shapes.add_textbox(Inches(7.1), Inches(1.8), Inches(5.1), Inches(5.0))
    tf_rc = tb_rc.text_frame
    tf_rc.word_wrap = True
    p = tf_rc.paragraphs[0]
    p.text = "🔥 2. Острая боль: Отдел гриля Profi"
    p.font.name = FONT_TITLE
    p.font.size = Pt(19)
    p.font.bold = True
    p.font.color.rgb = COLOR_TEXT_WHITE

    bullets_r = [
        ("Экстремальный режим:", "Выдержка в витрине при 85°C – 120°C до 6 часов. Пиковые температуры фасовки — до 250°C."),
        ("Дефекты текущей упаковки:", "Полиэтиленовые пакеты с фольгой размягчаются, текут горячим жиром и выделяют запах нагретой синтетики."),
        ("Эффект «парника»:", "Пар блокируется внутри пакета, размачивая хрустящую корочку гриля в кашу за 40 минут."),
        ("Скрытая токсичность:", "Жиростойкость текущих упаковок держится на PFAS, мигрирующих в горячую пищу."),
    ]
    for b_title, b_desc in bullets_r:
        p_b = tf_rc.add_paragraph()
        p_b.text = f"• {b_title} "
        p_b.font.bold = True
        p_b.font.size = Pt(12)
        p_b.font.color.rgb = COLOR_LEAF_VIBRANT
        r = p_b.add_run()
        r.text = b_desc
        r.font.bold = False
        r.font.color.rgb = COLOR_BG_BUTTER

    # =========================================================================
    # SLIDE 3: Hot Grill Annex (Comparison Cards)
    # =========================================================================
    slide3 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide3, COLOR_BG_BUTTER)
    add_site_header(slide3, "Готовая кулинария: Упаковка гриля высокой температуры", "СПЕЦИАЛЬНЫЙ ВЫЗОВ ХАКАТОНА")

    # Card A
    card_a = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.6), Inches(5.7), Inches(5.4))
    card_a.fill.solid()
    card_a.fill.fore_color.rgb = COLOR_CARD_LIGHT
    card_a.line.color.rgb = COLOR_BORDER_LIGHT
    card_a.line.width = Pt(1)

    tb_a = slide3.shapes.add_textbox(Inches(1.1), Inches(1.8), Inches(5.1), Inches(5.0))
    tf_a = tb_a.text_frame
    tf_a.word_wrap = True

    p = tf_a.paragraphs[0]
    p.text = "РЕШЕНИЕ А • МОНО-ПАКЕТ"
    p.font.name = FONT_BODY
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_LEAF_EMERALD

    p_t = tf_a.add_paragraph()
    p_t.text = "PackShift CPET Ultra-Flex"
    p_t.font.name = FONT_TITLE
    p_t.font.size = Pt(21)
    p_t.font.bold = True
    p_t.font.color.rgb = COLOR_FOREST_DEEP

    specs_a = [
        ("Материал:", "Кристаллический модифицированный rPET (кристалличность > 38%)."),
        ("Термостойкость:", "от -40°C до +250°C (Dual-Ovenable: разогрев прямо в упаковке)."),
        ("Витрина гриля:", "6+ часов при 95°C–120°C без усадки, запаха и деформации."),
        ("Жиростойкость:", "Высший класс Kit Test 12. 100% PFAS-Free."),
        ("Циркулярность:", "100% моно-материал. Поток ПЭТ-бутылок (RecyClass Class A)."),
        ("Экономика:", "Себестоимость на 14% ниже импортных фольгированных ламинатов."),
    ]
    for k, v in specs_a:
        p_sp = tf_a.add_paragraph()
        p_sp.text = f"{k} "
        p_sp.font.bold = True
        p_sp.font.size = Pt(12)
        p_sp.font.color.rgb = COLOR_FOREST_DEEP
        r = p_sp.add_run()
        r.text = v
        r.font.bold = False
        r.font.color.rgb = COLOR_TEXT_MUTED

    # Card B
    card_b = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.6), Inches(5.7), Inches(5.4))
    card_b.fill.solid()
    card_b.fill.fore_color.rgb = COLOR_FOREST_DEEP
    card_b.line.color.rgb = COLOR_LEAF_VIBRANT
    card_b.line.width = Pt(1.5)

    tb_b = slide3.shapes.add_textbox(Inches(7.1), Inches(1.8), Inches(5.1), Inches(5.0))
    tf_b = tb_b.text_frame
    tf_b.word_wrap = True

    p = tf_b.paragraphs[0]
    p.text = "РЕШЕНИЕ Б • БИО-КОМПОЗИТ (PROFI BIO)"
    p.font.name = FONT_BODY
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_LEAF_VIBRANT

    p_t = tf_b.add_paragraph()
    p_t.text = "LignoShield™ ThermoBox"
    p_t.font.name = FONT_TITLE
    p_t.font.size = Pt(21)
    p_t.font.bold = True
    p_t.font.color.rgb = COLOR_TEXT_WHITE

    specs_b = [
        ("Материал:", "100% формованная багасса сахарного тростника + био-барьер из водорослей."),
        ("Дышащая мембрана:", "Стравливает лишний пар, сохраняя корочку гриля хрустящей."),
        ("Жиробарьер:", "Альгинат натрия + хитозан на водной основе. 0% пластика, 0% фтора."),
        ("Термостойкость:", "До +220°C кратковременно, 6 часов в тепловой витрине."),
        ("Утилизация:", "Домашний компост за 60 дней (OK Compost HOME) или макулатура."),
        ("Позиционирование:", "Идеально для премиальной эко-кулинарии сети Profi."),
    ]
    for k, v in specs_b:
        p_sp = tf_b.add_paragraph()
        p_sp.text = f"{k} "
        p_sp.font.bold = True
        p_sp.font.size = Pt(12)
        p_sp.font.color.rgb = COLOR_LEAF_VIBRANT
        r = p_sp.add_run()
        r.text = v
        r.font.bold = False
        r.font.color.rgb = COLOR_BG_BUTTER

    # =========================================================================
    # SLIDE 4: 3 Store Departments (Site styling)
    # =========================================================================
    slide4 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide4, COLOR_BG_BUTTER)
    add_site_header(slide4, "Комплексная оптимизация отделов супермаркета", "КАТАЛОГ РЕШЕНИЙ ПО ОТДЕЛАМ")

    dept_cols = [
        ("🥬 Свежие овощи и зелень", "EMAP Laser-Breathe™", [
            ("Проблема:", "Конденсат, плесень, 28-35% списаний зелени и ягод."),
            ("Технология:", "Моно-пленка rBOPP с лазерной микроперфорацией. Баланс O₂ (3-5%) и CO₂ (5-8%)."),
            ("Эффект:", "+4 дня срока годности на полке."),
            ("Результат:", "Сокращение списаний в супермаркетах на 32%."),
            ("Материал:", "Моно-rBOPP • 100% Recyclable."),
        ]),
        ("🥩 Мясной и рыбный отдел", "Capillary Tray (Pad-Less)", [
            ("Проблема:", "Грязная салфетка-вкладыш исключает лоток из рециклинга."),
            ("Технология:", "Глубокоформованный моно-rPET лоток с капиллярным гидро-замком на 45 мл экссудата."),
            ("Эффект:", "0 синтетических салфеток-вкладышей."),
            ("Результат:", "100% чистый поток вторичной переработки."),
            ("Материал:", "85% rPET • RecyClass Class A."),
        ]),
        ("🥐 Пекарня и выпечка", "NatureVent Crisp™", [
            ("Проблема:", "Пакеты с пластиковыми окнами не перерабатываются."),
            ("Технология:", "Крафт FSC + прозрачное смотровое окно NatureFlex™ из древесной массы + био-воск."),
            ("Эффект:", "Сохранение хрустящей корочки круассанов."),
            ("Результат:", "100% Pulpable (переработка с бумагой)."),
            ("Материал:", "Крафт FSC + целлюлоза • 0% ПЭ."),
        ]),
    ]

    col_w = Inches(3.7)
    col_h = Inches(5.4)
    col_y = Inches(1.6)
    col_gap = Inches(0.3)

    for i, (dept_tag, tech_name, points) in enumerate(dept_cols):
        cx = Inches(0.8) + i * (col_w + col_gap)
        card = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, col_y, col_w, col_h)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD_LIGHT
        card.line.color.rgb = COLOR_BORDER_LIGHT
        card.line.width = Pt(1)

        tb = slide4.shapes.add_textbox(cx + Inches(0.2), col_y + Inches(0.2), col_w - Inches(0.4), col_h - Inches(0.4))
        tf = tb.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        p.text = dept_tag.upper()
        p.font.name = FONT_BODY
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = COLOR_LEAF_EMERALD

        p_tn = tf.add_paragraph()
        p_tn.text = tech_name
        p_tn.font.name = FONT_TITLE
        p_tn.font.size = Pt(18)
        p_tn.font.bold = True
        p_tn.font.color.rgb = COLOR_FOREST_DEEP

        for pk, pv in points:
            p_pt = tf.add_paragraph()
            p_pt.text = f"• {pk} "
            p_pt.font.bold = True
            p_pt.font.size = Pt(11.5)
            p_pt.font.color.rgb = COLOR_FOREST_DEEP
            r = p_pt.add_run()
            r.text = pv
            r.font.bold = False
            r.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 5: Calculator & Economics (Exact Website Calculator Card)
    # =========================================================================
    slide5 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide5, COLOR_BG_BUTTER)
    add_site_header(slide5, "Калькулятор ESG-эффекта и экономии для сети Profi", "ИНТЕРАКТИВНАЯ АНАЛИТИКА")

    # Big Deep Forest Green Box (identical to website calculator card)
    calc_box = slide5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.6), Inches(11.733), Inches(5.4))
    calc_box.fill.solid()
    calc_box.fill.fore_color.rgb = COLOR_FOREST_DEEP
    calc_box.line.color.rgb = COLOR_LEAF_VIBRANT
    calc_box.line.width = Pt(1.5)

    tb_cb = slide5.shapes.add_textbox(Inches(1.1), Inches(1.8), Inches(11.1), Inches(0.8))
    tf_cb = tb_cb.text_frame
    tf_cb.word_wrap = True
    p = tf_cb.paragraphs[0]
    p.text = "🎯 Моделирование эффекта на масштабе 100 супермаркетов Profi:"
    p.font.name = FONT_TITLE
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = COLOR_TEXT_WHITE

    calc_metrics = [
        ("1 840 т", "Первичный пластик", "Тонн первичных полимеров предотвращено в год благодаря rPET"),
        ("4 250 т", "Сокращение CO₂e", "Предотвращено выбросов парниковых газов (Scope 3 Profi)"),
        ("€1 280 000", "Экономия на EPR", "Прямая годовая экономия на налоге за счет моно-материалов"),
        ("€700 000", "Снижение списаний", "Сбережение продукции за счет продления свежести (+4 дня)"),
    ]

    cm_w = Inches(2.6)
    cm_h = Inches(2.7)
    cm_y = Inches(2.7)
    cm_gap = Inches(0.24)

    for i, (val, title, sub) in enumerate(calc_metrics):
        cx = Inches(1.1) + i * (cm_w + cm_gap)
        c_card = slide5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, cm_y, cm_w, cm_h)
        c_card.fill.solid()
        c_card.fill.fore_color.rgb = COLOR_FOREST_CARD
        c_card.line.color.rgb = COLOR_LEAF_VIBRANT if i == 2 else COLOR_LEAF_EMERALD
        c_card.line.width = Pt(2 if i == 2 else 1)

        tb_m = slide5.shapes.add_textbox(cx + Inches(0.12), cm_y + Inches(0.15), cm_w - Inches(0.24), cm_h - Inches(0.3))
        tf_m = tb_m.text_frame
        tf_m.word_wrap = True

        p_v = tf_m.paragraphs[0]
        p_v.text = val
        p_v.font.name = FONT_MONO
        p_v.font.size = Pt(24)
        p_v.font.bold = True
        p_v.font.color.rgb = COLOR_LEAF_VIBRANT

        p_t = tf_m.add_paragraph()
        p_t.text = title
        p_t.font.name = FONT_BODY
        p_t.font.size = Pt(13)
        p_t.font.bold = True
        p_t.font.color.rgb = COLOR_TEXT_WHITE

        p_s = tf_m.add_paragraph()
        p_s.text = sub
        p_s.font.name = FONT_BODY
        p_s.font.size = Pt(10.5)
        p_s.font.color.rgb = COLOR_BG_BUTTER

    # Bottom CapEx Strip
    tb_bot = slide5.shapes.add_textbox(Inches(1.1), Inches(5.7), Inches(11.1), Inches(1.0))
    tf_bot = tb_bot.text_frame
    tf_bot.word_wrap = True
    p_bot = tf_bot.paragraphs[0]
    p_bot.text = "💡 Нулевой CapEx для ритейлера: "
    p_bot.font.bold = True
    p_bot.font.size = Pt(13)
    p_bot.font.color.rgb = COLOR_LEAF_VIBRANT
    r_bot = p_bot.add_run()
    r_bot.text = "Упаковка PackShift полностью адаптирована под запайщики, весы и витрины сети Profi. Себестоимость моно-пакетов CPET на 14% ниже импортных фольгированных ламинатов."
    r_bot.font.bold = False
    r_bot.font.color.rgb = COLOR_TEXT_WHITE

    # =========================================================================
    # SLIDE 6: Standards & Certification (Site styling)
    # =========================================================================
    slide6 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide6, COLOR_BG_BUTTER)
    add_site_header(slide6, "Соответствие европейским и национальным регуляциям", "МЕЖДУНАРОДНАЯ СЕРТИФИКАЦИЯ")

    standards = [
        ("RecyClass", "Grade A Recyclability", "Подтвержденная совместимость с оптическими сепараторами и линиями вторичной переработки в ЕС."),
        ("EU PPWR", "PPWR 2030 Ready", "Упреждение европейских норм: >35% пищевого rPET, 0% первичного пластика в одноразовых пакетах."),
        ("Zero Toxic", "100% PFAS-Free", "Сертификация отсутствия пер- и полифторалкильных веществ в контакте с горячим маслом до 250°C."),
        ("Food Safe", "EFSA / FDA Approved", "Соответствие Regulation EU No 10/2011 для прямого контакта с горячими, кислыми и жирными продуктами."),
    ]

    st_w = Inches(2.75)
    st_h = Inches(5.4)
    st_y = Inches(1.6)
    st_gap = Inches(0.24)

    for i, (badge_txt, st_title, st_desc) in enumerate(standards):
        cx = Inches(0.8) + i * (st_w + st_gap)
        card = slide6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, st_y, st_w, st_h)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD_LIGHT
        card.line.color.rgb = COLOR_BORDER_LIGHT
        card.line.width = Pt(1)

        tb = slide6.shapes.add_textbox(cx + Inches(0.18), st_y + Inches(0.2), st_w - Inches(0.36), st_h - Inches(0.4))
        tf = tb.text_frame
        tf.word_wrap = True

        p_b = tf.paragraphs[0]
        p_b.text = badge_txt.upper()
        p_b.font.name = FONT_MONO
        p_b.font.size = Pt(11)
        p_b.font.bold = True
        p_b.font.color.rgb = COLOR_LEAF_EMERALD

        p_t = tf.add_paragraph()
        p_t.text = st_title
        p_t.font.name = FONT_TITLE
        p_t.font.size = Pt(17)
        p_t.font.bold = True
        p_t.font.color.rgb = COLOR_FOREST_DEEP

        p_d = tf.add_paragraph()
        p_d.text = st_desc
        p_d.font.name = FONT_BODY
        p_d.font.size = Pt(12)
        p_d.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 7: Pilot Roadmap (Site styling)
    # =========================================================================
    slide7 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide7, COLOR_BG_BUTTER)
    add_site_header(slide7, "План пилотного запуска в супермаркетах Profi", "ДОРОЖНАЯ КАРТА ВНЕДРЕНИЯ")

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
        ("ЭТАП 3", "Масштабирование на сеть", "Квартал 2", [
            "Полноформатное развертывание в сети магазинов Profi.",
            "Подключение автоматизированного дашборда учета сокращения выбросов Scope 3.",
            "Оптимизация логистики поставок моно-упаковки.",
        ]),
    ]

    st_w = Inches(3.7)
    st_h = Inches(3.8)
    st_y = Inches(1.6)
    st_gap = Inches(0.3)

    for i, (badge_txt, st_title, st_time, points) in enumerate(steps):
        cx = Inches(0.8) + i * (st_w + st_gap)
        card = slide7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, st_y, st_w, st_h)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD_LIGHT
        card.line.color.rgb = COLOR_LEAF_EMERALD if i == 0 else COLOR_BORDER_LIGHT
        card.line.width = Pt(1.5 if i == 0 else 1)

        tb = slide7.shapes.add_textbox(cx + Inches(0.18), st_y + Inches(0.18), st_w - Inches(0.36), st_h - Inches(0.36))
        tf = tb.text_frame
        tf.word_wrap = True

        p_b = tf.paragraphs[0]
        p_b.text = f"{badge_txt} • {st_time}"
        p_b.font.name = FONT_BODY
        p_b.font.size = Pt(11)
        p_b.font.bold = True
        p_b.font.color.rgb = COLOR_LEAF_EMERALD

        p_t = tf.add_paragraph()
        p_t.text = st_title
        p_t.font.name = FONT_TITLE
        p_t.font.size = Pt(18)
        p_t.font.bold = True
        p_t.font.color.rgb = COLOR_FOREST_DEEP

        for pt in points:
            p_pt = tf.add_paragraph()
            p_pt.text = f"• {pt}"
            p_pt.font.name = FONT_BODY
            p_pt.font.size = Pt(11.5)
            p_pt.font.color.rgb = COLOR_TEXT_MUTED

    # Action Banner (like site action-banner)
    banner = slide7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(5.7), Inches(11.733), Inches(1.3))
    banner.fill.solid()
    banner.fill.fore_color.rgb = COLOR_FOREST_DEEP
    banner.line.color.rgb = COLOR_LEAF_VIBRANT
    banner.line.width = Pt(1.5)
    tf_bn = banner.text_frame
    tf_bn.word_wrap = True
    p_bn = tf_bn.paragraphs[0]
    p_bn.text = "📦 Комплект из 200 тестовых образцов готов к передаче!"
    p_bn.font.bold = True
    p_bn.font.size = Pt(15)
    p_bn.font.color.rgb = COLOR_LEAF_VIBRANT

    r_bn = p_bn.add_run()
    r_bn.text = " Партия пакетов CPET Ultra-Flex и термобоксов LignoShield подготовлена для старта испытаний в сети Profi на следующей неделе."
    r_bn.font.bold = False
    r_bn.font.color.rgb = COLOR_TEXT_WHITE

    # =========================================================================
    # SLIDE 8: Team, Mission & Q&A
    # =========================================================================
    slide8 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide8, COLOR_BG_BUTTER)
    add_site_header(slide8, "Делаем циркулярность прибыльной для Profi уже сегодня", "КОМАНДА И КОНТАКТЫ")

    # Center Card
    center_box = slide8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(2.2), Inches(1.8), Inches(8.933), Inches(5.1))
    center_box.fill.solid()
    center_box.fill.fore_color.rgb = COLOR_CARD_LIGHT
    center_box.line.color.rgb = COLOR_LEAF_VIBRANT
    center_box.line.width = Pt(2)

    tb_cnt = slide8.shapes.add_textbox(Inches(2.5), Inches(2.1), Inches(8.333), Inches(4.5))
    tf_cnt = tb_cnt.text_frame
    tf_cnt.word_wrap = True

    p_c1 = tf_cnt.paragraphs[0]
    p_c1.alignment = PP_ALIGN.CENTER
    p_c1.text = "PackShift (Profi EcoPack)"
    p_c1.font.name = FONT_TITLE
    p_c1.font.size = Pt(30)
    p_c1.font.bold = True
    p_c1.font.color.rgb = COLOR_FOREST_DEEP

    p_c2 = tf_cnt.add_paragraph()
    p_c2.alignment = PP_ALIGN.CENTER
    p_c2.text = "0% первичного пластика • Защита маржинальности • Соответствие PPWR 2030"
    p_c2.font.name = FONT_TITLE
    p_c2.font.size = Pt(17)
    p_c2.font.bold = True
    p_c2.font.color.rgb = COLOR_LEAF_EMERALD

    p_c3 = tf_cnt.add_paragraph()
    p_c3.alignment = PP_ALIGN.CENTER
    p_c3.text = "\n🌐 Интерактивная веб-платформа и ESG-калькулятор сети Profi:"
    p_c3.font.size = Pt(14)
    p_c3.font.color.rgb = COLOR_TEXT_MUTED

    p_c4 = tf_cnt.add_paragraph()
    p_c4.alignment = PP_ALIGN.CENTER
    p_c4.text = "https://lisqrn.github.io/Hackaton-202636/"
    p_c4.font.name = FONT_MONO
    p_c4.font.size = Pt(20)
    p_c4.font.bold = True
    p_c4.font.color.rgb = COLOR_LEAF_EMERALD

    p_c5 = tf_cnt.add_paragraph()
    p_c5.alignment = PP_ALIGN.CENTER
    p_c5.text = "\nDeeptech Gigahack 2026 • Biomentorhub × Profi Track"
    p_c5.font.size = Pt(13)
    p_c5.font.color.rgb = COLOR_TEXT_MUTED

    p_c6 = tf_cnt.add_paragraph()
    p_c6.alignment = PP_ALIGN.CENTER
    p_c6.text = "\nБлагодарим за внимание! Готовы ответить на ваши вопросы."
    p_c6.font.name = FONT_TITLE
    p_c6.font.size = Pt(16)
    p_c6.font.bold = True
    p_c6.font.color.rgb = COLOR_FOREST_DEEP

    output_path = os.path.join(os.path.dirname(__file__), "PackShift_Profi_Gigahack2026.pptx")
    prs.save(output_path)
    print(f"Updated presentation saved to: {output_path}")

if __name__ == "__main__":
    create_deck()

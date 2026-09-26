import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

# 1. Colors & Design Palette (Butter, Forest Pine, Emerald - No brown)
COLOR_BG_BUTTER = RGBColor(0xFA, 0xF4, 0xE4)       # #FAF4E4 (Butter Yellow Canvas)
COLOR_FOREST_DEEP = RGBColor(0x12, 0x35, 0x24)     # #123524 (Deep Forest Pine)
COLOR_FOREST_CARD = RGBColor(0x15, 0x3E, 0x2B)     # #153E2B (Forest Card Background)
COLOR_LEAF_EMERALD = RGBColor(0x15, 0x80, 0x3D)    # #15803D (Leaf Emerald Accent)
COLOR_LEAF_VIBRANT = RGBColor(0x22, 0xC5, 0x5E)    # #22C55E (Vibrant Leaf Neon)
COLOR_TEXT_WHITE = RGBColor(0xFF, 0xFF, 0xFF)      # #FFFFFF (White)
COLOR_TEXT_MUTED = RGBColor(0x52, 0x6B, 0x5C)      # #526B5C (Forest Muted)
COLOR_CARD_LIGHT = RGBColor(0xFF, 0xFF, 0xFF)      # White Card on Butter Canvas
COLOR_BADGE_BG = RGBColor(0xFA, 0xF4, 0xE4)        # Translucent Badge Color

FONT_TITLE = "Plus Jakarta Sans"
FONT_BODY = "Plus Jakarta Sans"
FONT_MONO = "JetBrains Mono"

def create_deck():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    logo_path = os.path.join(os.path.dirname(__file__), "Проэкт Хакатон", "img", "logo.png")

    def set_slide_background(slide, color):
        bg_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg_shape.fill.solid()
        bg_shape.fill.fore_color.rgb = color
        bg_shape.line.color.rgb = color
        return bg_shape

    def add_header(slide, title_text, category_text="DEEPTECH GIGAHACK 2026 • BIOMENTORHUB × PROFI", dark_mode=False):
        # Category / Track Tag
        cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.45), Inches(9.0), Inches(0.35))
        tf_cat = cat_box.text_frame
        tf_cat.word_wrap = True
        p_cat = tf_cat.paragraphs[0]
        p_cat.text = category_text.upper()
        p_cat.font.name = FONT_BODY
        p_cat.font.size = Pt(11)
        p_cat.font.bold = True
        p_cat.font.color.rgb = COLOR_LEAF_EMERALD if not dark_mode else COLOR_LEAF_VIBRANT

        # Title
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.75), Inches(9.5), Inches(0.8))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        p_title = tf_title.paragraphs[0]
        p_title.text = title_text
        p_title.font.name = FONT_TITLE
        p_title.font.size = Pt(24)
        p_title.font.bold = True
        p_title.font.color.rgb = COLOR_FOREST_DEEP if not dark_mode else COLOR_TEXT_WHITE

        # Logo on top right
        if os.path.exists(logo_path):
            # Frosted glass-like backing container
            plate = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(10.6), Inches(0.45), Inches(2.0), Inches(0.75))
            plate.fill.solid()
            plate.fill.fore_color.rgb = COLOR_BADGE_BG
            plate.line.color.rgb = COLOR_LEAF_EMERALD if not dark_mode else COLOR_LEAF_VIBRANT
            plate.line.width = Pt(1)
            slide.shapes.add_picture(logo_path, Inches(10.7), Inches(0.52), width=Inches(1.8))

    # =========================================================================
    # SLIDE 1: Title Slide (Dark Forest Green Mode for Maximum Impact)
    # =========================================================================
    slide1 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide1, COLOR_FOREST_DEEP)

    # Hackathon Badge
    badge = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.9), Inches(1.0), Inches(5.8), Inches(0.45))
    badge.fill.solid()
    badge.fill.fore_color.rgb = COLOR_FOREST_CARD
    badge.line.color.rgb = COLOR_LEAF_VIBRANT
    badge.line.width = Pt(1)
    tf_b = badge.text_frame
    p_b = tf_b.paragraphs[0]
    p_b.text = "🌱 DEEPTECH GIGAHACK 2026 • BIOMENTORHUB × PROFI"
    p_b.font.name = FONT_BODY
    p_b.font.size = Pt(11)
    p_b.font.bold = True
    p_b.font.color.rgb = COLOR_LEAF_VIBRANT

    # Logo
    if os.path.exists(logo_path):
        plate = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(10.0), Inches(0.9), Inches(2.4), Inches(0.9))
        plate.fill.solid()
        plate.fill.fore_color.rgb = COLOR_BADGE_BG
        plate.line.color.rgb = COLOR_LEAF_VIBRANT
        plate.line.width = Pt(1)
        slide1.shapes.add_picture(logo_path, Inches(10.15), Inches(0.98), width=Inches(2.1))

    # Main Headline
    tb_title = slide1.shapes.add_textbox(Inches(0.9), Inches(1.7), Inches(11.5), Inches(1.8))
    tf_t = tb_title.text_frame
    tf_t.word_wrap = True
    p1 = tf_t.paragraphs[0]
    p1.text = "PackShift: Революция пищевой упаковки ритейла"
    p1.font.name = FONT_TITLE
    p1.font.size = Pt(36)
    p1.font.bold = True
    p1.font.color.rgb = COLOR_TEXT_WHITE

    p2 = tf_t.add_paragraph()
    p2.text = "0% первичного пластика • Термостойкость до 250°C • EU PPWR 2030 Ready"
    p2.font.name = FONT_TITLE
    p2.font.size = Pt(20)
    p2.font.bold = True
    p2.font.color.rgb = COLOR_LEAF_VIBRANT

    # Subtitle / Description
    tb_sub = slide1.shapes.add_textbox(Inches(0.9), Inches(3.6), Inches(11.0), Inches(0.9))
    tf_sub = tb_sub.text_frame
    tf_sub.word_wrap = True
    p_sub = tf_sub.paragraphs[0]
    p_sub.text = (
        "Инженерная система циркулярной упаковки для сети супермаркетов Profi. "
        "Биомиметические барьеры против жира, полное устранение токсичных PFAS и прямое снижение "
        "экологического сбора ритейлера (EPR Tax) без капитальных затрат (CapEx = 0)."
    )
    p_sub.font.name = FONT_BODY
    p_sub.font.size = Pt(15)
    p_sub.font.color.rgb = COLOR_BG_BUTTER

    # 4 Metric Cards at bottom
    metrics = [
        ("-88%", "Первичный пластик", "Замена на rPET и биополимеры"),
        ("250°C", "Вызов гриля", "До 6ч в витрине без протечек"),
        ("+4 дня", "Свежесть продуктов", "Лазерная микроперфорация EMAP"),
        ("Class A", "RecyClass & PPWR", "100% моно-материалы в сортировку"),
    ]
    card_w = Inches(2.7)
    card_h = Inches(1.8)
    card_y = Inches(4.9)
    gap = Inches(0.25)
    start_x = Inches(0.9)

    for i, (val, label, desc) in enumerate(metrics):
        cx = start_x + i * (card_w + gap)
        card = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, card_y, card_w, card_h)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_FOREST_CARD
        card.line.color.rgb = COLOR_LEAF_VIBRANT if i == 1 else COLOR_LEAF_EMERALD
        card.line.width = Pt(1.5 if i == 1 else 1)

        tb_m = slide1.shapes.add_textbox(cx + Inches(0.15), card_y + Inches(0.15), card_w - Inches(0.3), card_h - Inches(0.3))
        tf_m = tb_m.text_frame
        tf_m.word_wrap = True

        p_val = tf_m.paragraphs[0]
        p_val.text = val
        p_val.font.name = FONT_MONO
        p_val.font.size = Pt(28)
        p_val.font.bold = True
        p_val.font.color.rgb = COLOR_LEAF_VIBRANT

        p_lbl = tf_m.add_paragraph()
        p_lbl.text = label
        p_lbl.font.name = FONT_BODY
        p_lbl.font.size = Pt(13)
        p_lbl.font.bold = True
        p_lbl.font.color.rgb = COLOR_TEXT_WHITE

        p_dsc = tf_m.add_paragraph()
        p_dsc.text = desc
        p_dsc.font.name = FONT_BODY
        p_dsc.font.size = Pt(10)
        p_dsc.font.color.rgb = COLOR_BG_BUTTER

    # =========================================================================
    # SLIDE 2: Problem & Regulatory Pressure (Butter Canvas)
    # =========================================================================
    slide2 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide2, COLOR_BG_BUTTER)
    add_header(slide2, "Двойной вызов ритейла: Регуляторный капкан и боль витрины")

    # Left Column: Regulatory & Financial Pressure
    left_card = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.7), Inches(5.6), Inches(5.1))
    left_card.fill.solid()
    left_card.fill.fore_color.rgb = COLOR_CARD_LIGHT
    left_card.line.color.rgb = COLOR_FOREST_DEEP
    left_card.line.width = Pt(1)

    tb_lc = slide2.shapes.add_textbox(Inches(1.1), Inches(1.9), Inches(5.0), Inches(4.7))
    tf_lc = tb_lc.text_frame
    tf_lc.word_wrap = True

    p = tf_lc.paragraphs[0]
    p.text = "⚖️ 1. Регуляторное и налоговое давление"
    p.font.name = FONT_TITLE
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = COLOR_FOREST_DEEP

    bullets_l = [
        ("Директива EU PPWR 2030:", "Запрет многослойных неперерабатываемых полимеров. Штрафы ритейлерам за превышение квот первичного пластика."),
        ("Экологический сбор (EPR Tax):", "Налог на комбинированные ламинаты (ПЭ+фольга) вырос на 45%. Трудная упаковка становится прямым убытком."),
        ("Запрет PFAS (вечные токсины):", "Европейские регуляторы вводят полный бан на фторированные гидрофобизаторы в пищевом контакте."),
        ("CSRD отчетность Scope 3:", "Сеть Profi обязана публично декларировать и сокращать углеродный след упаковки по всей цепочке."),
    ]
    for b_title, b_desc in bullets_l:
        p_b = tf_lc.add_paragraph()
        p_b.text = f"• {b_title} "
        p_b.font.bold = True
        p_b.font.size = Pt(12)
        p_b.font.color.rgb = COLOR_FOREST_DEEP
        p_b_run = p_b.add_run()
        p_b_run.text = b_desc
        p_b_run.font.bold = False
        p_b_run.font.color.rgb = COLOR_TEXT_MUTED

    # Right Column: The Hot Food Pain Point
    right_card = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.7), Inches(5.7), Inches(5.1))
    right_card.fill.solid()
    right_card.fill.fore_color.rgb = COLOR_FOREST_DEEP
    right_card.line.color.rgb = COLOR_LEAF_VIBRANT
    right_card.line.width = Pt(1.5)

    tb_rc = slide2.shapes.add_textbox(Inches(7.1), Inches(1.9), Inches(5.1), Inches(4.7))
    tf_rc = tb_rc.text_frame
    tf_rc.word_wrap = True

    p = tf_rc.paragraphs[0]
    p.text = "🔥 2. Острая боль: Отдел гриля и кулинарии"
    p.font.name = FONT_TITLE
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = COLOR_TEXT_WHITE

    bullets_r = [
        ("Экстремальный режим:", "Выдержка в витрине при 85°C – 120°C до 6 часов. Пиковые температуры фасовки достигают 180°C – 250°C."),
        ("Дефекты текущей упаковки:", "Полиэтиленовые пакеты с фольгой размягчаются, деформируются, текут и выделяют запах нагретого полимера."),
        ("Эффект «парника»:", "Пар блокируется внутри пакета, размачивая хрустящую корочку гриля в кашу за 40 минут."),
        ("Скрытая токсичность:", "Жиростойкость текущих упаковок держится на PFAS-добавках, мигрирующих в горячую пищу."),
    ]
    for b_title, b_desc in bullets_r:
        p_b = tf_rc.add_paragraph()
        p_b.text = f"• {b_title} "
        p_b.font.bold = True
        p_b.font.size = Pt(12)
        p_b.font.color.rgb = COLOR_LEAF_VIBRANT
        p_b_run = p_b.add_run()
        p_b_run.text = b_desc
        p_b_run.font.bold = False
        p_b_run.font.color.rgb = COLOR_BG_BUTTER

    # =========================================================================
    # SLIDE 3: The Special Annex Challenge: Hot Grill Counter Solutions
    # =========================================================================
    slide3 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide3, COLOR_BG_BUTTER)
    add_header(slide3, "Специальный вызов: Инновации для Горячего Гриля (180°C – 250°C)")

    # Solution A Card (CPET Bag)
    sol_a = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.7), Inches(5.6), Inches(5.1))
    sol_a.fill.solid()
    sol_a.fill.fore_color.rgb = COLOR_CARD_LIGHT
    sol_a.line.color.rgb = COLOR_LEAF_EMERALD
    sol_a.line.width = Pt(1.5)

    tb_sa = slide3.shapes.add_textbox(Inches(1.1), Inches(1.9), Inches(5.0), Inches(4.7))
    tf_sa = tb_sa.text_frame
    tf_sa.word_wrap = True

    p = tf_sa.paragraphs[0]
    p.text = "РЕШЕНИЕ А: МОНО-ПАКЕТ"
    p.font.name = FONT_BODY
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_LEAF_EMERALD

    p_t = tf_sa.add_paragraph()
    p_t.text = "PackShift CPET Ultra-Flex"
    p_t.font.name = FONT_TITLE
    p_t.font.size = Pt(20)
    p_t.font.bold = True
    p_t.font.color.rgb = COLOR_FOREST_DEEP

    specs_a = [
        ("Материал:", "Модифицированный кристаллический rPET (степень кристалличности > 38%)."),
        ("Термостойкость:", "от -40°C до +250°C (Dual-Ovenable: духовка + СВЧ)."),
        ("Тепловая витрина:", "6+ часов при 95°C–120°C без усадки, протечек и запаха."),
        ("Жиростойкость:", "Высший класс Kit Test 12. 100% PFAS-Free."),
        ("Рециклинг:", "100% моно-материал. Поток ПЭТ-бутылок (RecyClass Class A)."),
        ("Экономика:", "Себестоимость на 14% ниже импортных алюминиевых ламинатов."),
    ]
    for k, v in specs_a:
        p_sp = tf_sa.add_paragraph()
        p_sp.text = f"{k} "
        p_sp.font.bold = True
        p_sp.font.size = Pt(11.5)
        p_sp.font.color.rgb = COLOR_FOREST_DEEP
        r = p_sp.add_run()
        r.text = v
        r.font.bold = False
        r.font.color.rgb = COLOR_TEXT_MUTED

    # Solution B Card (LignoShield BioBox)
    sol_b = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.7), Inches(5.7), Inches(5.1))
    sol_b.fill.solid()
    sol_b.fill.fore_color.rgb = COLOR_FOREST_DEEP
    sol_b.line.color.rgb = COLOR_LEAF_VIBRANT
    sol_b.line.width = Pt(1.5)

    tb_sb = slide3.shapes.add_textbox(Inches(7.1), Inches(1.9), Inches(5.1), Inches(4.7))
    tf_sb = tb_sb.text_frame
    tf_sb.word_wrap = True

    p = tf_sb.paragraphs[0]
    p.text = "РЕШЕНИЕ Б: БИО-КОМПОЗИТ (ДЛЯ PROFI BIO)"
    p.font.name = FONT_BODY
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_LEAF_VIBRANT

    p_t = tf_sb.add_paragraph()
    p_t.text = "LignoShield™ ThermoBox"
    p_t.font.name = FONT_TITLE
    p_t.font.size = Pt(20)
    p_t.font.bold = True
    p_t.font.color.rgb = COLOR_TEXT_WHITE

    specs_b = [
        ("Материал:", "100% формованная багасса сахарного тростника + альгинат/хитозан из водорослей."),
        ("Дышащая мембрана:", "Микропористая структура отводит влажный пар. Сохраняет корочку хрустящей."),
        ("Жиробарьер:", "Водный нано-барьер. 0% пластика, 0% фтора, 0% полиэтилена."),
        ("Температура:", "До +220°C кратковременно, 6 часов в витрине гриля."),
        ("Утилизация:", "Домашний компост за 60 дней (OK Compost HOME) или макулатура."),
        ("Статус:", "Идеальное позиционирование для премиальной эко-линейки Profi."),
    ]
    for k, v in specs_b:
        p_sp = tf_sb.add_paragraph()
        p_sp.text = f"{k} "
        p_sp.font.bold = True
        p_sp.font.size = Pt(11.5)
        p_sp.font.color.rgb = COLOR_LEAF_VIBRANT
        r = p_sp.add_run()
        r.text = v
        r.font.bold = False
        r.font.color.rgb = COLOR_BG_BUTTER

    # =========================================================================
    # SLIDE 4: 3 Store Departments: Vegetables, Meat, Bakery
    # =========================================================================
    slide4 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide4, COLOR_BG_BUTTER)
    add_header(slide4, "Оптимизация отделов Profi: Овощи, Мясо, Выпечка")

    dept_cols = [
        ("🥬 Свежие овощи", "EMAP Laser-Breathe™", [
            ("Проблема:", "Конденсат, плесень, 28-35% списаний ягод и зелени."),
            ("Инновация:", "Моно-пленка rBOPP с лазерной микроперфорацией. Динамический баланс O₂ (3-5%) и CO₂ (5-8%)."),
            ("Эффект:", "+4 дня к сроку годности на полке."),
            ("Ритейл:", "Снижение списаний в супермаркетах на 32%."),
        ]),
        ("🥩 Мясной отдел", "Capillary Tray (Pad-Less)", [
            ("Проблема:", "Грязная салфетка-вкладыш делает лоток неперерабатываемым."),
            ("Инновация:", "Глубокоформованный моно-rPET лоток с капиллярным гидро-замком. Сбор до 45 мл экссудата."),
            ("Эффект:", "0 синтетических салфеток-вкладышей."),
            ("Рециклинг:", "100% чистый поток сортировки и рециклинга."),
        ]),
        ("🥐 Пекарня", "NatureVent Crisp™", [
            ("Проблема:", "Полиэтиленовые окна пакетов сжигаются на полигоне."),
            ("Инновация:", "Крафт FSC + прозрачное смотровое био-окно NatureFlex™ из древесной массы + био-воск."),
            ("Эффект:", "Сохранение хрустящей корочки выпечки."),
            ("Рециклинг:", "100% Pulpable (полная переработка с бумагой)."),
        ]),
    ]

    col_w = Inches(3.7)
    col_h = Inches(5.1)
    col_y = Inches(1.7)
    col_gap = Inches(0.3)

    for i, (dept_tag, tech_name, points) in enumerate(dept_cols):
        cx = Inches(0.8) + i * (col_w + col_gap)
        card = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, col_y, col_w, col_h)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD_LIGHT
        card.line.color.rgb = COLOR_FOREST_DEEP
        card.line.width = Pt(1)

        tb = slide4.shapes.add_textbox(cx + Inches(0.2), col_y + Inches(0.2), col_w - Inches(0.4), col_h - Inches(0.4))
        tf = tb.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        p.text = dept_tag.upper()
        p.font.name = FONT_BODY
        p.font.size = Pt(12)
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
            p_pt.font.size = Pt(11)
            p_pt.font.color.rgb = COLOR_FOREST_DEEP
            r = p_pt.add_run()
            r.text = pv
            r.font.bold = False
            r.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 5: Economics & ESG Impact Calculator (Dark Forest Green)
    # =========================================================================
    slide5 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide5, COLOR_FOREST_DEEP)
    add_header(slide5, "Экономика и ESG-эффект: Калькулятор для сети Profi", dark_mode=True)

    # Context box
    ctx = slide5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.6), Inches(11.7), Inches(0.75))
    ctx.fill.solid()
    ctx.fill.fore_color.rgb = COLOR_FOREST_CARD
    ctx.line.color.rgb = COLOR_LEAF_VIBRANT
    ctx.line.width = Pt(1)
    tf_ctx = ctx.text_frame
    p_ctx = tf_ctx.paragraphs[0]
    p_ctx.text = "🎯 Моделирование эффекта на масштабе 100 супермаркетов Profi (базовый охват)"
    p_ctx.font.name = FONT_BODY
    p_ctx.font.size = Pt(14)
    p_ctx.font.bold = True
    p_ctx.font.color.rgb = COLOR_TEXT_WHITE

    calc_metrics = [
        ("1 840 т", "Первичный пластик", "Тонн первичных полимеров предотвращено в год благодаря rPET"),
        ("4 250 т", "Сокращение CO₂e", "Предотвращено выбросов парниковых газов (Scope 3 Profi)"),
        ("€1 280 000", "Экономия на EPR", "Прямая годовая экономия на налоге за счет моно-материалов"),
        ("€700 000", "Снижение списаний", "Сбережение продукции за счет продления свежести (+4 дня)"),
    ]

    cm_w = Inches(2.75)
    cm_h = Inches(3.0)
    cm_y = Inches(2.6)
    cm_gap = Inches(0.23)

    for i, (val, title, sub) in enumerate(calc_metrics):
        cx = Inches(0.8) + i * (cm_w + cm_gap)
        card = slide5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, cm_y, cm_w, cm_h)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_FOREST_CARD
        card.line.color.rgb = COLOR_LEAF_VIBRANT if i == 2 else COLOR_LEAF_EMERALD
        card.line.width = Pt(2 if i == 2 else 1)

        tb = slide5.shapes.add_textbox(cx + Inches(0.15), cm_y + Inches(0.2), cm_w - Inches(0.3), cm_h - Inches(0.4))
        tf = tb.text_frame
        tf.word_wrap = True

        p_v = tf.paragraphs[0]
        p_v.text = val
        p_v.font.name = FONT_MONO
        p_v.font.size = Pt(26)
        p_v.font.bold = True
        p_v.font.color.rgb = COLOR_LEAF_VIBRANT

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
        p_s.font.color.rgb = COLOR_BG_BUTTER

    # Bottom CapEx Bar
    capex_bar = slide5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(5.85), Inches(11.7), Inches(0.95))
    capex_bar.fill.solid()
    capex_bar.fill.fore_color.rgb = COLOR_FOREST_CARD
    capex_bar.line.color.rgb = COLOR_LEAF_VIBRANT
    capex_bar.line.width = Pt(1)
    tf_c = capex_bar.text_frame
    tf_c.word_wrap = True
    p_c = tf_c.paragraphs[0]
    p_c.text = "💡 Нулевой CapEx для ритейлера:"
    p_c.font.bold = True
    p_c.font.size = Pt(13)
    p_c.font.color.rgb = COLOR_LEAF_VIBRANT
    r_c = p_c.add_run()
    r_c.text = " Упаковка PackShift полностью адаптирована под стандартные упаковочные линии, запайщики лотков и витрины сети Profi. Переоборудование магазинов не требуется."
    r_c.font.bold = False
    r_c.font.color.rgb = COLOR_TEXT_WHITE

    # =========================================================================
    # SLIDE 6: Standards & Compliance (Butter Canvas)
    # =========================================================================
    slide6 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide6, COLOR_BG_BUTTER)
    add_header(slide6, "Стандарты и Сертификация: Регуляторный щит Profi")

    standards = [
        ("RecyClass", "Grade A Recyclability", "Подтвержденная совместимость с оптическими сепараторами и линиями вторичной переработки в ЕС."),
        ("EU PPWR", "PPWR 2030 Ready", "Упреждение европейских норм: >35% пищевого rPET, 0% первичного пластика в одноразовых пакетах."),
        ("Zero Toxic", "100% PFAS-Free", "Сертификация отсутствия пер- и полифторалкильных веществ в контакте с горячим маслом до 250°C."),
        ("Food Safe", "EFSA / FDA Approved", "Соответствие Regulation EU No 10/2011 для прямого контакта с горячими, кислыми и жирными продуктами."),
    ]

    st_w = Inches(2.75)
    st_h = Inches(5.1)
    st_y = Inches(1.7)
    st_gap = Inches(0.23)

    for i, (badge_txt, st_title, st_desc) in enumerate(standards):
        cx = Inches(0.8) + i * (st_w + st_gap)
        card = slide6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, st_y, st_w, st_h)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD_LIGHT
        card.line.color.rgb = COLOR_FOREST_DEEP
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
    # SLIDE 7: Pilot Implementation Roadmap (Butter Canvas)
    # =========================================================================
    slide7 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide7, COLOR_BG_BUTTER)
    add_header(slide7, "План пилотного внедрения в Profi: 6 недель до результата")

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
    st_h = Inches(3.6)
    st_y = Inches(1.7)
    st_gap = Inches(0.3)

    for i, (badge_txt, st_title, st_time, points) in enumerate(steps):
        cx = Inches(0.8) + i * (st_w + st_gap)
        card = slide7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, st_y, st_w, st_h)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD_LIGHT
        card.line.color.rgb = COLOR_LEAF_EMERALD if i == 0 else COLOR_FOREST_DEEP
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
            p_pt.font.size = Pt(11)
            p_pt.font.color.rgb = COLOR_TEXT_MUTED

    # Bottom Callout: 200 Samples Ready!
    callout = slide7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(5.6), Inches(11.7), Inches(1.2))
    callout.fill.solid()
    callout.fill.fore_color.rgb = COLOR_FOREST_DEEP
    callout.line.color.rgb = COLOR_LEAF_VIBRANT
    callout.line.width = Pt(1.5)
    tf_co = callout.text_frame
    tf_co.word_wrap = True

    p_co = tf_co.paragraphs[0]
    p_co.text = "📦 Готовность тестовой партии к передаче:"
    p_co.font.bold = True
    p_co.font.size = Pt(14)
    p_co.font.color.rgb = COLOR_LEAF_VIBRANT

    r_co = p_co.add_run()
    r_co.text = " Мы подготовили комплект из 200 образцов пакетов CPET Ultra-Flex и биобоксов LignoShield. Готовы передать их технологическому отделу Profi для старта испытаний на следующей неделе."
    r_co.font.bold = False
    r_co.font.color.rgb = COLOR_TEXT_WHITE

    # =========================================================================
    # SLIDE 8: Team & Call to Action (Dark Forest Green)
    # =========================================================================
    slide8 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide8, COLOR_FOREST_DEEP)

    # Logo on center top
    if os.path.exists(logo_path):
        plate = slide8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.2), Inches(0.8), Inches(3.0), Inches(1.0))
        plate.fill.solid()
        plate.fill.fore_color.rgb = COLOR_BADGE_BG
        plate.line.color.rgb = COLOR_LEAF_VIBRANT
        plate.line.width = Pt(1)
        slide8.shapes.add_picture(logo_path, Inches(5.35), Inches(0.88), width=Inches(2.7))

    tb_f = slide8.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(11.333), Inches(2.2))
    tf_f = tb_f.text_frame
    tf_f.word_wrap = True

    p_f1 = tf_f.paragraphs[0]
    p_f1.alignment = PP_ALIGN.CENTER
    p_f1.text = "Делаем циркулярность прибыльной для Profi уже сегодня"
    p_f1.font.name = FONT_TITLE
    p_f1.font.size = Pt(32)
    p_f1.font.bold = True
    p_f1.font.color.rgb = COLOR_TEXT_WHITE

    p_f2 = tf_f.add_paragraph()
    p_f2.alignment = PP_ALIGN.CENTER
    p_f2.text = "0% первичного пластика • Защита маржинальности • Соответствие PPWR 2030"
    p_f2.font.name = FONT_TITLE
    p_f2.font.size = Pt(18)
    p_f2.font.bold = True
    p_f2.font.color.rgb = COLOR_LEAF_VIBRANT

    # Links & Contact Box
    contact_card = slide8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(2.5), Inches(4.3), Inches(8.333), Inches(2.2))
    contact_card.fill.solid()
    contact_card.fill.fore_color.rgb = COLOR_FOREST_CARD
    contact_card.line.color.rgb = COLOR_LEAF_VIBRANT
    contact_card.line.width = Pt(1.5)

    tb_c = slide8.shapes.add_textbox(Inches(2.7), Inches(4.5), Inches(7.933), Inches(1.8))
    tf_c = tb_c.text_frame
    tf_c.word_wrap = True

    p_c1 = tf_c.paragraphs[0]
    p_c1.alignment = PP_ALIGN.CENTER
    p_c1.text = "🌐 Интерактивная платформа и ESG-калькулятор:"
    p_c1.font.bold = True
    p_c1.font.size = Pt(14)
    p_c1.font.color.rgb = COLOR_BG_BUTTER

    p_c2 = tf_c.add_paragraph()
    p_c2.alignment = PP_ALIGN.CENTER
    p_c2.text = "https://lisqrn.github.io/Hackaton-202636/"
    p_c2.font.name = FONT_MONO
    p_c2.font.bold = True
    p_c2.font.size = Pt(16)
    p_c2.font.color.rgb = COLOR_LEAF_VIBRANT

    p_c3 = tf_c.add_paragraph()
    p_c3.alignment = PP_ALIGN.CENTER
    p_c3.text = "Открытый репозиторий: github.com/denis100strike/Hackaton-202636 (PR #1)"
    p_c3.font.size = Pt(12)
    p_c3.font.color.rgb = COLOR_BG_BUTTER

    p_c4 = tf_c.add_paragraph()
    p_c4.alignment = PP_ALIGN.CENTER
    p_c4.text = "Спасибо за внимание! Готовы ответить на ваши вопросы."
    p_c4.font.bold = True
    p_c4.font.size = Pt(14)
    p_c4.font.color.rgb = COLOR_TEXT_WHITE

    # Save presentation
    output_path = os.path.join(os.path.dirname(__file__), "PackShift_Profi_Gigahack2026.pptx")
    prs.save(output_path)
    print(f"Presentation saved successfully to: {output_path}")

if __name__ == "__main__":
    create_deck()

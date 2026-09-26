import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

# ==============================================================================
# DEEPTECH DECK: Dark Graphite + Neon Emerald + Centered Scientific Schemas
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

    def add_centered_header(slide, tag_text, title_text, subtitle_text, slide_num):
        # Category Tag (Centered)
        tb_tag = slide.shapes.add_textbox(Inches(1.0), Inches(0.35), Inches(11.333), Inches(0.35))
        tf_t = tb_tag.text_frame
        p_t = tf_t.paragraphs[0]
        p_t.alignment = PP_ALIGN.CENTER
        p_t.text = tag_text.upper()
        p_t.font.name = FONT_MONO
        p_t.font.size = Pt(11)
        p_t.font.bold = True
        p_t.font.color.rgb = COLOR_NEON

        # Main Title (Centered)
        tb_tit = slide.shapes.add_textbox(Inches(1.0), Inches(0.7), Inches(11.333), Inches(0.65))
        tf_tit = tb_tit.text_frame
        p_tit = tf_tit.paragraphs[0]
        p_tit.alignment = PP_ALIGN.CENTER
        p_tit.text = title_text
        p_tit.font.name = FONT_TITLE
        p_tit.font.size = Pt(23)
        p_tit.font.bold = True
        p_tit.font.color.rgb = COLOR_TEXT_WHITE

        # Subtitle (Centered)
        tb_sub = slide.shapes.add_textbox(Inches(1.5), Inches(1.3), Inches(10.333), Inches(0.45))
        tf_sub = tb_sub.text_frame
        p_sub = tf_sub.paragraphs[0]
        p_sub.alignment = PP_ALIGN.CENTER
        p_sub.text = subtitle_text
        p_sub.font.name = FONT_BODY
        p_sub.font.size = Pt(12)
        p_sub.font.color.rgb = COLOR_TEXT_MUTED

        # Slide Number (Top Right)
        tb_num = slide.shapes.add_textbox(Inches(11.3), Inches(0.35), Inches(1.2), Inches(0.35))
        tf_num = tb_num.text_frame
        p_num = tf_num.paragraphs[0]
        p_num.alignment = PP_ALIGN.RIGHT
        p_num.text = f"0{slide_num} / 08"
        p_num.font.name = FONT_MONO
        p_num.font.size = Pt(11)
        p_num.font.bold = True
        p_num.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 1: Cover (Centered, Clean, Deeptech Venture Vision)
    # =========================================================================
    slide1 = prs.slides.add_slide(blank_layout)
    set_bg(slide1)

    tb_h = slide1.shapes.add_textbox(Inches(1.0), Inches(1.0), Inches(11.333), Inches(2.8))
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
    p_main.font.size = Pt(46)
    p_main.font.bold = True
    p_main.font.color.rgb = COLOR_TEXT_WHITE

    p_sub = tf_h.add_paragraph()
    p_sub.alignment = PP_ALIGN.CENTER
    p_sub.text = "ВЫСОКОКРИСТАЛЛИЧЕСКАЯ И БИОМИМЕТИЧЕСКАЯ МОНО-МАТЕРИАЛЬНАЯ ПЛАТФОРМА"
    p_sub.font.name = FONT_MONO
    p_sub.font.size = Pt(13)
    p_sub.font.bold = True
    p_sub.font.color.rgb = COLOR_NEON

    p_desc = tf_h.add_paragraph()
    p_desc.alignment = PP_ALIGN.CENTER
    p_desc.text = "\nРазрешая ключевой парадокс материаловедения: термостойкость до 250°C при 100% циркулярности"
    p_desc.font.name = FONT_BODY
    p_desc.font.size = Pt(14)
    p_desc.font.color.rgb = COLOR_TEXT_MUTED

    # 3 Centered Breakthrough Nodes
    breakthrough_nodes = [
        ("🧬 Нано-кристаллизация", "χc > 38%", "Модифицированный rPET. Термостойкость до 250°C без деформации в витрине"),
        ("🌿 Биомиметический барьер", "Kit 12 • 0% PFAS", "Ионно-сшитая матрица (альгинат + хитозан). Барьер горячих масел"),
        ("🔄 Чистая циркулярность", "RecyClass A", "100% моно-материалы: чистая сортировка в поток ПЭТ или домашний компост 60 дн"),
    ]
    card_w = Inches(3.6)
    card_h = Inches(2.6)
    card_y = Inches(4.3)
    gap = Inches(0.26)
    start_x = Inches(1.0)

    for i, (head, val, desc) in enumerate(breakthrough_nodes):
        cx = start_x + i * (card_w + gap)
        card = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, card_y, card_w, card_h)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD_GLOW if i == 0 else COLOR_CARD
        card.line.color.rgb = COLOR_EMERALD if i == 0 else COLOR_CARD_BORDER
        card.line.width = Pt(1.5 if i == 0 else 1)

        tb_c = slide1.shapes.add_textbox(cx + Inches(0.15), card_y + Inches(0.2), card_w - Inches(0.3), card_h - Inches(0.4))
        tf_c = tb_c.text_frame
        tf_c.word_wrap = True

        p_h = tf_c.paragraphs[0]
        p_h.alignment = PP_ALIGN.CENTER
        p_h.text = head
        p_h.font.name = FONT_TITLE
        p_h.font.size = Pt(14)
        p_h.font.bold = True
        p_h.font.color.rgb = COLOR_TEXT_WHITE

        p_v = tf_c.add_paragraph()
        p_v.alignment = PP_ALIGN.CENTER
        p_v.text = f"\n{val}\n"
        p_v.font.name = FONT_MONO
        p_v.font.size = Pt(22)
        p_v.font.bold = True
        p_v.font.color.rgb = COLOR_NEON

        p_d = tf_c.add_paragraph()
        p_d.alignment = PP_ALIGN.CENTER
        p_d.text = desc
        p_d.font.name = FONT_BODY
        p_d.font.size = Pt(11)
        p_d.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 2: The Packaging Trilemma (Scientific Challenge & Compromise)
    # =========================================================================
    slide2 = prs.slides.add_slide(blank_layout)
    set_bg(slide2)
    add_centered_header(slide2, "ФУНДАМЕНТАЛЬНЫЙ ВЫЗОВ • ТЕОРИЯ МАТЕРИАЛОВ",
                        "Материаловедческая трилемма пищевой упаковки",
                        "Почему современные полимеры не способны объединить температуру, переработку и себестоимость", 2)

    # Left: The Trilemma Radar Box
    c_tri = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(1.9), Inches(5.4), Inches(5.0))
    c_tri.fill.solid()
    c_tri.fill.fore_color.rgb = COLOR_CARD
    c_tri.line.color.rgb = COLOR_CARD_BORDER
    c_tri.line.width = Pt(1)

    tb_tri = slide2.shapes.add_textbox(Inches(1.2), Inches(2.1), Inches(5.0), Inches(4.6))
    tf_tri = tb_tri.text_frame
    tf_tri.word_wrap = True

    p = tf_tri.paragraphs[0]
    p.text = "🔺 ТРИЛЕММА УПАКОВОЧНЫХ МАТЕРИАЛОВ"
    p.font.name = FONT_MONO
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_NEON

    tri_points = [
        ("Вершина 1: Термостойкость (>200°C)", "Способность удерживать кипящий жир и пар до 6 часов в тепловых витринах Profi без деформации"),
        ("Вершина 2: Чистая циркулярность", "100% RecyClass Grade A: возможность переработки в чистом потоке бутылочного rPET или OK Compost HOME"),
        ("Вершина 3: Паритет себестоимости", "Целевая себестоимость 1:1 со стандартными пластиками без закупки дорогого фасовочного оборудования"),
        ("⚡ Прорыв PackShift", "Парето-оптимальная комбинация: нано-кристаллический rPET + природная ионно-сшитая полисахаридная матрица"),
    ]
    for h, b in tri_points:
        p_h = tf_tri.add_paragraph()
        p_h.text = f"\n{h}"
        p_h.font.name = FONT_TITLE
        p_h.font.size = Pt(12)
        p_h.font.bold = True
        p_h.font.color.rgb = COLOR_TEXT_WHITE if not "⚡" in h else COLOR_NEON

        p_b = tf_tri.add_paragraph()
        p_b.text = b
        p_b.font.name = FONT_BODY
        p_b.font.size = Pt(11)
        p_b.font.color.rgb = COLOR_TEXT_MUTED

    # Right: Failure of Existing Materials
    c_fail = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.9), Inches(5.5), Inches(5.0))
    c_fail.fill.solid()
    c_fail.fill.fore_color.rgb = COLOR_CARD
    c_fail.line.color.rgb = COLOR_CARD_BORDER
    c_fail.line.width = Pt(1)

    tb_fail = slide2.shapes.add_textbox(Inches(7.0), Inches(2.1), Inches(5.1), Inches(4.6))
    tf_fail = tb_fail.text_frame
    tf_fail.word_wrap = True

    p = tf_fail.paragraphs[0]
    p.text = "КРАХ ТРАДИЦИОННЫХ АЛЬТЕРНАТИВ"
    p.font.name = FONT_MONO
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0xF4, 0x3F, 0x5E)

    failures = [
        ("❌ Многослойные ламинаты (PET/Al/PE)", "Неразделимый пирог из 3 полимеров и клея. 0% вторичной переработки, распад клея при 85°C в витрине, штраф EPR €800/т."),
        ("⚠️ Традиционный алюминий", "Держит нагрев, но оставляет огромный углеродный след (11.5 кг CO₂/кг), непрозрачен на витрине и взрывается в бытовых СВЧ."),
        ("⚠️ Стандартные биопластики (PLA / PBAT)", "Плавятся уже при 55°C–60°C. Полностью неприменимы для горячих витрин гриля (85°C–120°C)."),
    ]
    for h, b in failures:
        p_h = tf_fail.add_paragraph()
        p_h.text = f"\n{h}"
        p_h.font.name = FONT_TITLE
        p_h.font.size = Pt(12)
        p_h.font.bold = True
        p_h.font.color.rgb = COLOR_TEXT_WHITE

        p_b = tf_fail.add_paragraph()
        p_b.text = b
        p_b.font.name = FONT_BODY
        p_b.font.size = Pt(11)
        p_b.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 3: Deeptech Core 1 — Polymer Engineering (High-χc CPET)
    # =========================================================================
    slide3 = prs.slides.add_slide(blank_layout)
    set_bg(slide3)
    add_centered_header(slide3, "ПОЛИМЕРНАЯ ИНЖЕНЕРИЯ • НАУЧНОЕ ЯДРО 1",
                        "Высококристаллический моно-rPET (CPET Ultra-Flex)",
                        "Управление кинетикой кристаллизации вторичного ПЭТ (χc > 38%) для удержания 250°C", 3)

    # Left: DMA Thermomechanical Chart Box
    c_dma = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(1.9), Inches(5.4), Inches(5.0))
    c_dma.fill.solid()
    c_dma.fill.fore_color.rgb = COLOR_CARD
    c_dma.line.color.rgb = COLOR_CARD_BORDER
    c_dma.line.width = Pt(1)

    tb_dma = slide3.shapes.add_textbox(Inches(1.2), Inches(2.1), Inches(5.0), Inches(4.6))
    tf_dma = tb_dma.text_frame
    tf_dma.word_wrap = True

    p = tf_dma.paragraphs[0]
    p.text = "📊 ТЕРМОМЕХАНИЧЕСКИЙ АНАЛИЗ (DMA)"
    p.font.name = FONT_MONO
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_NEON

    dma_bullets = [
        ("Температура стеклования (Tg ≈ 70°C)", "Обычный аморфный rPET теряет жесткость и модуль упругости при 70°C — упаковка плывет в тепловой витрине Profi."),
        ("Термо-плато PackShift (до 255°C)", "Нано-кристаллические сферолиты удерживают модуль упругости выше 1 200 МПа вплоть до точки плавления кристаллов Tm."),
        ("Стабильность геометрии лотка", "Усадка менее 1.2% после 6 часов выдержки при 120°C в термокамере."),
        ("100% NIR-детектируемость", "Оптические сортировщики распознают полимер как прозрачный чистый PET и отправляют в поток вторичного сырья."),
    ]
    for h, b in dma_bullets:
        p_h = tf_dma.add_paragraph()
        p_h.text = f"\n{h}"
        p_h.font.name = FONT_TITLE
        p_h.font.size = Pt(12)
        p_h.font.bold = True
        p_h.font.color.rgb = COLOR_TEXT_WHITE

        p_b = tf_dma.add_paragraph()
        p_b.text = b
        p_b.font.name = FONT_BODY
        p_b.font.size = Pt(11)
        p_b.font.color.rgb = COLOR_TEXT_MUTED

    # Right: Technical Specification
    c_spec = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.9), Inches(5.5), Inches(5.0))
    c_spec.fill.solid()
    c_spec.fill.fore_color.rgb = COLOR_CARD
    c_spec.line.color.rgb = COLOR_EMERALD
    c_spec.line.width = Pt(1.5)

    tb_spec = slide3.shapes.add_textbox(Inches(7.0), Inches(2.1), Inches(5.1), Inches(4.6))
    tf_spec = tb_spec.text_frame
    tf_spec.word_wrap = True

    p = tf_spec.paragraphs[0]
    p.text = "ИНЖЕНЕРНАЯ СПЕЦИФИКАЦИЯ МАТЕРИАЛА"
    p.font.name = FONT_MONO
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_TEXT_WHITE

    specs = [
        ("Степень кристалличности", "χc > 38% (индукция органическими нано-нуклеаторами)"),
        ("Рабочий температурный диапазон", "от -40°C до +250°C (Dual-Ovenable: морозилка / СВЧ / духовка)"),
        ("Жиростойкость по TAPPI T559", "Kit Test 12 / 12 без добавления фторорганики (Zero PFAS)"),
        ("Чистота вторичного потока", "99.2% выход при переработке в пищевой rPET гранулят"),
        ("Скорость термосварки", "1.1 – 1.3 секунды на типовых матрицах фасовки"),
    ]
    for h, b in specs:
        p_h = tf_spec.add_paragraph()
        p_h.text = f"\n• {h}:"
        p_h.font.name = FONT_TITLE
        p_h.font.size = Pt(12)
        p_h.font.bold = True
        p_h.font.color.rgb = COLOR_NEON

        p_b = tf_spec.add_paragraph()
        p_b.text = f"  {b}"
        p_b.font.name = FONT_BODY
        p_b.font.size = Pt(11)
        p_b.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 4: Deeptech Core 2 — Biomimetic Nanobarrier (LignoShield)
    # =========================================================================
    slide4 = prs.slides.add_slide(blank_layout)
    set_bg(slide4)
    add_centered_header(slide4, "БИОМИМЕТИКА • НАУЧНОЕ ЯДРО 2",
                        "Ионно-сшитая матрица LignoShield™",
                        "Паропроницаемый гидрофобный нано-барьер из полисахаридов водорослей и волокон багассы", 4)

    # Left: Cross-Section Architecture
    c_lay = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(1.9), Inches(5.4), Inches(5.0))
    c_lay.fill.solid()
    c_lay.fill.fore_color.rgb = COLOR_CARD
    c_lay.line.color.rgb = COLOR_CARD_BORDER
    c_lay.line.width = Pt(1)

    tb_lay = slide4.shapes.add_textbox(Inches(1.2), Inches(2.1), Inches(5.0), Inches(4.6))
    tf_lay = tb_lay.text_frame
    tf_lay.word_wrap = True

    p = tf_lay.paragraphs[0]
    p.text = "🔬 СТРУКТУРА МАТЕРИАЛА В РАЗРЕЗЕ"
    p.font.name = FONT_MONO
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_NEON

    layers = [
        ("Слой 1: Нанобарьер (Альгинат + Хитозан)", "Комплекс полиэлектролитов со сшивкой Ca²+. Угол смачивания >110°, блокирует кипящий жир без фтора (Kit 12)."),
        ("Слой 2: Диффузионная паропроницаемая мембрана", "Микропористая структура свободно отводит молекулы пара H₂O, предотвращая парниковый эффект и размокание."),
        ("Слой 3: Формованная лигноцеллюлоза (Багасса)", "Несущий каркас из жмыха сахарного тростника. 100% компостирование в саду за 60 дней (OK Compost HOME)."),
    ]
    for h, b in layers:
        p_h = tf_lay.add_paragraph()
        p_h.text = f"\n{h}"
        p_h.font.name = FONT_TITLE
        p_h.font.size = Pt(12)
        p_h.font.bold = True
        p_h.font.color.rgb = COLOR_TEXT_WHITE

        p_b = tf_lay.add_paragraph()
        p_b.text = b
        p_b.font.name = FONT_BODY
        p_b.font.size = Pt(11)
        p_b.font.color.rgb = COLOR_TEXT_MUTED

    # Right: Unique Value for Profi Hot Deli
    c_deli = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.9), Inches(5.5), Inches(5.0))
    c_deli.fill.solid()
    c_deli.fill.fore_color.rgb = COLOR_CARD
    c_deli.line.color.rgb = COLOR_CARD_BORDER
    c_deli.line.width = Pt(1)

    tb_deli = slide4.shapes.add_textbox(Inches(7.0), Inches(2.1), Inches(5.1), Inches(4.6))
    tf_deli = tb_deli.text_frame
    tf_deli.word_wrap = True

    p = tf_deli.paragraphs[0]
    p.text = "РЕШЕНИЕ ПРОБЛЕМЫ РАЗМОКАНИЯ КОРОЧКИ"
    p.font.name = FONT_MONO
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_TEXT_WHITE

    deli_points = [
        ("Проблема герметичных пакетов", "Горячая курица-гриль в витрине выделяет влагу. В обычном полиэтилене пар превращается в конденсат, и хрустящая корочка раскисает за 20 минут."),
        ("Технология Crisp Guard", "Селективный градиент мембраны выводит избыточный пар наружу, сохраняя аппетитную хрусткость корочки до 4 часов."),
        ("Чистый домашний компост", "Не требует сдачи на промышленный компостинг при 60°C. Полный распад в грунте на безопасный перегной за 2 месяца."),
    ]
    for h, b in deli_points:
        p_h = tf_deli.add_paragraph()
        p_h.text = f"\n{h}"
        p_h.font.name = FONT_TITLE
        p_h.font.size = Pt(12)
        p_h.font.bold = True
        p_h.font.color.rgb = COLOR_NEON

        p_b = tf_deli.add_paragraph()
        p_b.text = b
        p_b.font.name = FONT_BODY
        p_b.font.size = Pt(11)
        p_b.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 5: Experimental Validation & Benchmarks
    # =========================================================================
    slide5 = prs.slides.add_slide(blank_layout)
    set_bg(slide5)
    add_centered_header(slide5, "ЭКСПЕРИМЕНТАЛЬНЫЕ ДАННЫЕ • ВАЛИДАЦИЯ",
                        "Сравнительные лабораторные бенчмарки",
                        "Прямое физико-химическое и жизненное (LCA) сопоставление с традиционными материалами", 5)

    bench_cards = [
        ("ДЕФОРМАЦИЯ ПРИ 120°C (60 МИН)", "< 1.2%", "PackShift CPET", 
         "LDPE плавится на 100% за 10 мин.\nPLA биопластик деформируется на 88% за 15 мин.\nPackShift держит форму >6 часов."),
        ("УГЛЕРОДНЫЙ СЛЕД (LCA, КГ CO₂E/КГ)", "0.82", "PackShift rPET",
         "Первичный алюминий: 11.5 кг CO₂.\nМногослойный ламинат: 3.8 кг CO₂.\nPackShift LignoShield: всего 0.39 кг (-90%)."),
        ("ЖИРОСТОЙКОСТЬ (TAPPI KIT TEST)", "Kit 12 / 12", "0% фторорганики (PFAS-Free)",
         "Стандартная бумага: Kit 3–4 (промокание).\nPackShift Biomimetic: Kit 12 без протечек горячего масла 120°C за 8 часов."),
    ]

    card_w = Inches(3.6)
    card_h = Inches(4.8)
    card_y = Inches(1.9)
    gap = Inches(0.26)
    start_x = Inches(1.0)

    for i, (tag, val, sub, desc) in enumerate(bench_cards):
        cx = start_x + i * (card_w + gap)
        card = slide5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, card_y, card_w, card_h)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD
        card.line.color.rgb = COLOR_CARD_BORDER
        card.line.width = Pt(1)

        tb_c = slide5.shapes.add_textbox(cx + Inches(0.15), card_y + Inches(0.3), card_w - Inches(0.3), card_h - Inches(0.6))
        tf_c = tb_c.text_frame
        tf_c.word_wrap = True

        p_t = tf_c.paragraphs[0]
        p_t.alignment = PP_ALIGN.CENTER
        p_t.text = tag
        p_t.font.name = FONT_MONO
        p_t.font.size = Pt(11)
        p_t.font.bold = True
        p_t.font.color.rgb = COLOR_TEXT_MUTED

        p_v = tf_c.add_paragraph()
        p_v.alignment = PP_ALIGN.CENTER
        p_v.text = f"\n{val}\n"
        p_v.font.name = FONT_MONO
        p_v.font.size = Pt(28)
        p_v.font.bold = True
        p_v.font.color.rgb = COLOR_NEON

        p_s = tf_c.add_paragraph()
        p_s.alignment = PP_ALIGN.CENTER
        p_s.text = f"{sub}\n"
        p_s.font.name = FONT_TITLE
        p_s.font.size = Pt(12)
        p_s.font.bold = True
        p_s.font.color.rgb = COLOR_TEXT_WHITE

        p_d = tf_c.add_paragraph()
        p_d.alignment = PP_ALIGN.CENTER
        p_d.text = desc
        p_d.font.name = FONT_BODY
        p_d.font.size = Pt(11)
        p_d.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 6: Industrial Integration (Zero-CapEx Drop-In)
    # =========================================================================
    slide6 = prs.slides.add_slide(blank_layout)
    set_bg(slide6)
    add_centered_header(slide6, "ПРОМЫШЛЕННАЯ СОВМЕСТИМОСТЬ • DROP-IN",
                        "Интеграция в существующие упаковочные линии",
                        "Масштабирование без закупки нового оборудования и переоснащения фасовочных фабрик", 6)

    # Pipeline Flow across 4 steps
    pipe_steps = [
        ("ШАГ 01", "Сырьевой поток", "Пищевые rPET-хлопья и жмых багассы"),
        ("ШАГ 02", "Компаундинг", "Нано-нуклеация и формование рулонов"),
        ("ШАГ 03", "Термоформинг", "Линии Multivac, Ulma, Ilpra, Variovac"),
        ("ШАГ 04", "Витрины Profi", "Выкладка в тепловых витринах до 120°C"),
    ]
    p_w = Inches(2.6)
    p_h = Inches(1.8)
    p_y = Inches(1.9)
    p_gap = Inches(0.3)
    p_start_x = Inches(1.0)

    for i, (step_num, title, sub) in enumerate(pipe_steps):
        cx = p_start_x + i * (p_w + p_gap)
        card = slide6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, p_y, p_w, p_h)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD_GLOW if i == 1 else COLOR_CARD
        card.line.color.rgb = COLOR_EMERALD if i == 1 else COLOR_CARD_BORDER
        card.line.width = Pt(1.5 if i == 1 else 1)

        tb_p = slide6.shapes.add_textbox(cx + Inches(0.1), p_y + Inches(0.15), p_w - Inches(0.2), p_h - Inches(0.3))
        tf_p = tb_p.text_frame
        tf_p.word_wrap = True

        p1 = tf_p.paragraphs[0]
        p1.alignment = PP_ALIGN.CENTER
        p1.text = step_num
        p1.font.name = FONT_MONO
        p1.font.size = Pt(10)
        p1.font.bold = True
        p1.font.color.rgb = COLOR_NEON

        p2 = tf_p.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        p2.text = title
        p2.font.name = FONT_TITLE
        p2.font.size = Pt(13)
        p2.font.bold = True
        p2.font.color.rgb = COLOR_TEXT_WHITE

        p3 = tf_p.add_paragraph()
        p3.alignment = PP_ALIGN.CENTER
        p3.text = sub
        p3.font.name = FONT_BODY
        p3.font.size = Pt(10)
        p3.font.color.rgb = COLOR_TEXT_MUTED

    # 3 Implementation factors
    factors = [
        ("ZERO CAPEX ДЛЯ РИТЕЙЛА", "Нулевые инвестиции в оборудование", 
         "Ни сети Profi, ни сторонним фасовочным фабрикам не нужно покупать новые станки за миллионы евро. Пленка заряжается в существующие бобины."),
        ("ПАРАМЕТРЫ ТЕРМОСВАРКИ", "Окно сварки: 145–160°C", 
         "Время цикла запайки составляет 1.1–1.3 сек. Это на 100% соответствует текущим скоростям промышленных фасовочных линий."),
        ("СЕРТИФИКАЦИЯ ПИЩЕВОГО КОНТАКТА", "EFSA & FDA Approved", 
         "Полное соблюдение директив EC 1935/2004 и EU 10/2011 по миграции веществ при температурах свыше 100°C."),
    ]
    f_w = Inches(3.6)
    f_h = Inches(2.7)
    f_y = Inches(4.1)

    for i, (tag, head, desc) in enumerate(factors):
        cx = start_x + i * (f_w + gap)
        card = slide6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, f_y, f_w, f_h)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD
        card.line.color.rgb = COLOR_CARD_BORDER
        card.line.width = Pt(1)

        tb_f = slide6.shapes.add_textbox(cx + Inches(0.15), f_y + Inches(0.2), f_w - Inches(0.3), f_h - Inches(0.4))
        tf_f = tb_f.text_frame
        tf_f.word_wrap = True

        p1 = tf_f.paragraphs[0]
        p1.text = tag
        p1.font.name = FONT_MONO
        p1.font.size = Pt(10)
        p1.font.bold = True
        p1.font.color.rgb = COLOR_NEON

        p2 = tf_f.add_paragraph()
        p2.text = f"\n{head}\n"
        p2.font.name = FONT_TITLE
        p2.font.size = Pt(13)
        p2.font.bold = True
        p2.font.color.rgb = COLOR_TEXT_WHITE

        p3 = tf_f.add_paragraph()
        p3.text = desc
        p3.font.name = FONT_BODY
        p3.font.size = Pt(11)
        p3.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 7: Scaled Unit Economics & PPWR Shield
    # =========================================================================
    slide7 = prs.slides.add_slide(blank_layout)
    set_bg(slide7)
    add_centered_header(slide7, "ЭКОНОМИКА МАСШТАБА • РЕГУЛЯТОРНАЯ БЕЗОПАСНОСТЬ",
                        "Unit-экономика и защита ритейлера от рисков",
                        "Финансовая модель окупаемости, страховка от штрафов EU PPWR и снижение списаний", 7)

    econ_nodes = [
        ("1 : 1", "Паритет себестоимости", "Выход на паритет с ископаемым пластиком благодаря дешевому rPET сырью"),
        ("€800 / т", "Защита от налога EPR", "Освобождение от штрафов за неперерабатываемую упаковку (EU Plastic Levy)"),
        ("-30%", "Снижение списаний", "Сохранение органолептики и корочки продлевает товарный вид блюд на 3+ ч"),
        ("Class A", "RecyClass & PPWR", "Опережение регуляторных директив ЕС на 4 года (PPWR 2030)"),
    ]
    card_w = Inches(2.6)
    card_h = Inches(2.2)
    card_y = Inches(1.9)
    gap = Inches(0.3)
    start_x = Inches(1.0)

    for i, (val, title, desc) in enumerate(econ_nodes):
        cx = start_x + i * (card_w + gap)
        card = slide7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, card_y, card_w, card_h)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD
        card.line.color.rgb = COLOR_CARD_BORDER
        card.line.width = Pt(1)

        tb_c = slide7.shapes.add_textbox(cx + Inches(0.1), card_y + Inches(0.15), card_w - Inches(0.2), card_h - Inches(0.3))
        tf_c = tb_c.text_frame
        tf_c.word_wrap = True

        p1 = tf_c.paragraphs[0]
        p1.alignment = PP_ALIGN.CENTER
        p1.text = val
        p1.font.name = FONT_MONO
        p1.font.size = Pt(26)
        p1.font.bold = True
        p1.font.color.rgb = COLOR_NEON

        p2 = tf_c.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        p2.text = title
        p2.font.name = FONT_TITLE
        p2.font.size = Pt(12)
        p2.font.bold = True
        p2.font.color.rgb = COLOR_TEXT_WHITE

        p3 = tf_c.add_paragraph()
        p3.alignment = PP_ALIGN.CENTER
        p3.text = desc
        p3.font.name = FONT_BODY
        p3.font.size = Pt(10)
        p3.font.color.rgb = COLOR_TEXT_MUTED

    # Big Banner for 100 Stores
    c_ban = slide7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(4.5), Inches(11.333), Inches(2.2))
    c_ban.fill.solid()
    c_ban.fill.fore_color.rgb = COLOR_CARD_GLOW
    c_ban.line.color.rgb = COLOR_EMERALD
    c_ban.line.width = Pt(1.5)

    tb_b = slide7.shapes.add_textbox(Inches(1.3), Inches(4.7), Inches(10.7), Inches(1.8))
    tf_b = tb_b.text_frame
    tf_b.word_wrap = True

    p = tf_b.paragraphs[0]
    p.text = "СОВОКУПНЫЙ ЭФФЕКТ ДЛЯ СЕТИ ИЗ 100 МАГАЗИНОВ PROFI:"
    p.font.name = FONT_MONO
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = COLOR_NEON

    p_m = tf_b.add_paragraph()
    p_m.text = "\nЭкономия свыше €1.98 млн в год на эко-сборах и сохраненной продукции"
    p_m.font.name = FONT_TITLE
    p_m.font.size = Pt(18)
    p_m.font.bold = True
    p_m.font.color.rgb = COLOR_TEXT_WHITE

    p_d = tf_b.add_paragraph()
    p_d.text = "\n• -1 840 тонн первичного пластика в год   • -4 250 тонн CO₂e предотвращенных выбросов   • 100% циркулярность"
    p_d.font.name = FONT_MONO
    p_d.font.size = Pt(12)
    p_d.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 8: Technology Readiness (TRL 4 -> 6) & The Ask
    # =========================================================================
    slide8 = prs.slides.add_slide(blank_layout)
    set_bg(slide8)
    add_centered_header(slide8, "ДОРОЖНАЯ КАРТА • ПАРТНЁРСТВО С PROFI",
                        "Статус готовности и предложение о сотрудничестве",
                        "От лабораторного подтверждения TRL 4 к запуску пилота в тепловых витринах сети", 8)

    # TRL Timeline Steps
    trl_steps = [
        ("ТЕКУЩИЙ СТАТУС", "TRL 4 (Лаборатория)", "Синтез компаунда подтвержден. 200 образцов готовы к тестам прямо сейчас"),
        ("НЕДЕЛИ 1–2", "TRL 5 (Тест витрин)", "Испытания в витринах Profi при 85°C–120°C с замером миграции"),
        ("НЕДЕЛИ 3–5", "TRL 6 (Пилот)", "Пилотная выкладка в 5 флагманских супермаркетах Profi"),
        ("НЕДЕЛЯ 6+", "TRL 7 (Масштаб)", "Подписание контракта на тиражирование по всей сети ритейлера"),
    ]
    p_w = Inches(2.6)
    p_h = Inches(1.8)
    p_y = Inches(1.9)
    p_gap = Inches(0.3)
    p_start_x = Inches(1.0)

    for i, (tag, head, desc) in enumerate(trl_steps):
        cx = p_start_x + i * (p_w + p_gap)
        card = slide8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, p_y, p_w, p_h)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD_GLOW if i == 0 else COLOR_CARD
        card.line.color.rgb = COLOR_EMERALD if i == 0 else COLOR_CARD_BORDER
        card.line.width = Pt(1.5 if i == 0 else 1)

        tb_t = slide8.shapes.add_textbox(cx + Inches(0.1), p_y + Inches(0.15), p_w - Inches(0.2), p_h - Inches(0.3))
        tf_t = tb_t.text_frame
        tf_t.word_wrap = True

        p1 = tf_t.paragraphs[0]
        p1.alignment = PP_ALIGN.CENTER
        p1.text = tag
        p1.font.name = FONT_MONO
        p1.font.size = Pt(10)
        p1.font.bold = True
        p1.font.color.rgb = COLOR_NEON

        p2 = tf_t.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        p2.text = head
        p2.font.name = FONT_TITLE
        p2.font.size = Pt(12)
        p2.font.bold = True
        p2.font.color.rgb = COLOR_TEXT_WHITE

        p3 = tf_t.add_paragraph()
        p3.alignment = PP_ALIGN.CENTER
        p3.text = desc
        p3.font.name = FONT_BODY
        p3.font.size = Pt(10)
        p3.font.color.rgb = COLOR_TEXT_MUTED

    # THE ASK Box
    c_ask = slide8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(4.1), Inches(11.333), Inches(2.6))
    c_ask.fill.solid()
    c_ask.fill.fore_color.rgb = COLOR_CARD_GLOW
    c_ask.line.color.rgb = COLOR_EMERALD
    c_ask.line.width = Pt(2)

    tb_a = slide8.shapes.add_textbox(Inches(1.3), Inches(4.3), Inches(10.7), Inches(2.2))
    tf_a = tb_a.text_frame
    tf_a.word_wrap = True

    p = tf_a.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.text = "НАШ ЗАПРОС К ЖЮРИ И СЕТИ PROFI (THE ASK):"
    p.font.name = FONT_MONO
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_NEON

    p_h = tf_a.add_paragraph()
    p_h.alignment = PP_ALIGN.CENTER
    p_h.text = "\nПроведение 2-недельного натурного тестирования готовой партии из 200 образцов в реальных тепловых витринах Profi"
    p_h.font.name = FONT_TITLE
    p_h.font.size = Pt(18)
    p_h.font.bold = True
    p_h.font.color.rgb = COLOR_TEXT_WHITE

    p_desc = tf_a.add_paragraph()
    p_desc.alignment = PP_ALIGN.CENTER
    p_desc.text = "\nМы предоставляем готовую упаковку и научную команду. Сеть Profi предоставляет доступ к витринам для валидации органолептики и термостойкости. Давайте создадим первую в Европе 100% циркулярную витрину горячей кулинарии вместе!"
    p_desc.font.name = FONT_BODY
    p_desc.font.size = Pt(12)
    p_desc.font.color.rgb = COLOR_TEXT_MUTED

    output_path = os.path.join(os.path.dirname(__file__), "PackShift_Profi_Gigahack2026.pptx")
    prs.save(output_path)
    print(f"Presentation successfully saved to: {output_path}")

if __name__ == "__main__":
    create_deck()

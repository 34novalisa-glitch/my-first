"""Builds the test-task presentation in Estate Invest brand colours."""
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

NAVY = RGBColor(0x0A, 0x19, 0x2C)
RED = RGBColor(0xEE, 0x27, 0x24)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT = RGBColor(0xF4, 0xF5, 0xF7)
GREY = RGBColor(0x5B, 0x66, 0x75)
MUTED = RGBColor(0xA9, 0xB3, 0xC1)
LINE = RGBColor(0xDD, 0xE1, 0xE6)
FONT = "Arial"

IN = 914400
W, H = int(13.333 * IN), int(7.5 * IN)

prs = Presentation()
prs.slide_width, prs.slide_height = W, H
BLANK = prs.slide_layouts[6]


def i(x):
    return int(x * IN)


def rect(slide, x, y, w, h, color, shape=MSO_SHAPE.RECTANGLE):
    s = slide.shapes.add_shape(shape, i(x), i(y), i(w), i(h))
    s.fill.solid()
    s.fill.fore_color.rgb = color
    s.line.fill.background()
    s.shadow.inherit = False
    return s


def text(slide, x, y, w, h, content, size=14, color=NAVY, bold=False,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=1.15, caps=False):
    """content: str or list of paragraphs; a paragraph may be a list of (text, bold, color) runs."""
    tb = slide.shapes.add_textbox(i(x), i(y), i(w), i(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    paras = content if isinstance(content, list) else [content]
    for n, para in enumerate(paras):
        p = tf.paragraphs[0] if n == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        if n:
            p.space_before = Pt(size * 0.45)
        runs = para if isinstance(para, list) else [(para, bold, color)]
        for r_text, r_bold, r_color in runs:
            r = p.add_run()
            r.text = r_text.upper() if caps else r_text
            r.font.name = FONT
            r.font.size = Pt(size)
            r.font.bold = r_bold
            r.font.color.rgb = r_color
    return tb


def logo(slide, x, y, s=0.5, color=RED):
    """Brand mark: three horizontal bars + one vertical bar."""
    bar_h = s * 0.22
    gap = (s - 3 * bar_h) / 2
    for k in range(3):
        rect(slide, x, y + k * (bar_h + gap), s * 0.72, bar_h, color)
    rect(slide, x + s * 0.82, y, s * 0.22, s, color)


def frame(title, num, section, dark=False):
    slide = prs.slides.add_slide(BLANK)
    bg = NAVY if dark else WHITE
    rect(slide, 0, 0, 13.333, 7.5, bg)
    fg = WHITE if dark else NAVY
    rect(slide, 0.6, 0.55, 0.5, 0.06, RED)
    text(slide, 0.6, 0.72, 9, 0.3, section, 11, MUTED if dark else GREY, caps=True)
    text(slide, 0.6, 1.0, 11.5, 0.8, title, 30, fg, bold=True)
    logo(slide, 12.35, 0.55, 0.38)
    text(slide, 11.6, 6.95, 1.13, 0.3, f"{num:02d}", 11, MUTED if dark else GREY,
         align=PP_ALIGN.RIGHT)
    text(slide, 0.6, 6.95, 6, 0.3, "Estate Invest · Unicorn Nova · тестовое задание", 10,
         MUTED if dark else GREY)
    return slide


def card(slide, x, y, w, h, head, body, dark=False, accent=RED, head_size=16, body_size=12.5,
         tag=None):
    rect(slide, x, y, w, h, NAVY if dark else LIGHT)
    rect(slide, x, y, 0.07, h, accent)
    cy = y + 0.25
    if tag:
        text(slide, x + 0.3, cy, w - 0.5, 0.3, tag, 10, RED, bold=True, caps=True)
        cy += 0.32
    text(slide, x + 0.3, cy, w - 0.5, 0.6, head, head_size, WHITE if dark else NAVY, bold=True)
    text(slide, x + 0.3, cy + 0.48 + (0.25 if len(head) > 34 and w < 4.5 else 0), w - 0.55,
         h - 1.0, body, body_size, MUTED if dark else GREY, spacing=1.2)


def table(slide, x, y, col_w, rows, head_size=11.5, body_size=11, row_h=None, first_bold=True):
    """rows[0] is the header. Drawn with shapes for full styling control."""
    total_w = sum(col_w)
    rh = row_h or [0.45] + [(6.6 - y - 0.45) / (len(rows) - 1)] * (len(rows) - 1)
    cy = y
    for r, row in enumerate(rows):
        h = rh[r]
        if r == 0:
            rect(slide, x, cy, total_w, h, NAVY)
        elif r % 2 == 0:
            rect(slide, x, cy, total_w, h, LIGHT)
        cx = x
        for c, val in enumerate(row):
            if r == 0:
                text(slide, cx + 0.15, cy, col_w[c] - 0.25, h, val, head_size, WHITE, bold=True,
                     anchor=MSO_ANCHOR.MIDDLE, caps=True)
            else:
                b = first_bold and c == 0
                text(slide, cx + 0.15, cy + 0.08, col_w[c] - 0.3, h - 0.12, val, body_size,
                     NAVY if b else GREY, bold=b, anchor=MSO_ANCHOR.MIDDLE, spacing=1.12)
            cx += col_w[c]
        cy += h
    rect(slide, x, cy, total_w, 0.03, RED)


# 1. Cover -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.333, 7.5, NAVY)
s.shapes.add_picture("assets/helmet.jpg", i(9.65), 0, height=H)
rect(s, 9.45, 0, 0.2, 7.5, RED)
logo(s, 0.8, 0.8, 0.6)
text(s, 0.8, 2.15, 8.3, 0.4, "Тестовое задание", 14, MUTED, caps=True)
text(s, 0.8, 2.6, 8.4, 1.9, ["SMM / Content Creator /", "Shooting & Editing Specialist"],
     38, WHITE, bold=True, spacing=1.05)
rect(s, 0.8, 4.55, 0.8, 0.06, RED)
text(s, 0.8, 4.8, 8, 0.5, "Estate Invest  ·  Unicorn Nova", 20, WHITE)
text(s, 0.8, 6.4, 8, 0.4, "Елизавета Краснова  ·  2026", 13, MUTED)

# 2. Content role ------------------------------------------------------
s = frame("Какую роль играет контент", 2, "01 · Контент-стратегия")
card(s, 0.6, 2.1, 5.9, 4.5, "Estate Invest",
     ["Контент работает как фильтр доверия.",
      "Человек покупает объект за $150k–1M+ в другой стране и часто ни разу его не видел. "
      "Ему нужно поверить в три вещи: рынок понятен, объект реален, люди за брендом — профессионалы.",
      [("Контент = первая встреча с агентом до звонка: ", True, NAVY),
       ("объясняет рынок, честно показывает объекты, отвечает на страхи (налоги, ВНЖ, рассрочка, перепродажа).",
        False, GREY)]],
     tag="международная недвижимость", body_size=16)
card(s, 6.83, 2.1, 5.9, 4.5, "Unicorn Nova",
     ["Контент работает как доказательство результата.",
      "Продукт — трансформация, а её лучше всего продаёт визуал: было → процесс → стало → цифры.",
      [("Контент = портфолио в движении: ", True, NAVY),
       ("экспертиза команды, качество работы и логика «вложил в ремонт → продал или сдал дороже».",
        False, GREY)]],
     tag="интерьерные инвестпроекты", body_size=16)

# 3. Differences -------------------------------------------------------
s = frame("Чем отличается подача", 3, "01 · Контент-стратегия")
table(s, 0.6, 2.05, [2.6, 4.75, 4.75], [
    ["", "Estate Invest", "Unicorn Nova"],
    ["Главная эмоция", "Спокойствие и уверенность: «я понимаю, куда вкладываю»",
     "Восхищение и азарт: «смотри, что из этого сделали»"],
    ["Тон", "Экспертный, аналитичный, премиальный", "Живой, визуальный, «на объекте»"],
    ["Герой контента", "Рынок, город, объект, эксперт-агент", "Процесс, команда, детали, до/после"],
    ["Визуал", "Дрон, панорамы, виды, лайфстайл локации",
     "Крупные планы материалов, таймлапсы, reveal"],
    ["Ключевой формат", "Разборы, сравнения, обзоры-туры", "Трансформации, backstage, сторис со стройки"],
], body_size=13)

# 4. Goals -------------------------------------------------------------
s = frame("3 ключевые цели соцсетей", 4, "01 · Контент-стратегия", dark=True)
goals = [
    ("01", "Доверие", "Лица команды, реальные объекты и сделки, прозрачные цифры.",
     "сохранения · время просмотра · повторные визиты в профиль"),
    ("02", "Лиды", "Охват → заявки через лид-магниты (гайд, подборка, калькулятор) "
                   "и ключевые слова в Direct.", "заявки · переходы · CPL"),
    ("03", "Узнаваемость в нише", "Стать аккаунтом, который присылают друг другу: "
                                  "«скинь, тут понятно объяснили про Дубай».",
     "репосты · охват не по подписчикам · рост аудитории"),
]
for k, (n, h, b, m) in enumerate(goals):
    x = 0.6 + k * 4.1
    rect(s, x, 2.2, 3.85, 4.3, RGBColor(0x13, 0x26, 0x3F))
    text(s, x + 0.35, 2.45, 2, 0.8, n, 40, RED, bold=True)
    text(s, x + 0.35, 3.4, 3.2, 0.5, h, 20, WHITE, bold=True)
    text(s, x + 0.35, 4.0, 3.2, 1.4, b, 13.5, MUTED, spacing=1.2)
    rect(s, x + 0.35, 5.5, 3.15, 0.02, GREY)
    text(s, x + 0.35, 5.62, 3.2, 0.7, m, 11, WHITE, caps=True, spacing=1.2)

# 5. Weekly plan -------------------------------------------------------
s = frame("Контент-план на неделю: Estate Invest", 5, "02 · Контент-план")
text(s, 0.6, 1.7, 12, 0.35,
     [[("Ось недели: ", True, RED), ("«Дубай или Турция: куда вкладывать в 2027»", False, GREY)]], 13)
table(s, 0.6, 2.2, [0.75, 1.75, 2.2, 7.4], [
    ["День", "Тип", "Формат", "Тема и смысл"],
    ["Пн", "Экспертный", "Карусель 8 слайдов",
     "«Дубай vs Анталья/Алания: честное сравнение». Цена входа, доходность, налоги, ВНЖ. CTA: «СРАВНЕНИЕ» в Direct"],
    ["Вт", "Вовлекающий", "Reels", "«Угадай цену квартиры»: тур 15 секунд → «пиши в комменты» → ответ в конце"],
    ["Ср", "Доверие", "Reels", "«Один день брокера в Дубае»: показ, стройка, переговоры, закат из пентхауса"],
    ["Чт", "Экспертный", "Карусель", "«5 ошибок при покупке off-plan»: застройщик, план платежей, сервисный сбор"],
    ["Пт", "Продающий", "Карусель фото + видео",
     "«Объект недели»: 3 причины купить, рассрочка, локация. CTA: «ОБЪЕКТ» в Direct"],
    ["Сб", "Лайфстайл", "Reels", "«Выходной, если ты живёшь в Алании»: продаём не метры, а жизнь"],
], body_size=11.5)

# 6. Stories -----------------------------------------------------------
s = frame("Сторис-план на 4 дня", 6, "02 · Контент-план")
days = [
    ("Пн", "Поддержка карусели", ["Анонс разбора + ссылка", "Опрос: Дубай / Турция",
                                   "Слайдер: «готовы к покупке?» 🏠", "Итоги + тизер на завтра"]),
    ("Ср", "Backstage", ["«Едем на показ, угадай район»", "Видео со стройки + гео",
                         "Квиз: сколько этажей в башне?", "Окно вопросов о покупке"]),
    ("Пт", "Продажа", ["Тизер: объект, который бронируют быстрее всех", "Видео-тур 3 × 15 секунд",
                       "Квиз: минимальный первый взнос?", "Ссылка «Получить подборку»",
                       "Отсчёт до эфира"]),
    ("Вс", "Лайфстайл", ["Закат из объекта + музыка", "Опрос: море или город?",
                         "Ответы на вопросы голосом", "Анонс темы недели"]),
]
for k, (d, h, items) in enumerate(days):
    x = 0.6 + k * 3.07
    rect(s, x, 2.1, 2.88, 0.85, NAVY)
    text(s, x + 0.25, 2.1, 0.8, 0.85, d, 22, RED, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    text(s, x + 0.95, 2.1, 1.85, 0.85, h, 12.5, WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    rect(s, x, 2.95, 2.88, 3.65, LIGHT)
    for j, it in enumerate(items):
        yy = 3.15 + j * 0.68
        rect(s, x + 0.25, yy + 0.08, 0.1, 0.1, RED)
        text(s, x + 0.5, yy, 2.25, 0.65, it, 11.5, NAVY, spacing=1.1)

# 7. Reels ideas -------------------------------------------------------
s = frame("Идеи Reels / Shorts", 7, "02 · Контент-план", dark=True)
reels = [
    ("$300 000 в Дубае vs $300 000 в Алании",
     "Сплит-экран: метраж, вид, бассейн, доход от аренды за одну сумму. Финал: «А ты бы что выбрал(а)?»",
     "сравнения собирают споры в комментариях"),
    ("«Угадай цену»",
     "Динамичный тур, пауза, ответ. Повторяемая рубрика-игра, ради которой возвращаются.",
     "комментарии + досмотры"),
    ("Off-plan за 30 секунд",
     "Эксперт в кадре + графика. Короткий ликбез на один вопрос.",
     "сохранения и пересылки"),
]
for k, (h, b, why) in enumerate(reels):
    x = 0.6 + k * 4.1
    rect(s, x, 2.1, 3.85, 4.4, RGBColor(0x13, 0x26, 0x3F))
    rect(s, x + 0.35, 2.4, 0.55, 0.9, RED, MSO_SHAPE.ROUNDED_RECTANGLE)
    text(s, x + 0.35, 2.4, 0.55, 0.9, "9:16", 9, WHITE, bold=True, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE)
    text(s, x + 0.35, 3.55, 3.2, 0.9, h, 17, WHITE, bold=True)
    text(s, x + 0.35, 4.45, 3.2, 1.3, b, 12.5, MUTED, spacing=1.2)
    text(s, x + 0.35, 5.9, 3.2, 0.5, [[("→ ", True, RED), (why, False, WHITE)]], 11.5)

# 8. Platforms ---------------------------------------------------------
s = frame("Платформенная адаптация", 8, "03 · Платформы")
plats = [
    ("Instagram", "Эстетика и доверие", "Карусели 4:5 · Reels 9:16 · сторис",
     "Красиво, структурно, «сохрани себе». Кейсы, туры, интерактивы в сторис."),
    ("TikTok", "Нативность и эмоция", "Вертикаль 9:16 · 15–45 секунд · субтитры",
     "Просто, быстро, с характером, как от друга, а не бренда. Тренды, реакции, споры."),
    ("YouTube", "Глубина и поиск", "Горизонталь 16:9, 8–20 минут + Shorts",
     "Подробно и спокойно: главы, цифры на экране. Shorts — воронка в длинные видео."),
]
for k, (p, what, fmt, how) in enumerate(plats):
    x = 0.6 + k * 4.1
    rect(s, x, 2.1, 3.85, 0.9, NAVY)
    text(s, x + 0.3, 2.1, 3.3, 0.9, p, 20, WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    rect(s, x, 3.0, 3.85, 3.6, LIGHT)
    for j, (lab, val) in enumerate([("Что работает", what), ("Форматы", fmt), ("Подача", how)]):
        yy = 3.2 + j * 1.1
        text(s, x + 0.3, yy, 3.3, 0.3, lab, 10, RED, bold=True, caps=True)
        text(s, x + 0.3, yy + 0.3, 3.3, 0.8, val, 12, NAVY, spacing=1.12)

# 9. One idea three ways -----------------------------------------------
s = frame("Одна идея в трёх подачах", 9, "03 · Платформы")
text(s, 0.6, 1.7, 12, 0.35, [[("Смысл: ", True, RED),
     ("«Сколько реально стоит владеть квартирой в Дубае, кроме цены покупки»", False, GREY)]], 13)
ways = [
    ("Instagram", "Карусель, 7 слайдов",
     "Обложка: «Купил квартиру в Дубае. Что дальше платить?» → по слайду на статью расходов → "
     "итог в год → «Сохрани, пригодится»."),
    ("TikTok", "Reels, 25 секунд",
     "Брокер в кадре: «Тебе сказали, что квартира стоит 1,2 млн дирхам? Это не всё…» → "
     "быстрые цифры → «Напиши бюджет, посчитаю»."),
    ("YouTube", "Видео, 12 минут + 3 Shorts",
     "«Полная стоимость владения в Дубае: реальные цифры». Главы, таблица, сравнение районов, "
     "интервью с управляющей компанией."),
]
for k, (p, f, b) in enumerate(ways):
    card(s, 0.6 + k * 4.1, 2.25, 3.85, 4.35, p, [[(f, True, RED)], b], body_size=15.5)

# 10-11. Creative --------------------------------------------------------
creative = [
    ("Estate Invest", "«Ключи по видеосвязи»",
     "«Он купил квартиру в Дубае и впервые увидел её сегодня»",
     "Человек открывает дверь, камера за спиной, свет из панорамного окна",
     "Реальная история клиента (с согласия), купившего объект удалённо: звонок → видеопоказы → "
     "сделка → день, когда он заходит в квартиру. Финал — реакция и одна фраза.",
     "Закрывает главный страх («как купить, не видя?») через эмоцию, а не аргументы. "
     "Сторителлинг с героем досматривают. Можно сделать рубрикой «Истории клиентов»."),
    ("Unicorn Nova", "«Квартира, которую никто не хотел»",
     "«Эту квартиру 8 месяцев никто не покупал. Мы купили и сделали вот это»",
     "Облезлые стены, старый линолеум, тусклый свет",
     "Трансформация за 30–40 секунд: честное «до» → ускоренный процесс → reveal под музыку "
     "с прохода из той же точки → цифры: покупка, ремонт, срок, цена продажи.",
     "«До/после» — один из самых досматриваемых форматов. Цифры превращают красивый ролик "
     "в инвесткейс, а сериал по неделям возвращает аудиторию."),
]
for k, (proj, name, hook, vis, desc, why) in enumerate(creative):
    s = frame(f"{proj}: {name}", 10 + k, "04 · Креативный блок")
    rect(s, 0.6, 2.05, 12.13, 1.45, NAVY)
    rect(s, 0.6, 2.05, 0.1, 1.45, RED)
    text(s, 0.95, 2.18, 3, 0.3, "Хук", 10, RED, bold=True, caps=True)
    text(s, 0.95, 2.48, 11.5, 0.9, hook, 22, WHITE, bold=True)
    for j, (lab, val) in enumerate([("Визуал", vis), ("Описание", desc), ("Почему зайдёт", why)]):
        x = 0.6 + j * 4.1
        rect(s, x, 3.75, 3.85, 0.04, RED if j == 2 else NAVY)
        text(s, x, 3.95, 3.85, 0.3, lab, 11, RED, bold=True, caps=True)
        text(s, x, 4.3, 3.85, 2.3, val, 15, NAVY if j == 0 else GREY, spacing=1.22)

# 12. Skills -----------------------------------------------------------
s = frame("Навыки визуала и продакшена", 12, "05 · Навыки")
skills = [
    ("Дизайн", ["Figma: макеты каруселей и сторис, сетка профиля, компоненты и стили",
                "Canva: бренд-кит, шаблоны для потока, Magic Resize",
                "After Effects: моушн, анимированный текст, инфографика"]),
    ("Монтаж", ["CapCut: вертикальные видео, субтитры, тренды, бит-монтаж",
                "Premiere Pro, DaVinci Resolve: длинные ролики, цветокоррекция",
                "AI-инструменты для визуалов, видео и текстов"]),
    ("Съёмка", ["iPhone, полный цикл: люди, «говорящая голова», лайфстайл, детали",
                "YouTube-проект о путешествиях: ~7 тыс. подписчиков, 500+ видео",
                "Живу в Дубае: снимаю объекты на месте"]),
]
for k, (h, items) in enumerate(skills):
    x = 0.6 + k * 4.1
    rect(s, x, 2.1, 3.85, 4.5, LIGHT)
    rect(s, x, 2.1, 3.85, 0.07, RED)
    text(s, x + 0.3, 2.4, 3.3, 0.5, h, 20, NAVY, bold=True)
    for j, it in enumerate(items):
        yy = 3.15 + j * 1.1
        rect(s, x + 0.3, yy + 0.08, 0.1, 0.1, RED)
        text(s, x + 0.55, yy, 3.05, 1.0, it, 12.5, GREY, spacing=1.15)

# 13. Shooting day timing ---------------------------------------------
s = frame("Съёмочный день: объект / интерьер", 13, "05 · Продакшен")
text(s, 0.6, 1.7, 12, 0.35, [[("Накануне: ", True, RED),
     ("бриф и цель · shot list и раскадровка хука · свет по сторонам света · стайлинг объекта · "
      "техника, стабилизатор, дрон", False, GREY)]], 12)
table(s, 0.6, 2.2, [2.4, 9.73], [
    ["Время", "Задача"],
    ["08:00–09:30", "Мягкий утренний свет: экстерьер, фасад, виды, дрон"],
    ["09:30–12:00", "Интерьер: общие планы комнат (16:9 + 9:16), проходы на стабилизаторе"],
    ["12:00–13:00", "Детали: материалы, фурнитура, техника, текстуры"],
    ["13:00–14:00", "Перерыв и проверка отснятого"],
    ["14:00–15:30", "Лайфстайл с моделью / агентом: кофе на балконе, вход в квартиру"],
    ["15:30–16:30", "«Говорящая голова» агента: обзор, 3 причины, CTA"],
    ["16:30–18:00", "Инфраструктура: бассейн, лобби, спортзал, пляж, район"],
    ["Золотой час", "Виды на закате, вечерний свет в интерьере, сторис в реальном времени"],
], body_size=12)

# 14. Shots & accents --------------------------------------------------
s = frame("Обязательные кадры и акценты", 14, "05 · Продакшен", dark=True)
shots = ["Хук: открытие двери / проход к окну / дрон-подлёт",
         "Общий план каждой комнаты из угла",
         "Плавный тур одним дублем",
         "Вид из окна днём и на закате",
         "10–15 деталей: кухня, ванная, материалы",
         "Лобби, бассейн, фитнес, парковка",
         "Район: пляж, метро, кафе рядом",
         "Люди: агент и лайфстайл-модель",
         "Backstage для сторис"]
text(s, 0.6, 1.95, 5, 0.3, "Обязательные кадры", 12, RED, bold=True, caps=True)
for j, sh in enumerate(shots):
    yy = 2.4 + j * 0.47
    text(s, 0.6, yy, 0.5, 0.4, f"{j + 1:02d}", 12, RED, bold=True)
    text(s, 1.1, yy, 4.9, 0.4, sh, 13, WHITE)
acc = [("Свет и воздух", "естественный свет, светлый цвет, окна без пересвета (HDR)"),
       ("Ровная геометрия", "прямые вертикали, камера на уровне груди"),
       ("Эмоция места", "не метры, а «как здесь живут»: утро, кофе, вечер"),
       ("Премиальность", "панорамы, медленные движения, атмосферный звук"),
       ("Детали качества", "материалы крупно подтверждают цену"),
       ("Единый стиль", "фирменные цвета и шрифты, одинаковые переходы")]
text(s, 6.7, 1.95, 6, 0.3, "Акценты в визуале", 12, RED, bold=True, caps=True)
for j, (h, b) in enumerate(acc):
    yy = 2.4 + j * 0.7
    rect(s, 6.7, yy, 6.03, 0.6, RGBColor(0x13, 0x26, 0x3F))
    rect(s, 6.7, yy, 0.06, 0.6, RED)
    text(s, 6.95, yy, 5.7, 0.6, [[(h + " — ", True, WHITE), (b, False, MUTED)]], 12,
         anchor=MSO_ANCHOR.MIDDLE)

# 15. Bonus ------------------------------------------------------------
s = frame("Бонус: англоязычная страница Estate Invest", 15, "06 · Бонус")
text(s, 0.6, 1.95, 5.8, 0.3, "Что изменю в подаче", 12, RED, bold=True, caps=True)
left = [("Данные вместо «мечты»", "ROI, rental yield, service charges, сравнение с рынком страны инвестора"),
        ("Лица и экспертиза", "эксперт в кадре на хорошем английском: market updates, Q&A, законы"),
        ("Корректность", "без «гарантированной доходности», ссылки на DLD / RERA, дисклеймеры"),
        ("LinkedIn + YouTube", "для B2B и состоятельной аудитории: отчёты, кейсы, обзоры районов")]
right = [("Язык", "живой английский, а не перевод; C1/C2 (CAE), 10 лет преподавания: пишу и вычитываю сама"),
         ("Сегменты", "«Buying in Dubai from the UK», «Golden Visa explained», «Istanbul vs Dubai»"),
         ("Нейтральность", "без локальных мемов; субтитры всегда; цены в USD / AED / EUR"),
         ("Воронка", "PDF-гайд на английском, автоответы в Direct, звонок в любом часовом поясе")]
text(s, 6.93, 1.95, 5.8, 0.3, "Как адаптирую стиль", 12, RED, bold=True, caps=True)
for col, items in ((0.6, left), (6.93, right)):
    for j, (h, b) in enumerate(items):
        yy = 2.4 + j * 1.07
        rect(s, col, yy, 5.8, 0.95, LIGHT)
        rect(s, col, yy, 0.06, 0.95, RED)
        text(s, col + 0.3, yy + 0.13, 5.3, 0.3, h, 13.5, NAVY, bold=True)
        text(s, col + 0.3, yy + 0.45, 5.3, 0.5, b, 11.5, GREY, spacing=1.1)

# 16. Final ------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.333, 7.5, NAVY)
rect(s, 0, 7.3, 13.333, 0.2, RED)
logo(s, 6.27, 1.7, 0.8)
text(s, 0.6, 3.0, 12.13, 1.0, "Спасибо!", 44, WHITE, bold=True, align=PP_ALIGN.CENTER)
text(s, 0.6, 4.1, 12.13, 0.5, "Готова подготовить макеты каруселей, сторис и первый съёмочный день",
     16, MUTED, align=PP_ALIGN.CENTER)
text(s, 0.6, 5.3, 12.13, 0.4, "Елизавета Краснова", 14, WHITE, bold=True, align=PP_ALIGN.CENTER)

prs.save("test-task-presentation.pptx")
print("saved", len(prs.slides), "slides")

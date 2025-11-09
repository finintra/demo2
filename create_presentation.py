#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import requests
from io import BytesIO

# Створюємо презентацію
prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(7.5)

# Функція для додавання слайду
def add_title_slide(title, subtitle, bg_color=(15, 12, 41)):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout

    # Background
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(*bg_color)

    # Title
    title_box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(2))
    tf = title_box.text_frame
    tf.text = title
    tf.paragraphs[0].font.size = Pt(54)
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER

    # Subtitle
    if subtitle:
        subtitle_box = slide.shapes.add_textbox(Inches(1), Inches(4.5), Inches(8), Inches(1))
        stf = subtitle_box.text_frame
        stf.text = subtitle
        stf.paragraphs[0].font.size = Pt(28)
        stf.paragraphs[0].font.color.rgb = RGBColor(200, 200, 200)
        stf.paragraphs[0].alignment = PP_ALIGN.CENTER

    return slide

# Slide 1: Title
add_title_slide(
    "Low-Code та n8n:",
    "Ваша нова суперсила (без жодного рядка коду)"
)

# Slide 2: Вам це знайомо?
slide = prs.slides.add_slide(prs.slide_layouts[6])
background = slide.background.fill
background.solid()
background.fore_color.rgb = RGBColor(249, 250, 251)

title = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
tf = title.text_frame
tf.text = "Вам це знайомо?"
tf.paragraphs[0].font.size = Pt(44)
tf.paragraphs[0].font.bold = True
tf.paragraphs[0].font.color.rgb = RGBColor(31, 41, 55)

# Bullets
content = slide.shapes.add_textbox(Inches(0.5), Inches(1.8), Inches(5.5), Inches(4))
tf = content.text_frame
problems = [
    "✓ Копіювати дані з сайту в Google Таблицю?",
    "✓ Вручну перевіряти пошту кожні 15 хвилин?",
    "✓ Губити дедлайни, що прийшли на email?",
    "✓ Робити ту саму монотонну дію знову і знову?"
]
for i, problem in enumerate(problems):
    p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
    p.text = problem
    p.font.size = Pt(20)
    p.font.color.rgb = RGBColor(55, 65, 81)
    p.space_after = Pt(15)

conclusion = slide.shapes.add_textbox(Inches(0.5), Inches(6), Inches(5.5), Inches(0.8))
tf = conclusion.text_frame
tf.text = "Це називається... Рутина."
tf.paragraphs[0].font.size = Pt(24)
tf.paragraphs[0].font.bold = True
tf.paragraphs[0].font.color.rgb = RGBColor(220, 38, 38)

# Slide 3: Що таке Low-Code / No-Code?
slide = prs.slides.add_slide(prs.slide_layouts[6])
background = slide.background.fill
background.solid()
background.fore_color.rgb = RGBColor(239, 246, 255)

title = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
tf = title.text_frame
tf.text = "Що таке Low-Code / No-Code?"
tf.paragraphs[0].font.size = Pt(40)
tf.paragraphs[0].font.bold = True
tf.paragraphs[0].font.color.rgb = RGBColor(31, 41, 55)
tf.paragraphs[0].alignment = PP_ALIGN.CENTER

subtitle = slide.shapes.add_textbox(Inches(0.5), Inches(1.3), Inches(9), Inches(0.5))
stf = subtitle.text_frame
stf.text = "Уявіть, що ви будуєте будинок:"
stf.paragraphs[0].font.size = Pt(20)
stf.paragraphs[0].font.color.rgb = RGBColor(107, 114, 128)
stf.paragraphs[0].alignment = PP_ALIGN.CENTER

# Three boxes
boxes = [
    {"title": "High-Code", "text": "Ви самі виготовляєте кожну цеглину", "note": "Потрібні роки навчання", "color": (239, 68, 68), "left": 0.5},
    {"title": "No-Code", "text": "Ви збираєте будинок з готових блоків LEGO", "note": "Швидко і просто", "color": (34, 197, 94), "left": 3.5},
    {"title": "Low-Code", "text": "LEGO + пластилін для унікальних деталей", "note": "Гнучкість + простота", "color": (168, 85, 247), "left": 6.5}
]

for box in boxes:
    # Box
    box_shape = slide.shapes.add_shape(1, Inches(box["left"]), Inches(2.5), Inches(2.8), Inches(3))
    box_shape.fill.solid()
    box_shape.fill.fore_color.rgb = RGBColor(255, 255, 255)
    box_shape.line.color.rgb = RGBColor(*box["color"])
    box_shape.line.width = Pt(3)

    # Title
    title_box = slide.shapes.add_textbox(Inches(box["left"] + 0.2), Inches(2.7), Inches(2.4), Inches(0.5))
    tf = title_box.text_frame
    tf.text = box["title"]
    tf.paragraphs[0].font.size = Pt(20)
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = RGBColor(*box["color"])

    # Text
    text_box = slide.shapes.add_textbox(Inches(box["left"] + 0.2), Inches(3.3), Inches(2.4), Inches(1.5))
    tf = text_box.text_frame
    tf.text = box["text"]
    tf.word_wrap = True
    tf.paragraphs[0].font.size = Pt(14)
    tf.paragraphs[0].font.color.rgb = RGBColor(75, 85, 99)

    # Note
    note_box = slide.shapes.add_textbox(Inches(box["left"] + 0.2), Inches(4.9), Inches(2.4), Inches(0.4))
    tf = note_box.text_frame
    tf.text = box["note"]
    tf.paragraphs[0].font.size = Pt(11)
    tf.paragraphs[0].font.italic = True
    tf.paragraphs[0].font.color.rgb = RGBColor(107, 114, 128)

# Conclusion
conclusion = slide.shapes.add_textbox(Inches(1.5), Inches(6.2), Inches(7), Inches(0.8))
tf = conclusion.text_frame
tf.text = "n8n — це ваш набір LEGO та пластиліну для автоматизацій"
tf.paragraphs[0].font.size = Pt(22)
tf.paragraphs[0].font.bold = True
tf.paragraphs[0].font.color.rgb = RGBColor(126, 34, 206)
tf.paragraphs[0].alignment = PP_ALIGN.CENTER

# Функція для створення слайду-прикладу
def add_example_slide(number, title, steps, result):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide.background.fill
    background.solid()
    background.fore_color.rgb = RGBColor(249, 250, 251)

    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    tf = title_box.text_frame
    tf.text = f'Приклад {number}: "{title}"'
    tf.paragraphs[0].font.size = Pt(38)
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = RGBColor(31, 41, 55)
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER

    # Steps
    step_colors = [(59, 130, 246), (249, 115, 22), (34, 197, 94), (168, 85, 247), (234, 179, 8)]
    box_width = (10 - 0.5 * 2 - 0.3 * (len(steps) - 1)) / len(steps)

    for i, step in enumerate(steps):
        left = 0.5 + i * (box_width + 0.3)

        # Box
        box = slide.shapes.add_shape(1, Inches(left), Inches(2), Inches(box_width), Inches(2.5))
        box.fill.solid()
        box.fill.fore_color.rgb = RGBColor(255, 255, 255)
        color = step_colors[i % len(step_colors)]
        box.line.color.rgb = RGBColor(*color)
        box.line.width = Pt(3)

        # Step type
        type_box = slide.shapes.add_textbox(Inches(left + 0.1), Inches(2.2), Inches(box_width - 0.2), Inches(0.4))
        tf = type_box.text_frame
        tf.text = step["type"]
        tf.paragraphs[0].font.size = Pt(16)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.color.rgb = RGBColor(*color)

        # Description
        desc_box = slide.shapes.add_textbox(Inches(left + 0.1), Inches(2.7), Inches(box_width - 0.2), Inches(1.5))
        tf = desc_box.text_frame
        tf.text = step["desc"]
        tf.word_wrap = True
        tf.paragraphs[0].font.size = Pt(12)
        tf.paragraphs[0].font.color.rgb = RGBColor(75, 85, 99)

        # Arrow
        if i < len(steps) - 1:
            arrow_left = left + box_width + 0.05
            arrow = slide.shapes.add_textbox(Inches(arrow_left), Inches(3.2), Inches(0.2), Inches(0.3))
            tf = arrow.text_frame
            tf.text = "→"
            tf.paragraphs[0].font.size = Pt(24)
            tf.paragraphs[0].font.color.rgb = RGBColor(156, 163, 175)

    # Result
    result_box = slide.shapes.add_textbox(Inches(1), Inches(5.5), Inches(8), Inches(1))
    result_box.fill.solid()
    result_box.fill.fore_color.rgb = RGBColor(219, 234, 254)
    tf = result_box.text_frame
    tf.text = f"✓ Результат: {result}"
    tf.paragraphs[0].font.size = Pt(20)
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = RGBColor(37, 99, 235)
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER

# Examples 1-8
examples = [
    {
        "number": 1,
        "title": "Розумний список завдань",
        "steps": [
            {"type": "ТРИГЕР", "desc": "n8n перевіряє Gmail кожні 30 хв"},
            {"type": "ЛОГІКА", "desc": "Якщо лист від викладача містить 'дедлайн'"},
            {"type": "ДІЯ 1", "desc": "Створити завдання в Trello/Notion"},
            {"type": "ДІЯ 2", "desc": "Створити подію в Google Календарі"}
        ],
        "result": "Ви ніколи не пропустите дедлайн"
    },
    {
        "number": 2,
        "title": "Куратор контенту",
        "steps": [
            {"type": "ТРИГЕР", "desc": "Кожного ранку о 9:00"},
            {"type": "ДІЯ 1", "desc": "Прочитати нові статті з RSS"},
            {"type": "ДІЯ 2", "desc": "AI вибирає 3 найкращі"},
            {"type": "ДІЯ 3", "desc": "Надіслати на Email/Slack"}
        ],
        "result": "Година роботи → 5 хвилин"
    },
    {
        "number": 3,
        "title": "Збір відгуків",
        "steps": [
            {"type": "ТРИГЕР", "desc": "Заповнено Google Форму"},
            {"type": "ДІЯ 1", "desc": "Записати в Google Таблицю"},
            {"type": "ДІЯ 2", "desc": "AI аналізує: позитивний/негативний?"},
            {"type": "ЛОГІКА", "desc": "Якщо негативний → email"}
        ],
        "result": "Миттєва реакція на проблеми"
    },
    {
        "number": 4,
        "title": "Моніторинг вакансій",
        "steps": [
            {"type": "ТРИГЕР", "desc": "Кожні 4 години"},
            {"type": "ДІЯ 1", "desc": "Перевірити RSS Work.ua, DOU"},
            {"type": "ЛОГІКА", "desc": "Якщо є Junior/Intern"},
            {"type": "ДІЯ 2", "desc": "Додати в Google Таблицю"}
        ],
        "result": "Ваша особиста дошка вакансій"
    },
    {
        "number": 5,
        "title": "Синхронізація Notion/Календаря",
        "steps": [
            {"type": "ТРИГЕР", "desc": "Кожну годину"},
            {"type": "ЛОГІКА", "desc": "Новий запис з датою дедлайну?"},
            {"type": "ДІЯ 1", "desc": "Взяти назву та дату з Notion"},
            {"type": "ДІЯ 2", "desc": "Створити подію в Календарі"}
        ],
        "result": "Календар = ваші плани в Notion"
    },
    {
        "number": 6,
        "title": "Авто-'Список для читання'",
        "steps": [
            {"type": "ТРИГЕР", "desc": "Поставили зірочку листу в Gmail"},
            {"type": "ДІЯ 1", "desc": "Взяти посилання з листа"},
            {"type": "ДІЯ 2", "desc": "AI створює короткий підсумок"},
            {"type": "ДІЯ 3", "desc": "Зберегти в Notion"}
        ],
        "result": "База знань, а не звалище посилань"
    },
    {
        "number": 7,
        "title": "Трекер згадок у соцмережах",
        "steps": [
            {"type": "ТРИГЕР", "desc": "Кожні 6 годин"},
            {"type": "ДІЯ 1", "desc": "Шукати на Reddit ваш проєкт"},
            {"type": "ЛОГІКА", "desc": "Знайдено новий пост?"},
            {"type": "ДІЯ 2", "desc": "Надіслати в Discord/Slack"}
        ],
        "result": "Миттєва реакція на згадки"
    },
    {
        "number": 8,
        "title": "Ранковий бриф",
        "steps": [
            {"type": "ТРИГЕР", "desc": "Кожен день о 7:00"},
            {"type": "ДІЯ 1", "desc": "Взяти прогноз погоди"},
            {"type": "ДІЯ 2", "desc": "Взяти події з Календаря"},
            {"type": "ДІЯ 3", "desc": "Надіслати один Email"}
        ],
        "result": "Один лист з планом на день"
    }
]

for example in examples:
    add_example_slide(example["number"], example["title"], example["steps"], example["result"])

print("Створено слайди 1-11 (титульний + 8 прикладів)")

# Збереження
prs.save('/home/user/demo2/n8n-presentation.pptx')
print("✓ Презентацію збережено як n8n-presentation.pptx")
print(f"Всього слайдів: {len(prs.slides)}")

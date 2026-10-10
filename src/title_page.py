"""Титульный лист.

Два режима:
  1. Программный (fallback) -- собирается "с нуля" строго по шаблону ЛЭТИ.
  2. Из шаблона -- загружается пользовательский .docx (например,
     официальный титульник кафедры) и подставляются значения через
     docxtpl: Jinja2-плейсхолдеры вида `{{ key }}`.
"""
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT

from .inline import set_run_font


# ---------------------------------------------------------------------
#  Программный титульник (fallback, по шаблону ЛЭТИ)
# ---------------------------------------------------------------------
def _para(doc, text="", size=14, bold=False,
          align=WD_ALIGN_PARAGRAPH.CENTER, space_after=0):
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.line_spacing = 1.0
    pf.space_before = Pt(0)
    pf.space_after = Pt(space_after)
    if text:
        run = p.add_run(text)
        set_run_font(run, size=size, bold=bold)
    return p


def add_title_page(doc, meta):
    """Программный титульный лист по шаблону ЛЭТИ."""

    # --- Шапка ------------------------------------------------------
    _para(doc, meta.get("ministry", "МИНОБРНАУКИ РОССИИ"), size=12)
    for line in (meta.get("university") or "").split("\n"):
        line = line.strip()
        if line:
            _para(doc, line, size=12)
    dept = meta.get("department") or ""
    if dept:
        _para(doc, "Кафедра {0}".format(dept), size=12)

    # --- Пустое место -----------------------------------------------
    for _ in range(6):
        _para(doc, "")

    # --- Заголовок --------------------------------------------------
    _para(doc, "ОТЧЕТ", size=14, bold=True)
    _para(doc, "по лабораторной работе № {0}".format(
        meta.get("lab_number", "")), size=14)
    _para(doc, "по дисциплине «{0}»".format(
        meta.get("discipline", "")), size=14)
    _para(doc, "Тема: {0}".format(meta.get("topic", "")), size=14)

    # --- Пустое место -----------------------------------------------
    for _ in range(3):
        _para(doc, "")

    # --- Студент / Преподаватель (две колонки через tab stop) -------
    def two_col(left, right):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf = p.paragraph_format
        pf.first_line_indent = Cm(0)
        pf.line_spacing = 1.0
        pf.tab_stops.add_tab_stop(Cm(9.0), WD_TAB_ALIGNMENT.LEFT)
        r = p.add_run("{0}\t{1}".format(left, right))
        set_run_font(r, size=14)

    two_col("Студент(ка) гр. {0}".format(meta.get("group", "")),
            "Преподаватель")
    two_col(str(meta.get("student", "")),
            str(meta.get("teacher", "")))

    # --- Пустое место -----------------------------------------------
    for _ in range(4):
        _para(doc, "")

    # --- Город / Год ------------------------------------------------
    _para(doc, meta.get("city", "Санкт-Петербург"), size=14)
    _para(doc, str(meta.get("year", "")), size=14)

    doc.add_page_break()


# ---------------------------------------------------------------------
#  Титульник из пользовательского .docx (docxtpl)
# ---------------------------------------------------------------------
def load_title_template(path, context):
    """Загружает .docx-шаблон титульника и подставляет {{ плейсхолдеры }}.

    Возвращает объект DocxTemplate -- подкласс python-docx.Document,
    поэтому дальше с ним можно работать тем же API (add_paragraph и т.п.).

    Доступные плейсхолдеры:
        {{ ministry }}    {{ university }}  {{ department }}
        {{ lab_number }}  {{ discipline }}  {{ topic }}
        {{ student }}     {{ group }}       {{ teacher }}
        {{ city }}        {{ year }}
    """
    from docxtpl import DocxTemplate
    tpl = DocxTemplate(str(path))
    tpl.render(context or {})
    return tpl
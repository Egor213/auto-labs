"""Функции добавления элементов в документ.

Все функции добавления параграфа защищены от ситуации, когда
требуемый стиль отсутствует в документе (актуально при работе
с пользовательским .docx-титульником): в этом случае используется
обычный абзац + ручное форматирование run'а.

Списки рисуются вручную (без опоры на стили List Bullet / List
Number): в шаблонах Word для этих стилей нужны дополнительные
определения numbering.xml, которых у пользовательских шаблонов
может не быть.
"""
import re
from pathlib import Path

from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from .formula import add_latex_formula

from .inline import add_inline, set_run_font, typo

try:
    from PIL import Image
    _HAS_PIL = True
except ImportError:
    _HAS_PIL = False


# ---------------------------------------------------------------------
#  Хелпер: абзац с заданным стилем и откатом на Normal
# ---------------------------------------------------------------------
def _safe_paragraph(doc, style_name):
    """doc.add_paragraph(style=...) с откатом, если стиля нет."""
    if style_name:
        try:
            return doc.add_paragraph(style=style_name)
        except KeyError:
            pass
    return doc.add_paragraph()


# ---------------------------------------------------------------------
#  Заголовки, абзацы, списки
# ---------------------------------------------------------------------
def add_heading(doc, text, level):
    level = max(1, min(level, 3))
    style_name = "Heading {0}".format(level)
    p = _safe_paragraph(doc, style_name)

    # Если стиль отсутствовал -- хоть как-то отформатируем вручную.
    if p.style is None or p.style.name != style_name:
        p.alignment = (WD_ALIGN_PARAGRAPH.CENTER if level == 1
                       else WD_ALIGN_PARAGRAPH.JUSTIFY)
        run = p.add_run(typo(text))
        set_run_font(run, size=14, bold=True)
        return p

    add_inline(p, text)
    return p


def add_paragraph_multiline(doc, lines):
    """Один абзац, строки внутри разделены soft-break (<w:br/>)."""
    p = doc.add_paragraph()
    for idx, line in enumerate(lines):
        if idx > 0:
            p.add_run().add_break()
        add_inline(p, line)
    return p


def add_paragraph(doc, text):
    p = doc.add_paragraph()
    add_inline(p, text)
    return p


from .numbering import register_numbering, apply_numbering


def add_list_item(doc, text, prefix="\u2013 ", indent_cm=0.75,
                  hanging_cm=0.5):
    """Ручной пункт списка (fallback, если нумерация не зарегистрировалась)."""
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(indent_cm)
    pf.first_line_indent = Cm(-hanging_cm)
    pf.line_spacing = 1.5
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if prefix:
        run = p.add_run(prefix)
        set_run_font(run, size=14)
    add_inline(p, text)
    return p


def add_list_block(doc, items, ordered):
    """Добавляет блок списка с настоящей Word-нумерацией.

    items  -- список строк без префикса;
    ordered -- True для нумерованного, False для маркированного.

    Один вызов = один экземпляр нумерации, поэтому следующий
    нумерованный список автоматически начнётся с 1.
    """
    kind = "ordered" if ordered else "bullet"

    try:
        num_id = register_numbering(doc, kind)
    except Exception as e:
        print("[WARN] Не удалось зарегистрировать нумерацию: {0}".format(e))
        for item in items:
            prefix = "" if ordered else "\u2013 "
            add_list_item(doc, item, prefix=prefix)
        return

    for item in items:
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = 1.5
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        apply_numbering(p, num_id, level=0)
        add_inline(p, item)

# ---------------------------------------------------------------------
#  Подписи
# ---------------------------------------------------------------------
def add_caption_above(doc, caption, align=None):
    p = doc.add_paragraph()
    p.alignment = align if align is not None else WD_ALIGN_PARAGRAPH.LEFT
    pf = p.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.line_spacing = 1.0
    pf.space_before = Pt(6)
    pf.space_after = Pt(3)
    run = p.add_run(typo(caption))
    set_run_font(run, size=14)
    return p


def add_caption_below(doc, caption, align=None):
    p = doc.add_paragraph()
    p.alignment = align if align is not None else WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.line_spacing = 1.0
    pf.space_before = Pt(3)
    pf.space_after = Pt(6)
    run = p.add_run(typo(caption))
    set_run_font(run, size=14)
    return p


# ---------------------------------------------------------------------
#  Листинг
# ---------------------------------------------------------------------
def add_code_block(doc, code, caption=None):
    if caption:
        add_caption_above(doc, caption)

    for line in code.split("\n"):
        p = _safe_paragraph(doc, "Listing")
        pf = p.paragraph_format
        pf.first_line_indent = Cm(0)
        pf.line_spacing = 1.0
        run = p.add_run(line if line else " ")
        set_run_font(run, name="Courier New", size=11)

    tail = doc.add_paragraph()
    tail.paragraph_format.first_line_indent = Cm(0)
    tail.paragraph_format.line_spacing = 1.0
    tail.paragraph_format.space_after = Pt(0)


# ---------------------------------------------------------------------
#  Рисунок
# ---------------------------------------------------------------------
def add_image(doc, path, caption=""):
    path = Path(path)
    if not path.exists():
        print("[WARN] Изображение не найдено: {0}".format(path))
        if caption:
            add_caption_below(doc, caption)
        return

    width_cm = 15.0
    if _HAS_PIL:
        try:
            with Image.open(path) as im:
                w, h = im.size
                dpi = im.info.get("dpi", (96, 96))[0] or 96
                native_cm = w / dpi * 2.54
                width_cm = min(native_cm, 16.0)
        except Exception:
            pass

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.space_before = Pt(6)
    run = p.add_run()
    run.add_picture(str(path), width=Cm(width_cm))

    if caption:
        add_caption_below(doc, caption)


# ---------------------------------------------------------------------
#  Таблица
# ---------------------------------------------------------------------
_SEP = re.compile(r"^:?-+:?$")


def _parse_table_row(line):
    s = line.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    return [c.strip() for c in s.split("|")]


def add_table(doc, table_lines, caption=None):
    if caption:
        add_caption_above(doc, caption)

    rows = [_parse_table_row(l) for l in table_lines]
    if not rows:
        return

    header = rows[0]
    data = []
    for r in rows[1:]:
        if r and all(_SEP.match(c) for c in r if c):
            continue
        data.append(r)

    n_cols = len(header)
    fs = 12 if len(data) > 5 else 14

    table = doc.add_table(rows=1 + len(data), cols=n_cols)
    try:
        table.style = "Table Grid"
    except KeyError:
        pass
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for j, txt in enumerate(header):
        cell = table.rows[0].cells[j]
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.line_spacing = 1.0
        run = p.add_run(txt)
        set_run_font(run, size=fs, bold=True)

    for i, row in enumerate(data):
        for j in range(n_cols):
            txt = row[j] if j < len(row) else ""
            cell = table.rows[i + 1].cells[j]
            cell.text = ""
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.line_spacing = 1.0
            run = p.add_run(txt)
            set_run_font(run, size=fs)

    tail = doc.add_paragraph()
    tail.paragraph_format.first_line_indent = Cm(0)
    tail.paragraph_format.line_spacing = 1.0
    tail.paragraph_format.space_after = Pt(0)


# ---------------------------------------------------------------------
#  Формула
# ---------------------------------------------------------------------
def add_formula(doc, formula_text):
    """Добавляет выключную формулу (по центру, отдельным абзацем)."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    
    # Убираем $$, если они есть, и пробелы по краям
    latex = formula_text.strip()
    if latex.startswith('$$') and latex.endswith('$$'):
        latex = latex[2:-2].strip()
    
    add_latex_formula(p, latex, inline=False)
    return p
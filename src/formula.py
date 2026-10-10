"""Вставка формул в DOCX как нативных объектов Word (OMML)."""
from docx.shared import Pt
from docx.text.paragraph import Paragraph

try:
    import math2docx
    _HAS_MATH2DOCX = True
except ImportError:
    _HAS_MATH2DOCX = False


def add_latex_formula(paragraph: Paragraph, latex: str, inline: bool = False):
    """
    Вставляет формулу LaTeX в абзац как объект Word (OMML).

    Args:
        paragraph: Абзац python-docx, куда вставляется формула.
        latex: Строка с формулой в формате LaTeX.
        inline: Сохранён для совместимости, в текущей версии
                math2docx не используется (формула всегда встраивается
                в переданный абзац).
    """
    if not _HAS_MATH2DOCX:
        print("[WARN] math2docx не установлен, формула вставлена как текст. "
              "Установите: pip install math2docx")
        _fallback(paragraph, latex)
        return

    try:
        # В текущей версии math2docx нет параметра inline.
        # Вставляем формулу в переданный абзац; выключная она или
        # внутристрочная — определяется контекстом абзаца.
        math2docx.add_math(paragraph, latex)
    except Exception as e:
        print("[WARN] Не удалось вставить формулу '{0}': {1}".format(latex, e))
        _fallback(paragraph, latex)


def _fallback(paragraph: Paragraph, latex: str):
    """Резервный вариант: вывести формулу как текст."""
    run = paragraph.add_run(latex)
    run.font.name = "Courier New"
    run.font.size = Pt(11)
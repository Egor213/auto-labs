"""Разбор тела Markdown-документа.

Особенности:
    * блок подряд идущих непустых "обычных" строк склеивается в ОДИН
      абзац Word; переносы между строками становятся soft-break
      (<w:br/>, Shift+Enter в Word);
    * пустая строка = конец абзаца;
    * списки (маркированные/нумерованные) рисуются вручную,
      нумерация ведётся счётчиком и не зависит от стилей Word;
    * спец. элементы (заголовки, листинги, таблицы, рисунки, формулы)
      разбираются построчно.
"""
import re
from pathlib import Path

from .elements import (
    add_heading, add_paragraph_multiline, add_list_item, add_code_block,
    add_image, add_table, add_formula,
)

_RE_HEADING = re.compile(r"^(#{1,6})\s+(.*)$")
_RE_IMG = re.compile(r"^!\[([^\]]*)\]\(([^)]+)\)\s*$")
_RE_BULLET = re.compile(r"^[-*+]\s+(.*)$")
_RE_ORDERED = re.compile(r"^\d+\.\s+(.*)$")
_RE_TABLE_CAP = re.compile(r"^Таблица\s+\d+\s*[-]", re.IGNORECASE)
_RE_LISTING_CAP = re.compile(r"^Листинг\s+\d+\s*[-]", re.IGNORECASE)
_RE_TABLE_SEP = re.compile(r"^\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)+\|?\s*$")


def _is_special_start(s: str) -> bool:
    """Начало ли это спец-элемента (не часть обычного абзаца)."""
    if not s:
        return True
    if s.startswith("```"):
        return True
    if s.startswith("$$"):
        return True
    if _RE_IMG.match(s):
        return True
    if _RE_TABLE_CAP.match(s):
        return True
    if _RE_HEADING.match(s):
        return True
    if _RE_BULLET.match(s):
        return True
    if _RE_ORDERED.match(s):
        return True
    return False


def _is_table_start(lines, i):
    """Строка i начинает таблицу (pipe + строка-разделитель ниже)."""
    s = lines[i].strip()
    if "|" not in s:
        return False
    return i + 1 < len(lines) and bool(_RE_TABLE_SEP.match(lines[i + 1].strip()))


def process_body(doc, body, base_dir):
    lines = body.split("\n")
    n = len(lines)
    i = 0
    ordered_index = 0  # счётчик для нумерованных списков

    while i < n:
        stripped = lines[i].strip()

        # пустая строка
        if not stripped:
            i += 1
            continue

        # ---- типографический счётчик нумерованных списков -----------
        m_ordered = _RE_ORDERED.match(stripped)
        if m_ordered:
            ordered_index += 1
        else:
            ordered_index = 0

        # ``` - блок кода / листинг
        if stripped.startswith("```"):
            info = stripped[3:].strip()
            caption = None
            if ":" in info:
                _, caption = info.split(":", 1)
                caption = caption.strip()
            if caption is None and i > 0:
                prev = lines[i - 1].strip()
                if _RE_LISTING_CAP.match(prev):
                    caption = prev

            code_lines = []
            i += 1
            while i < n and not lines[i].strip().startswith("```"):
                code_lines.append(lines[i])
                i += 1
            i += 1  # пропускаем закрывающий ```
            add_code_block(doc, "\n".join(code_lines), caption)
            continue

        # $$ - формула (display)
        if stripped.startswith("$$"):
            if stripped.endswith("$$") and len(stripped) > 4:
                add_formula(doc, stripped[2:-2].strip())
                i += 1
                continue
            buf = [stripped[2:]]
            i += 1
            while i < n and "$$" not in lines[i]:
                buf.append(lines[i])
                i += 1
            if i < n:
                buf.append(lines[i].split("$$")[0])
                i += 1
            add_formula(doc, "\n".join(buf).strip())
            continue

        # ![Рисунок N -- Имя](path)
        m = _RE_IMG.match(stripped)
        if m:
            caption = m.group(1)
            src = m.group(2)
            add_image(doc, (Path(base_dir) / src).resolve(), caption)
            i += 1
            continue

        # "Таблица N -- Имя" + pipe-таблица
        if _RE_TABLE_CAP.match(stripped):
            j = i + 1
            while j < n and not lines[j].strip():
                j += 1
            if j < n and "|" in lines[j]:
                caption = stripped
                tlines = []
                while j < n and "|" in lines[j]:
                    tlines.append(lines[j])
                    j += 1
                add_table(doc, tlines, caption)
                i = j
                continue

        # pipe-таблица без подписи
        if _is_table_start(lines, i):
            tlines = []
            j = i
            while j < n and "|" in lines[j]:
                tlines.append(lines[j])
                j += 1
            add_table(doc, tlines, None)
            i = j
            continue

        # # Заголовок
        m = _RE_HEADING.match(stripped)
        if m:
            add_heading(doc, m.group(2).strip(), len(m.group(1)))
            i += 1
            continue

        # - пункт (маркированный)
        m = _RE_BULLET.match(stripped)
        if m:
            add_list_item(doc, m.group(1), prefix="\u2013 ")
            i += 1
            continue

        # 1. пункт (нумерованный) -- номер берём из счётчика
        if m_ordered:
            add_list_item(doc, m_ordered.group(1),
                          prefix="{0}. ".format(ordered_index))
            i += 1
            continue

        # обычный абзац: собираем подряд идущие непустые "обычные" строки
        buf = [stripped]
        i += 1
        while i < n:
            nxt = lines[i].strip()
            if not nxt:
                break
            if _is_special_start(nxt):
                break
            if _is_table_start(lines, i):
                break
            buf.append(nxt)
            i += 1
        add_paragraph_multiline(doc, buf)
"""Основной конвертер Markdown -> DOCX."""
import argparse
import sys
from pathlib import Path

from docx import Document

from .frontmatter import parse as parse_frontmatter
from .styles import setup_page, setup_styles, setup_page_numbers
from .title_page import add_title_page, load_title_template
from .parser import process_body


def convert(md_path, docx_path=None):
    """Конвертирует .md в .docx. Возвращает путь к созданному файлу."""
    md_path = Path(md_path)
    if not md_path.exists():
        raise FileNotFoundError(md_path)

    text = md_path.read_text(encoding="utf-8")
    meta, body = parse_frontmatter(text)

    # --- Определяем титульник --------------------------------------
    template_docx = str(meta.get("template_docx") or "").strip()
    doc = None

    if template_docx:
        tmpl_path = Path(template_docx)
        if not tmpl_path.is_absolute():
            tmpl_path = (md_path.parent / tmpl_path).resolve()
        if tmpl_path.exists():
            try:
                doc = load_title_template(tmpl_path, meta)
                print("[OK] Использую титульный шаблон: {0}".format(tmpl_path))
            except Exception as e:
                print("[WARN] Не удалось загрузить шаблон {0}: {1}".format(
                    tmpl_path, e))
                doc = None
        else:
            print("[WARN] template_docx не найден: {0}".format(tmpl_path))

    # --- Fallback: встроенный титульник -----------------------------
    if doc is None:
        doc = Document()
        add_title_page(doc, meta)

    # --- Всегда применяем оформление --------------------------------
    # Это нужно в т.ч. для случая template_docx, потому что в
    # пользовательском шаблоне может не быть нужных стилей (Listing,
    # Heading 1/2/3), а также полей и нумерации страниц.
    try:
        setup_page(doc)
    except Exception as e:
        print("[WARN] setup_page: {0}".format(e))
    try:
        setup_styles(doc)
    except Exception as e:
        print("[WARN] setup_styles: {0}".format(e))
    try:
        setup_page_numbers(doc)
    except Exception as e:
        print("[WARN] setup_page_numbers: {0}".format(e))

    # --- Тело -------------------------------------------------------
    process_body(doc, body, md_path.parent)

    if docx_path is None:
        docx_path = md_path.with_suffix(".docx")
    docx_path = Path(docx_path)
    doc.save(str(docx_path))
    print("[OK] {0} -> {1}".format(md_path, docx_path))
    return docx_path


def main(argv=None):
    ap = argparse.ArgumentParser(description="MD -> DOCX для отчетов ЛЭТИ")
    ap.add_argument("input", nargs="?", help="Входной .md файл")
    ap.add_argument("output", nargs="?", help="Выходной .docx (опционально)")
    args = ap.parse_args(argv)

    if not args.input:
        ap.print_help()
        sys.exit(1)
    convert(args.input, args.output)


if __name__ == "__main__":
    main()
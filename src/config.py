"""Загрузка значений по умолчанию из template.yaml."""
from pathlib import Path

try:
    import yaml
except ImportError:
    yaml = None  # type: ignore

_HERE = Path(__file__).parent
_TEMPLATE = _HERE / "template.yaml"

FALLBACK = {
    "ministry": "МИНОБРНАУКИ РОССИИ",
    "university": (
        "САНКТ-ПЕТЕРБУРГСКИЙ ГОСУДАРСТВЕННЫЙ\n"
        "ЭЛЕКТРОТЕХНИЧЕСКИЙ УНИВЕРСИТЕТ\n"
        "«ЛЭТИ» ИМ. В.И. УЛЬЯНОВА (ЛЕНИНА)"
    ),
    "department": "МО ЭВМ",
    "city": "Санкт-Петербург",
    "year": "2024",
    # Путь к .docx-титульнику с {{ плейсхолдерами }}. Пусто = встроенный.
    "template_docx": "",
}


def load_defaults() -> dict:
    """Возвращает словарь значений по умолчанию из template.yaml."""
    if yaml is None or not _TEMPLATE.exists():
        return dict(FALLBACK)
    try:
        with _TEMPLATE.open(encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}
    except Exception:
        return dict(FALLBACK)
    merged = dict(FALLBACK)
    for k, v in data.items():
        if v is not None:
            merged[k] = v
    return merged
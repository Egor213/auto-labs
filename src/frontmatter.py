"""Разбор YAML frontmatter и слияние с template.yaml."""
import re

try:
    import yaml
except ImportError:
    yaml = None  # type: ignore

from .config import load_defaults

_FM_RE = re.compile(r"^---\s*\n(.*?)\n---\s*\n?(.*)$", re.DOTALL)


def parse(text: str):
    """Возвращает (meta, body)."""
    text = text.lstrip("\ufeff")
    defaults = load_defaults()

    if not text.startswith("---"):
        return defaults, text

    m = _FM_RE.match(text)
    if not m:
        return defaults, text

    meta = {}
    if yaml is not None:
        try:
            meta = yaml.safe_load(m.group(1)) or {}
        except Exception:
            meta = {}

    merged = dict(defaults)
    for k, v in meta.items():
        if v is not None:
            merged[k] = v
    return merged, m.group(2)

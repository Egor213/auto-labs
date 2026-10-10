"""Управление нумерацией списков через numbering.xml.

Создаёт abstractNum + num в numbering part документа, чтобы Word
применял настоящую автонумерацию вместо ручных цифр.
"""
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

try:
    from docx.parts.numbering import NumberingPart
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    _HAS_PART = True
except ImportError:
    _HAS_PART = False


def _ensure_numbering_part(doc):
    """Возвращает NumberingPart, создавая его при необходимости."""
    try:
        return doc.part.numbering_part
    except NotImplementedError:
        pass

    if not _HAS_PART:
        raise RuntimeError(
            "numbering part недоступен, обновите python-docx: "
            "pip install -U python-docx"
        )
    part = NumberingPart.new()
    doc.part.relate_to(part, RT.NUMBERING)
    return part


def _next_ids(root):
    """Следующие свободные abstractNumId и numId."""
    a_ids = [int(el.get(qn("w:abstractNumId")))
             for el in root.findall(qn("w:abstractNum"))]
    n_ids = [int(el.get(qn("w:numId")))
             for el in root.findall(qn("w:num"))]
    # numId=0 зарезервирован -- он означает "без нумерации"
    return max(a_ids, default=-1) + 1, max(n_ids, default=0) + 1


def _make_abstract_num(aid, kind):
    """Создаёт <w:abstractNum> для 'ordered' или 'bullet'."""
    el = OxmlElement("w:abstractNum")
    el.set(qn("w:abstractNumId"), str(aid))

    mlt = OxmlElement("w:multiLevelType")
    mlt.set(qn("w:val"), "singleLevel")
    el.append(mlt)

    lvl = OxmlElement("w:lvl")
    lvl.set(qn("w:ilvl"), "0")

    if kind == "ordered":
        fmt_val = "decimal"
        text_val = "%1."
    else:
        fmt_val = "bullet"
        text_val = "\u2013"  # en-dash

    start = OxmlElement("w:start"); start.set(qn("w:val"), "1")
    fmt = OxmlElement("w:numFmt"); fmt.set(qn("w:val"), fmt_val)
    txt = OxmlElement("w:lvlText"); txt.set(qn("w:val"), text_val)
    jc = OxmlElement("w:lvlJc"); jc.set(qn("w:val"), "left")

    lvl.append(start)
    lvl.append(fmt)
    lvl.append(txt)
    lvl.append(jc)

    # Отступы: left=2см, hanging=1см (номер слева от текста).
    ppr = OxmlElement("w:pPr")
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"), "1134")
    ind.set(qn("w:hanging"), "567")
    ppr.append(ind)
    lvl.append(ppr)

    # Шрифт номера/маркера.
    rpr = OxmlElement("w:rPr")
    rfonts = OxmlElement("w:rFonts")
    for a in ("w:ascii", "w:hAnsi", "w:cs"):
        rfonts.set(qn(a), "Times New Roman")
    rpr.append(rfonts)
    sz = OxmlElement("w:sz"); sz.set(qn("w:val"), "28")  # 14 pt
    rpr.append(sz)
    lvl.append(rpr)

    el.append(lvl)
    return el


def _make_num(nid, aid):
    """Создаёт <w:num> -- экземпляр нумерации, ссылающийся на abstractNum."""
    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(nid))
    ref = OxmlElement("w:abstractNumId")
    ref.set(qn("w:val"), str(aid))
    num.append(ref)
    return num


def register_numbering(doc, kind):
    """Регистрирует новый экземпляр нумерации в документе.

    kind: 'ordered' | 'bullet'
    Возвращает num_id, который нужно повесить на абзацы списка.
    """
    part = _ensure_numbering_part(doc)
    root = part.element
    aid, nid = _next_ids(root)

    abstract = _make_abstract_num(aid, kind)
    # abstractNum должен идти перед num в XML.
    first_num = root.find(qn("w:num"))
    if first_num is not None:
        first_num.addprevious(abstract)
    else:
        root.append(abstract)

    root.append(_make_num(nid, aid))
    return nid


def apply_numbering(paragraph, num_id, level=0):
    """Проставляет numPr на абзац: <w:numPr><w:ilvl/><w:numId/></w:numPr>."""
    pPr = paragraph._p.get_or_add_pPr()

    for old in pPr.findall(qn("w:numPr")):
        pPr.remove(old)

    numPr = OxmlElement("w:numPr")
    ilvl = OxmlElement("w:ilvl"); ilvl.set(qn("w:val"), str(level))
    numId = OxmlElement("w:numId"); numId.set(qn("w:val"), str(num_id))
    numPr.append(ilvl)
    numPr.append(numId)

    # numPr должен идти сразу после pStyle (порядок в OOXML-схеме).
    pStyle = pPr.find(qn("w:pStyle"))
    if pStyle is not None:
        pStyle.addnext(numPr)
    else:
        pPr.insert(0, numPr)
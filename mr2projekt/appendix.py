"""Приложение к отчёту: исходные отчёты «Эколога», как в ручной сборке «мр2».

Отдельный документ, в котором подряд идут:
  * отчёт МР целиком;
  * первый лист СГ и результаты расчёта по веществам;
  * первый лист СС и результаты расчёта по веществам.
Во всех результатах строки расчётных точек отсортированы по номеру точки.
Документ строится на основе файла МР, поэтому оформление «Эколога» сохраняется.
"""

from __future__ import annotations

import copy
from dataclasses import dataclass, field
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import Table

from .docx_utils import normalize_space
from .mr_parser import (
    REPORT_TITLE,
    _calc_kind,
    _is_results_table,
)


@dataclass
class _Block:
    """Один отчёт «Эколога» внутри файла: элементы тела документа подряд."""

    kind: str | None
    elements: list = field(default_factory=list)


def _text(element) -> str:
    return normalize_space("".join(node.text or "" for node in element.iter(qn("w:t"))))


def _first_cell_text(element) -> str:
    tc = next(element.iter(qn("w:tc")), None)
    return _text(tc) if tc is not None else ""


def _calc_line(tbl) -> str:
    for tr in tbl.iter(qn("w:tr")):
        text = _first_cell_text(tr)
        if text.startswith("Расчет:"):
            return text
    return ""


def _blocks(document) -> list[_Block]:
    blocks: list[_Block] = []
    for element in document.element.body.iterchildren():
        if element.tag == qn("w:sectPr"):
            continue
        if element.tag == qn("w:tbl") and _first_cell_text(element).startswith(
            REPORT_TITLE
        ):
            calc = _calc_line(element)
            if calc or not blocks:
                blocks.append(_Block(_calc_kind(calc)))
        if not blocks:
            blocks.append(_Block(None))
        blocks[-1].elements.append(element)
    return blocks


def _first_sheet_and_results(block: _Block, document) -> list:
    """Шапка отчёта (первый лист) и всё, начиная с результатов по веществам."""
    elements = block.elements
    tables = [index for index, item in enumerate(elements) if item.tag == qn("w:tbl")]
    if not tables:
        return []
    head_end = tables[0] + 1
    # абзац-разделитель раздела сразу после шапки — часть первого листа
    if head_end < len(elements) and elements[head_end].tag == qn("w:p"):
        head_end += 1
    start = next(
        (
            index
            for index in tables[1:]
            if _is_results_table(Table(elements[index], document))
        ),
        None,
    )
    if start is None:
        return elements[:head_end]
    return elements[:head_end] + elements[start:]


def _point_number(tr) -> int | None:
    text = _first_cell_text(tr)
    if text.isdigit() and 1 <= int(text) <= 100:
        return int(text)
    return None


def sort_result_rows(elements: list) -> None:
    """Строки точек каждого вещества — по возрастанию номера точки.

    Сортируются только разделы «Результаты расчета по веществам»; строки
    переставляются по уже существующим местам, поэтому разбивка на таблицы
    и страницы остаётся как в отчёте «Эколога».
    """
    groups: list[list] = []
    current: list | None = None
    in_results = False
    for element in elements:
        if element.tag != qn("w:tbl"):
            continue
        first = _first_cell_text(element)
        if first.startswith(REPORT_TITLE):
            in_results = False
        elif first.startswith("Результаты расчета"):
            in_results = True
        if not in_results:
            continue
        for tr in element.findall(qn("w:tr")):
            text = _first_cell_text(tr)
            if text.startswith("Вещество:"):
                current = []
                groups.append(current)
                continue
            if current is not None and _point_number(tr) is not None:
                current.append(tr)

    for rows in groups:
        ordered = sorted(rows, key=_point_number)
        if ordered == rows:
            continue
        slots = []
        for tr in rows:
            slot = OxmlElement("w:tr")
            tr.addprevious(slot)
            slots.append(slot)
        for tr in rows:
            tr.getparent().remove(tr)
        for slot, tr in zip(slots, ordered):
            slot.addprevious(tr)
            slot.getparent().remove(slot)


def _section_break(sectPr) -> object:
    """Абзац, завершающий раздел с параметрами sectPr (дальше — новая страница)."""
    p = OxmlElement("w:p")
    p.get_or_add_pPr().append(copy.deepcopy(sectPr))
    return p


def _part(path: Path, kind: str) -> list | None:
    """Первый лист и результаты отчёта kind из файла (копии элементов)."""
    document = Document(str(path))
    block = next((item for item in _blocks(document) if item.kind == kind), None)
    if block is None:
        return None
    elements = [copy.deepcopy(item) for item in _first_sheet_and_results(block, document)]
    sort_result_rows(elements)
    return elements


APPENDIX_FONT = "Times New Roman"
FONT_ATTRS = ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia")
THEME_ATTRS = ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme")


def _set_font(document, font: str) -> None:
    """Один шрифт во всём документе: в тексте, стилях и умолчаниях."""
    roots = [document.element.body, document.styles.element]
    for root in roots:
        for rFonts in root.iter(qn("w:rFonts")):
            for attr in THEME_ATTRS:
                rFonts.attrib.pop(qn(attr), None)
            for attr in FONT_ATTRS:
                rFonts.set(qn(attr), font)
    # умолчания документа: шрифт для текста без явно заданного
    defaults = document.styles.element.find(qn("w:docDefaults"))
    if defaults is not None:
        rPr = defaults.find(f"{qn('w:rPrDefault')}/{qn('w:rPr')}")
        if rPr is not None and rPr.find(qn("w:rFonts")) is None:
            rFonts = OxmlElement("w:rFonts")
            for attr in FONT_ATTRS:
                rFonts.set(qn(attr), font)
            rPr.insert(0, rFonts)


def appendix_path(output_path: str | Path) -> Path:
    output_path = Path(output_path)
    return output_path.with_name(f"{output_path.stem} — приложение.docx")


def build_appendix(
    output_path: str | Path,
    mr_path: str | Path,
    annual_path: str | Path | None = None,
    daily_path: str | Path | None = None,
) -> list[str]:
    """Собрать приложение: МР целиком, затем первый лист и результаты СГ и СС."""
    warnings: list[str] = []
    document = Document(str(mr_path))
    body = document.element.body
    blocks = _blocks(document)
    main = next((item for item in blocks if item.kind == "max"), blocks[0] if blocks else None)
    if main is None:
        return ["Приложение: в файле МР нет отчёта «Эколога», приложение не создано"]

    # из сводного файла оставляем только сам МР
    for block in blocks:
        if block is not main:
            for element in block.elements:
                body.remove(element)
    sort_result_rows(main.elements)

    body_sectPr = body.find(qn("w:sectPr"))
    for kind, path, title in (
        ("annual", annual_path or mr_path, "СГ"),
        ("daily", daily_path or mr_path, "СС"),
    ):
        elements = _part(Path(path), kind)
        if elements is None:
            warnings.append(f"Приложение: нет отчёта {title}, он не добавлен")
            continue
        # каждый отчёт — с новой страницы, как в «Экологе»
        body_sectPr.addprevious(_section_break(body_sectPr))
        for element in elements:
            body_sectPr.addprevious(element)

    _set_font(document, APPENDIX_FONT)
    document.save(str(output_path))
    return warnings

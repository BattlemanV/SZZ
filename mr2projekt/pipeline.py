from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from .appendix import appendix_path, build_appendix
from .generator import generate
from .mr_parser import parse_mr
from .reference import load_reference


@dataclass(frozen=True)
class ConversionResult:
    summary: dict[str, object]


def default_output_path(mr_path: str | Path) -> Path:
    """Отчёт сохраняется рядом с отчётом МР."""
    mr_path = Path(mr_path)
    return mr_path.with_name(f"Отчет СЗЗ ({mr_path.stem}).docx")


def convert(
    mr_path: str | Path,
    template_path: str | Path,
    output_path: str | Path,
    *,
    annual_path: str | Path | None = None,
    daily_path: str | Path | None = None,
    vocw_dir: str | Path | None = None,
    with_appendix: bool = True,
) -> ConversionResult:
    """mr_path — отчёт МР; annual_path (СГ) и daily_path (СС) — отдельные отчёты.

    Если СГ/СС не заданы, их результаты ищутся в сводном файле МР.
    """
    mr_path = Path(mr_path).expanduser().resolve()
    template_path = Path(template_path).expanduser().resolve()
    output_path = Path(output_path).expanduser().resolve()

    if not mr_path.is_file():
        raise FileNotFoundError(f"Не найден отчёт МР: {mr_path}")
    extra: dict[str, Path | None] = {}
    for key, value, title in (
        ("annual_path", annual_path, "СГ"),
        ("daily_path", daily_path, "СС"),
    ):
        extra[key] = Path(value).expanduser().resolve() if value else None
        if extra[key] is not None and not extra[key].is_file():
            raise FileNotFoundError(f"Не найден отчёт {title}: {extra[key]}")
    if not template_path.is_file():
        raise FileNotFoundError(f"Не найден шаблон отчёта: {template_path}")

    output_path.parent.mkdir(parents=True, exist_ok=True)

    reference = load_reference(vocw_dir)
    model = parse_mr(
        mr_path,
        reference.substances,
        reference.group_members,
        hazard_classes=reference.hazard_classes,
        **extra,
    )
    if reference.vocw is None:
        model.warnings.insert(
            0,
            "Справочник веществ VocW не найден — нормативы и группы суммации "
            f"взяты из отчёта МР ({reference.vocw_error})",
        )
    result = generate(
        model,
        template_path=template_path,
        output_path=output_path,
        labels=reference.table52_labels,
        intro_names=reference.intro_names,
    )
    # исходные отчёты «Эколога» — отдельным файлом рядом с отчётом
    appendix: Path | None = None
    if with_appendix:
        appendix = appendix_path(output_path)
        result["warnings"].extend(
            build_appendix(appendix, mr_path, extra["annual_path"], extra["daily_path"])
        )

    summary: dict[str, object] = {
        "mr": str(mr_path),
        "annual": str(extra["annual_path"] or ""),
        "daily": str(extra["daily_path"] or ""),
        "template": str(template_path),
        "vocw": str(reference.vocw.path) if reference.vocw else "",
        "output": str(output_path),
        "appendix": str(appendix) if appendix else "",
        "sources": len(model.sources),
        "organized": result["organized"],
        "unorganized": result["unorganized"],
        "substances": len(model.substances),
        "groups": len(model.groups),
        "points": result["points"],
        "stats": {
            key: round(value, 10) if isinstance(value, float) else value
            for key, value in result["stats"].items()
        },
        "warnings": result["warnings"],
    }

    return ConversionResult(summary=summary)

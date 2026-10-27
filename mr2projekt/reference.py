from __future__ import annotations

import csv
import io
import json
from dataclasses import dataclass, field
from pathlib import Path

from . import settings
from .vocw import Norms, Vocw, load_vocw


DATA_DIR = Path(__file__).resolve().parent.parent / "data"
HAZARD_FILE = "классы опасности.csv"
HAZARD_VALUES = {"1", "2", "3", "4", "Нет"}


def _load_json(name: str, default):
    path = DATA_DIR / name
    if not path.exists():
        return default
    return json.loads(path.read_text(encoding="utf-8"))


def hazard_file() -> Path:
    return settings.app_dir() / HAZARD_FILE


def load_hazard_classes(path: Path | None = None) -> dict[str, str]:
    """Классы опасности пользователя: «код;класс» (Excel, UTF-8 или cp1251)."""
    path = path or hazard_file()
    if not path.is_file():
        return {}
    raw = path.read_bytes()
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = raw.decode("cp1251")
    classes: dict[str, str] = {}
    for row in csv.reader(io.StringIO(text), delimiter=";"):
        if len(row) < 2:
            continue
        code, value = row[0].strip(), row[1].strip()
        if not code.isdigit() or len(code) > 4:
            continue  # заголовок, комментарии
        value = "Нет" if value.lower() == "нет" else value
        if value in HAZARD_VALUES:
            # Excel убирает ведущие нули: 301 -> 0301
            classes[code.zfill(4)] = value
    return classes


@dataclass(frozen=True)
class Reference:
    vocw: Vocw | None
    table52_labels: dict[str, str]
    hazard_classes: dict[str, str] = field(default_factory=dict)
    vocw_error: str = ""
    # привычные названия веществ для перечня в разделе 3.1
    intro_names: dict[str, str] = field(default_factory=dict)

    @property
    def substances(self) -> dict[str, Norms]:
        return self.vocw.substances if self.vocw else {}

    @property
    def group_members(self) -> dict[str, list[str]]:
        return self.vocw.group_members if self.vocw else {}


def load_reference(vocw_dir: str | Path | None = None) -> Reference:
    """Справочник VocW необязателен: без него нормативы берутся из отчёта МР."""
    labels = _load_json("table52_labels.json", {})
    vocw: Vocw | None = None
    error = ""
    try:
        vocw = load_vocw(vocw_dir or settings.vocw_dir())
    except (OSError, ValueError) as exc:
        error = str(exc)
    return Reference(
        vocw=vocw,
        table52_labels={str(key): str(value) for key, value in labels.items()},
        hazard_classes=load_hazard_classes(),
        vocw_error=error,
        intro_names={
            str(key): str(value)
            for key, value in _load_json("intro_names.json", {}).items()
        },
    )

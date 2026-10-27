"""Настройки программы: каталог справочника VocW и расположение файлов.

settings.ini лежит рядом с программой; его записывает установщик,
а окно программы обновляет, если путь к справочнику поменяли.
"""

from __future__ import annotations

import configparser
import sys
from pathlib import Path

from .vocw import DEFAULT_VOCW_DIR


SETTINGS_NAME = "settings.ini"
TEMPLATE_NAME = "отчет выход.docx"
SECTION = "paths"


def app_dir() -> Path:
    """Каталог программы: рядом с exe в установленной версии, иначе — проект."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def settings_path() -> Path:
    return app_dir() / SETTINGS_NAME


def template_path() -> Path:
    return app_dir() / TEMPLATE_NAME


def _read() -> configparser.ConfigParser:
    config = configparser.ConfigParser()
    path = settings_path()
    if path.is_file():
        try:
            config.read(path, encoding="utf-8-sig")
        except UnicodeDecodeError:
            config.read(path, encoding="cp1251")
    return config


def vocw_dir() -> Path:
    value = _read().get(SECTION, "vocw", fallback="").strip()
    return Path(value) if value else DEFAULT_VOCW_DIR


def save_vocw_dir(path: str | Path) -> None:
    config = _read()
    if not config.has_section(SECTION):
        config.add_section(SECTION)
    config.set(SECTION, "vocw", str(path))
    with settings_path().open("w", encoding="utf-8") as handle:
        config.write(handle)

"""Справочник веществ УПРЗА «Эколог» (VocW, таблицы Advantage Database .ADT)."""

from __future__ import annotations

import struct
from dataclasses import dataclass
from pathlib import Path


DEFAULT_VOCW_DIR = Path(r"C:\integral.ltd\VocW")

HEADER_SIZE = 400
FIELD_SIZE = 200
FIELD_NAME_SIZE = 129
RECORD_HEADER_SIZE = 5

TYPE_LOGICAL = 1
TYPE_CHAR = 4
TYPE_DOUBLE = 10
TYPE_INTEGER = 11
TYPE_SHORT = 12
TYPE_AUTOINC = 15

NULL_SHORT = -32768

# PR_VALUE — единица норматива (пусто = мг/м3); список из формы VOCW.exe.
UNITS = {
    1: ("мкг/м3", 1e-3),
    2: ("нг/м3", 1e-6),
    3: ("пг/м3", 1e-9),
    4: ("микробных тел/м3", None),
    5: ("микробных тел/л", None),
    6: ("клеток/м3 (ПДК среднесменная)", None),
    7: ("волокон/мл воздуха", None),
}


@dataclass(frozen=True)
class Norms:
    code: str
    name: str
    pdk_mr: float | None
    pdk_ss: float | None
    pdk_sg: float | None
    obuv: float | None
    hazard_class: int | None
    unit: str = "мг/м3"

    @property
    def hazard_label(self) -> str:
        if self.hazard_class in {1, 2, 3, 4}:
            return str(self.hazard_class)
        return "Нет"

    @property
    def max_norm(self) -> tuple[str, float] | None:
        """Норматив для расчёта максимальных концентраций."""
        if self.pdk_mr is not None:
            return "ПДК м/р", self.pdk_mr
        if self.obuv is not None:
            return "ОБУВ", self.obuv
        return None

    @property
    def annual_norm(self) -> tuple[str, float] | None:
        """Норматив для расчёта среднегодовых концентраций."""
        if self.pdk_sg is not None:
            return "ПДК с/г", self.pdk_sg
        if self.pdk_ss is not None:
            return "ПДК с/с", self.pdk_ss
        return None

    @property
    def daily_norm(self) -> tuple[str, float] | None:
        """Норматив для расчёта среднесуточных концентраций."""
        if self.pdk_ss is not None:
            return "ПДК с/с", self.pdk_ss
        return None


@dataclass(frozen=True)
class Vocw:
    path: Path
    substances: dict[str, Norms]
    group_names: dict[str, str]
    group_members: dict[str, list[str]]


def _code(value: int | None) -> str:
    return "" if value is None else f"{value:04d}"


def _decode(field_type: int, raw: bytes):
    if field_type == TYPE_CHAR:
        return raw.split(b"\0", 1)[0].decode("cp1251").strip()
    if field_type == TYPE_DOUBLE:
        # NULL хранится как 0x8000000000000020 (отрицательное денормализованное число)
        if raw[-1] & 0x80 and raw[:-1].strip(b"\0") in {b"", b"\x20"}:
            return None
        return struct.unpack("<d", raw)[0]
    if field_type in {TYPE_INTEGER, TYPE_AUTOINC}:
        return struct.unpack("<i", raw)[0]
    if field_type == TYPE_SHORT:
        value = struct.unpack("<h", raw)[0]
        return None if value == NULL_SHORT else value
    if field_type == TYPE_LOGICAL:
        return raw in {b"T", b"t", b"Y", b"y"}
    return None


def read_adt(path: str | Path) -> list[dict[str, object]]:
    """Прочитать все записи таблицы .ADT (без мемо-полей)."""
    data = Path(path).read_bytes()
    if not data.startswith(b"Advantage Table"):
        raise ValueError(f"Не таблица Advantage: {path}")
    count, data_offset, record_size = (
        struct.unpack_from("<I", data, offset)[0] for offset in (0x18, 0x20, 0x24)
    )

    fields: list[tuple[str, int, int, int]] = []
    position = HEADER_SIZE
    while position < data_offset and data[position] != 0:
        name = data[position : position + FIELD_NAME_SIZE].split(b"\0", 1)[0]
        field_type, offset, length = struct.unpack_from(
            "<HIH", data, position + FIELD_NAME_SIZE
        )
        fields.append((name.decode("ascii"), field_type, offset, length))
        position += FIELD_SIZE

    records: list[dict[str, object]] = []
    for index in range(count):
        start = data_offset + index * record_size
        record = data[start : start + record_size]
        if len(record) < record_size:
            break
        records.append(
            {
                name: _decode(field_type, record[offset : offset + length])
                for name, field_type, offset, length in fields
            }
        )
    return records


def _positive(value, factor: float = 1.0) -> float | None:
    if isinstance(value, float) and value > 0:
        return float(f"{value * factor:.12g}")
    return None


def _database_dir(path: Path) -> Path:
    if (path / "FPR.ADT").is_file():
        return path
    return path / "DataBase"


def load_vocw(path: str | Path | None = None) -> Vocw:
    root = Path(path) if path else DEFAULT_VOCW_DIR
    database = _database_dir(root.expanduser())
    if not (database / "FPR.ADT").is_file():
        raise FileNotFoundError(
            f"Не найден справочник веществ VocW (FPR.ADT): {database}"
        )

    substances: dict[str, Norms] = {}
    for row in read_adt(database / "FPR.ADT"):
        code = _code(row.get("PR_COD4"))
        if not code:
            continue
        unit, factor = UNITS.get(row.get("PR_VALUE"), ("мг/м3", 1.0))
        if factor is None:
            # Не массовая единица: в мг/м3 не пересчитать, значение как есть.
            factor = 1.0
        else:
            unit = "мг/м3"
        substances[code] = Norms(
            code=code,
            name=str(row.get("AP_PRNAM") or ""),
            pdk_mr=_positive(row.get("PR_PDKMR"), factor),
            pdk_ss=_positive(row.get("PR_PDKSS"), factor),
            pdk_sg=_positive(row.get("PR_PDKSG"), factor),
            obuv=_positive(row.get("PR_OBUV"), factor),
            hazard_class=row.get("PR_CL_OP"),
            unit=unit,
        )

    group_names = {
        _code(row.get("PR_COD4")): str(row.get("AP_PRNAM") or "")
        for row in read_adt(database / "GrSum.ADT")
        if row.get("PR_COD4") is not None
    }
    group_members: dict[str, list[str]] = {}
    for row in read_adt(database / "FGRPR.ADT"):
        group = _code(row.get("GR_COD4"))
        member = _code(row.get("PR_COD4"))
        if group and member:
            group_members.setdefault(group, []).append(member)
    for members in group_members.values():
        members.sort()

    return Vocw(
        path=database,
        substances=substances,
        group_names=group_names,
        group_members=group_members,
    )

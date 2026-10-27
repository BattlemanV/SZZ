"""Точка входа программы «Отчёт СЗЗ» (из неё собирается exe).

Файлы перетащены на ярлык/exe — сразу формируется отчёт (drop.py);
запуск без файлов — открывается окно программы (gui.py).
`--selftest <файл.json>` — самопроверка установленной программы.
"""

from __future__ import annotations

import json
import sys
import tempfile
import traceback
from pathlib import Path


def selftest(report: Path) -> int:
    """Проверить справочник, шаблон и формирование отчёта без данных МР."""
    result: dict[str, object] = {"ok": False}
    try:
        from docx import Document

        from mr2projekt import settings
        from mr2projekt.generator import generate
        from mr2projekt.mr_parser import MRModel, ResultPoint, Substance
        from mr2projekt.reference import load_reference
        from mr2projekt.table52 import find_result_tables

        reference = load_reference()
        result["vocw"] = str(reference.vocw.path) if reference.vocw else ""
        result["vocw_error"] = reference.vocw_error
        result["vocw_substances"] = len(reference.substances)
        result["hazard_classes_file"] = len(reference.hazard_classes)
        result["template"] = str(settings.template_path())

        norms = reference.substances.get("0301")
        model = MRModel(path=Path("selftest"))
        model.substances = [
            Substance(
                code="0301",
                name="Азота диоксид",
                g_s=0.1,
                t_year=1.0,
                hazard=norms.hazard_label if norms else "3",
                max_type="ПДК м/р",
                max_value="0.2",
            )
        ]
        model.results = {
            "max": {
                "0301": [
                    ResultPoint(fraction=0.5 + index, mg_m3=0.1, number=index + 1)
                    for index in range(3)
                ]
            }
        }
        model.points_count = 3
        with tempfile.TemporaryDirectory() as folder:
            output = Path(folder) / "selftest.docx"
            generate(
                model,
                settings.template_path(),
                output,
                labels=reference.table52_labels,
            )
            result["tables52"] = len(find_result_tables(Document(str(output))))
        result["ok"] = True
    except Exception:
        result["error"] = traceback.format_exc()
    report.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return 0 if result["ok"] else 1


def main(argv: list[str]) -> int:
    if len(argv) == 2 and argv[0] == "--selftest":
        return selftest(Path(argv[1]))
    if argv:
        import drop

        return drop.main(argv)
    import gui

    gui.main()
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))

"""Формирование отчёта перетаскиванием файлов на ярлык.

Перетащите на ярлык отчёты «Эколога» МР, СГ и СС (или один сводный МР):
программа сама определит, какой файл какой, и сохранит отчёт рядом с МР.
"""

from __future__ import annotations

import os
import sys
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

import tkinter as tk  # noqa: E402
from tkinter import messagebox  # noqa: E402

from mr2projekt import settings  # noqa: E402
from mr2projekt.mr_parser import classify_reports  # noqa: E402
from mr2projekt.pipeline import convert, default_output_path  # noqa: E402


TITLE = "Отчёт СЗЗ"
MAX_WARNINGS = 15


def _summary(reports, summary: dict, warnings: list[str]) -> str:
    lines = [
        f"Отчёт сохранён:\n{summary['output']}",
        "",
        f"Приложение (отчёты «Эколога»):\n{summary['appendix']}",
        "",
        f"МР: {reports.mr.name}",
        f"СГ: {reports.annual.name if reports.annual else 'из отчёта МР'}",
        f"СС: {reports.daily.name if reports.daily else 'из отчёта МР'}",
    ]
    if reports.skipped:
        names = ", ".join(path.name for path in reports.skipped)
        lines.append(f"Пропущены (не отчёты «Эколога»): {names}")
    if warnings:
        lines += ["", "Предупреждения:"]
        lines += [f"• {item}" for item in warnings[:MAX_WARNINGS]]
        if len(warnings) > MAX_WARNINGS:
            lines.append(f"… и ещё {len(warnings) - MAX_WARNINGS}")
    lines += ["", "Открыть отчёт?"]
    return "\n".join(lines)


def main(argv: list[str]) -> int:
    root = tk.Tk()
    root.withdraw()

    paths = [Path(arg) for arg in argv if arg.lower().endswith(".docx")]
    if not paths:
        messagebox.showinfo(
            TITLE,
            "Перетащите на ярлык отчёты «Эколога» МР, СГ и СС (.docx) "
            "или один сводный отчёт МР.",
        )
        return 1

    output: Path | None = None
    try:
        reports = classify_reports(paths)
        output = default_output_path(reports.mr)
        result = convert(
            reports.mr,
            settings.template_path(),
            output,
            annual_path=reports.annual,
            daily_path=reports.daily,
        )
    except PermissionError as exc:
        messagebox.showerror(
            TITLE,
            f"Не удалось записать файл:\n{exc.filename or output}\n\n"
            "Вероятно, он открыт в Word. Закройте его и перетащите файлы снова.",
        )
        return 2
    except Exception as exc:  # показать любую ошибку, консоли нет
        details = traceback.format_exc()
        root.clipboard_clear()
        root.clipboard_append(details)
        root.update()
        messagebox.showerror(
            TITLE,
            f"{exc or type(exc).__name__}\n\n(подробности скопированы в буфер обмена)",
        )
        return 2

    warnings = list(result.summary.get("warnings", []))
    if messagebox.askyesno(
        TITLE,
        _summary(reports, result.summary, warnings),
        icon="warning" if warnings else "question",
    ):
        os.startfile(str(output))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))

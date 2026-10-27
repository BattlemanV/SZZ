from __future__ import annotations

import os
import threading
import traceback
import webbrowser
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from mr2projekt import settings
from mr2projekt.pipeline import ConversionResult, convert, default_output_path


class Application:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Отчёт СЗЗ")
        self.root.minsize(780, 520)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        self.mr_var = tk.StringVar()
        self.sg_var = tk.StringVar()
        self.ss_var = tk.StringVar()
        self.template_var = tk.StringVar(value=str(settings.template_path()))
        self.output_var = tk.StringVar()
        self.vocw_var = tk.StringVar(value=str(settings.vocw_dir()))
        self.status_var = tk.StringVar(value="Выберите файлы и нажмите «Сформировать».")

        self._running = False
        self.last_output: Path | None = None

        frame = ttk.Frame(root, padding=12)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(8, weight=1)

        self._file_row(
            frame, 0, "Отчёт МР (максимальные):", self.mr_var,
            lambda: self._browse_report(self.mr_var, "Выберите отчёт МР"),
        )
        self._file_row(
            frame, 1, "Отчёт СГ (среднегодовые):", self.sg_var,
            lambda: self._browse_report(self.sg_var, "Выберите отчёт СГ"),
        )
        self._file_row(
            frame, 2, "Отчёт СС (среднесуточные):", self.ss_var,
            lambda: self._browse_report(self.ss_var, "Выберите отчёт СС"),
        )
        ttk.Label(
            frame,
            text="Если СГ или СС не указаны, они берутся из сводного отчёта МР.",
            foreground="gray",
        ).grid(row=3, column=1, columnspan=2, sticky="w")
        self._file_row(
            frame, 4, "Шаблон отчёта:", self.template_var, self._browse_template
        )
        self._file_row(
            frame, 5, "Результат (пусто — рядом с МР):", self.output_var, self._browse_output
        )
        self._file_row(
            frame, 6, "Справочник VocW (необязательно):", self.vocw_var, self._browse_vocw
        )

        actions = ttk.Frame(frame)
        actions.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(12, 0))
        self.run_button = ttk.Button(
            actions, text="Сформировать", command=self._start
        )
        self.run_button.pack(side="left")
        ttk.Label(actions, textvariable=self.status_var).pack(
            side="left", padx=(12, 0)
        )

        log_frame = ttk.Frame(frame)
        log_frame.grid(row=8, column=0, columnspan=3, sticky="nsew", pady=(12, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log = tk.Text(
            log_frame,
            height=14,
            wrap="word",
            state="disabled",
            font=("Consolas", 9),
        )
        self.log.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.log.configure(yscrollcommand=scroll.set)
        self.log.bind("<Control-c>", self._copy_log)
        self.log.bind("<Control-C>", self._copy_log)

        bottom = ttk.Frame(frame)
        bottom.grid(row=9, column=0, columnspan=3, sticky="ew", pady=(12, 0))
        self.open_result_button = ttk.Button(
            bottom,
            text="Открыть результат",
            command=self._open_result,
            state="disabled",
        )
        self.open_result_button.pack(side="left")
        ttk.Button(
            bottom, text="Копировать лог", command=self._copy_log
        ).pack(side="left", padx=(8, 0))
        ttk.Button(bottom, text="Выход", command=self.root.destroy).pack(
            side="right"
        )

    def _file_row(
        self,
        parent: ttk.Frame,
        row: int,
        label: str,
        variable: tk.StringVar,
        command,
    ) -> None:
        ttk.Label(parent, text=label).grid(
            row=row, column=0, sticky="w", padx=(0, 8), pady=4
        )
        ttk.Entry(parent, textvariable=variable).grid(
            row=row, column=1, sticky="ew", pady=4
        )
        ttk.Button(parent, text="Обзор…", command=command).grid(
            row=row, column=2, padx=(8, 0), pady=4
        )

    def _browse_report(self, variable: tk.StringVar, title: str) -> None:
        current = variable.get().strip() or self.mr_var.get()
        path = filedialog.askopenfilename(
            title=title,
            initialdir=str(Path(current).parent),
            filetypes=[("Word documents", "*.docx")],
        )
        if path:
            variable.set(path)

    def _browse_template(self) -> None:
        path = filedialog.askopenfilename(
            title="Выберите шаблон отчёта",
            initialdir=str(Path(self.template_var.get()).parent),
            filetypes=[("Word documents", "*.docx")],
        )
        if path:
            self.template_var.set(path)

    def _browse_output(self) -> None:
        current = self.output_var.get().strip()
        mr_text = self.mr_var.get().strip()
        if not current and mr_text:
            current = str(default_output_path(mr_text))
        path = filedialog.asksaveasfilename(
            title="Куда сохранить результат",
            initialdir=str(Path(current).parent) if current else None,
            initialfile=Path(current).name if current else "",
            defaultextension=".docx",
            filetypes=[("Word documents", "*.docx")],
        )
        if path:
            self.output_var.set(path)

    def _browse_vocw(self) -> None:
        path = filedialog.askdirectory(
            title="Каталог справочника VocW",
            initialdir=self.vocw_var.get(),
        )
        if path:
            self.vocw_var.set(path)

    def _log(self, text: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _clear_log(self) -> None:
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def _copy_text(self, text: str) -> None:
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.root.update_idletasks()

    def _copy_log(self, _event=None):
        try:
            text = self.log.get("sel.first", "sel.last")
        except tk.TclError:
            text = self.log.get("1.0", "end-1c")
        self._copy_text(text)
        self.status_var.set("Текст скопирован в буфер обмена")
        return "break"

    def _start(self) -> None:
        if self._running:
            return

        mr_path = Path(self.mr_var.get().strip())
        sg_text = self.sg_var.get().strip()
        ss_text = self.ss_var.get().strip()
        sg_path = Path(sg_text) if sg_text else None
        ss_path = Path(ss_text) if ss_text else None
        template_path = Path(self.template_var.get().strip())
        if not self.mr_var.get().strip():
            messagebox.showerror("Ошибка", "Укажите отчёт МР.")
            return
        output_text = self.output_var.get().strip()
        output_path = Path(output_text) if output_text else default_output_path(mr_path)
        vocw_dir = Path(self.vocw_var.get().strip() or settings.vocw_dir())
        if output_path.suffix.lower() != ".docx":
            output_path = output_path.with_suffix(".docx")

        for title, path in (("МР", mr_path), ("СГ", sg_path), ("СС", ss_path)):
            if path is not None and not path.is_file():
                messagebox.showerror("Ошибка", f"Не найден отчёт {title}:\n{path}")
                return
        if not template_path.is_file():
            messagebox.showerror(
                "Ошибка", f"Не найден шаблон отчёта:\n{template_path}"
            )
            return
        inputs = {mr_path, template_path, sg_path, ss_path} - {None}
        if output_path.resolve() in {path.resolve() for path in inputs}:
            messagebox.showerror(
                "Ошибка",
                "Результат нельзя сохранять поверх отчётов или шаблона.",
            )
            return

        self._running = True
        self.run_button.state(["disabled"])
        self.open_result_button.state(["disabled"])
        self.status_var.set("Обработка…")
        self._clear_log()
        self._log(f"МР: {mr_path}")
        self._log(f"СГ: {sg_path or 'из отчёта МР'}")
        self._log(f"СС: {ss_path or 'из отчёта МР'}")
        self._log(f"Шаблон: {template_path}")
        self._log(f"Результат: {output_path}")
        self._log(f"Справочник VocW: {vocw_dir}")
        self._log("")

        thread = threading.Thread(
            target=self._worker,
            args=(
                mr_path,
                sg_path,
                ss_path,
                template_path,
                output_path,
                vocw_dir,
            ),
            daemon=True,
        )
        thread.start()

    def _worker(
        self,
        mr_path: Path,
        sg_path: Path | None,
        ss_path: Path | None,
        template_path: Path,
        output_path: Path,
        vocw_dir: Path,
    ) -> None:
        try:
            result = convert(
                mr_path,
                template_path,
                output_path,
                annual_path=sg_path,
                daily_path=ss_path,
                vocw_dir=vocw_dir,
            )
        except Exception as exc:
            self.root.after(0, self._finish_error, exc, traceback.format_exc())
            return
        self.root.after(0, self._finish_ok, result)

    def _finish_ok(self, result: ConversionResult) -> None:
        self._running = False
        self.run_button.state(["!disabled"])
        summary = result.summary
        warnings = list(summary.get("warnings", []))
        self.last_output = Path(str(summary["output"]))
        self.open_result_button.state(["!disabled"])

        self._log("Готово.")
        self._log(f"Источников: {summary['sources']}")
        self._log(f"Веществ: {summary['substances']}")
        self._log(f"Групп: {summary['groups']}")
        self._log(f"Точек: {summary['points']}")
        self._log(f"Файл: {self.last_output}")
        if summary.get("appendix"):
            self._log(f"Приложение: {summary['appendix']}")

        vocw_text = self.vocw_var.get().strip()
        if vocw_text and Path(vocw_text) != settings.vocw_dir():
            try:
                settings.save_vocw_dir(vocw_text)
                self._log(f"Каталог справочника запомнен: {vocw_text}")
            except OSError as exc:
                self._log(f"Не удалось сохранить каталог справочника: {exc}")

        if warnings:
            self._log("")
            self._log("Предупреждения:")
            for warning in warnings:
                self._log(f"  • {warning}")

        self.status_var.set("Готово" + (" (с предупреждениями)" if warnings else ""))

    def _finish_error(self, exc: Exception, details: str) -> None:
        self._running = False
        self.run_button.state(["!disabled"])
        self._log("Ошибка:")
        self._log(details)
        self._copy_text(self.log.get("1.0", "end-1c"))
        self.status_var.set("Ошибка — текст скопирован в буфер обмена")
        self._log("")
        self._log("Текст ошибки скопирован в буфер обмена.")
        message = str(exc) or type(exc).__name__
        messagebox.showerror(
            "Ошибка (текст скопирован)",
            f"{message}\n\nПодробности — в логе приложения.",
        )

    def _open_result(self) -> None:
        if self.last_output is not None:
            self._open_path(self.last_output)

    def _open_path(self, path: Path) -> None:
        if not path.exists():
            messagebox.showerror("Ошибка", f"Файл не найден:\n{path}")
            return
        if hasattr(os, "startfile"):
            os.startfile(str(path))
        else:
            webbrowser.open(path.resolve().as_uri())


def main() -> None:
    root = tk.Tk()
    Application(root)
    root.mainloop()


if __name__ == "__main__":
    main()

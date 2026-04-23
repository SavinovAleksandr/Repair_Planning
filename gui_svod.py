# -*- coding: utf-8 -*-
"""
GUI-обёртка для сборщика сводного графика ремонтов.

Окно даёт несколько кнопок:
  • «Сформировать сводный из проектов» — объединяет Арх + Коми (без сортировки
    и оформления) в новый файл «Сводный график …xlsx» в корне.
  • «Всё и сразу (пересобрать с нуля)» — запускает полный конвейер:
    объединение → приоритеты → нормализация → оглавление → высоты → Гант.
  • Чекбоксы + кнопка «Выполнить отмеченное» — точечные стадии над уже
    существующим сводником (без пересборки с нуля):
        ☑ Расстановка по приоритетам (полный rebuild из свода)
        ☑ Нормализация текста H/N
        ☑ Оглавление
        ☑ Фиксация высот + wrap
        ☑ Диаграмма Ганта
  • «Откатить к предыдущей версии» — восстановление из backups/.
  • «Открыть файл» / «Открыть папку» — быстрые ярлыки.

Все операции выполняются в отдельном потоке, чтобы окно не зависало; лог
и статус-строка обновляются из очереди сообщений.
"""
from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import traceback
from datetime import datetime
from pathlib import Path

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

import build_svod as bs


# ------------------------------------------------------------------ КОНСТАНТЫ

PAD = 8
ROOT_DIR = bs.ROOT


# ------------------------------------------------------------------ ХЕЛПЕРЫ


def open_in_system(path: Path) -> None:
    """Открывает файл/папку в системе: Windows — start, macOS — open, Linux — xdg-open."""
    if not path.exists():
        messagebox.showwarning("Нет файла", f"Не найден:\n{path}")
        return
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
        else:
            subprocess.run(["xdg-open", str(path)], check=False)
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось открыть:\n{path}\n\n{e}")


# ---------------------------------------------------------------- ОКНО


class SvodApp(tk.Tk):
    """Главное окно GUI."""

    def __init__(self) -> None:
        super().__init__()
        self.title("График ремонтов — сводный")
        self.geometry("820x620")
        self.minsize(700, 560)

        # Очередь «GUI ← рабочий поток» для лога и статуса.
        self.msg_q: queue.Queue[tuple[str, str]] = queue.Queue()
        self.worker: threading.Thread | None = None

        # Переменные чекбоксов «к существующему своднику».
        self.var_sort = tk.BooleanVar(value=True)
        self.var_norm = tk.BooleanVar(value=True)
        self.var_toc = tk.BooleanVar(value=True)
        self.var_heights = tk.BooleanVar(value=True)
        self.var_gantt = tk.BooleanVar(value=True)
        # Нормализация в «Всё и сразу».
        self.var_all_norm = tk.BooleanVar(value=True)

        self._build_ui()
        self._refresh_status()
        # Периодический опрос очереди сообщений.
        self.after(120, self._pump_messages)

    # ----------------------------------------------------- UI

    def _build_ui(self) -> None:
        s = ttk.Style(self)
        # На macOS по умолчанию тема 'aqua'. Если недоступна — 'clam'.
        try:
            s.theme_use(s.theme_use())
        except Exception:
            pass

        # Верхняя полоса: путь к рабочей папке + кнопка «Открыть папку».
        top = ttk.Frame(self, padding=(PAD, PAD, PAD, 0))
        top.pack(fill=tk.X)
        ttk.Label(top, text="Папка:").pack(side=tk.LEFT)
        self.path_var = tk.StringVar(value=str(ROOT_DIR))
        ttk.Entry(top, textvariable=self.path_var, state="readonly").pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 4))
        ttk.Button(top, text="Открыть папку",
                   command=lambda: open_in_system(ROOT_DIR)).pack(side=tk.LEFT)

        # Блок «Собрать из проектов».
        g1 = ttk.LabelFrame(self, text="Сборка из проектов",
                            padding=(PAD, PAD, PAD, PAD))
        g1.pack(fill=tk.X, padx=PAD, pady=(PAD, 0))

        btn_merge = ttk.Button(
            g1, text="Сформировать сводный (только объединение)",
            command=self._on_merge_only)
        btn_merge.pack(side=tk.LEFT, padx=(0, PAD))

        btn_all = ttk.Button(
            g1, text="Всё и сразу (пересобрать с нуля)",
            command=self._on_all,
            style="Accent.TButton")
        btn_all.pack(side=tk.LEFT)

        ttk.Checkbutton(
            g1, text="С нормализацией текста",
            variable=self.var_all_norm,
        ).pack(side=tk.LEFT, padx=(PAD, 0))

        # Блок «Применить к существующему своднику».
        g2 = ttk.LabelFrame(
            self, text="Применить к существующему своднику",
            padding=(PAD, PAD, PAD, PAD))
        g2.pack(fill=tk.X, padx=PAD, pady=(PAD, 0))

        grid = ttk.Frame(g2)
        grid.pack(fill=tk.X)
        ttk.Checkbutton(grid, text="Расстановка по приоритетам",
                        variable=self.var_sort).grid(
            row=0, column=0, sticky="w", padx=(0, 16), pady=2)
        ttk.Checkbutton(grid, text="Нормализация текста (H/N)",
                        variable=self.var_norm).grid(
            row=0, column=1, sticky="w", padx=(0, 16), pady=2)
        ttk.Checkbutton(grid, text="Оглавление (TOC)",
                        variable=self.var_toc).grid(
            row=1, column=0, sticky="w", padx=(0, 16), pady=2)
        ttk.Checkbutton(grid, text="Фиксация высот + wrap",
                        variable=self.var_heights).grid(
            row=1, column=1, sticky="w", padx=(0, 16), pady=2)
        ttk.Checkbutton(grid, text="Диаграмма Ганта",
                        variable=self.var_gantt).grid(
            row=2, column=0, sticky="w", padx=(0, 16), pady=2)

        btn_apply = ttk.Button(
            g2, text="Выполнить отмеченное",
            command=self._on_apply_selected,
            style="Accent.TButton")
        btn_apply.pack(anchor="e", pady=(PAD, 0))

        # Откат + открыть файл.
        g3 = ttk.Frame(self, padding=(PAD, PAD, PAD, 0))
        g3.pack(fill=tk.X)
        ttk.Button(g3, text="Откатить к предыдущей версии",
                   command=self._on_restore).pack(side=tk.LEFT)
        ttk.Button(g3, text="Открыть сводник в Excel",
                   command=self._on_open_svod).pack(side=tk.LEFT, padx=(PAD, 0))

        # Статус-строка.
        sf = ttk.Frame(self, padding=(PAD, PAD, PAD, 0))
        sf.pack(fill=tk.X)
        self.status_var = tk.StringVar(value="")
        ttk.Label(sf, textvariable=self.status_var, foreground="#555").pack(
            side=tk.LEFT)

        # Лог.
        logf = ttk.LabelFrame(self, text="Лог", padding=(PAD, PAD, PAD, PAD))
        logf.pack(fill=tk.BOTH, expand=True, padx=PAD, pady=PAD)
        self.log = scrolledtext.ScrolledText(
            logf, wrap="word", height=14, font=("Menlo", 10))
        self.log.pack(fill=tk.BOTH, expand=True)
        self.log.configure(state="disabled")

        # Кнопку-акцент обустроим покрасивее, где тема поддерживает.
        try:
            s.configure("Accent.TButton", font=("Helvetica", 11, "bold"))
        except Exception:
            pass

    # ----------------------------------------------------- СТАТУС/ЛОГ

    def _refresh_status(self) -> None:
        svod = bs.find_existing_svod(ROOT_DIR)
        if svod is None:
            self.status_var.set("Сводник в папке не найден. "
                                "Доступна только «Сборка из проектов».")
        else:
            mt = datetime.fromtimestamp(svod.stat().st_mtime).strftime(
                "%Y-%m-%d %H:%M:%S")
            self.status_var.set(f"Текущий сводник: {svod.name} · обновлён {mt}")

    def _log(self, msg: str) -> None:
        ts = datetime.now().strftime("%H:%M:%S")
        self.log.configure(state="normal")
        self.log.insert("end", f"[{ts}] {msg}\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _push(self, kind: str, text: str) -> None:
        """Из рабочего потока кладёт сообщение в очередь для GUI."""
        self.msg_q.put((kind, text))

    def _pump_messages(self) -> None:
        try:
            while True:
                kind, text = self.msg_q.get_nowait()
                if kind == "log":
                    self._log(text)
                elif kind == "error":
                    self._log(f"ОШИБКА: {text}")
                    messagebox.showerror("Ошибка", text)
                elif kind == "done":
                    self._log(text)
                    self._refresh_status()
                    self._enable_buttons(True)
        except queue.Empty:
            pass
        self.after(120, self._pump_messages)

    def _enable_buttons(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        for child in self.winfo_children():
            self._walk_state(child, state)

    def _walk_state(self, widget, state: str) -> None:
        for w in widget.winfo_children():
            try:
                if isinstance(w, (ttk.Button, ttk.Checkbutton)):
                    w.configure(state=state)
            except Exception:
                pass
            self._walk_state(w, state)

    # ----------------------------------------------------- ДЕЙСТВИЯ

    def _run_in_thread(self, fn, *args, **kwargs) -> None:
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("Идёт работа",
                                "Дождитесь завершения текущей операции.")
            return
        self._enable_buttons(False)
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

        def _target():
            try:
                fn(*args, **kwargs)
                self._push("done", "Готово.")
            except Exception as e:
                tb = traceback.format_exc()
                self._push("log", tb)
                self._push("error", str(e))
                self._push("done", "Завершено с ошибкой.")

        self.worker = threading.Thread(target=_target, daemon=True)
        self.worker.start()

    def _log_fn(self, msg: str) -> None:
        """Функция-логгер, которая передаётся стадиям build_svod."""
        self._push("log", str(msg))

    # --- кнопки -----------------------------------------------------------

    def _on_merge_only(self) -> None:
        """Чистое объединение проектов без сортировки/TOC/Ганта/высот."""
        svod = bs.find_existing_svod(ROOT_DIR)
        if svod is not None:
            ok = messagebox.askyesno(
                "Перезаписать сводник?",
                f"В папке уже есть файл:\n{svod.name}\n\n"
                "«Сформировать» пересоздаст его из проектов Арх/Коми — "
                "ручные правки в нём пропадут (старая версия попадёт в backups/).\n\n"
                "Продолжить?",
            )
            if not ok:
                return

        opts = bs.NormOptions(enabled=self.var_all_norm.get())
        stats = bs.NormStats()

        def run():
            self._push("log", "=== Сформировать сводный (только объединение) ===")
            out = bs.stage_full_rebuild(
                ROOT_DIR, None, opts, stats, log=self._log_fn,
                apply_sort=False, apply_toc=False,
                apply_heights=False, apply_gantt=False,
            )
            self._push("log", f"Готов файл: {out.name}")
            self._report_norm(stats)

        self._run_in_thread(run)

    def _on_all(self) -> None:
        """Полная пересборка с нуля."""
        svod = bs.find_existing_svod(ROOT_DIR)
        if svod is not None:
            ok = messagebox.askyesno(
                "Перезаписать сводник?",
                f"В папке уже есть файл:\n{svod.name}\n\n"
                "Пересобрать его заново из проектов (ручные правки пропадут, "
                "старая версия уедет в backups/)?",
            )
            if not ok:
                return

        opts = bs.NormOptions(enabled=self.var_all_norm.get())
        stats = bs.NormStats()

        def run():
            self._push("log", "=== Всё и сразу: пересборка с нуля ===")
            out = bs.stage_full_rebuild(
                ROOT_DIR, None, opts, stats, log=self._log_fn,
                apply_sort=True, apply_toc=True,
                apply_heights=True, apply_gantt=True,
            )
            self._push("log", f"Готов файл: {out.name}")
            self._report_norm(stats)

        self._run_in_thread(run)

    def _on_apply_selected(self) -> None:
        """Применить отмеченные стадии к существующему своднику."""
        svod = bs.find_existing_svod(ROOT_DIR)
        if svod is None:
            messagebox.showerror(
                "Нет сводника",
                f"В папке {ROOT_DIR} не найден «Сводный график …xlsx».\n\n"
                "Сначала нажмите «Сформировать» или «Всё и сразу».",
            )
            return

        do_sort = self.var_sort.get()
        do_norm = self.var_norm.get()
        do_toc = self.var_toc.get()
        do_heights = self.var_heights.get()
        do_gantt = self.var_gantt.get()
        if not any([do_sort, do_norm, do_toc, do_heights, do_gantt]):
            messagebox.showinfo("Ничего не выбрано",
                                "Отметьте хотя бы одну стадию.")
            return

        opts = bs.NormOptions(enabled=True)
        stats = bs.NormStats()

        def run():
            nonlocal svod
            self._push("log", "=== Применить отмеченные стадии ===")
            # Если нужна расстановка — делаем полный rebuild из существующего:
            # он заодно пересчитает всё остальное (TOC/высоты/Гант).
            if do_sort:
                self._push("log", "→ Расстановка по приоритетам (rebuild)")
                svod = bs.stage_rebuild_from_existing(
                    svod, None, opts, stats, log=self._log_fn)
                # После rebuild остальные стадии-inplace могут быть не нужны —
                # запустим их только если пользователь всё ещё хочет
                # (например, выключил в rebuild нормализацию).
            if do_norm and not do_sort:
                # Если был sort, нормализация уже применена в rebuild.
                self._push("log", "→ Нормализация текста")
                bs.stage_normalize_inplace(svod, opts, stats, log=self._log_fn)
            if do_toc and not do_sort:
                self._push("log", "→ Оглавление")
                bs.stage_build_toc_inplace(svod, log=self._log_fn)
            if do_heights and not do_sort:
                self._push("log", "→ Фиксация высот + wrap")
                bs.stage_set_heights_inplace(svod, log=self._log_fn)
            if do_gantt and not do_sort:
                self._push("log", "→ Диаграмма Ганта")
                bs.stage_build_gantt_inplace(
                    svod, datetime.now().year, log=self._log_fn)
            self._report_norm(stats)

        self._run_in_thread(run)

    def _on_restore(self) -> None:
        """Откат свода к последней резервной копии."""
        svod = bs.find_existing_svod(ROOT_DIR)
        if svod is None:
            messagebox.showerror(
                "Нет сводника",
                f"В папке {ROOT_DIR} нет файла «Сводный график …xlsx» — "
                "не к чему откатываться.",
            )
            return
        ok = messagebox.askyesno(
            "Откат",
            "Заменить текущий сводник на последнюю резервную копию из backups/?\n\n"
            "Текущая версия будет положена в backups/ на случай, если понадобится.",
        )
        if not ok:
            return

        def run():
            self._push("log", "=== Откат к предыдущей версии ===")
            restored = bs.restore_latest_backup(svod, log=self._log_fn)
            if restored is None:
                raise RuntimeError("Нет подходящих копий в backups/.")

        self._run_in_thread(run)

    def _on_open_svod(self) -> None:
        svod = bs.find_existing_svod(ROOT_DIR)
        if svod is None:
            messagebox.showinfo("Нет сводника",
                                "В папке не найден сводный график.")
            return
        open_in_system(svod)

    # --- вспомогательное ---------------------------------------------------

    def _report_norm(self, stats: bs.NormStats) -> None:
        """Выводит в лог краткую статистику нормализации."""
        if not stats.counts:
            return
        self._push("log", "Нормализация:")
        for label, n in sorted(stats.counts.items(),
                               key=lambda kv: (-kv[1], kv[0])):
            self._push("log", f"  • {label}: {n}")


def main() -> None:
    app = SvodApp()
    app.mainloop()


if __name__ == "__main__":
    main()

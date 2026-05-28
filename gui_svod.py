# -*- coding: utf-8 -*-
"""
GUI-обёртка для сборщика сводного графика ремонтов.

Copyright (c) 2026 Савинов Александр, Сыктывкар. Все права защищены.

Две основные кнопки:
  • «Всё и сразу» — полный конвейер (объединение → приоритеты → нормализация
    → оглавление → высоты → Гант → сравнение с проектами).
  • «Выполнить отмеченное» — только отмеченные чекбоксы.

Чекбоксы: только объединение, приоритеты, нормализация, оглавление,
высоты, Гант, сравнение с проектами.

Также: откат из backups/, открыть сводник / папку.
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

ROOT_DIR = Path(__file__).resolve().parent

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

import build_svod as bs

# Всегда работаем из папки gui_svod.py (важно для pythonw / двойного клика).
os.chdir(ROOT_DIR)
bs.ROOT = ROOT_DIR

# ------------------------------------------------------------------ КОНСТАНТЫ

PAD = 8
ERROR_LOG = ROOT_DIR / "gui_error.log"
LOG_FONT = ("Consolas", 10) if sys.platform.startswith("win") else ("Menlo", 10)
COPYRIGHT = "© Савинов Александр, Сыктывкар, 2026"


# ------------------------------------------------------------------ ХЕЛПЕРЫ


def bring_window_to_front(root: tk.Tk) -> None:
    """Поднимает окно на передний план (важно при запуске через pythonw / bat)."""
    try:
        root.update_idletasks()
        if root.state() == "iconic":
            root.deiconify()
        root.state("normal")
        root.lift()
        root.attributes("-topmost", True)
        root.focus_force()
        root.after(250, lambda: root.attributes("-topmost", False))
    except tk.TclError:
        pass
    if not sys.platform.startswith("win"):
        return
    try:
        import ctypes

        hwnd = root.winfo_id()
        u32 = ctypes.windll.user32
        u32.ShowWindow(hwnd, 9)  # SW_RESTORE
        u32.AllowSetForegroundWindow(-1)
        u32.SetForegroundWindow(hwnd)
    except Exception:
        pass


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

        # Переменные чекбоксов стадий.
        self.var_merge = tk.BooleanVar(value=False)
        self.var_sort = tk.BooleanVar(value=True)
        self.var_norm = tk.BooleanVar(value=True)
        self.var_toc = tk.BooleanVar(value=True)
        self.var_heights = tk.BooleanVar(value=True)
        self.var_gantt = tk.BooleanVar(value=True)
        self.var_diff = tk.BooleanVar(value=False)

        self._build_ui()
        self._refresh_status()
        bring_window_to_front(self)
        self.after(150, lambda: bring_window_to_front(self))
        self.after(600, lambda: bring_window_to_front(self))
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

        # Блок операций: чекбоксы + две кнопки.
        g_ops = ttk.LabelFrame(self, text="Операции", padding=(PAD, PAD, PAD, PAD))
        g_ops.pack(fill=tk.X, padx=PAD, pady=(PAD, 0))

        grid = ttk.Frame(g_ops)
        grid.pack(fill=tk.X)
        ttk.Checkbutton(grid, text="Только объединение",
                        variable=self.var_merge).grid(
            row=0, column=0, sticky="w", padx=(0, 16), pady=2)
        ttk.Checkbutton(grid, text="Расстановка по приоритетам",
                        variable=self.var_sort).grid(
            row=0, column=1, sticky="w", padx=(0, 16), pady=2)
        ttk.Checkbutton(grid, text="Нормализация текста (H/N)",
                        variable=self.var_norm).grid(
            row=1, column=0, sticky="w", padx=(0, 16), pady=2)
        ttk.Checkbutton(grid, text="Оглавление (TOC)",
                        variable=self.var_toc).grid(
            row=1, column=1, sticky="w", padx=(0, 16), pady=2)
        ttk.Checkbutton(grid, text="Фиксация высот + wrap",
                        variable=self.var_heights).grid(
            row=2, column=0, sticky="w", padx=(0, 16), pady=2)
        ttk.Checkbutton(grid, text="Диаграмма Ганта",
                        variable=self.var_gantt).grid(
            row=2, column=1, sticky="w", padx=(0, 16), pady=2)
        ttk.Checkbutton(grid, text="Сравнение с проектами",
                        variable=self.var_diff).grid(
            row=3, column=0, columnspan=2, sticky="w", pady=2)

        btns = ttk.Frame(g_ops)
        btns.pack(fill=tk.X, pady=(PAD, 0))
        ttk.Button(
            btns, text="Всё и сразу",
            command=self._on_all,
            style="Accent.TButton",
        ).pack(side=tk.LEFT, padx=(0, PAD))
        ttk.Button(
            btns, text="Выполнить отмеченное",
            command=self._on_apply_selected,
        ).pack(side=tk.LEFT)

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
            logf, wrap="word", height=14, font=LOG_FONT)
        self.log.pack(fill=tk.BOTH, expand=True)
        self.log.configure(state="disabled")

        ttk.Label(
            self,
            text=COPYRIGHT,
            foreground="#888",
            font=("", 9),
        ).pack(side=tk.BOTTOM, anchor="e", padx=PAD, pady=(0, 4))

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
                                "Отметьте «Только объединение» или «Всё и сразу».")
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

    def _confirm_overwrite(self, svod: Path | None) -> bool:
        if svod is None:
            return True
        return messagebox.askyesno(
            "Перезаписать сводник?",
            f"В папке уже есть файл:\n{svod.name}\n\n"
            "Операция пересоздаст его — ручные правки пропадут "
            "(старая версия попадёт в backups/).\n\n"
            "Продолжить?",
        )

    def _run_diff(self, svod: Path) -> None:
        p_komi = bs.find_file(bs.FILE_KOMI)
        p_arkh = bs.find_file(bs.FILE_ARKH)
        if not p_komi and not p_arkh:
            raise RuntimeError(
                "Для сравнения положите в папку файлы "
                f"«{bs.FILE_KOMI}» и/или «{bs.FILE_ARKH}»."
            )
        self._push("log", "→ Сравнение с проектами")
        stats = bs.stage_build_diff_inplace(
            svod, ROOT_DIR, None, log=self._log_fn)
        self._push(
            "log",
            f"  изменено {stats.modified}, новых {stats.new_in_svod}, "
            f"удалено из проектов {stats.deleted_from_source}. "
            f"Вкладки «{bs.DIFF_SHEET_NAME}» и «{bs.DIFF_CLEAN_SHEET_NAME}».",
        )

    def _on_all(self) -> None:
        """Полный конвейер: всё сразу."""
        svod = bs.find_existing_svod(ROOT_DIR)
        if not self._confirm_overwrite(svod):
            return

        opts = bs.NormOptions(enabled=True)
        stats = bs.NormStats()

        def run():
            self._push("log", "=== Всё и сразу ===")
            out = bs.stage_full_rebuild(
                ROOT_DIR, None, opts, stats, log=self._log_fn,
                apply_sort=True, apply_toc=True,
                apply_heights=True, apply_gantt=True,
            )
            self._push("log", f"Готов файл: {out.name}")
            self._run_diff(out)
            self._report_norm(stats)

        self._run_in_thread(run)

    def _on_apply_selected(self) -> None:
        """Выполнить только отмеченные стадии."""
        do_merge = self.var_merge.get()
        do_sort = self.var_sort.get()
        do_norm = self.var_norm.get()
        do_toc = self.var_toc.get()
        do_heights = self.var_heights.get()
        do_gantt = self.var_gantt.get()
        do_diff = self.var_diff.get()

        if not any([do_merge, do_sort, do_norm, do_toc,
                    do_heights, do_gantt, do_diff]):
            messagebox.showinfo("Ничего не выбрано",
                                "Отметьте хотя бы одну операцию.")
            return

        svod = bs.find_existing_svod(ROOT_DIR)
        if do_merge:
            if not self._confirm_overwrite(svod):
                return
        elif svod is None:
            messagebox.showerror(
                "Нет сводника",
                f"В папке {ROOT_DIR} не найден «Сводный график …xlsx».\n\n"
                "Отметьте «Только объединение» или нажмите «Всё и сразу».",
            )
            return

        opts = bs.NormOptions(enabled=do_norm)
        stats = bs.NormStats()

        def run():
            nonlocal svod
            self._push("log", "=== Выполнить отмеченное ===")

            if do_merge:
                self._push("log", "→ Объединение проектов Арх/Коми")
                svod = bs.stage_full_rebuild(
                    ROOT_DIR, None, opts, stats, log=self._log_fn,
                    apply_sort=do_sort,
                    apply_toc=do_toc and not do_sort,
                    apply_heights=do_heights and not do_sort,
                    apply_gantt=do_gantt and not do_sort,
                )
                self._push("log", f"  файл: {svod.name}")
            elif do_sort:
                self._push("log", "→ Расстановка по приоритетам (rebuild)")
                svod = bs.stage_rebuild_from_existing(
                    svod, None, opts, stats, log=self._log_fn)

            if do_norm and not do_merge and not do_sort:
                self._push("log", "→ Нормализация текста")
                bs.stage_normalize_inplace(svod, opts, stats, log=self._log_fn)
            if do_toc and not do_merge and not do_sort:
                self._push("log", "→ Оглавление")
                bs.stage_build_toc_inplace(svod, log=self._log_fn)
            if do_heights and not do_merge and not do_sort:
                self._push("log", "→ Фиксация высот + wrap")
                bs.stage_set_heights_inplace(svod, log=self._log_fn)
            if do_gantt and not do_merge and not do_sort:
                self._push("log", "→ Диаграмма Ганта")
                bs.stage_build_gantt_inplace(svod, None, log=self._log_fn)
            if do_diff:
                self._run_diff(svod)

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


def _report_startup_error(text: str) -> None:
    """Пишет ошибку в gui_error.log и показывает окно (если tkinter жив)."""
    try:
        ERROR_LOG.write_text(text, encoding="utf-8")
    except OSError:
        pass
    try:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Ошибка запуска",
            f"Не удалось открыть окно программы.\n\n"
            f"Подробности сохранены в:\n{ERROR_LOG}\n\n"
            f"{text[:800]}",
        )
        root.destroy()
    except Exception:
        pass


def main() -> None:
    try:
        app = SvodApp()
        app.mainloop()
    except Exception:
        _report_startup_error(traceback.format_exc())
        raise SystemExit(1)


if __name__ == "__main__":
    main()

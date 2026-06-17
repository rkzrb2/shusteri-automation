#!/usr/bin/env python3
"""
Точка входа для GUI-режима Shusteri Automation.

Запуск:
    python ui_run.py
"""
import logging
import os
import sys
from pathlib import Path

# Когда запущено как .exe (PyInstaller --onefile), CWD ставим рядом с исполняемым
# файлом — чтобы папки input/, output/, presets/ находились там же.
if getattr(sys, "frozen", False):
    os.chdir(Path(sys.executable).parent)

# Создаём рабочие папки при первом запуске
for _d in ["input", "output", "presets", "выгрузка честный знак"]:
    Path(_d).mkdir(exist_ok=True)

# Настройка логирования до импорта UI
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
    datefmt="%H:%M:%S",
    handlers=[
        logging.FileHandler("automation.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)

try:
    import customtkinter as ctk
except ImportError:
    print("Ошибка: customtkinter не установлен.")
    print("Установите: pip install customtkinter")
    sys.exit(1)

import tkinter as tk


def _patch_ctk_entry_clipboard():
    """
    На Windows CustomTkinter перехватывает <Control-v/c/x/a> на уровне
    внутреннего tk.Entry и не всегда корректно обрабатывает их.
    Патчим CTkEntry.__init__ чтобы подменять эти привязки на рабочие
    до создания любого виджета.
    """
    _orig = ctk.CTkEntry.__init__

    def _new_init(self, *args, **kwargs):
        _orig(self, *args, **kwargs)
        inner = self._entry  # внутренний tk.Entry

        def paste(_):
            try:
                text = inner.clipboard_get()
                if inner.selection_present():
                    inner.delete("sel.first", "sel.last")
                inner.insert("insert", text)
            except tk.TclError:
                pass
            return "break"

        def copy(_):
            try:
                if inner.selection_present():
                    inner.clipboard_clear()
                    inner.clipboard_append(inner.selection_get())
            except tk.TclError:
                pass
            return "break"

        def cut(_):
            try:
                if inner.selection_present():
                    inner.clipboard_clear()
                    inner.clipboard_append(inner.selection_get())
                    inner.delete("sel.first", "sel.last")
            except tk.TclError:
                pass
            return "break"

        def select_all(_):
            inner.select_range(0, "end")
            inner.icursor("end")
            return "break"

        inner.bind("<Control-v>", paste)
        inner.bind("<Control-c>", copy)
        inner.bind("<Control-x>", cut)
        inner.bind("<Control-a>", select_all)

    ctk.CTkEntry.__init__ = _new_init


_patch_ctk_entry_clipboard()

from ui.app import App


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()

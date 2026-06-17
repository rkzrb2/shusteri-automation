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
    level=logging.DEBUG,
    format="%(asctime)s  %(levelname)-7s  %(name)s  %(message)s",
    datefmt="%H:%M:%S",
    handlers=[
        logging.FileHandler("automation.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
# Подавляем шум от сторонних библиотек, оставляем clipboard DEBUG
logging.getLogger("PIL").setLevel(logging.WARNING)
logging.getLogger("customtkinter").setLevel(logging.WARNING)

try:
    import customtkinter as ctk
except ImportError:
    print("Ошибка: customtkinter не установлен.")
    print("Установите: pip install customtkinter")
    sys.exit(1)

import tkinter as tk
import ctypes

_clip_log = logging.getLogger("clipboard")


def _win_clipboard_get() -> str | None:
    """Читает CF_UNICODETEXT из Windows clipboard напрямую через WinAPI."""
    CF_UNICODETEXT = 13
    try:
        user32   = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32
        if not user32.OpenClipboard(0):
            return None
        try:
            h = user32.GetClipboardData(CF_UNICODETEXT)
            if not h:
                return None
            p = kernel32.GlobalLock(h)
            if not p:
                return None
            try:
                return ctypes.wstring_at(p)
            finally:
                kernel32.GlobalUnlock(h)
        finally:
            user32.CloseClipboard()
    except Exception as e:
        _clip_log.warning(f"WinAPI clipboard error: {e}")
        return None


def _patch_ctk_entry_clipboard():
    """
    На Windows Tk конвертирует Ctrl+V в виртуальное событие <<Paste>>
    до того как оно доходит до виджета. Патчим CTkEntry.__init__,
    чтобы заменить стандартный обработчик на наш с WinAPI fallback.
    """
    _orig = ctk.CTkEntry.__init__

    def _new_init(self, *args, **kwargs):
        _orig(self, *args, **kwargs)
        inner = self._entry

        def paste(_):
            _clip_log.debug("paste() вызван")
            text = None
            try:
                text = inner.clipboard_get()
                _clip_log.debug(f"clipboard_get OK, len={len(text)}")
            except tk.TclError as e:
                _clip_log.warning(f"clipboard_get TclError: {e} — пробуем WinAPI")
                text = _win_clipboard_get()
                if text:
                    _clip_log.debug(f"WinAPI OK, len={len(text)}")
                else:
                    _clip_log.warning("WinAPI тоже вернул None")
            if text:
                try:
                    if inner.selection_present():
                        inner.delete("sel.first", "sel.last")
                    inner.insert("insert", text)
                except tk.TclError as e:
                    _clip_log.error(f"insert failed: {e}")
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

        for seq in ("<Control-v>", "<<Paste>>"):
            inner.bind(seq, paste)
        for seq in ("<Control-c>", "<<Copy>>"):
            inner.bind(seq, copy)
        for seq in ("<Control-x>", "<<Cut>>"):
            inner.bind(seq, cut)
        for seq in ("<Control-a>", "<<SelectAll>>"):
            inner.bind(seq, select_all)

    ctk.CTkEntry.__init__ = _new_init


_patch_ctk_entry_clipboard()

from ui.app import App


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()

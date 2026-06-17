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
    import customtkinter  # noqa: F401
except ImportError:
    print("Ошибка: customtkinter не установлен.")
    print("Установите: pip install customtkinter")
    sys.exit(1)

from ui.app import App


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()

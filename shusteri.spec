# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec для сборки Shusteri Automation в один .exe файл.

Сборка:
    pyinstaller shusteri.spec
"""
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# Данные customtkinter (темы, изображения)
datas = collect_data_files("customtkinter")

hidden_imports = [
    # UI
    "customtkinter",
    "PIL._tkinter_finder",
    # Данные / сериализация
    "yaml",
    "pydantic",
    "pydantic.deprecated.class_validators",
    "pydantic.deprecated.config",
    "pydantic.v1",
    # Excel
    "openpyxl",
    "openpyxl.styles",
    "openpyxl.styles.fills",
    "openpyxl.styles.borders",
    "openpyxl.styles.alignment",
    "openpyxl.utils",
    "pandas",
    "xlsxwriter",
    # Стандартные модули которые PyInstaller может не найти
    "decimal",
    "logging",
    "threading",
    "queue",
]
hidden_imports += collect_submodules("customtkinter")

a = Analysis(
    ["ui_run.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "matplotlib",
        "scipy",
        "sklearn",
        "notebook",
        "jupyter",
        "IPython",
        "PyQt5",
        "PyQt6",
        "PySide2",
        "PySide6",
        "wx",
        "gi",
    ],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="ShusteriAutomation",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,          # UPX может ломать на некоторых антивирусах
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,      # Без консольного окна
    disable_windowed_traceback=False,
    icon=None,          # Заменить на путь к .ico когда будет готов
)

"""
Вкладка «История» — список файлов из папки output/.
"""
import subprocess
import sys
from pathlib import Path
from datetime import datetime

import customtkinter as ctk



class HistoryView(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent, corner_radius=0, fg_color="transparent")
        self._body_font  = ctk.CTkFont(size=12)
        self._small_font = ctk.CTkFont(size=11)
        self._build()
        self._load()

    def _build(self):
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=16, pady=(16, 8))
        ctk.CTkLabel(header, text="История генераций", font=ctk.CTkFont(size=16, weight="bold")).pack(side="left")
        ctk.CTkButton(header, text="↻  Обновить", width=110, font=self._body_font, command=self._load).pack(side="right")

        # Заголовок таблицы
        col_header = ctk.CTkFrame(self, corner_radius=0, fg_color=("gray80", "gray22"))
        col_header.pack(fill="x", padx=16)
        for text, w in [("Дата изменения", 150), ("Имя файла", 0), ("Размер", 70)]:
            ctk.CTkLabel(col_header, text=text, font=self._small_font, width=w, anchor="w").pack(
                side="left", padx=8, pady=6
            )

        # Скроллируемый список
        self._scroll = ctk.CTkScrollableFrame(self, corner_radius=10, fg_color="transparent")
        self._scroll.pack(fill="both", expand=True, padx=16, pady=(0, 16))

    def _load(self):
        for w in self._scroll.winfo_children():
            w.destroy()

        out_dir = Path("output")
        if not out_dir.exists():
            ctk.CTkLabel(self._scroll, text="Папка output/ не найдена", font=self._body_font, text_color="gray55").pack(pady=20)
            return

        files = sorted(
            [f for f in out_dir.glob("*.xlsx") if not f.name.startswith("~")],
            key=lambda f: f.stat().st_mtime, reverse=True,
        )
        if not files:
            ctk.CTkLabel(self._scroll, text="Нет сгенерированных файлов", font=self._body_font, text_color="gray55").pack(pady=20)
            return

        for f in files:
            self._add_row(f)

    def _add_row(self, f: Path):
        row = ctk.CTkFrame(self._scroll, corner_radius=6)
        row.pack(fill="x", pady=2)

        mtime   = datetime.fromtimestamp(f.stat().st_mtime).strftime("%d.%m.%Y %H:%M")
        size_kb = f.stat().st_size / 1024

        ctk.CTkLabel(row, text=mtime, font=self._small_font, text_color="gray55", width=150, anchor="w").pack(side="left", padx=8, pady=6)
        ctk.CTkLabel(row, text=f.name, font=self._body_font, anchor="w").pack(side="left", padx=4, pady=6, expand=True, fill="x")
        ctk.CTkLabel(row, text=f"{size_kb:.0f} KB", font=self._small_font, text_color="gray55", width=70).pack(side="left", padx=4)

        ctk.CTkButton(
            row, text="📂", width=36, height=28, font=self._body_font,
            fg_color="transparent", hover_color=("gray75", "gray30"),
            command=lambda path=f: self._open_file(path),
        ).pack(side="left", padx=(0, 4), pady=4)

        ctk.CTkButton(
            row, text="📁", width=36, height=28, font=self._body_font,
            fg_color="transparent", hover_color=("gray75", "gray30"),
            command=lambda path=f: self._open_folder(path),
        ).pack(side="left", padx=(0, 6), pady=4)

    def _open_file(self, f: Path):
        if sys.platform == "win32":
            subprocess.Popen(["start", "", str(f)], shell=True)

    def _open_folder(self, f: Path):
        if sys.platform == "win32":
            subprocess.Popen(["explorer", "/select,", str(f)])

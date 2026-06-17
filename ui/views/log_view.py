"""
Вкладка «Лог» — просмотр automation.log в реальном времени.
"""
import logging
from pathlib import Path
from datetime import datetime

import customtkinter as ctk

_LOG_FILE  = Path("automation.log")

# Цвета для уровней лога
_LEVEL_COLORS = {
    "ERROR":    "#e05252",
    "WARNING":  "#e0a040",
    "INFO":     None,        # цвет по умолчанию
    "DEBUG":    "#808080",
}


class UILogHandler(logging.Handler):
    """Передаёт записи лога в LogView через callback."""
    def __init__(self, callback):
        super().__init__()
        self._cb = callback

    def emit(self, record):
        try:
            self._cb(self.format(record), record.levelname)
        except Exception:
            pass


class LogView(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent, corner_radius=0, fg_color="transparent")
        self._auto_scroll = ctk.BooleanVar(value=True)
        self._body_font   = ctk.CTkFont(size=12)
        self._mono_font   = ctk.CTkFont(family="Consolas", size=11)
        self._build()
        self._attach_handler()
        self._load_existing()

    def _build(self):
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=16, pady=(16, 8))
        ctk.CTkLabel(header, text="Лог работы", font=ctk.CTkFont(size=16, weight="bold")).pack(side="left")

        btn_frame = ctk.CTkFrame(header, fg_color="transparent")
        btn_frame.pack(side="right")
        ctk.CTkCheckBox(btn_frame, text="Авто-скролл", variable=self._auto_scroll, font=self._body_font).pack(side="left", padx=(0, 8))
        ctk.CTkButton(btn_frame, text="💾  Сохранить", width=100, font=self._body_font, command=self._save).pack(side="left", padx=(0, 8))
        ctk.CTkButton(btn_frame, text="🗑️  Очистить", width=100, font=self._body_font,
                      fg_color="transparent", border_width=1, command=self._clear).pack(side="left")

        self._textbox = ctk.CTkTextbox(
            self, font=self._mono_font, state="disabled", wrap="none",
        )
        self._textbox.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        # Теги цвета (через базовый tk.Text)
        txt = self._textbox._textbox
        txt.tag_config("ERROR",   foreground="#e05252")
        txt.tag_config("WARNING", foreground="#e0a040")
        txt.tag_config("DEBUG",   foreground="#808080")

    def _attach_handler(self):
        handler = UILogHandler(self._append)
        handler.setFormatter(logging.Formatter("%(asctime)s  %(levelname)-7s  %(message)s", datefmt="%H:%M:%S"))
        logging.getLogger().addHandler(handler)

    def _load_existing(self):
        if _LOG_FILE.exists():
            try:
                lines = _LOG_FILE.read_text(encoding="utf-8", errors="replace").splitlines()
                for line in lines[-200:]:   # последние 200 строк
                    level = "INFO"
                    for lvl in _LEVEL_COLORS:
                        if f" {lvl} " in line or f" {lvl}\t" in line:
                            level = lvl
                            break
                    self._append(line, level)
            except Exception:
                pass

    def _append(self, text: str, level: str = "INFO"):
        txt = self._textbox._textbox
        self._textbox.configure(state="normal")
        tag = level if level in _LEVEL_COLORS and _LEVEL_COLORS[level] else ""
        txt.insert("end", text + "\n", tag or ())
        self._textbox.configure(state="disabled")
        if self._auto_scroll.get():
            self._textbox._textbox.see("end")

    def _clear(self):
        self._textbox.configure(state="normal")
        self._textbox.delete("1.0", "end")
        self._textbox.configure(state="disabled")

    def _save(self):
        from tkinter import filedialog
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            initialfile=f"shusteri_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
        )
        if path:
            content = self._textbox._textbox.get("1.0", "end")
            Path(path).write_text(content, encoding="utf-8")

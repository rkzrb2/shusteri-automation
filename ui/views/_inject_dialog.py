"""
Диалог «Добавить КИЗ в готовый файл» — вызывается из вкладки Генерация.
"""
import threading
import queue
from pathlib import Path
from datetime import datetime

import customtkinter as ctk
from tkinter import filedialog



class InjectKIZDialog(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Добавить КИЗ коды в готовый файл")
        self.geometry("560x480")
        self.grab_set()

        self._target_file: Path | None = None
        self._km_file:     Path | None = None
        self._log_queue: queue.Queue = queue.Queue()
        self._body_font  = ctk.CTkFont(size=12)
        self._small_font = ctk.CTkFont(size=11)
        self._build()
        self._refresh_km_list()

    def _build(self):
        pad = {"padx": 16, "pady": 6}

        # Целевой файл
        ctk.CTkLabel(self, text="Целевой файл (из папки output/):", font=self._body_font, anchor="w").pack(fill="x", **pad)
        row1 = ctk.CTkFrame(self, fg_color="transparent")
        row1.pack(fill="x", padx=16, pady=(0, 4))
        ctk.CTkButton(row1, text="📂  Из output/", font=self._body_font, command=self._pick_output).pack(side="left", padx=(0, 8))
        ctk.CTkButton(row1, text="Обзор...", font=self._body_font, fg_color="transparent", border_width=1, command=self._pick_dialog).pack(side="left")
        self._target_lbl = ctk.CTkLabel(self, text="Файл не выбран", font=self._small_font, text_color="gray55", anchor="w")
        self._target_lbl.pack(fill="x", padx=16)

        ctk.CTkFrame(self, height=1, fg_color="gray30").pack(fill="x", padx=16, pady=10)

        # Файл маркировки
        ctk.CTkLabel(self, text="Файл маркировки (Честный знак):", font=self._body_font, anchor="w").pack(fill="x", **pad)
        km_row = ctk.CTkFrame(self, fg_color="transparent")
        km_row.pack(fill="x", padx=16, pady=(0, 4))
        self._km_dropdown = ctk.CTkOptionMenu(km_row, font=self._body_font, width=260, command=self._on_km_selected)
        self._km_dropdown.pack(side="left", padx=(0, 6))
        ctk.CTkButton(km_row, text="↻", width=32, font=self._body_font, command=self._refresh_km_list).pack(side="left", padx=(0, 6))
        ctk.CTkButton(km_row, text="📂  Обзор...", width=110, font=self._body_font,
                      fg_color="transparent", border_width=1,
                      command=self._browse_km_file).pack(side="left")

        ctk.CTkFrame(self, height=1, fg_color="gray30").pack(fill="x", padx=16, pady=10)

        # Режим сохранения
        self._overwrite_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            self, text="Перезаписать исходный файл  (иначе — сохранить как *_with_kiz.xlsx)",
            variable=self._overwrite_var, font=self._small_font,
        ).pack(anchor="w", padx=16, pady=(0, 10))

        # Кнопка запуска
        self._run_btn = ctk.CTkButton(
            self, text="🏷️  Добавить КИЗ коды", height=40,
            font=ctk.CTkFont(size=13, weight="bold"), command=self._run,
        )
        self._run_btn.pack(fill="x", padx=16, pady=(0, 8))

        # Лог
        self._log = ctk.CTkTextbox(self, height=140, font=ctk.CTkFont(family="Consolas", size=11), state="disabled")
        self._log.pack(fill="both", expand=True, padx=16, pady=(0, 12))

    # ------------------------------------------------------------------
    def _pick_output(self):
        out_dir = Path("output")
        files = sorted(
            [f for f in out_dir.glob("*.xlsx") if not f.name.startswith("~")],
            key=lambda f: f.stat().st_mtime, reverse=True,
        ) if out_dir.exists() else []

        if not files:
            self._target_lbl.configure(text="Нет файлов в output/", text_color="orange")
            return
        if len(files) == 1:
            self._set_target(files[0]); return

        win = ctk.CTkToplevel(self)
        win.title("Выберите файл")
        win.geometry("480x280")
        win.grab_set()
        scroll = ctk.CTkScrollableFrame(win)
        scroll.pack(fill="both", expand=True, padx=12, pady=12)
        for f in files:
            mtime = datetime.fromtimestamp(f.stat().st_mtime).strftime("%d.%m.%Y %H:%M")
            row = ctk.CTkFrame(scroll, corner_radius=6, cursor="hand2")
            row.pack(fill="x", pady=2)
            ctk.CTkLabel(row, text=f.name, font=self._body_font, anchor="w").pack(side="left", padx=8, pady=6)
            ctk.CTkLabel(row, text=mtime, font=self._small_font, text_color="gray55").pack(side="right", padx=8)
            for w in (row, *row.winfo_children()):
                w.bind("<Button-1>", lambda e, file=f, w=win: (self._set_target(file), w.destroy()))

    def _pick_dialog(self):
        path = filedialog.askopenfilename(
            title="Выберите файл Спецификации",
            filetypes=[("Excel files", "*.xlsx *.xls")],
            initialdir=str(Path("output")) if Path("output").exists() else ".",
        )
        if path:
            self._set_target(Path(path))

    def _set_target(self, f: Path):
        self._target_file = f
        self._target_lbl.configure(text=f"✅  {f.name}", text_color=("gray20", "gray85"))

    def _refresh_km_list(self):
        km_dir = Path("выгрузка честный знак")
        files = sorted(
            [f for f in km_dir.glob("*.xlsx") if not f.name.startswith("~")],
            key=lambda f: f.stat().st_mtime, reverse=True,
        ) if km_dir.exists() else []
        if files:
            self._km_files_map = {f.name: f for f in files}
            self._km_dropdown.configure(values=[f.name for f in files])
            self._km_dropdown.set(files[0].name)
            self._km_file = files[0]
        else:
            self._km_dropdown.configure(values=["— нет файлов —"])
            self._km_dropdown.set("— нет файлов —")
            self._km_files_map = {}
            self._km_file = None

    def _browse_km_file(self):
        from tkinter import filedialog
        path = filedialog.askopenfilename(
            title="Выберите файл маркировки (Честный знак)",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if not path:
            return
        f = Path(path)
        if not hasattr(self, "_km_files_map"):
            self._km_files_map = {}
        self._km_files_map[f.name] = f
        current_values = list(self._km_dropdown.cget("values") or [])
        if f.name not in current_values:
            current_values = [v for v in current_values if v != "— нет файлов —"]
            current_values.insert(0, f.name)
            self._km_dropdown.configure(values=current_values)
        self._km_dropdown.set(f.name)
        self._km_file = f

    def _on_km_selected(self, value: str):
        self._km_file = self._km_files_map.get(value)

    # ------------------------------------------------------------------
    def _run(self):
        if not self._target_file or not self._target_file.exists():
            self._log_append("❌  Выберите целевой файл.")
            return
        if not self._km_file or not self._km_file.exists():
            self._log_append("❌  Выберите файл маркировки.")
            return

        output_file = self._target_file if self._overwrite_var.get() else (
            self._target_file.parent / (self._target_file.stem + "_with_kiz" + self._target_file.suffix)
        )

        self._run_btn.configure(state="disabled")
        threading.Thread(
            target=self._do_inject,
            args=(self._target_file, self._km_file, output_file),
            daemon=True,
        ).start()
        self._poll()

    def _do_inject(self, target, km_file, output_file):
        from src.km_loader    import KMLoader
        from src.kiz_injector import KIZInjector
        def log(m): self._log_queue.put(("msg", m))
        try:
            log(f"📥  Загрузка КМ справочника из {km_file.name}...")
            loader = KMLoader(str(km_file))
            log(f"✅  {loader.total_codes} кодов, {len(loader.articles)} артикулов")
            log(f"✏️   Запись в {target.name}...")
            stats = KIZInjector(loader).inject(target, output_file)
            log(f"✅  Обновлено строк: {stats['rows_updated']}")
            log(f"   Добавлено кодов: {stats['codes_added']}")
            log(f"   Пропущено строк: {stats['rows_skipped']}")
            log(f"📎  Сохранено: {output_file.name}")
            self._log_queue.put(("done", None))
        except Exception as e:
            log(f"❌  Ошибка: {e}")
            self._log_queue.put(("fail", None))

    def _poll(self):
        try:
            while True:
                kind, _ = self._log_queue.get_nowait()
                if kind == "msg":
                    self._log_append(_)
                elif kind in ("done", "fail"):
                    self._run_btn.configure(state="normal")
                    return
        except queue.Empty:
            pass
        self.after(100, self._poll)

    def _log_append(self, text: str):
        self._log.configure(state="normal")
        self._log.insert("end", text + "\n")
        self._log.see("end")
        self._log.configure(state="disabled")

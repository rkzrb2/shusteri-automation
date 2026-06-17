"""
Вкладка «Пресеты» — просмотр и редактирование пресетов клиентов.
"""
from pathlib import Path
import yaml
import customtkinter as ctk


# Поля для редактирования: (путь в YAML через точки, метка)
_FIELDS = [
    ("preset_name",              "Название пресета"),
    ("seller.name",              "Продавец (рус.)"),
    ("seller.name_en",           "Продавец (англ.)"),
    ("seller.address",           "Адрес продавца (рус.)"),
    ("seller.address_en",        "Адрес продавца (англ.)"),
    ("buyer.name",               "Покупатель (рус.)"),
    ("buyer.address",            "Адрес покупателя"),
    ("contract.number",          "Номер договора"),
    ("contract.date",            "Дата договора"),
    ("delivery.terms",           "Условия поставки"),
    ("delivery.currency",        "Валюта"),
    ("delivery.country_of_origin",    "Страна происхождения (рус.)"),
    ("delivery.country_of_origin_en", "Страна происхождения (англ.)"),
]


def _get_nested(data: dict, path: str):
    keys = path.split(".")
    for k in keys:
        data = data.get(k, "") if isinstance(data, dict) else ""
    return data or ""


def _set_nested(data: dict, path: str, value: str):
    keys = path.split(".")
    for k in keys[:-1]:
        data = data.setdefault(k, {})
    data[keys[-1]] = value


class PresetsView(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent, corner_radius=0, fg_color="transparent")
        self._lbl_font   = ctk.CTkFont(size=11, weight="bold")
        self._body_font  = ctk.CTkFont(size=12)
        self._small_font = ctk.CTkFont(size=11)
        self._presets: list[dict] = []
        self._selected_idx: int | None = None
        self._entries: dict[str, ctk.CTkEntry] = {}
        self._build()
        self._load_presets()

    # ------------------------------------------------------------------
    def _build(self):
        # Заголовок
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=16, pady=(16, 8))
        ctk.CTkLabel(header, text="Управление пресетами", font=ctk.CTkFont(size=16, weight="bold")).pack(side="left")
        ctk.CTkButton(header, text="+ Новый пресет", width=130, font=self._body_font, command=self._new_preset).pack(side="right")

        # Основная область: список слева + редактор справа
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=16, pady=4)

        # Список пресетов
        list_frame = ctk.CTkFrame(main, width=200, corner_radius=10)
        list_frame.pack(side="left", fill="y", padx=(0, 12))
        list_frame.pack_propagate(False)

        ctk.CTkLabel(list_frame, text="Клиенты", font=self._lbl_font, text_color="gray55").pack(pady=(10, 4), padx=10, anchor="w")
        self._list_scroll = ctk.CTkScrollableFrame(list_frame, fg_color="transparent")
        self._list_scroll.pack(fill="both", expand=True, padx=6, pady=(0, 8))

        # Редактор
        edit_frame = ctk.CTkScrollableFrame(main, corner_radius=10)
        edit_frame.pack(side="left", fill="both", expand=True)

        self._edit_title = ctk.CTkLabel(edit_frame, text="Выберите пресет", font=ctk.CTkFont(size=14, weight="bold"))
        self._edit_title.pack(anchor="w", padx=14, pady=(12, 8))

        self._fields_frame = ctk.CTkFrame(edit_frame, fg_color="transparent")
        self._fields_frame.pack(fill="x", padx=10)
        self._fields_frame.columnconfigure(1, weight=1)

        for row_idx, (path, label) in enumerate(_FIELDS):
            ctk.CTkLabel(self._fields_frame, text=label + ":", font=self._small_font, anchor="e").grid(
                row=row_idx, column=0, sticky="e", padx=(0, 8), pady=4
            )
            entry = ctk.CTkEntry(self._fields_frame, font=self._body_font)
            entry.grid(row=row_idx, column=1, sticky="ew", pady=4)
            self._entries[path] = entry

        # Кнопки сохранить/удалить
        btn_row = ctk.CTkFrame(edit_frame, fg_color="transparent")
        btn_row.pack(anchor="e", padx=14, pady=12)
        ctk.CTkButton(btn_row, text="Удалить пресет", width=130, font=self._body_font,
                      fg_color="transparent", border_width=1,
                      hover_color=("red80", "red30"),
                      command=self._delete_preset).pack(side="left", padx=(0, 8))
        ctk.CTkButton(btn_row, text="💾  Сохранить", width=120, font=self._body_font,
                      command=self._save_preset).pack(side="left")

    # ------------------------------------------------------------------
    def _load_presets(self):
        for w in self._list_scroll.winfo_children():
            w.destroy()
        self._presets = []

        presets_dir = Path("presets")
        if not presets_dir.exists():
            return

        for f in sorted(presets_dir.glob("*.yaml")):
            try:
                data = yaml.safe_load(f.read_text(encoding="utf-8"))
                self._presets.append({"file": f, "data": data})
            except Exception:
                pass

        for idx, p in enumerate(self._presets):
            name = p["data"].get("preset_name", p["file"].stem)
            btn = ctk.CTkButton(
                self._list_scroll, text=name, anchor="w", height=36,
                corner_radius=8, font=self._body_font,
                fg_color="transparent", hover_color=("gray80", "gray25"),
                command=lambda i=idx: self._select(i),
            )
            btn.pack(fill="x", pady=2)

        if self._presets:
            self._select(0)

    def _select(self, idx: int):
        self._selected_idx = idx
        p = self._presets[idx]
        self._edit_title.configure(text=p["data"].get("preset_name", p["file"].stem))
        for path, _ in _FIELDS:
            val = _get_nested(p["data"], path)
            e = self._entries[path]
            e.delete(0, "end")
            e.insert(0, str(val))

    def _save_preset(self):
        if self._selected_idx is None:
            return
        p = self._presets[self._selected_idx]
        data = p["data"]
        for path, _ in _FIELDS:
            val = self._entries[path].get().strip()
            _set_nested(data, path, val)
        with open(p["file"], "w", encoding="utf-8") as f:
            yaml.dump(data, f, allow_unicode=True, sort_keys=False)
        self._load_presets()

    def _delete_preset(self):
        if self._selected_idx is None:
            return
        p = self._presets[self._selected_idx]
        win = ctk.CTkToplevel(self)
        win.title("Подтверждение")
        win.geometry("360x140")
        win.grab_set()
        ctk.CTkLabel(win, text=f"Удалить пресет «{p['data'].get('preset_name', '')}»?", font=self._body_font).pack(expand=True, pady=20)
        row = ctk.CTkFrame(win, fg_color="transparent")
        row.pack(pady=(0, 16))
        ctk.CTkButton(row, text="Отмена", width=90, fg_color="transparent", border_width=1, command=win.destroy).pack(side="left", padx=8)
        def do_delete():
            p["file"].unlink(missing_ok=True)
            win.destroy()
            self._selected_idx = None
            self._load_presets()
        ctk.CTkButton(row, text="Удалить", width=90, command=do_delete).pack(side="left")

    def _new_preset(self):
        template = {
            "preset_name": "Новый клиент",
            "description": "",
            "seller": {"name": "", "name_en": "", "address": "", "address_en": ""},
            "buyer":  {"name": "", "address": "", "address_en": ""},
            "contract": {"number": "", "date": ""},
            "delivery": {"terms": "", "currency": "CNY", "country_of_origin": "Китай", "country_of_origin_en": "China"},
        }
        idx = len(self._presets) + 1
        new_file = Path("presets") / f"new_preset_{idx}.yaml"
        new_file.parent.mkdir(exist_ok=True)
        with open(new_file, "w", encoding="utf-8") as f:
            yaml.dump(template, f, allow_unicode=True, sort_keys=False)
        self._load_presets()
        self._select(len(self._presets) - 1)

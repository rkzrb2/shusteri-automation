"""
Главное окно приложения Shusteri Automation.
Сайдбар + переключение вкладок.
"""
import customtkinter as ctk

from ui.views.generation import GenerationView
from ui.views.presets    import PresetsView
from ui.views.history    import HistoryView
from ui.views.log_view   import LogView

# Тема по умолчанию
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

_NAV_ITEMS = [
    ("🔄  Генерация",  "generation"),
    ("👤  Пресеты",    "presets"),
    ("📋  История",    "history"),
    ("📝  Лог",        "log"),
]

_SIDEBAR_W = 190


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Shusteri Automation")
        self.geometry("980x720")
        self.minsize(820, 600)

        self._active_view = None
        self._nav_buttons: dict[str, ctk.CTkButton] = {}

        self._build_sidebar()
        self._build_content()
        self._show_view("generation")

    # ------------------------------------------------------------------
    # Sidebar
    # ------------------------------------------------------------------
    def _build_sidebar(self):
        sidebar = ctk.CTkFrame(self, width=_SIDEBAR_W, corner_radius=0)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        # Логотип
        ctk.CTkLabel(
            sidebar,
            text="Shusteri\nAutomation",
            font=ctk.CTkFont(size=16, weight="bold"),
            justify="center",
        ).pack(pady=(24, 6))

        ctk.CTkLabel(sidebar, text="v2.0", font=ctk.CTkFont(size=11), text_color="gray").pack(pady=(0, 20))

        ctk.CTkFrame(sidebar, height=1, fg_color="gray30").pack(fill="x", padx=16, pady=(0, 16))

        # Навигационные кнопки
        for label, key in _NAV_ITEMS:
            btn = ctk.CTkButton(
                sidebar,
                text=label,
                anchor="w",
                height=40,
                corner_radius=8,
                fg_color="transparent",
                hover_color=("gray80", "gray25"),
                font=ctk.CTkFont(size=13),
                command=lambda k=key: self._show_view(k),
            )
            btn.pack(fill="x", padx=12, pady=3)
            self._nav_buttons[key] = btn

        # Переключатель темы внизу
        sidebar.pack_propagate(False)
        bottom = ctk.CTkFrame(sidebar, fg_color="transparent")
        bottom.pack(side="bottom", fill="x", padx=12, pady=16)

        ctk.CTkLabel(bottom, text="Тема", font=ctk.CTkFont(size=11), text_color="gray").pack()
        ctk.CTkSegmentedButton(
            bottom,
            values=["🌙 Dark", "☀️ Light"],
            command=self._toggle_theme,
            font=ctk.CTkFont(size=11),
        ).pack(fill="x", pady=(4, 0))

    def _toggle_theme(self, value: str):
        mode = "dark" if "Dark" in value else "light"
        ctk.set_appearance_mode(mode)

    # ------------------------------------------------------------------
    # Content area
    # ------------------------------------------------------------------
    def _build_content(self):
        self._content = ctk.CTkFrame(self, corner_radius=0, fg_color=("gray92", "gray14"))
        self._content.pack(side="left", fill="both", expand=True)

        # Создаём все вьюхи заранее (hidden)
        self._views = {
            "generation": GenerationView(self._content),
            "presets":    PresetsView(self._content),
            "history":    HistoryView(self._content),
            "log":        LogView(self._content),
        }

    def _show_view(self, key: str):
        # Скрываем текущую
        if self._active_view:
            self._views[self._active_view].pack_forget()

        # Подсвечиваем активную кнопку
        for k, btn in self._nav_buttons.items():
            if k == key:
                btn.configure(fg_color=("gray75", "gray30"))
            else:
                btn.configure(fg_color="transparent")

        # Показываем новую
        self._views[key].pack(fill="both", expand=True, padx=0, pady=0)
        self._active_view = key

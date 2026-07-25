from __future__ import annotations

import subprocess
import sys
import threading
import webbrowser
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List

import tkinter as tk
from tkinter import messagebox

import customtkinter as ctk
from PIL import Image

try:
    from CTkMessagebox import CTkMessagebox
except ImportError:  # Fallback, falls Paket nicht vorhanden
    CTkMessagebox = None


# ---------------------------------------------------------------------------
# Helper für Messagebox
# ---------------------------------------------------------------------------

def show_message(title: str, message: str, icon: str = "info"):
    if CTkMessagebox is not None:
        CTkMessagebox(title=title, message=message, icon=icon)
    else:
        if icon == "cancel":
            messagebox.showerror(title, message)
        elif icon == "check":
            messagebox.showinfo(title, message)
        else:
            messagebox.showinfo(title, message)


# ---------------------------------------------------------------------------
# Pfade & Branding
# ---------------------------------------------------------------------------

if getattr(sys, "frozen", False):
    APP_DIR = Path(sys.executable).resolve().parent
    BUNDLE_DIR = Path(sys._MEIPASS)
else:
    APP_DIR = Path(__file__).resolve().parent
    BUNDLE_DIR = APP_DIR

# PS-Script liegt im Bundle (mit --add-data eingebunden)
OFFICE_PS = BUNDLE_DIR / "office_installer.ps1"

# Logo kann ebenfalls aus dem Bundle geladen werden
LOGO_PATH = BUNDLE_DIR / "logo.png"
ICON_PATH = BUNDLE_DIR / "office.ico"

APP_NAME = "Office Installer – Retail"
BRAND_TEXT = "© 2026 SD-ITLab – MIT licensed"
BRAND_URL = "https://sd-itlab.de"
LOGO_URL = "https://sd-itlab.de"
README_URL = BRAND_URL  # ggf. auf eigene Doku anpassen

ACCENT = "#3B82F6"


# ---------------------------------------------------------------------------
# Sprach-Optionen (Anzeigename -> Code)
# ---------------------------------------------------------------------------

LANG_OPTIONS = [
    ("Deutsch (de-de)", "de-de"),
    ("English (en-us)", "en-us"),
    ("Polnisch (pl-pl)", "pl-pl"),
    ("Italienisch (it-it)", "it-it"),
    ("Französisch (fr-fr)", "fr-fr"),
    ("Kroatisch (hr-hr)", "hr-hr"),
    ("Türkisch (tr-tr)", "tr-tr"),
    ("Russisch (ru-ru)", "ru-ru"),
]

LANG_DISPLAY_VALUES = [label for label, _ in LANG_OPTIONS]
LANG_MAP = {label: code for label, code in LANG_OPTIONS}


# ---------------------------------------------------------------------------
# Office-Presets (Retail only)
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class OfficePreset:
    key: str
    label: str
    product_id: str
    channel: str
    shortcut_apps: List[str]
    exclude_apps: List[str]


OFFICE_PRODUCTS: Dict[str, OfficePreset] = {
    # --- Microsoft 365 (eigene Gruppe) -------------------------------------
    "m365_family": OfficePreset(
        key="m365_family",
        label="Microsoft 365 Family",
        product_id="O365HomePremRetail",
        channel="Current",
        shortcut_apps=["Word", "Excel", "PowerPoint", "Outlook"],
        exclude_apps=["Access", "Publisher", "Teams"],
    ),
    "m365_single": OfficePreset(
        key="m365_single",
        label="Microsoft 365 Single",
        product_id="O365HomePremRetail",
        channel="Current",
        shortcut_apps=["Word", "Excel", "PowerPoint", "Outlook"],
        exclude_apps=["Access", "Publisher", "Teams"],
    ),

    # --- 2024 ---------------------------------------------------------------
    "2024_home_student": OfficePreset(  # Key bleibt, Label geändert
        key="2024_home_student",
        label="Office 2024 Home",
        product_id="Home2024Retail",
        channel="Current",
        shortcut_apps=["Word", "Excel", "PowerPoint"],
        exclude_apps=["Outlook", "Access", "Publisher", "Teams"],
    ),
    "2024_business": OfficePreset(
        key="2024_business",
        label="Office 2024 Home & Business",
        product_id="HomeBusiness2024Retail",
        channel="Current",
        shortcut_apps=["Word", "Excel", "PowerPoint", "Outlook"],
        exclude_apps=["Access", "Publisher", "Teams"],
    ),
    "2024_professional": OfficePreset(
        key="2024_professional",
        label="Office 2024 Professional Plus",
        product_id="ProPlus2024Retail",
        channel="Current",
        shortcut_apps=[
            "Word",
            "Excel",
            "PowerPoint",
            "Outlook",
            "Access",
            "Publisher",
        ],
        exclude_apps=["Teams"],
    ),

    # --- 2021 ---------------------------------------------------------------
    "2021_home_student": OfficePreset(
        key="2021_home_student",
        label="Office 2021 Home & Student",
        product_id="HomeStudent2021Retail",
        channel="Current",
        shortcut_apps=["Word", "Excel", "PowerPoint"],
        exclude_apps=["Outlook", "Access", "Publisher", "Teams"],
    ),
    "2021_business": OfficePreset(
        key="2021_business",
        label="Office 2021 Home & Business",
        product_id="HomeBusiness2021Retail",
        channel="Current",
        shortcut_apps=["Word", "Excel", "PowerPoint", "Outlook"],
        exclude_apps=["Access", "Publisher", "Teams"],
    ),
    "2021_professional": OfficePreset(
        key="2021_professional",
        label="Office 2021 Professional Plus",
        product_id="ProPlus2021Retail",
        channel="Current",
        shortcut_apps=[
            "Word",
            "Excel",
            "PowerPoint",
            "Outlook",
            "Access",
            "Publisher",
        ],
        exclude_apps=["Teams"],
    ),

    # --- 2019 ---------------------------------------------------------------
    "2019_home_student": OfficePreset(
        key="2019_home_student",
        label="Office 2019 Home & Student",
        product_id="HomeStudentRetail",
        channel="Current",
        shortcut_apps=["Word", "Excel", "PowerPoint"],
        exclude_apps=["Outlook", "Access", "Publisher", "Teams"],
    ),
    "2019_business": OfficePreset(
        key="2019_business",
        label="Office 2019 Home & Business",
        product_id="HomeBusinessRetail",
        channel="Current",
        shortcut_apps=["Word", "Excel", "PowerPoint", "Outlook"],
        exclude_apps=["Access", "Publisher", "Teams"],
    ),
    "2019_professional": OfficePreset(
        key="2019_professional",
        label="Office 2019 Professional Plus",
        product_id="ProPlusRetail",
        channel="Current",
        shortcut_apps=[
            "Word",
            "Excel",
            "PowerPoint",
            "Outlook",
            "Access",
            "Publisher",
        ],
        exclude_apps=["Teams"],
    ),

    # --- 2016 ---------------------------------------------------------------
    "2016_home_student": OfficePreset(
        key="2016_home_student",
        label="Office 2016 Home & Student",
        product_id="HomeStudentRetail",
        channel="Current",
        shortcut_apps=["Word", "Excel", "PowerPoint"],
        exclude_apps=["Outlook", "Access", "Publisher", "Teams"],
    ),
    "2016_business": OfficePreset(
        key="2016_business",
        label="Office 2016 Home & Business",
        product_id="HomeBusinessRetail",
        channel="Current",
        shortcut_apps=["Word", "Excel", "PowerPoint", "Outlook"],
        exclude_apps=["Access", "Publisher", "Teams"],
    ),
    "2016_professional": OfficePreset(
        key="2016_professional",
        label="Office 2016 Professional",
        product_id="ProfessionalRetail",
        channel="Current",
        shortcut_apps=[
            "Word",
            "Excel",
            "PowerPoint",
            "Outlook",
            "Access",
            "Publisher",
        ],
        exclude_apps=["Teams"],
    ),
}

# Seitenleiste: (Label, interner Key)
YEARS = [
    ("Microsoft 365", "m365"),
    ("2024", "2024"),
    ("2021", "2021"),
    ("2019", "2019"),
    ("2016", "2016"),
]


# ---------------------------------------------------------------------------
# GUI-Komponenten
# ---------------------------------------------------------------------------

class EditionCard(ctk.CTkFrame):
    def __init__(
        self,
        master,
        title: str,
        description: str,
        variable: ctk.StringVar,
        value: str,
        on_select=None,
        **kwargs,
    ):
        super().__init__(master, **kwargs)
        self.variable = variable
        self.value = value
        self.on_select = on_select
        self._trace_id = None  # wird gleich gesetzt
        self.configure(corner_radius=12, fg_color="#F9FAFB")

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.title_label = ctk.CTkLabel(
            self,
            text=title,
            font=ctk.CTkFont(size=14, weight="bold"),
            anchor="w",
        )
        self.title_label.grid(row=0, column=0, sticky="ew", padx=16, pady=(10, 2))

        self.desc_label = ctk.CTkLabel(
            self,
            text=description,
            font=ctk.CTkFont(size=11),
            text_color="#6B7280",
            anchor="w",
            justify="left",
            wraplength=420,
        )
        self.desc_label.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 10))

        # Klickbar machen
        for w in (self, self.title_label, self.desc_label):
            w.bind("<Button-1>", self._on_click)

        # Style initial
        self._update_style(self.variable.get() == self.value)

        # Trace registrieren und ID merken, damit wir ihn später entfernen können
        self._trace_id = self.variable.trace_add("write", self._on_variable_changed)

    def _on_click(self, _event=None):
        self.variable.set(self.value)
        if callable(self.on_select):
            self.on_select()

    def _on_variable_changed(self, *args):
        """Wird von trace_add aufgerufen, wenn sich der StringVar-Wert ändert."""
        # Wenn das Widget bereits zerstört ist: nichts tun
        if not self.winfo_exists():
            return
        self._update_style(self.variable.get() == self.value)

    def _update_style(self, selected: bool):
        """Optik je nach Auswahlzustand anpassen."""
        try:
            if selected:
                self.configure(fg_color="#DBEAFE")
                self.title_label.configure(text_color="#1D4ED8")
            else:
                self.configure(fg_color="#F9FAFB")
                self.title_label.configure(text_color="#111827")
        except tk.TclError:
            # Falls Tk zwischenzeitlich schon zerstört wurde, Fehler ignorieren
            pass

    def destroy(self):
        """Beim Zerstören den trace wieder deregistrieren, damit keine Ghost-Callbacks bleiben."""
        try:
            if self._trace_id is not None:
                self.variable.trace_remove("write", self._trace_id)
        except Exception:
            pass
        super().destroy()

# ---------------------------------------------------------------------------
# Hauptfenster
# ---------------------------------------------------------------------------

class OfficeInstallerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.title(APP_NAME)

        # Fenster-Icon setzen
        if ICON_PATH.exists():
            try:
                self.iconbitmap(str(ICON_PATH))
            except Exception:
                pass

        # Feste Größe, mittig
        width, height = 1000, 540
        self.geometry(f"{width}x{height}")
        self.resizable(False, False)
        self.update_idletasks()
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        x = (screen_w // 2) - (width // 2)
        y = (screen_h // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")

        self._installing = False
        self._logo_img = None

        # State
        # Default: 2024 Home
        self.selected_year = ctk.StringVar(value="2024")
        self.selected_preset = ctk.StringVar(value="2024_home_student")

        # Sprache als Anzeige- und Code-Var getrennt
        self.selected_lang_display = ctk.StringVar(value=LANG_DISPLAY_VALUES[0])
        self.selected_lang_code = ctk.StringVar(value=LANG_OPTIONS[0][1])

        self.selected_arch = ctk.StringVar(value="64")

        # Grid: Sidebar | Main | RightPanel  +  Bottombar
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0)
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(2, weight=0)

        self._build_sidebar()
        self._build_main()
        self._build_right_panel()
        self._build_bottom()

        self._update_edition_cards()
        self.bind_all("<Up>", self._on_key_up)
        self.bind_all("<Down>", self._on_key_down)
        self.bind_all("<Left>", self._on_key_left)
        self.bind_all("<Right>", self._on_key_right)
        self.bind_all("<Return>", self._on_key_activate)
        self.bind_all("<space>", self._on_key_activate)

            # ------------------------------------------------------------------ Tastatur-Navigation

    def _set_cards_focus(self, state: bool):
        """Aktiviert/deaktiviert den 'Editionsbereich ist im Fokus'-Modus."""
        self._cards_focus = bool(state)

    def _get_current_card_index(self) -> int:
        """Index der aktuell gewählten EditionCard anhand selected_preset."""
        current_key = self.selected_preset.get()
        for idx, card in enumerate(self.edition_cards):
            if card.value == current_key:
                return idx
        return 0

    def _select_card_by_index(self, idx: int):
        """Hilfsfunktion: Edition per Index auswählen."""
        if not self.edition_cards:
            return
        idx = idx % len(self.edition_cards)
        self.selected_preset.set(self.edition_cards[idx].value)

    def _change_year_by_offset(self, delta: int):
        """Wechselt das Jahr in der linken Leiste relativ (nur im Cards-Fokus)."""
        year_keys = [key for (_label, key) in YEARS]
        current = self.selected_year.get()
        if current not in year_keys:
            return

        idx = year_keys.index(current)
        idx = (idx + delta) % len(year_keys)
        new_year = year_keys[idx]
        self._on_year_selected(new_year)

    def _on_key_up(self, event):
        # Nur reagieren, wenn wir im Editions-Fokus sind und keine Installation läuft
        if not self._cards_focus or self._installing:
            return  # Standard-Verhalten beibehalten

        idx = self._get_current_card_index()
        self._select_card_by_index(idx - 1)
        return "break"

    def _on_key_down(self, event):
        if not self._cards_focus or self._installing:
            return

        idx = self._get_current_card_index()
        self._select_card_by_index(idx + 1)
        return "break"

    def _on_key_left(self, event):
        if not self._cards_focus or self._installing:
            return

        self._change_year_by_offset(-1)
        return "break"

    def _on_key_right(self, event):
        if not self._cards_focus or self._installing:
            return

        self._change_year_by_offset(+1)
        return "break"

    def _on_key_activate(self, event):
        """
        Enter/Space:
        - Wenn wir im Editions-Fokus sind -> Installation starten.
        - Sonst: nichts tun (Buttons/Combos verhalten sich normal).
        """
        if not self._cards_focus or self._installing:
            return  # Standard-Verhalten für Buttons/Combos beibehalten

        self._on_install_clicked()
        return "break"
    
    def _cycle_combobox(self, combobox: ctk.CTkComboBox, direction: int):
        """Wechselt den Wert der Combobox per Pfeiltaste."""
        values = list(combobox.cget("values"))
        if not values:
            return

        current = combobox.get()
        try:
            idx = values.index(current)
        except ValueError:
            idx = 0

        idx = (idx + direction) % len(values)
        combobox.set(values[idx])

    # ------------------------------------------------------------------ UI --

    def _build_sidebar(self):
        sidebar = ctk.CTkFrame(self, fg_color="#F3F4F6")
        sidebar.grid(row=0, column=0, sticky="nsw")
        sidebar.grid_columnconfigure(0, weight=1)

        header = ctk.CTkLabel(
            sidebar,
            text="Office-Version",
            font=ctk.CTkFont(size=14, weight="bold"),
        )
        header.grid(row=0, column=0, padx=12, pady=(10, 6), sticky="w")

        self.year_buttons: Dict[str, ctk.CTkButton] = {}
        for idx, (label, key) in enumerate(YEARS, start=1):
            btn = ctk.CTkButton(
                sidebar,
                text=label,
                command=lambda k=key: self._on_year_selected(k),
                fg_color="#E5E7EB",
                text_color="#111827",
                hover_color="#D1D5DB",
                corner_radius=8,
                height=32,
            )
            btn.grid(row=idx, column=0, padx=12, pady=4, sticky="ew")
            self.year_buttons[key] = btn

        sidebar.grid_rowconfigure(len(YEARS) + 1, weight=1)
        self._update_year_buttons()

    def _build_main(self):
        main = ctk.CTkFrame(self, fg_color="white")
        main.grid(row=0, column=1, sticky="nsew", padx=(12, 8), pady=(12, 0))
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(0, weight=0)
        main.grid_rowconfigure(1, weight=0)
        main.grid_rowconfigure(2, weight=1)

        title = ctk.CTkLabel(
            main,
            text="Office Installer",
            font=ctk.CTkFont(size=20, weight="bold"),
            anchor="w",
        )
        title.grid(row=0, column=0, sticky="ew", padx=8, pady=(6, 2))

        subtitle = ctk.CTkLabel(
            main,
            text=(
                "Wähle Office-Version und Edition. Die Installation läuft "
                "unbeaufsichtigt im Hintergrund."
            ),
            font=ctk.CTkFont(size=11),
            text_color="#6B7280",
            anchor="w",
        )
        subtitle.grid(row=1, column=0, sticky="ew", padx=8, pady=(0, 10))

        cards_frame = ctk.CTkFrame(main, fg_color="white")
        cards_frame.grid(row=2, column=0, sticky="nsew", padx=4, pady=(0, 8))
        cards_frame.grid_columnconfigure(0, weight=1)

        self.cards_frame = cards_frame
        self.edition_cards: List[EditionCard] = []

        # Fokus-Anker für Keyboard-Navigation der Editionen
        self.cards_focus = tk.Frame(cards_frame, takefocus=1, width=1, height=1)
        self.cards_focus.grid(row=0, column=0, sticky="w")

        # Wenn dieser Frame den Fokus bekommt / verliert, schalten wir den "Card-Fokus"-Modus
        self._cards_focus = False
        self.cards_focus.bind("<FocusIn>", lambda e: self._set_cards_focus(True))
        self.cards_focus.bind("<FocusOut>", lambda e: self._set_cards_focus(False))

    def _build_right_panel(self):
        right = ctk.CTkFrame(self, fg_color="#F3F4F6")
        right.grid(row=0, column=2, sticky="nse", padx=(8, 12), pady=(12, 0))
        right.grid_columnconfigure(0, weight=1)

        # Logo-Box
        logo_box = ctk.CTkFrame(
            right,
            fg_color="#EFF4FF",
            corner_radius=18,
            width=240,
            height=140,
            border_width=1,
            border_color="#E5E7EB",
        )
        logo_box.grid(row=0, column=0, sticky="n", padx=4, pady=(4, 6))
        logo_box.grid_propagate(False)

        self.logo_label = ctk.CTkLabel(logo_box, text="")
        self.logo_label.place(relx=0.5, rely=0.5, anchor="center")
        self.logo_label.configure(cursor="hand2")
        self.logo_label.bind("<Button-1>", lambda e: self._open_url(LOGO_URL))

        self._load_logo()

        # Installationsoptionen (Architektur / Language)
        opt_title = ctk.CTkLabel(
            right,
            text="Installationsoptionen",
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        opt_title.grid(row=1, column=0, sticky="w", padx=6, pady=(8, 2))

        opt_frame = ctk.CTkFrame(right, fg_color="transparent")
        opt_frame.grid(row=2, column=0, sticky="w", padx=6, pady=(0, 6))

        # Architektur zuerst
        arch_label = ctk.CTkLabel(
            opt_frame,
            text="Architektur:",
            font=ctk.CTkFont(size=11),
            anchor="w",
        )
        arch_label.grid(row=0, column=0, sticky="w", pady=(0, 2))

        arch_box = ctk.CTkComboBox(
            opt_frame,
            values=["64", "32", "ARM64"],
            variable=self.selected_arch,
            width=100,
            state="readonly",
        )
        arch_box.grid(row=0, column=1, sticky="w", padx=(4, 0))
                # ↑ / ↓ für Architektur
        arch_box.bind(
            "<Up>",
            lambda e, cb=arch_box: (
                self._cycle_combobox(cb, -1),
                "break",
            )[-1],
        )
        arch_box.bind(
            "<Down>",
            lambda e, cb=arch_box: (
                self._cycle_combobox(cb, +1),
                "break",
            )[-1],
        )

        # Language danach (mit ausgeschriebenem Namen)
        lang_label = ctk.CTkLabel(
            opt_frame,
            text="Language:",
            font=ctk.CTkFont(size=11),
            anchor="w",
        )
        lang_label.grid(row=1, column=0, sticky="w", pady=(6, 2))

        lang_box = ctk.CTkComboBox(
            opt_frame,
            values=LANG_DISPLAY_VALUES,
            variable=self.selected_lang_display,
            width=160,
            command=self._on_language_changed,
            state="readonly",
        )
        lang_box.grid(row=1, column=1, sticky="w", padx=(4, 0))
        lang_box.bind(
            "<Up>",
            lambda e, cb=lang_box: (
                self._cycle_combobox(cb, -1),
                "break",
            )[-1],
        )
        lang_box.bind(
            "<Down>",
            lambda e, cb=lang_box: (
                self._cycle_combobox(cb, +1),
                "break",
            )[-1],
        )

    def _build_bottom(self):
        # Untere Leiste
        bottom = ctk.CTkFrame(self, fg_color="#F9FAFB")
        bottom.grid(
            row=1,
            column=0,
            columnspan=3,
            sticky="ew",
            padx=2,
            pady=(0, 0),
        )

        # Fixe Mindesthöhe, damit es nicht so gequetscht aussieht
        bottom.configure(height=72)
        bottom.grid_propagate(False)  # Frame behält die gesetzte Höhe

        # Spalten: 0 = Progress/Status (streckt), 1 = Brand, 2–4 = Buttons
        bottom.grid_columnconfigure(0, weight=1)
        for col in (1, 2, 3, 4):
            bottom.grid_columnconfigure(col, weight=0)
        bottom.grid_rowconfigure(0, weight=0)
        bottom.grid_rowconfigure(1, weight=0)

        # Progressbar oben in der Leiste, über komplette Breite
        self.progress = ctk.CTkProgressBar(
            bottom,
            progress_color=ACCENT,
            fg_color="#E5E7EB",
            height=10,
            corner_radius=6,
        )
        self.progress.set(0.0)
        self.progress.grid(
            row=0,
            column=0,
            columnspan=5,
            sticky="ew",
            padx=(4, 4),       # minimal Luft damit der Punkt nicht am Rand klebt
            pady=(10, 4),
        )

        # Status links
        self.status_label = ctk.CTkLabel(
            bottom,
            text="Bereit.",
            text_color="#6B7280",
        )
        self.status_label.grid(
            row=1,
            column=0,
            sticky="w",
            padx=(4, 0),
            pady=(4, 10),
        )

        # Branding mittig
        self.footer_brand = ctk.CTkLabel(
            bottom,
            text=BRAND_TEXT,
            font=ctk.CTkFont(size=10),
            text_color="#6B7280",
            cursor="hand2",
        )
        self.footer_brand.grid(
            row=1,
            column=1,
            sticky="e",
            padx=(0, 12),
            pady=(4, 10),
        )
        self.footer_brand.bind("<Button-1>", lambda e: self._open_url(BRAND_URL))
        self.footer_brand.bind(
            "<Enter>",
            lambda e: self.footer_brand.configure(text_color=ACCENT),
        )
        self.footer_brand.bind(
            "<Leave>",
            lambda e: self.footer_brand.configure(text_color="#6B7280"),
        )

        # Buttons rechts
        self.btn_readme = ctk.CTkButton(
            bottom,
            text="Readme",
            width=100,
            command=lambda: self._open_url(README_URL),
        )
        self.btn_readme.grid(row=1, column=2, padx=(0, 6), pady=(4, 10))

        self.btn_install = ctk.CTkButton(
            bottom,
            text="Installieren",
            width=110,
            command=self._on_install_clicked,
        )
        self.btn_install.grid(row=1, column=3, padx=6, pady=(4, 10))

        self.btn_cancel = ctk.CTkButton(
            bottom,
            text="Abbrechen",
            width=110,
            fg_color="#E5E7EB",
            text_color="#111827",
            hover_color="#D1D5DB",
            command=self._on_cancel_clicked,
        )
        self.btn_cancel.grid(row=1, column=4, padx=(0, 4), pady=(4, 10))

    # ------------------------------------------------------------------ Helper

    def _open_url(self, url: str):
        try:
            webbrowser.open(url, new=2)
        except Exception:
            pass

    def _load_logo(self):
        """Versucht logo.png zu laden, sonst Text-Fallback."""
        if LOGO_PATH.exists():
            try:
                img = Image.open(LOGO_PATH)
                max_width, max_height = 220, 90
                img_ratio = img.width / img.height
                box_ratio = max_width / max_height

                if img_ratio > box_ratio:
                    new_w = max_width
                    new_h = int(max_width / img_ratio)
                else:
                    new_h = max_height
                    new_w = int(max_height * img_ratio)

                img = img.resize((new_w, new_h), Image.LANCZOS)
                self._logo_img = ctk.CTkImage(
                    light_image=img,
                    dark_image=img,
                    size=(new_w, new_h),
                )
                self.logo_label.configure(image=self._logo_img, text="")
                return
            except Exception:
                pass

        self.logo_label.configure(
            text="SD-ITLab",
            justify="center",
            text_color="#6B7280",
        )

    def _on_language_changed(self, choice: str):
        """Callback, wenn im Language-Combo etwas gewählt wird."""
        code = LANG_MAP.get(choice, "de-de")
        self.selected_lang_code.set(code)

    # ------------------------------------------------------------------ State/Updates

    def _update_year_buttons(self):
        for key, btn in self.year_buttons.items():
            if self.selected_year.get() == key:
                btn.configure(
                    fg_color="#2563EB",
                    text_color="white",
                    hover_color="#1D4ED8",
                )
            else:
                btn.configure(
                    fg_color="#E5E7EB",
                    text_color="#111827",
                    hover_color="#D1D5DB",
                )

    def _on_year_selected(self, year_key: str):
        if self._installing:
            return
        self.selected_year.set(year_key)

        if year_key == "m365":
            # Default dort: Single
            self.selected_preset.set("m365_single")
        else:
            self.selected_preset.set(f"{year_key}_home_student")

        self._update_year_buttons()
        self._update_edition_cards()

    def _update_edition_cards(self):
        for card in self.edition_cards:
            card.destroy()
        self.edition_cards.clear()

        year = self.selected_year.get()

        if year == "m365":
            # Reihenfolge: Single, dann Family
            presets_for_year = ["m365_single", "m365_family"]
        else:
            presets_for_year = [
                f"{year}_home_student",
                f"{year}_business",
                f"{year}_professional",
            ]

        descriptions = {
            "home_student": "Enthält Word, Excel, PowerPoint – ideal für Privat / Schule.",
            "business": "Home & Business: zusätzlich Outlook – ideal für Office-Arbeitsplätze.",
            "professional": "Professional Plus: inkl. Access & Publisher – für erweiterte Anforderungen.",
            "single": "Abonnement: Microsoft 365 Single – 1 Benutzer mit je 5 Geräten, persönliche Lizenz.",
            "family": "Abonnement: Microsoft 365 Family – bis zu 6 Benutzer mit je 5 Geräten.",
        }

        row = 1
        for key in presets_for_year:
            preset = OFFICE_PRODUCTS[key]
            suffix = key.split("_", 1)[1]
            desc_text = descriptions.get(suffix, "")
            card = EditionCard(
                self.cards_frame,
                title=preset.label,
                description=desc_text,
                variable=self.selected_preset,
                value=key,
            )
            card.grid(row=row, column=0, padx=4, pady=4, sticky="ew")
            self.edition_cards.append(card)
            row += 1

    # ------------------------------------------------------------------ Installation

    def _on_cancel_clicked(self):
        if self._installing:
            show_message(
                "Hinweis",
                "Eine Installation läuft bereits. Bitte warte, bis sie abgeschlossen ist.",
                icon="info",
            )
            return
        self.destroy()

    def _on_install_clicked(self):
        if self._installing:
            return

        preset_key = self.selected_preset.get()
        if preset_key not in OFFICE_PRODUCTS:
            show_message("Fehler", "Bitte eine Office-Version auswählen.", icon="cancel")
            return

        if not OFFICE_PS.exists():
            show_message(
                "Fehlende Datei",
                f"PowerShell-Script nicht gefunden:\n{OFFICE_PS}",
                icon="cancel",
            )
            return

        self._installing = True
        self.btn_install.configure(state="disabled")
        self.btn_cancel.configure(state="disabled")
        self.status_label.configure(text="Starte Office-Installation …")
        self.progress.start()

        thread = threading.Thread(target=self._run_install_thread, daemon=True)
        thread.start()

    def _run_install_thread(self):
        preset = OFFICE_PRODUCTS[self.selected_preset.get()]
        lang = self.selected_lang_code.get()
        arch = self.selected_arch.get()

        cmd = [
            "powershell.exe",
            "-NoLogo",
            "-NoProfile",
            "-WindowStyle", "Hidden",
            "-ExecutionPolicy", "Bypass",
            "-File", str(OFFICE_PS),
            "-ProductID", preset.product_id,
            "-Language", lang,
            "-Arch", arch,
            "-Channel", preset.channel,
        ]

        if preset.exclude_apps:
            cmd += ["-ExcludeApps", ",".join(preset.exclude_apps)]

        if preset.shortcut_apps:
            cmd += ["-ShortcutApps", ",".join(preset.shortcut_apps)]

        kwargs = dict(
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )

        if sys.platform == "win32":
            kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW

        try:
            result = subprocess.run(cmd, **kwargs)
        except Exception as exc:
            self.after(
                0,
                self._install_finished_error,
                f"Fehler beim Starten von PowerShell:\n{exc}",
            )
            return

        rc = result.returncode

        if rc == 0:
            self.after(0, self._install_finished_ok)
        elif rc == 2001:
            # eigener Code aus der PS1: Office bereits installiert
            self.after(
                0,
                self._install_finished_error,
                "Auf diesem System ist bereits eine Office-Version installiert.\n"
                "Bitte deinstalliere diese zuerst, bevor du eine neue Edition installierst.",
            )
        else:
            msg = (result.stderr or result.stdout or "").strip()
            if msg:
                msg = "\n".join(msg.splitlines()[:10])
            else:
                msg = f"Office-Installation meldete ExitCode {rc}."
            self.after(0, self._install_finished_error, msg)

    def _install_finished_ok(self):
        self._installing = False
        self.progress.stop()
        self.progress.set(0)
        self.btn_install.configure(state="normal")
        self.btn_cancel.configure(state="normal")
        self.status_label.configure(text="Fertig ✅ Office wurde installiert.")
        show_message("Fertig", "Office wurde erfolgreich installiert.", icon="check")

    def _install_finished_error(self, msg: str):
        self._installing = False
        self.progress.stop()
        self.progress.set(0)
        self.btn_install.configure(state="normal")
        self.btn_cancel.configure(state="normal")
        self.status_label.configure(text="Fehler bei der Installation.")
        show_message("Fehler", msg, icon="cancel")


# ---------------------------------------------------------------------------

if __name__ == "__main__":
    app = OfficeInstallerApp()
    app.mainloop()

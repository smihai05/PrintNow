"""
PrintNow - Minimal batch printing tool for Windows.

Requirements:
    pip install customtkinter pywin32 pypdf pyinstaller

Also needed: SumatraPDF.exe (portable) placed next to this file.
Download: https://www.sumatrapdfreader.org/download-free-pdf-viewer
"""

import ctypes
import json
import os
import sys
import subprocess
import threading
from ctypes import wintypes
from concurrent.futures import ThreadPoolExecutor

import customtkinter as ctk
from tkinter import filedialog, messagebox
import win32print

try:
    from pypdf import PdfReader
except ImportError:  # page counts will show "-" if pypdf is missing
    PdfReader = None

APP_NAME = "PrintNow"
SUPPORTED_EXT = (".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff")

# ---- Setting options ----
DUPLEX_OPTIONS = ["One-sided", "Two-sided (long edge)", "Two-sided (short edge)"]
COLOR_OPTIONS = ["Color", "Black & white"]
ORIENTATION_OPTIONS = ["Auto", "Portrait", "Landscape"]
PAPER_OPTIONS = ["Printer default", "A3", "A4", "A5", "Letter", "Legal"]
SCALE_OPTIONS = ["Shrink to fit", "Fit to page", "Actual size"]
NO_LAYOUT = "(no layout)"

DEFAULT_SETTINGS = {
    "copies": "1",
    "duplex": DUPLEX_OPTIONS[0],
    "color": COLOR_OPTIONS[0],
    "orientation": ORIENTATION_OPTIONS[0],
    "paper": PAPER_OPTIONS[0],
    "scale": SCALE_OPTIONS[0],
}

# ---- Theme colors (minimal / futuristic) ----
BG = "#0a0e14"
PANEL = "#111823"
PANEL_2 = "#0d131c"
BORDER = "#1c2a3a"
ACCENT = "#00e5ff"
ACCENT_HOVER = "#00b8cc"
TEXT = "#e6f1ff"
MUTED = "#6b7f95"
PAGES_COLOR = "#9fb3c8"
DANGER = "#ff4d6d"
SUCCESS = "#00e58f"


# ------------------------------------------------------------------ helpers
def resource_path(name):
    """Return path to a bundled resource (works for script and PyInstaller exe)."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


def find_sumatra():
    """Look for SumatraPDF.exe inside the bundle or next to the exe/script."""
    candidates = [
        resource_path("SumatraPDF.exe"),
        os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "SumatraPDF.exe"),
    ]
    for path in candidates:
        if os.path.isfile(path):
            return path
    return None


def list_printers():
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    try:
        return [p[2] for p in win32print.EnumPrinters(flags)]
    except Exception:
        return []


def default_printer():
    try:
        return win32print.GetDefaultPrinter()
    except Exception:
        return None


def count_pages(path):
    """Return (pages, error). pages is None if it could not be determined."""
    if not path.lower().endswith(".pdf"):
        return 1, None  # images
    if PdfReader is None:
        return None, "pypdf is not installed (run: py -m pip install pypdf)"
    try:
        reader = PdfReader(path)
        if reader.is_encrypted:
            reader.decrypt("")
        return len(reader.pages), None
    except Exception as e:
        return None, f"{os.path.basename(path)}: {e}"


def pages_text(n):
    if n is None:
        return "? pages"
    return "1 page" if n == 1 else f"{n} pages"


# ------------------------------------------------------------------- config
def config_path():
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    folder = os.path.join(base, APP_NAME)
    os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, "config.json")


def load_config():
    try:
        with open(config_path(), "r", encoding="utf-8") as f:
            data = json.load(f)
        layouts = data.get("layouts", {})
        default = data.get("default")
        if not isinstance(layouts, dict):
            layouts = {}
        if default not in layouts:
            default = None
        return {"layouts": layouts, "default": default}
    except Exception:
        return {"layouts": {}, "default": None}


def save_config(config):
    try:
        with open(config_path(), "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
        return True
    except Exception:
        return False


# ---------------------------------------------------------------------- app
class PrintNowApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("dark")

        self.title(APP_NAME)
        self.configure(fg_color=BG)
        self._setup_window()

        try:
            icon = resource_path("print.ico")
            self.after(200, lambda: self.iconbitmap(icon))
        except Exception:
            pass

        self.folder = ""
        self.file_rows = []  # list of (BooleanVar, path, status_label, pages_label)
        self.page_cache = {}  # path -> page count (or None)
        self.pool = ThreadPoolExecutor(max_workers=4)
        self.is_printing = False

        self.config_data = load_config()
        self.current_layout = None

        self._build_header()
        self._build_footer()  # packed at the bottom first, so it is always visible
        self._build_body()
        self._load_printers()
        self._load_default_layout()

        if PdfReader is None:
            self.after(600, self._warn_no_pypdf)

    def _setup_window(self):
        """Size the window to fit the visible work area (above the taskbar) and center it."""
        try:
            scaling = self._get_window_scaling()
        except Exception:
            scaling = 1.0

        left, top = 0, 0
        wa_w, wa_h = self.winfo_screenwidth(), self.winfo_screenheight()
        try:
            rect = wintypes.RECT()
            if ctypes.windll.user32.SystemParametersInfoW(48, 0, ctypes.byref(rect), 0):  # SPI_GETWORKAREA
                left, top = rect.left, rect.top
                wa_w, wa_h = rect.right - rect.left, rect.bottom - rect.top
        except Exception:
            pass

        width = int(min(960, wa_w / scaling - 40))
        height = int(min(680, wa_h / scaling - 70))
        x = int(left + (wa_w - width * scaling) / 2)
        y = int(top + (wa_h - (height + 40) * scaling) / 2)
        self.geometry(f"{width}x{height}+{max(x, left)}+{max(y, top)}")
        self.minsize(min(860, width), min(600, height))

    def _warn_no_pypdf(self):
        messagebox.showwarning(
            APP_NAME,
            "The 'pypdf' library is not installed, so page counts are unavailable.\n\n"
            "Install it with:\npy -m pip install pypdf\n\n"
            "Then restart PrintNow.",
        )

    # ------------------------------------------------------------------ UI
    def _build_header(self):
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=28, pady=(14, 4))

        ctk.CTkLabel(
            header, text="PRINT", font=("Consolas", 26, "bold"), text_color=TEXT
        ).pack(side="left")
        ctk.CTkLabel(
            header, text="NOW", font=("Consolas", 26, "bold"), text_color=ACCENT
        ).pack(side="left")
        ctk.CTkLabel(
            header,
            text="  //  batch print console",
            font=("Consolas", 13),
            text_color=MUTED,
        ).pack(side="left", pady=(10, 0))

        line = ctk.CTkFrame(self, height=1, fg_color=BORDER)
        line.pack(fill="x", padx=28, pady=(0, 10))

    def _panel(self, parent):
        return ctk.CTkFrame(
            parent,
            fg_color=PANEL,
            corner_radius=14,
            border_width=1,
            border_color=BORDER,
        )

    def _section_title(self, parent, text):
        ctk.CTkLabel(
            parent, text=text.upper(), font=("Consolas", 12, "bold"), text_color=ACCENT
        ).pack(anchor="w", padx=18, pady=(10, 4))

    def _build_body(self):
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=28, pady=(0, 0))
        body.grid_columnconfigure(0, weight=3)
        body.grid_columnconfigure(1, weight=2)
        body.grid_rowconfigure(0, weight=1)

        # ---------- Left: folder + files ----------
        left = self._panel(body)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        self._section_title(left, "01 / Source folder")

        folder_row = ctk.CTkFrame(left, fg_color="transparent")
        folder_row.pack(fill="x", padx=18)
        self.folder_label = ctk.CTkLabel(
            folder_row,
            text="No folder selected",
            font=("Consolas", 12),
            text_color=MUTED,
            anchor="w",
        )
        self.folder_label.pack(side="left", fill="x", expand=True)
        self._button(folder_row, "Browse", self.choose_folder, width=100).pack(
            side="right"
        )

        self.subfolders_var = ctk.BooleanVar(value=False)
        ctk.CTkSwitch(
            left,
            text="Include subfolders",
            variable=self.subfolders_var,
            command=self.refresh_files,
            progress_color=ACCENT,
            font=("Consolas", 12),
            text_color=TEXT,
        ).pack(anchor="w", padx=18, pady=(10, 0))

        self._section_title(left, "02 / Files")

        tools = ctk.CTkFrame(left, fg_color="transparent")
        tools.pack(fill="x", padx=18)
        self._button(tools, "Select all", lambda: self.set_all(True), ghost=True, width=100).pack(side="left")
        self._button(tools, "Clear", lambda: self.set_all(False), ghost=True, width=80).pack(side="left", padx=8)
        self.count_label = ctk.CTkLabel(
            tools, text="0 selected", font=("Consolas", 12), text_color=MUTED
        )
        self.count_label.pack(side="right")

        self.files_frame = ctk.CTkScrollableFrame(
            left, fg_color=PANEL_2, corner_radius=10, border_width=1, border_color=BORDER
        )
        self.files_frame.pack(fill="both", expand=True, padx=18, pady=(10, 18))

        # ---------- Right: printer + layout + settings ----------
        right = self._panel(body)
        right.grid(row=0, column=1, sticky="nsew", padx=(10, 0))

        # Printer
        self._section_title(right, "03 / Printer")
        printer_row = ctk.CTkFrame(right, fg_color="transparent")
        printer_row.pack(fill="x", padx=18)
        self.printer_var = ctk.StringVar()
        self.printer_menu = self._option(printer_row, self.printer_var, [""])
        self.printer_menu.pack(side="left", fill="x", expand=True)
        self._button(printer_row, "↻", self._load_printers, ghost=True, width=36).pack(
            side="right", padx=(8, 0)
        )

        # Layout presets
        self._section_title(right, "04 / Layout")
        layout_box = ctk.CTkFrame(right, fg_color="transparent")
        layout_box.pack(fill="x", padx=18)

        self.layout_var = ctk.StringVar(value=NO_LAYOUT)
        self.layout_menu = self._option(
            layout_box, self.layout_var, [NO_LAYOUT], command=self.on_layout_selected
        )
        self.layout_menu.pack(fill="x")

        layout_buttons = ctk.CTkFrame(layout_box, fg_color="transparent")
        layout_buttons.pack(fill="x", pady=(8, 0))
        for col in range(4):
            layout_buttons.grid_columnconfigure(col, weight=1, uniform="lb")
        self._button(layout_buttons, "Save", self.save_layout, ghost=True, width=60).grid(
            row=0, column=0, sticky="ew", padx=(0, 4)
        )
        self._button(layout_buttons, "Save as", self.save_layout_as, ghost=True, width=60).grid(
            row=0, column=1, sticky="ew", padx=4
        )
        self._button(layout_buttons, "Delete", self.delete_layout, ghost=True, width=60).grid(
            row=0, column=2, sticky="ew", padx=4
        )
        self._button(layout_buttons, "★ Default", self.toggle_default_layout, ghost=True, width=60).grid(
            row=0, column=3, sticky="ew", padx=(4, 0)
        )

        # Settings
        self._section_title(right, "05 / Settings")
        settings = ctk.CTkFrame(right, fg_color="transparent")
        settings.pack(fill="x", padx=18)
        settings.grid_columnconfigure(1, weight=1)

        # Copies
        self._label(settings, "Copies").grid(row=0, column=0, sticky="w", pady=3)
        copies_box = ctk.CTkFrame(settings, fg_color="transparent")
        copies_box.grid(row=0, column=1, sticky="e", pady=3)
        self.copies_var = ctk.StringVar(value=DEFAULT_SETTINGS["copies"])
        self._button(copies_box, "−", lambda: self.change_copies(-1), ghost=True, width=32).pack(side="left")
        ctk.CTkEntry(
            copies_box,
            textvariable=self.copies_var,
            width=54,
            justify="center",
            font=("Consolas", 13),
            fg_color=PANEL_2,
            border_color=BORDER,
            text_color=TEXT,
        ).pack(side="left", padx=6)
        self._button(copies_box, "+", lambda: self.change_copies(1), ghost=True, width=32).pack(side="left")

        # Sides
        self.duplex_var = ctk.StringVar(value=DEFAULT_SETTINGS["duplex"])
        self._label(settings, "Sides").grid(row=1, column=0, sticky="w", pady=3)
        self._option(settings, self.duplex_var, DUPLEX_OPTIONS).grid(
            row=1, column=1, sticky="e", pady=5
        )

        # Color
        self.color_var = ctk.StringVar(value=DEFAULT_SETTINGS["color"])
        self._label(settings, "Color mode").grid(row=2, column=0, sticky="w", pady=3)
        self._option(settings, self.color_var, COLOR_OPTIONS).grid(
            row=2, column=1, sticky="e", pady=5
        )

        # Orientation
        self.orientation_var = ctk.StringVar(value=DEFAULT_SETTINGS["orientation"])
        self._label(settings, "Orientation").grid(row=3, column=0, sticky="w", pady=3)
        self._option(settings, self.orientation_var, ORIENTATION_OPTIONS).grid(
            row=3, column=1, sticky="e", pady=5
        )

        # Paper
        self.paper_var = ctk.StringVar(value=DEFAULT_SETTINGS["paper"])
        self._label(settings, "Paper size").grid(row=4, column=0, sticky="w", pady=3)
        self._option(settings, self.paper_var, PAPER_OPTIONS).grid(
            row=4, column=1, sticky="e", pady=5
        )

        # Scaling
        self.scale_var = ctk.StringVar(value=DEFAULT_SETTINGS["scale"])
        self._label(settings, "Scaling").grid(row=5, column=0, sticky="w", pady=3)
        self._option(settings, self.scale_var, SCALE_OPTIONS).grid(
            row=5, column=1, sticky="e", pady=5
        )

    def _build_footer(self):
        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.pack(side="bottom", fill="x", padx=28, pady=(10, 14))

        self.status_label = ctk.CTkLabel(
            footer, text="Ready.", font=("Consolas", 12), text_color=MUTED, anchor="w"
        )
        self.status_label.pack(side="left", fill="x", expand=True)

        self.progress = ctk.CTkProgressBar(
            footer, width=220, height=6, progress_color=ACCENT, fg_color=BORDER
        )
        self.progress.set(0)
        self.progress.pack(side="left", padx=16)

        self.print_btn = ctk.CTkButton(
            footer,
            text="PRINT  ▶",
            command=self.start_printing,
            width=170,
            height=42,
            corner_radius=10,
            fg_color=ACCENT,
            hover_color=ACCENT_HOVER,
            text_color="#001418",
            font=("Consolas", 15, "bold"),
        )
        self.print_btn.pack(side="right")

    # ----------------------------------------------------------- UI helpers
    def _button(self, parent, text, command, ghost=False, width=110):
        return ctk.CTkButton(
            parent,
            text=text,
            command=command,
            width=width,
            height=32,
            corner_radius=8,
            fg_color="transparent" if ghost else ACCENT,
            hover_color=BORDER if ghost else ACCENT_HOVER,
            border_width=1 if ghost else 0,
            border_color=BORDER,
            text_color=TEXT if ghost else "#001418",
            font=("Consolas", 12, "bold"),
        )

    def _option(self, parent, variable, values, command=None):
        return ctk.CTkOptionMenu(
            parent,
            variable=variable,
            values=values,
            command=command,
            width=200,
            height=32,
            corner_radius=8,
            fg_color=PANEL_2,
            button_color=BORDER,
            button_hover_color=ACCENT_HOVER,
            dropdown_fg_color=PANEL_2,
            dropdown_hover_color=BORDER,
            dropdown_text_color=TEXT,
            text_color=TEXT,
            font=("Consolas", 12),
            dropdown_font=("Consolas", 12),
        )

    def _label(self, parent, text):
        return ctk.CTkLabel(parent, text=text, font=("Consolas", 12), text_color=MUTED)

    def set_status(self, text, color=MUTED):
        self.status_label.configure(text=text, text_color=color)

    # -------------------------------------------------------------- Printers
    def _load_printers(self):
        printers = list_printers()
        if not printers:
            printers = ["No printers found"]
        self.printer_menu.configure(values=printers)
        default = default_printer()
        self.printer_var.set(default if default in printers else printers[0])

    # --------------------------------------------------------------- Layouts
    def _layout_display(self, name):
        return f"★ {name}" if name == self.config_data["default"] else name

    def _refresh_layout_menu(self):
        names = sorted(self.config_data["layouts"].keys(), key=str.lower)
        values = [NO_LAYOUT] + [self._layout_display(n) for n in names]
        self.layout_menu.configure(values=values)
        if self.current_layout in self.config_data["layouts"]:
            self.layout_var.set(self._layout_display(self.current_layout))
        else:
            self.current_layout = None
            self.layout_var.set(NO_LAYOUT)

    def get_current_settings(self):
        try:
            copies = str(max(1, min(99, int(self.copies_var.get()))))
        except ValueError:
            copies = "1"
        return {
            "copies": copies,
            "duplex": self.duplex_var.get(),
            "color": self.color_var.get(),
            "orientation": self.orientation_var.get(),
            "paper": self.paper_var.get(),
            "scale": self.scale_var.get(),
        }

    def apply_settings(self, settings):
        def pick(key, options):
            value = settings.get(key, DEFAULT_SETTINGS[key])
            return value if value in options else DEFAULT_SETTINGS[key]

        copies = str(settings.get("copies", "1"))
        self.copies_var.set(copies if copies.isdigit() and int(copies) >= 1 else "1")
        self.duplex_var.set(pick("duplex", DUPLEX_OPTIONS))
        self.color_var.set(pick("color", COLOR_OPTIONS))
        self.orientation_var.set(pick("orientation", ORIENTATION_OPTIONS))
        self.paper_var.set(pick("paper", PAPER_OPTIONS))
        self.scale_var.set(pick("scale", SCALE_OPTIONS))

    def _load_default_layout(self):
        default = self.config_data["default"]
        if default and default in self.config_data["layouts"]:
            self.current_layout = default
            self.apply_settings(self.config_data["layouts"][default])
        self._refresh_layout_menu()

    def on_layout_selected(self, choice):
        if choice == NO_LAYOUT:
            self.current_layout = None
            return
        name = choice[2:] if choice.startswith("★ ") else choice
        if name in self.config_data["layouts"]:
            self.current_layout = name
            self.apply_settings(self.config_data["layouts"][name])
            self.set_status(f"Layout '{name}' loaded.", ACCENT)

    def save_layout(self):
        if not self.current_layout:
            self.save_layout_as()
            return
        self.config_data["layouts"][self.current_layout] = self.get_current_settings()
        save_config(self.config_data)
        self._refresh_layout_menu()
        self.set_status(f"Layout '{self.current_layout}' updated.", SUCCESS)

    def save_layout_as(self):
        dialog = ctk.CTkInputDialog(text="Layout name:", title="Save layout")
        name = dialog.get_input()
        if not name:
            return
        name = name.strip().lstrip("★").strip()
        if not name or name == NO_LAYOUT:
            messagebox.showwarning(APP_NAME, "Please enter a valid name.")
            return
        self.config_data["layouts"][name] = self.get_current_settings()
        self.current_layout = name
        save_config(self.config_data)
        self._refresh_layout_menu()
        self.set_status(f"Layout '{name}' saved.", SUCCESS)

    def delete_layout(self):
        name = self.current_layout
        if not name:
            messagebox.showinfo(APP_NAME, "Select a layout to delete.")
            return
        if not messagebox.askyesno(APP_NAME, f"Delete layout '{name}'?"):
            return
        self.config_data["layouts"].pop(name, None)
        if self.config_data["default"] == name:
            self.config_data["default"] = None
        self.current_layout = None
        save_config(self.config_data)
        self._refresh_layout_menu()
        self.set_status(f"Layout '{name}' deleted.", MUTED)

    def toggle_default_layout(self):
        name = self.current_layout
        if not name:
            messagebox.showinfo(APP_NAME, "Select or save a layout first.")
            return
        if self.config_data["default"] == name:
            self.config_data["default"] = None
            self.set_status(f"'{name}' is no longer the default layout.", MUTED)
        else:
            self.config_data["default"] = name
            self.set_status(f"'{name}' will load automatically on startup.", SUCCESS)
        save_config(self.config_data)
        self._refresh_layout_menu()

    # ----------------------------------------------------------------- Files
    def choose_folder(self):
        folder = filedialog.askdirectory(title="Select a folder")
        if folder:
            self.folder = folder
            self.folder_label.configure(text=folder, text_color=TEXT)
            self.refresh_files()

    def refresh_files(self):
        for widget in self.files_frame.winfo_children():
            widget.destroy()
        self.file_rows.clear()

        if not self.folder:
            return

        found = []
        if self.subfolders_var.get():
            for root, _dirs, files in os.walk(self.folder):
                for f in files:
                    if f.lower().endswith(SUPPORTED_EXT):
                        found.append(os.path.join(root, f))
        else:
            for f in os.listdir(self.folder):
                full = os.path.join(self.folder, f)
                if os.path.isfile(full) and f.lower().endswith(SUPPORTED_EXT):
                    found.append(full)

        found.sort(key=lambda p: p.lower())

        if not found:
            ctk.CTkLabel(
                self.files_frame,
                text="No printable files found.",
                font=("Consolas", 12),
                text_color=MUTED,
            ).pack(pady=20)

        for path in found:
            var = ctk.BooleanVar(value=False)
            rel = os.path.relpath(path, self.folder)

            row = ctk.CTkFrame(self.files_frame, fg_color="transparent")
            row.pack(fill="x", pady=3, padx=6)

            status = ctk.CTkLabel(
                row, text="", width=100, anchor="e", font=("Consolas", 12, "bold"), text_color=MUTED
            )
            status.pack(side="right")

            pages = ctk.CTkLabel(
                row, text="", width=80, anchor="e", font=("Consolas", 12), text_color=PAGES_COLOR
            )
            pages.pack(side="right", padx=(0, 8))

            ctk.CTkCheckBox(
                row,
                text=rel,
                variable=var,
                command=lambda v=var, p=path, pl=pages: self.on_file_toggle(v, p, pl),
                checkbox_width=18,
                checkbox_height=18,
                corner_radius=4,
                fg_color=ACCENT,
                hover_color=ACCENT_HOVER,
                border_color=BORDER,
                checkmark_color="#001418",
                text_color=TEXT,
                font=("Consolas", 12),
            ).pack(side="left", fill="x", expand=True)

            self.file_rows.append((var, path, status, pages))

        self.update_count()

    def on_file_toggle(self, var, path, pages_label):
        if var.get():
            self.show_pages(var, path, pages_label)
        else:
            pages_label.configure(text="")
        self.update_count()

    def show_pages(self, var, path, pages_label):
        if path in self.page_cache:
            pages_label.configure(text=pages_text(self.page_cache[path]))
            return
        pages_label.configure(text="...")
        self.pool.submit(self._count_task, var, path, pages_label)

    def _count_task(self, var, path, pages_label):
        n, err = count_pages(path)
        try:
            self.after(0, self._apply_pages, var, path, pages_label, n, err)
        except Exception:
            pass

    def _apply_pages(self, var, path, pages_label, n, err=None):
        self.page_cache[path] = n
        try:
            if var.get():
                pages_label.configure(text=pages_text(n))
            if err:
                self.set_status(f"Could not read page count - {err}", DANGER)
            self.update_count()
        except Exception:
            pass  # row no longer exists

    def set_all(self, value):
        for var, path, _status, pages in self.file_rows:
            var.set(value)
            if value:
                self.show_pages(var, path, pages)
            else:
                pages.configure(text="")
        self.update_count()

    def update_count(self):
        selected = [path for var, path, _s, _p in self.file_rows if var.get()]
        total_pages = sum(self.page_cache.get(path) or 0 for path in selected)
        text = f"{len(selected)} selected"
        if total_pages:
            text += f"  ·  {total_pages} pages"
        self.count_label.configure(text=text)

    def set_file_status(self, label, text, color):
        label.configure(text=text, text_color=color)

    # -------------------------------------------------------------- Printing
    def change_copies(self, delta):
        try:
            value = int(self.copies_var.get())
        except ValueError:
            value = 1
        self.copies_var.set(str(max(1, min(99, value + delta))))

    def build_print_settings(self):
        parts = []

        try:
            copies = max(1, min(99, int(self.copies_var.get())))
        except ValueError:
            copies = 1
        if copies > 1:
            parts.append(f"{copies}x")

        duplex = self.duplex_var.get()
        if duplex == "Two-sided (long edge)":
            parts.append("duplexlong")
        elif duplex == "Two-sided (short edge)":
            parts.append("duplexshort")
        else:
            parts.append("simplex")

        parts.append("color" if self.color_var.get() == "Color" else "monochrome")

        orientation = self.orientation_var.get()
        if orientation == "Portrait":
            parts.append("portrait")
        elif orientation == "Landscape":
            parts.append("landscape")

        paper = self.paper_var.get()
        if paper != "Printer default":
            parts.append(f"paper={paper}")

        scale = self.scale_var.get()
        parts.append({"Shrink to fit": "shrink", "Fit to page": "fit", "Actual size": "noscale"}[scale])

        return ",".join(parts)

    def start_printing(self):
        if self.is_printing:
            return

        selected = [(path, status) for var, path, status, _p in self.file_rows if var.get()]
        printer = self.printer_var.get()

        if not selected:
            messagebox.showwarning(APP_NAME, "Please select at least one file.")
            return
        if not printer or printer == "No printers found":
            messagebox.showwarning(APP_NAME, "Please select a printer.")
            return

        sumatra = find_sumatra()
        if not sumatra:
            messagebox.showerror(
                APP_NAME,
                "SumatraPDF.exe was not found.\n"
                "Place it next to PrintNow.exe and try again.",
            )
            return

        settings = self.build_print_settings()
        self.is_printing = True
        self.print_btn.configure(state="disabled", text="PRINTING...")
        self.progress.set(0)

        for _path, status in selected:
            self.set_file_status(status, "Queued", MUTED)

        threading.Thread(
            target=self._print_worker,
            args=(sumatra, printer, settings, selected),
            daemon=True,
        ).start()

    def _print_worker(self, sumatra, printer, settings, files):
        total = len(files)
        errors = []

        for index, (path, status) in enumerate(files, start=1):
            name = os.path.basename(path)
            self.after(0, self.set_status, f"Printing {index}/{total}: {name}", ACCENT)
            self.after(0, self.set_file_status, status, "Printing...", ACCENT)
            failed = False

            try:
                result = subprocess.run(
                    [
                        sumatra,
                        "-print-to", printer,
                        "-print-settings", settings,
                        "-silent",
                        path,
                    ],
                    creationflags=subprocess.CREATE_NO_WINDOW,
                    timeout=300,
                )
                if result.returncode != 0:
                    failed = True
            except Exception:
                failed = True

            if failed:
                errors.append(name)
                self.after(0, self.set_file_status, status, "Failed ✗", DANGER)
                self.after(0, self.set_status, f"{name} failed", DANGER)
            else:
                self.after(0, self.set_file_status, status, "Done ✓", SUCCESS)
                self.after(0, self.set_status, f"{name} finished", SUCCESS)

            self.after(0, self.progress.set, index / total)

        self.after(0, self._print_done, total, errors)

    def _print_done(self, total, errors):
        self.is_printing = False
        self.print_btn.configure(state="normal", text="PRINT  ▶")

        if errors:
            self.set_status(
                f"Done with {len(errors)} error(s): {', '.join(errors[:3])}"
                + ("..." if len(errors) > 3 else ""),
                DANGER,
            )
        else:
            self.set_status(f"Done. {total} file(s) sent to printer.", SUCCESS)


if __name__ == "__main__":
    app = PrintNowApp()
    app.mainloop()
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, simpledialog, filedialog
import threading
import time
import traceback
import json
import os
import re
import sys
import webbrowser
import requests
import pyperclip
import gspread
from pynput import mouse

# === CONFIGURATION (defaults — all editable in the UI) ===
APP_NAME = "AutoPaster"
CURRENT_VERSION = "v1.0.0"                 # bump this (and the GitHub release tag) for each release
GITHUB_REPO = "username/repository"        # TODO: your public GitHub repo as "owner/repo"
CREDENTIALS_FILENAME = "credentials.json"
SETTINGS_FILENAME = "auto_paste_settings.json"
REQUIRED_CREDENTIAL_KEYS = ("type", "project_id", "private_key", "client_email", "token_uri")
DEFAULT_SETTINGS = {
    "spreadsheet": "",
    "spreadsheets": [""],   # saved list shown in the dropdown
    "tab": "Auto-Paste",
    "tabs": {},                              # last tab used, per spreadsheet
    "start_row": "2",
    "col_mode": "range",        # "range" = start column + count, "custom" = explicit list
    "start_col": "A",
    "items_per_row": 5,
    "custom_cols": "A, C, E",
    "strip_links": False,
    "link_mode": "text",        # "text" = convert to plain text, "discard" = leave cell blank
}


# =========================
# Paths
# =========================
def app_dir():
    # Folder of the .exe when frozen by PyInstaller, otherwise the script's folder
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def user_config_dir():
    """Per-user folder outside the project (never committed, not shipped with the exe).
    Windows: %APPDATA%\\AutoPaster   macOS/Linux: ~/.config/autopaster"""
    if sys.platform == "win32":
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
        path = os.path.join(base, APP_NAME)
    else:
        base = os.environ.get("XDG_CONFIG_HOME") or os.path.expanduser("~/.config")
        path = os.path.join(base, APP_NAME.lower())
    os.makedirs(path, exist_ok=True)
    return path


def credentials_path():
    return os.path.join(user_config_dir(), CREDENTIALS_FILENAME)


def has_credentials():
    return os.path.isfile(credentials_path())


def validate_credentials(text):
    """Parse and sanity-check a service-account key. Returns the dict or raises ValueError."""
    try:
        info = json.loads(text)
    except ValueError as e:
        raise ValueError(f"Not valid JSON: {e}")
    if not isinstance(info, dict):
        raise ValueError("Expected a JSON object")
    if info.get("type") != "service_account":
        raise ValueError("This is not a service-account key (\"type\" must be \"service_account\")")
    missing = [k for k in REQUIRED_CREDENTIAL_KEYS if not info.get(k)]
    if missing:
        raise ValueError(f"Missing field(s): {', '.join(missing)}")
    try:
        # Confirms the private key actually parses, without contacting Google
        from google.oauth2.service_account import Credentials
        Credentials.from_service_account_info(info)
    except Exception as e:
        raise ValueError(f"The private key could not be loaded: {e}")
    return info


def restrict_to_current_user(path):
    """Best effort: make the file readable/writable by the current user only."""
    try:
        if sys.platform == "win32":
            import subprocess
            user = os.environ.get("USERNAME")
            if user:
                subprocess.run(["icacls", path, "/inheritance:r", "/grant:r", f"{user}:F"],
                               capture_output=True, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        else:
            os.chmod(path, 0o600)
    except Exception:
        pass


def save_credentials(info):
    # Write to a temp file and swap it in, so a crash never leaves a half-written key
    path = credentials_path()
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(info, f, indent=2)
    restrict_to_current_user(tmp)
    os.replace(tmp, path)
    restrict_to_current_user(path)


def service_account_email():
    try:
        with open(credentials_path(), encoding="utf-8") as f:
            return json.load(f).get("client_email", "")
    except (OSError, ValueError):
        return ""


def open_spreadsheet(client, ref):
    """Open by full URL, by spreadsheet ID, or by exact file name."""
    if "docs.google.com" in ref:
        return client.open_by_url(ref)
    if re.fullmatch(r"[A-Za-z0-9_-]{30,}", ref):
        return client.open_by_key(ref)
    return client.open(ref)


# =========================
# Column helpers
# =========================
def col_to_index(letters):
    letters = letters.strip().upper()
    if not re.fullmatch(r"[A-Z]{1,3}", letters):
        raise ValueError(f"'{letters}' is not a valid column letter")
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def index_to_col(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def parse_column_spec(spec):
    """'A, C, E-G' -> [1, 3, 5, 6, 7]. Order is preserved (it's the paste order)."""
    cols = []
    for part in re.split(r"[,\s;]+", spec.strip()):
        if not part:
            continue
        if "-" in part or ":" in part:
            a, b = re.split(r"[-:]", part, maxsplit=1)
            a, b = col_to_index(a), col_to_index(b)
            step = 1 if b >= a else -1
            cols.extend(range(a, b + step, step))
        else:
            cols.append(col_to_index(part))
    if not cols:
        raise ValueError("Enter at least one column")
    dupes = sorted({index_to_col(c) for c in cols if cols.count(c) > 1})
    if dupes:
        raise ValueError(f"Duplicate column(s): {', '.join(dupes)}")
    return cols


def build_row_updates(row, cols, values):
    """Group adjacent target columns into ranges so a whole row is one batch request."""
    data = []
    i = 0
    while i < len(cols):
        j = i
        while j + 1 < len(cols) and cols[j + 1] == cols[j] + 1:
            j += 1
        rng = f"{index_to_col(cols[i])}{row}:{index_to_col(cols[j])}{row}"
        data.append({"range": rng, "values": [values[i:j + 1]]})
        i = j + 1
    return data


# =========================
# Link detection / stripping
# =========================
URL_RE = re.compile(r"(?:https?://|ftp://|www\.|mailto:)[^\s<>\"']+", re.I)
HYPERLINK_FORMULA_RE = re.compile(
    r'^\s*=\s*HYPERLINK\(\s*"([^"]*)"\s*(?:[,;]\s*"([^"]*)"\s*)?\)\s*$', re.I)
MD_LINK_RE = re.compile(r"\[([^\]]*)\]\([^)\s]+\)")
ANCHOR_RE = re.compile(r"<a\s[^>]*href\s*=", re.I)


def get_clipboard_html():
    """Return the HTML fragment on the Windows clipboard ('' if none/unavailable).
    Rich copies (browsers, Word, Sheets) put the real <a href> here, while the
    plain-text clipboard only holds the visible label."""
    if sys.platform != "win32":
        return ""
    try:
        import ctypes
        from ctypes import wintypes
        user32, kernel32 = ctypes.windll.user32, ctypes.windll.kernel32
        user32.RegisterClipboardFormatW.restype = wintypes.UINT
        user32.GetClipboardData.argtypes = [wintypes.UINT]
        user32.GetClipboardData.restype = wintypes.HANDLE
        kernel32.GlobalLock.argtypes = [wintypes.HGLOBAL]
        kernel32.GlobalLock.restype = ctypes.c_void_p
        kernel32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
        kernel32.GlobalSize.argtypes = [wintypes.HGLOBAL]
        kernel32.GlobalSize.restype = ctypes.c_size_t

        fmt = user32.RegisterClipboardFormatW("HTML Format")
        if not user32.IsClipboardFormatAvailable(fmt):
            return ""
        for _ in range(5):  # clipboard may be briefly held by another app
            if user32.OpenClipboard(None):
                break
            time.sleep(0.02)
        else:
            return ""
        try:
            handle = user32.GetClipboardData(fmt)
            if not handle:
                return ""
            ptr = kernel32.GlobalLock(handle)
            if not ptr:
                return ""
            try:
                raw = ctypes.string_at(ptr, kernel32.GlobalSize(handle))
            finally:
                kernel32.GlobalUnlock(handle)
        finally:
            user32.CloseClipboard()

        html = raw.split(b"\0", 1)[0].decode("utf-8", errors="ignore")
        m = re.search(r"<!--StartFragment-->(.*)<!--EndFragment-->", html, re.S)
        return m.group(1) if m else html
    except Exception:
        return ""


def contains_link(text, html=""):
    return bool(HYPERLINK_FORMULA_RE.match(text) or MD_LINK_RE.search(text)
                or URL_RE.search(text) or (html and ANCHOR_RE.search(html)))


def strip_links(text):
    """Keep the human-readable part of a link, drop the URL itself."""
    m = HYPERLINK_FORMULA_RE.match(text)
    if m:
        return m.group(2) or ""
    text = MD_LINK_RE.sub(r"\1", text)
    text = URL_RE.sub("", text)
    return re.sub(r"[ \t]{2,}", " ", text).strip()


def preview(value, limit=30):
    value = value.replace("\n", " ").replace("\r", "")
    return value if len(value) <= limit else value[:limit - 1] + "…"


# =========================
# Settings persistence
# =========================
def load_settings():
    settings = dict(DEFAULT_SETTINGS)
    # Fall back to the old location next to the app so existing setups carry over
    for folder in (user_config_dir(), app_dir()):
        try:
            with open(os.path.join(folder, SETTINGS_FILENAME), encoding="utf-8") as f:
                settings.update(json.load(f))
            break
        except (OSError, ValueError):
            continue
    return settings


def save_settings(settings):
    try:
        with open(os.path.join(user_config_dir(), SETTINGS_FILENAME), "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2)
    except OSError:
        pass


class CredentialsDialog:
    """Modal window to paste or browse for a Google service-account JSON key."""

    def __init__(self, parent, first_run=False):
        self.saved = False
        self.win = tk.Toplevel(parent)
        self.win.title("Google Credentials Setup")
        self.win.geometry("540x420")
        self.win.transient(parent)

        intro = ("Welcome! To write to Google Sheets, this app needs your service-account key (JSON).\n\n"
                 if first_run else "Replace the stored service-account key (JSON).\n\n")
        intro += ("Click 'Browse…' to pick the .json file, or paste its contents below.\n"
                  f"It will be saved privately for your Windows user only at:\n{credentials_path()}")
        ttk.Label(self.win, text=intro, wraplength=510, justify="left").pack(anchor="w", padx=12, pady=(12, 6))

        self.text = scrolledtext.ScrolledText(self.win, width=64, height=12, font=("Consolas", 9))
        self.text.pack(fill=tk.BOTH, expand=True, padx=12)

        btns = tk.Frame(self.win)
        btns.pack(fill=tk.X, padx=12, pady=10)
        tk.Button(btns, text="Browse…", command=self.browse).pack(side=tk.LEFT)
        tk.Button(btns, text="Cancel", command=self.win.destroy, width=10).pack(side=tk.RIGHT)
        tk.Button(btns, text="Save", command=self.save, width=10, bg="#90ee90").pack(side=tk.RIGHT, padx=6)

        self.win.grab_set()
        self.win.focus_set()
        parent.wait_window(self.win)

    def browse(self):
        path = filedialog.askopenfilename(parent=self.win, title="Select service-account JSON key",
                                          filetypes=[("JSON key", "*.json"), ("All files", "*.*")])
        if not path:
            return
        try:
            with open(path, encoding="utf-8") as f:
                content = f.read()
        except OSError as e:
            messagebox.showerror("Error", f"Could not read file:\n{e}", parent=self.win)
            return
        self.text.delete("1.0", tk.END)
        self.text.insert("1.0", content)

    def save(self):
        try:
            info = validate_credentials(self.text.get("1.0", tk.END).strip())
            save_credentials(info)
        except ValueError as e:
            messagebox.showerror("Invalid Credentials", str(e), parent=self.win)
            return
        except OSError as e:
            messagebox.showerror("Error", f"Could not save credentials:\n{e}", parent=self.win)
            return
        # Don't keep the private key around in the widget
        self.text.delete("1.0", tk.END)
        self.saved = True
        messagebox.showinfo("Saved", f"Credentials saved for {info['client_email']}.\n\n"
                            "Share each spreadsheet (Editor) with that email.\n"
                            "You can now delete the original .json file you downloaded.", parent=self.win)
        self.win.destroy()


# =========================
# Check for Updates (GitHub Releases)
# =========================
_update_check_running = False


def parse_version(tag):
    """'v1.2.3', '1.2', 'v2.0.0-beta' -> (1, 2, 3), (1, 2, 0), (2, 0, 0)"""
    core = re.split(r"[-+]", tag.strip().lstrip("vV"), maxsplit=1)[0]
    nums = [int(n) for n in re.findall(r"\d+", core)[:3]]
    return tuple(nums + [0] * (3 - len(nums)))


def check_for_updates(parent=None):
    """Ask GitHub for the latest release and offer to download it if it's newer.
    The network call runs in a background thread so the window never freezes."""
    global _update_check_running
    parent = parent or tk._default_root
    if _update_check_running:
        return
    if GITHUB_REPO == "username/repository" or GITHUB_REPO.count("/") != 1:
        messagebox.showerror("Update Check", "GITHUB_REPO is not configured.\n"
                             "Set it to your repository as \"owner/repo\" in Auto.py.", parent=parent)
        return
    _update_check_running = True

    def fetch():
        url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
        try:
            resp = requests.get(url, timeout=10, headers={
                "Accept": "application/vnd.github+json",
                "User-Agent": f"{APP_NAME}/{CURRENT_VERSION}",
            })
            if resp.status_code == 404:
                raise RuntimeError(f"No published release found for '{GITHUB_REPO}'.\n\n"
                                   "Check the repository name, and that the repo is public "
                                   "and has at least one (non-draft) release.")
            if resp.status_code == 403 and resp.headers.get("X-RateLimit-Remaining") == "0":
                raise RuntimeError("GitHub's rate limit was reached. Please try again in a while.")
            resp.raise_for_status()
            result = ("ok", resp.json())
        except requests.exceptions.RequestException as e:
            result = ("error", f"Could not reach GitHub. Check your internet connection.\n\nDetails: {e}")
        except (RuntimeError, ValueError) as e:
            result = ("error", str(e))
        parent.after(0, show_result, *result)

    def show_result(status, payload):
        global _update_check_running
        _update_check_running = False
        if status == "error":
            messagebox.showerror("Update Check Failed", payload, parent=parent)
            return

        latest = payload.get("tag_name") or ""
        if not latest or parse_version(latest) <= parse_version(CURRENT_VERSION):
            messagebox.showinfo("Up to Date", f"You're running the latest version ({CURRENT_VERSION}).",
                                parent=parent)
            return

        # Prefer a direct .exe download from the release assets, else the release page
        exe_url = next((a.get("browser_download_url") for a in payload.get("assets", [])
                        if a.get("name", "").lower().endswith(".exe")), None)
        download_url = exe_url or payload.get("html_url") or f"https://github.com/{GITHUB_REPO}/releases/latest"

        notes = (payload.get("body") or "").strip()
        if len(notes) > 400:
            notes = notes[:400].rstrip() + "…"
        message = f"A new version is available!\n\nInstalled: {CURRENT_VERSION}\nLatest: {latest}\n"
        if notes:
            message += f"\nWhat's new:\n{notes}\n"
        message += "\nDownload it now?" if exe_url else "\nOpen the release page to download it?"

        if messagebox.askyesno("Update Available", message, parent=parent):
            if download_url.startswith("https://"):
                webbrowser.open(download_url)

    threading.Thread(target=fetch, daemon=True).start()


class AutoPasterApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"Auto Paster {CURRENT_VERSION} | Made by Safeer")
        self.root.geometry("580x760")
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        self.running = False
        self.client = None
        self.spreadsheet = None
        self.sheet = None
        self.skip_requested = False

        self.row_buffer = []
        self.current_position = 0  # Position inside row
        self.target_cols = []      # Column indices for the active run (snapshot at Start)

        s = load_settings()

        # Tk variables
        self.spreadsheet_var = tk.StringVar(value=s["spreadsheet"])
        self.tab_var = tk.StringVar(value="")
        self.preferred_tab = s["tab"]
        self.spreadsheets = [x for x in s["spreadsheets"] if x]
        if s["spreadsheet"] and s["spreadsheet"] not in self.spreadsheets:
            self.spreadsheets.append(s["spreadsheet"])
        self.tabs_by_spreadsheet = dict(s["tabs"])
        self.col_mode_var = tk.StringVar(value=s["col_mode"])
        self.start_col_var = tk.StringVar(value=s["start_col"])
        self.items_per_row_var = tk.StringVar(value=str(s["items_per_row"]))
        self.custom_cols_var = tk.StringVar(value=s["custom_cols"])
        self.strip_links_var = tk.BooleanVar(value=s["strip_links"])
        self.link_mode_var = tk.StringVar(value=s["link_mode"])

        # Mouse listener
        self.mouse_listener = mouse.Listener(on_click=self.on_click)
        self.mouse_listener.start()

        # ---------- UI ----------
        self.status_label = tk.Label(root, text="Status: Disconnected", fg="red", font=("Arial", 10, "bold"))
        self.status_label.pack(pady=(10, 5))

        # --- Connection / Sheet picker ---
        conn = ttk.LabelFrame(root, text="Spreadsheet & Tab")
        conn.pack(fill=tk.X, padx=12, pady=5)

        ttk.Label(conn, text="Spreadsheet:").grid(row=0, column=0, sticky="w", padx=5, pady=4)
        self.spreadsheet_combo = ttk.Combobox(conn, textvariable=self.spreadsheet_var, state="readonly",
                                              values=self.spreadsheets, width=30)
        self.spreadsheet_combo.grid(row=0, column=1, sticky="we", padx=5, pady=4)
        self.spreadsheet_combo.bind("<<ComboboxSelected>>", self.on_spreadsheet_selected)
        self.btn_connect = tk.Button(conn, text="1. Connect", command=self.connect_to_sheets)
        self.btn_connect.grid(row=0, column=2, sticky="we", padx=5, pady=4)

        manage = tk.Frame(conn)
        manage.grid(row=1, column=1, columnspan=2, sticky="w", padx=5, pady=(0, 4))
        self.btn_add_sheet = tk.Button(manage, text="+ Add Spreadsheet", command=self.add_spreadsheet)
        self.btn_add_sheet.pack(side=tk.LEFT)
        self.btn_remove_sheet = tk.Button(manage, text="− Remove", command=self.remove_spreadsheet)
        self.btn_remove_sheet.pack(side=tk.LEFT, padx=6)
        self.btn_credentials = tk.Button(manage, text="🔑 Credentials…", command=self.change_credentials)
        self.btn_credentials.pack(side=tk.LEFT)

        ttk.Label(conn, text="Target Tab:").grid(row=2, column=0, sticky="w", padx=5, pady=4)
        self.tab_combo = ttk.Combobox(conn, textvariable=self.tab_var, state="disabled", width=30)
        self.tab_combo.grid(row=2, column=1, sticky="we", padx=5, pady=4)
        self.tab_combo.bind("<<ComboboxSelected>>", self.on_tab_selected)
        self.btn_refresh = tk.Button(conn, text="⟳ Refresh", command=self.refresh_tabs, state=tk.DISABLED)
        self.btn_refresh.grid(row=2, column=2, sticky="we", padx=5, pady=4)
        conn.columnconfigure(1, weight=1)

        # --- Row layout / column mapping ---
        layout = ttk.LabelFrame(root, text="Row Layout")
        layout.pack(fill=tk.X, padx=12, pady=5)

        ttk.Label(layout, text="Start Row:").grid(row=0, column=0, sticky="w", padx=5, pady=4)
        self.row_entry = ttk.Entry(layout, width=8)
        self.row_entry.insert(0, s["start_row"])
        self.row_entry.grid(row=0, column=1, sticky="w", padx=5, pady=4)

        self.rb_range = ttk.Radiobutton(layout, text="Consecutive columns", value="range",
                                        variable=self.col_mode_var, command=self.update_layout_state)
        self.rb_range.grid(row=1, column=0, columnspan=4, sticky="w", padx=5, pady=(6, 0))

        ttk.Label(layout, text="Start column:").grid(row=2, column=0, sticky="w", padx=(25, 5), pady=2)
        self.start_col_entry = ttk.Entry(layout, textvariable=self.start_col_var, width=6)
        self.start_col_entry.grid(row=2, column=1, sticky="w", padx=5, pady=2)
        ttk.Label(layout, text="Items per row:").grid(row=2, column=2, sticky="w", padx=5, pady=2)
        self.items_spin = ttk.Spinbox(layout, from_=1, to=200, textvariable=self.items_per_row_var, width=6)
        self.items_spin.grid(row=2, column=3, sticky="w", padx=5, pady=2)

        self.rb_custom = ttk.Radiobutton(layout, text="Specific columns", value="custom",
                                         variable=self.col_mode_var, command=self.update_layout_state)
        self.rb_custom.grid(row=3, column=0, columnspan=4, sticky="w", padx=5, pady=(6, 0))

        ttk.Label(layout, text="Columns:").grid(row=4, column=0, sticky="w", padx=(25, 5), pady=2)
        self.custom_cols_entry = ttk.Entry(layout, textvariable=self.custom_cols_var, width=28)
        self.custom_cols_entry.grid(row=4, column=1, columnspan=3, sticky="we", padx=5, pady=2)
        ttk.Label(layout, text="e.g.  A, C, E   or   A-C, F   (paste order follows this list)",
                  foreground="gray").grid(row=5, column=1, columnspan=3, sticky="w", padx=5)

        self.layout_preview = ttk.Label(layout, text="", foreground="#1a5fb4")
        self.layout_preview.grid(row=6, column=0, columnspan=4, sticky="w", padx=5, pady=(6, 4))

        for var in (self.col_mode_var, self.start_col_var, self.items_per_row_var, self.custom_cols_var):
            var.trace_add("write", lambda *_: self.update_layout_preview())

        # --- Hyperlink filtering ---
        links = ttk.LabelFrame(root, text="Hyperlink Filtering")
        links.pack(fill=tk.X, padx=12, pady=5)

        self.chk_strip = ttk.Checkbutton(links, text="Ignore / Strip Links", variable=self.strip_links_var,
                                         command=self.update_link_state)
        self.chk_strip.grid(row=0, column=0, columnspan=2, sticky="w", padx=5, pady=4)
        self.rb_link_text = ttk.Radiobutton(links, text="Convert to plain text (keep label, drop URL)",
                                            value="text", variable=self.link_mode_var)
        self.rb_link_text.grid(row=1, column=0, sticky="w", padx=(25, 5), pady=2)
        self.rb_link_discard = ttk.Radiobutton(links, text="Ignore link (stay on the same cell)",
                                               value="discard", variable=self.link_mode_var)
        self.rb_link_discard.grid(row=2, column=0, sticky="w", padx=(25, 5), pady=(2, 6))

        # --- Start / Stop ---
        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=8)

        self.btn_start = tk.Button(btn_frame, text="2. Start", command=self.start, state=tk.DISABLED,
                                   bg="#90ee90", width=12)
        self.btn_start.pack(side=tk.LEFT, padx=10)

        self.btn_stop = tk.Button(btn_frame, text="Stop", command=self.stop, state=tk.DISABLED,
                                  bg="#ffcccb", width=12)
        self.btn_stop.pack(side=tk.LEFT, padx=10)

        self.btn_update = tk.Button(btn_frame, text="Check for Updates",
                                    command=lambda: check_for_updates(self.root))
        self.btn_update.pack(side=tk.LEFT, padx=10)

        self.log_area = scrolledtext.ScrolledText(root, width=66, height=14, state='disabled')
        self.log_area.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 10))

        # Controls that must be locked while a run is active
        self.config_widgets = [
            self.row_entry, self.rb_range, self.rb_custom, self.start_col_entry, self.items_spin,
            self.custom_cols_entry, self.chk_strip, self.rb_link_text, self.rb_link_discard,
        ]

        self.update_layout_state()
        self.update_link_state()

        # First launch: ask once for the Google key (stored per-user, outside the project)
        if not has_credentials():
            self.root.after(300, lambda: self.prompt_for_credentials(first_run=True))

    # =========================
    # Logging (safe to call from the worker thread)
    # =========================
    def log(self, message):
        self.root.after(0, self._append_log, message)

    def _append_log(self, message):
        self.log_area.config(state='normal')
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state='disabled')

    # =========================
    # UI state helpers
    # =========================
    def update_layout_state(self):
        custom = self.col_mode_var.get() == "custom"
        if not self.running:
            self.start_col_entry.config(state="disabled" if custom else "normal")
            self.items_spin.config(state="disabled" if custom else "normal")
            self.custom_cols_entry.config(state="normal" if custom else "disabled")
        self.update_layout_preview()

    def update_link_state(self):
        state = "normal" if self.strip_links_var.get() and not self.running else "disabled"
        self.rb_link_text.config(state=state)
        self.rb_link_discard.config(state=state)

    def update_layout_preview(self):
        try:
            cols = self.get_target_columns()
            names = [index_to_col(c) for c in cols]
            shown = ", ".join(names) if len(names) <= 12 else ", ".join(names[:12]) + ", …"
            self.layout_preview.config(text=f"→ Each row fills: {shown}  ({len(cols)} items, then new row)",
                                       foreground="#1a5fb4")
        except ValueError as e:
            self.layout_preview.config(text=f"⚠ {e}", foreground="red")

    def get_target_columns(self):
        if self.col_mode_var.get() == "custom":
            return parse_column_spec(self.custom_cols_var.get())
        start = col_to_index(self.start_col_var.get())
        try:
            count = int(self.items_per_row_var.get())
        except ValueError:
            raise ValueError("Items per row must be a whole number")
        if count < 1:
            raise ValueError("Items per row must be at least 1")
        return list(range(start, start + count))

    def set_config_locked(self, locked):
        for w in self.config_widgets:
            w.config(state="disabled" if locked else "normal")
        self.spreadsheet_combo.config(state="disabled" if locked else "readonly")
        for w in (self.btn_connect, self.btn_add_sheet, self.btn_remove_sheet, self.btn_credentials):
            w.config(state=tk.DISABLED if locked else tk.NORMAL)
        self.tab_combo.config(state="disabled" if locked else "readonly")
        self.btn_refresh.config(state=tk.DISABLED if locked else tk.NORMAL)
        if not locked:
            self.update_layout_state()
            self.update_link_state()

    def current_settings(self):
        try:
            items = int(self.items_per_row_var.get())
        except ValueError:
            items = DEFAULT_SETTINGS["items_per_row"]
        return {
            "spreadsheet": self.spreadsheet_var.get().strip(),
            "spreadsheets": self.spreadsheets,
            "tab": self.tab_var.get() or self.preferred_tab,
            "tabs": self.tabs_by_spreadsheet,
            "start_row": self.row_entry.get().strip(),
            "col_mode": self.col_mode_var.get(),
            "start_col": self.start_col_var.get().strip().upper(),
            "items_per_row": items,
            "custom_cols": self.custom_cols_var.get().strip(),
            "strip_links": self.strip_links_var.get(),
            "link_mode": self.link_mode_var.get(),
        }

    # =========================
    # Spreadsheet list
    # =========================
    def add_spreadsheet(self):
        email = service_account_email()
        hint = f"\n\nThe sheet must be shared (Editor) with:\n{email}" if email else ""
        ref = simpledialog.askstring(
            "Add Spreadsheet",
            "Enter the spreadsheet's exact name, its URL, or its ID:" + hint,
            parent=self.root)
        ref = (ref or "").strip()
        if not ref:
            return
        if ref not in self.spreadsheets:
            self.spreadsheets.append(ref)
            self.spreadsheet_combo.config(values=self.spreadsheets)
        self.spreadsheet_var.set(ref)
        save_settings(self.current_settings())
        self.connect_to_sheets()

    def remove_spreadsheet(self):
        ref = self.spreadsheet_var.get()
        if not ref:
            return
        if not messagebox.askyesno("Remove Spreadsheet", f"Remove '{ref}' from the list?\n"
                                   "(This only removes it from this app, not from Google Drive.)"):
            return
        self.spreadsheets = [x for x in self.spreadsheets if x != ref]
        self.tabs_by_spreadsheet.pop(ref, None)
        self.spreadsheet_combo.config(values=self.spreadsheets)
        self.spreadsheet_var.set(self.spreadsheets[0] if self.spreadsheets else "")
        self.disconnect()
        save_settings(self.current_settings())
        self.log(f"Removed '{ref}' from the list.")

    # =========================
    # Credentials
    # =========================
    def prompt_for_credentials(self, first_run=False):
        dialog = CredentialsDialog(self.root, first_run=first_run)
        if dialog.saved:
            self.client = None  # rebuild the client with the new key on next connect
            self.log(f"🔑 Credentials saved for {service_account_email()}")
        return dialog.saved

    def change_credentials(self):
        if self.prompt_for_credentials():
            self.disconnect()

    def on_spreadsheet_selected(self, event=None):
        self.connect_to_sheets()

    def disconnect(self):
        self.spreadsheet = None
        self.sheet = None
        self.tab_var.set("")
        self.tab_combo.config(values=[], state="disabled")
        self.btn_refresh.config(state=tk.DISABLED)
        self.btn_start.config(state=tk.DISABLED)
        self.btn_connect.config(text="1. Connect")
        self.status_label.config(text="Status: Disconnected", fg="red")

    # =========================
    # Connect to Sheets
    # =========================
    def connect_to_sheets(self):
        name = self.spreadsheet_var.get().strip()
        if not name:
            messagebox.showerror("Error", "Add or select a spreadsheet first")
            return
        if not has_credentials() and not self.prompt_for_credentials(first_run=True):
            return
        self.disconnect()
        # Each spreadsheet remembers its own last-used tab
        self.preferred_tab = self.tabs_by_spreadsheet.get(name, self.preferred_tab)
        try:
            # Modern gspread authentication using google-auth backend
            if self.client is None:
                self.client = gspread.service_account(filename=credentials_path())
            self.spreadsheet = open_spreadsheet(self.client, name)
            name = self.spreadsheet.title

            self.status_label.config(text=f"Status: Connected to '{name}' ✅", fg="green")
            self.btn_connect.config(text="Reconnect")
            self.log(f"Connected to '{name}'.")
            self.refresh_tabs()

        except Exception as e:
            # Print full detailed error trace to terminal
            print("\n--- FULL CONNECTION ERROR TRACEBACK ---")
            traceback.print_exc()
            print("----------------------------------------\n")

            error_details = str(e) or type(e).__name__
            if hasattr(e, 'response') and hasattr(e.response, 'text'):
                error_details = f"{e}\n\nResponse Text:\n{e.response.text[:300]}"

            if isinstance(e, gspread.SpreadsheetNotFound):
                email = service_account_email()
                error_details = ("Spreadsheet not found. Check the name/URL, and make sure the sheet "
                                 f"is shared (Editor) with:\n{email or 'the service account email'}")

            messagebox.showerror("Connection Error", f"Details:\n{error_details}")

    # =========================
    # Sheet / Tab picker
    # =========================
    def refresh_tabs(self):
        if self.spreadsheet is None:
            return
        try:
            titles = [ws.title for ws in self.spreadsheet.worksheets()]
        except Exception as e:
            messagebox.showerror("Error", f"Could not load tabs:\n{e}")
            return

        self.tab_combo.config(values=titles, state="readonly")
        self.btn_refresh.config(state=tk.NORMAL)

        # Keep the current/remembered tab if it still exists
        wanted = self.tab_var.get() or self.preferred_tab
        self.tab_var.set(wanted if wanted in titles else (titles[0] if titles else ""))
        self.log(f"Found {len(titles)} tab(s): {', '.join(titles)}")
        self.on_tab_selected()

    def on_tab_selected(self, event=None):
        title = self.tab_var.get()
        if not title or self.spreadsheet is None:
            return
        try:
            self.sheet = self.spreadsheet.worksheet(title)
            self.preferred_tab = title
            self.tabs_by_spreadsheet[self.spreadsheet_var.get()] = title
            self.btn_start.config(state=tk.NORMAL)
            self.log(f"Target tab: '{title}'")
        except Exception as e:
            self.sheet = None
            self.btn_start.config(state=tk.DISABLED)
            messagebox.showerror("Error", f"Could not open tab '{title}':\n{e}")

    # =========================
    # Start / Stop
    # =========================
    def start(self):
        if self.sheet is None:
            messagebox.showerror("Error", "Select a target tab first")
            return
        try:
            self.row = int(self.row_entry.get())
            if self.row < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Enter valid row number")
            return
        try:
            self.target_cols = self.get_target_columns()
        except ValueError as e:
            messagebox.showerror("Column Setup Error", str(e))
            return

        # Grow the tab if the mapping points past its last column
        try:
            needed = max(self.target_cols) - self.sheet.col_count
            if needed > 0:
                self.sheet.add_cols(needed)
                self.log(f"Added {needed} column(s) to '{self.sheet.title}' to fit the mapping.")
        except Exception as e:
            self.log(f"⚠ Could not check/extend sheet columns: {e}")

        # Snapshot link settings so the worker thread never touches Tk variables
        self.strip_links_enabled = self.strip_links_var.get()
        self.link_mode = self.link_mode_var.get()
        save_settings(self.current_settings())

        self.running = True
        self.skip_requested = False
        self.row_buffer = []
        self.current_position = 0

        self.set_config_locked(True)
        self.btn_start.config(state=tk.DISABLED)
        self.btn_stop.config(state=tk.NORMAL)

        cols = ", ".join(index_to_col(c) for c in self.target_cols)
        link_info = ("off" if not self.strip_links_enabled
                     else "convert to text" if self.link_mode == "text" else "ignore link, stay on cell")
        self.log(f"▶ Started on '{self.sheet.title}' row {self.row} | columns: {cols} | link filter: {link_info}")

        threading.Thread(target=self.auto_paste_loop, daemon=True).start()

    def stop(self):
        self.running = False
        self.set_config_locked(False)
        self.btn_start.config(state=tk.NORMAL if self.sheet else tk.DISABLED)
        self.btn_stop.config(state=tk.DISABLED)
        if self.row_buffer:
            self.log(f"⚠ {len(self.row_buffer)} unsaved value(s) in the incomplete row were discarded.")
        self.row_entry.delete(0, tk.END)
        self.row_entry.insert(0, str(getattr(self, "row", 2)))
        save_settings(self.current_settings())
        self.log("Stopped.")

    def on_close(self):
        self.running = False
        save_settings(self.current_settings())
        try:
            self.mouse_listener.stop()
        except Exception:
            pass
        self.root.destroy()

    # =========================
    # Right Click = Skip Cell
    # =========================
    def on_click(self, x, y, button, pressed):
        if pressed and button == mouse.Button.right and self.running:
            self.skip_requested = True

    # =========================
    # Link filtering
    # =========================
    def filter_value(self, value):
        """Apply the hyperlink filter. Returns (None, note) when the copy should be
        ignored entirely, so the paste position stays on the same cell."""
        if not self.strip_links_enabled:
            return value, None

        html = get_clipboard_html()
        if not contains_link(value, html):
            return value, None

        if self.link_mode == "discard":
            return None, "🔗 Link ignored, staying on same cell"

        cleaned = strip_links(value)
        if cleaned == value:
            # Rich-text link: plain-text clipboard already holds just the label
            return value, "🔗 Link converted to plain text"
        if not cleaned:
            return None, "🔗 Bare URL ignored, staying on same cell"
        return cleaned, "🔗 URL stripped → plain text"

    # =========================
    # Main Loop
    # =========================
    def auto_paste_loop(self):
        self.last_clipboard = pyperclip.paste()
        total = len(self.target_cols)

        while self.running:
            time.sleep(0.3)

            # Handle Skip Cell
            if self.skip_requested:
                self.skip_requested = False
                col = index_to_col(self.target_cols[self.current_position])

                self.row_buffer.append("")  # Empty value
                self.current_position += 1

                self.log(f"⏭ Skipped {col}{self.row} ({self.current_position}/{total})")

                self.check_and_commit_row()
                continue

            try:
                current_clip = pyperclip.paste()
            except Exception:
                continue  # clipboard temporarily locked by another app

            if current_clip != self.last_clipboard:
                self.last_clipboard = current_clip
                col = index_to_col(self.target_cols[self.current_position])

                value, note = self.filter_value(current_clip)
                if note:
                    self.log(note)
                if value is None:
                    continue  # link ignored, keep waiting for this same cell

                self.row_buffer.append(value)
                self.current_position += 1

                self.log(f"✔ {col}{self.row} ← '{preview(value)}' ({self.current_position}/{total})")

                self.check_and_commit_row()

    # =========================
    # Commit Row if Full
    # =========================
    def check_and_commit_row(self):
        if self.current_position >= len(self.target_cols):
            try:
                # raw=True: values are stored exactly as text, never parsed as formulas/links
                self.sheet.batch_update(build_row_updates(self.row, self.target_cols, self.row_buffer), raw=True)
                self.log(f"🚀 Row {self.row} updated!")

            except Exception as e:
                self.log(f"❌ Error: {e}")

            # Reset for next row
            self.row += 1
            self.row_buffer = []
            self.current_position = 0
            self.log(f"➡ Moving to Row {self.row}")


# =============================
# Run App
# =============================
if __name__ == "__main__":
    root = tk.Tk()
    app = AutoPasterApp(root)
    root.mainloop()

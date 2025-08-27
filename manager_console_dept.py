# -*- coding: utf-8 -*-
import os
import sys
import socket
import re
import configparser
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess

APP_TITLE = "Operator Department Updater — Manager Console"
ACCENT = "#E8AE1C"  # gold accent

# ----------------------- Paths & config -----------------------
def app_dir():
    return os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))

CONFIG_PATH = os.path.join(app_dir(), "config.ini")

def FOLDERS(base):
    return {
        "emails": os.path.join(base, "emails"),
        "logs": os.path.join(base, "program_logs"),
        "data_loader": os.path.join(base, "DataLoad_21.1.x"),
    }

# ----------------------- Config I/O -----------------------
def load_config():
    cfg = configparser.ConfigParser()
    if not os.path.exists(CONFIG_PATH):
        cfg["EMAIL"] = {"from_address": "no-reply@northbay.ca", "recipients": ""}
        cfg["PATHS"] = {"base_dir": r"***************local address****************"}
        cfg["SERVER"] = {"host": "*********", "port": "2000"}
        cfg["UPLOAD"] = {"fadataloader_user": "************", "fadataloader_pass": ""}
        cfg["RETENTION"] = {"days": "30"}
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            cfg.write(f)
    else:
        cfg.read(CONFIG_PATH, encoding="utf-8")
        if not cfg.has_section("RETENTION"):
            cfg.add_section("RETENTION")
            cfg.set("RETENTION", "days", cfg.get("RETENTION", "days", fallback="30"))
    return cfg

def normalize_recipients(text):
    parts = [p.strip() for p in re.split(r"[;,\n]", text) if p.strip()]
    return ", ".join(parts)

def save_config(values):
    cfg = configparser.ConfigParser()
    cfg["EMAIL"] = {
        "from_address": values["from_address"].strip(),
        "recipients": normalize_recipients(", ".join(values["recipients_list"])),
    }
    cfg["PATHS"] = {"base_dir": values["base_dir"].strip()}
    cfg["SERVER"] = {"host": values["host"].strip(), "port": str(values["port"]).strip()}
    cfg["UPLOAD"] = {"fadataloader_user": values["fa_user"].strip(), "fadataloader_pass": values["fa_pass"]}
    cfg["RETENTION"] = {"days": str(values["retain_days"]).strip()}
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        cfg.write(f)

# ----------------------- Utilities -----------------------
def test_server(host, port, timeout=3):
    try:
        with socket.create_connection((host, int(port)), timeout=timeout):
            return True, None
    except Exception as e:
        return False, str(e)

def open_path(p):
    if os.path.isdir(p) or os.path.isfile(p):
        try:
            if sys.platform.startswith("win"):
                os.startfile(p)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", p])
            else:
                subprocess.Popen(["xdg-open", p])
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Failed to open: {p}\n{e}")
    else:
        messagebox.showwarning(APP_TITLE, f"Not found: {p}")

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

# ----------------------- Widgets -----------------------
class EmailList(ttk.Frame):
    """Recipients list with an 'Add Email' button, matching the sample layout.
       No pink (invalid) on load; only mark invalid when non-empty and malformed.
    """
    def __init__(self, parent, emails=None):
        super().__init__(parent)
        self.rows = []

        bar = ttk.Frame(self)
        bar.pack(fill="x", pady=(0,6))
        ttk.Button(bar, text="Add Email", style="Accent.TButton", command=self.add_row).pack(side="left")

        self.list_frame = ttk.Frame(self)
        self.list_frame.pack(fill="x")

        if emails:
            for e in emails:
                self.add_row(e)
        else:
            self.add_row("")

    def add_row(self, value=""):
        row = ttk.Frame(self.list_frame)
        row.pack(fill="x", pady=6)
        var = tk.StringVar(value=value)
        entry = ttk.Entry(row, textvariable=var)
        entry.pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="×", width=3, command=lambda r=row: self._remove_row(r)).pack(side="left", padx=6)

        def on_change(*_):
            v = var.get().strip()
            if v and not EMAIL_RE.match(v):
                entry.configure(style="Invalid.TEntry")
            else:
                entry.configure(style="TEntry")
        var.trace_add("write", on_change)
        on_change()

        self.rows.append((row, var, entry))

    def _remove_row(self, row):
        for i, (r, var, entry) in enumerate(self.rows):
            if r is row:
                r.destroy()
                del self.rows[i]
                break
        if not self.rows:
            self.add_row("")

    def get_emails(self):
        emails = []
        for _, var, _ in self.rows:
            v = var.get().strip()
            if v:
                emails.append(v)
        return emails

# ----------------------- App -----------------------
class ManagerConsole(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.pack(fill="both", expand=True)
        master.title(APP_TITLE)
        master.minsize(760, 520)

        self.cfg = load_config()
        self._init_styles(master)
        self.vars = self._build_vars()

        header = ttk.Frame(self)
        header.pack(fill="x", padx=10, pady=(10,0))
        ttk.Label(header, text=APP_TITLE, style="Header.TLabel").pack(side="left")

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=8, pady=8)
        nb.add(self._email_tab(nb), text="Email")
        nb.add(self._paths_tab(nb), text="Paths")
        nb.add(self._server_tab(nb), text="Server")
        nb.add(self._upload_tab(nb), text="FA Data Loader")
        nb.add(self._retention_tab(nb), text="Retention")

        btns = ttk.Frame(self)
        btns.pack(fill="x", padx=8, pady=(0,10))
        ttk.Button(btns, text="Save", style="Accent.TButton", command=self.on_save).pack(side="right", padx=4)
        ttk.Button(btns, text="Reload", command=self.on_reload).pack(side="right", padx=4)
        ttk.Button(btns, text="Open config.ini", command=lambda: open_path(CONFIG_PATH)).pack(side="left")

    # ----- styles & vars -----
    def _init_styles(self, root):
        style = ttk.Style(root)
        base_theme = "clam" if "clam" in style.theme_names() else style.theme_use()
        style.theme_use(base_theme)
        style.configure("TFrame", padding=6)
        style.configure("TButton", padding=(10,6))
        style.configure("TLabel", padding=2)
        style.configure("Header.TLabel", font=("Segoe UI", 13, "bold"))
        style.configure("Accent.TButton", background=ACCENT, foreground="black")
        style.map("Accent.TButton", background=[("active", ACCENT)], foreground=[("disabled", "#777")])
        style.configure("Invalid.TEntry", fieldbackground="#ffecec")

    def _build_vars(self):
        v = {
            "from_address": tk.StringVar(value=self.cfg.get("EMAIL", "from_address", fallback="no-reply@northbay.ca")),
            "recipients": tk.StringVar(value=self.cfg.get("EMAIL", "recipients", fallback="")),
            "base_dir": tk.StringVar(value=self.cfg.get("PATHS", "base_dir", fallback=r"\\v-arisfleet\arisdata\NewDepartmentUpdater")),
            "host": tk.StringVar(value=self.cfg.get("SERVER", "host", fallback="v-fleetfocustest")),
            "port": tk.StringVar(value=self.cfg.get("SERVER", "port", fallback="2000")),
            "fa_user": tk.StringVar(value=self.cfg.get("UPLOAD", "fadataloader_user", fallback="SYSADMIN-ARIS")),
            "fa_pass": tk.StringVar(value=self.cfg.get("UPLOAD", "fadataloader_pass", fallback="")),
            "retain_days": tk.StringVar(value=self.cfg.get("RETENTION", "days", fallback="30")),
        }
        return v

    # ----- tabs -----
    def _email_tab(self, parent):
        f = ttk.Frame(parent)
        # From address top
        ttk.Label(f, text="From address:").pack(anchor="w")
        ttk.Entry(f, textvariable=self.vars["from_address"]).pack(fill="x")
        # Recipients list (no label; button lives inside widget)
        emails = [e.strip() for e in re.split(r"[;,]", self.vars["recipients"].get()) if e.strip()]
        self.email_list = EmailList(f, emails=emails)
        self.email_list.pack(fill="x", pady=(8,0))
        self._email_tab_parent = f
        return f

    def _paths_tab(self, parent):
        f = ttk.Frame(parent)
        left = ttk.Frame(f); left.pack(side="left", anchor="n", padx=(0,10))
        right = ttk.Frame(f); right.pack(side="left", fill="both", expand=True)

        ttk.Button(left, text="Open Email Reports",
                   command=lambda: open_path(FOLDERS(self.vars["base_dir"].get())["emails"])
        ).pack(fill="x")
        ttk.Button(left, text="Open Recorded Logs",
                   command=lambda: open_path(FOLDERS(self.vars["base_dir"].get())["logs"])
        ).pack(fill="x", pady=(6,0))

        ttk.Label(right, text="Base directory:").pack(anchor="w")
        row = ttk.Frame(right); row.pack(fill="x")
        ttk.Entry(row, textvariable=self.vars["base_dir"]).pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="Browse…", command=self.pick_base_dir).pack(side="left", padx=6)
        return f

    def _server_tab(self, parent):
        f = ttk.Frame(parent)
        f.columnconfigure(1, weight=1)
        ttk.Label(f, text="AssetWorks host:").grid(row=0, column=0, sticky="w")
        ttk.Entry(f, textvariable=self.vars["host"]).grid(row=0, column=1, sticky="we")
        ttk.Label(f, text="Port:").grid(row=1, column=0, sticky="w")
        ttk.Entry(f, textvariable=self.vars["port"], width=10).grid(row=1, column=1, sticky="w")
        ttk.Button(f, text="Test Connection", command=self.on_test_server).grid(row=2, column=0, pady=6)
        return f

    def _upload_tab(self, parent):
        f = ttk.Frame(parent)
        f.columnconfigure(1, weight=1)
        ttk.Label(f, text="FA Data Loader user:").grid(row=0, column=0, sticky="w")
        ttk.Entry(f, textvariable=self.vars["fa_user"]).grid(row=0, column=1, sticky="we")
        ttk.Label(f, text="Password:").grid(row=1, column=0, sticky="w")
        pw = ttk.Entry(f, textvariable=self.vars["fa_pass"], show="*")
        pw.grid(row=1, column=1, sticky="we")
        show_var = tk.BooleanVar(value=False)
        def toggle():
            pw.configure(show='' if show_var.get() else '*')
        ttk.Checkbutton(f, text="Show", variable=show_var, command=toggle).grid(row=1, column=2, sticky="w")
        return f

    def _retention_tab(self, parent):
        f = ttk.Frame(parent)
        f.columnconfigure(1, weight=1)
        ttk.Label(f, text="Keep generated files for (days):").grid(row=0, column=0, sticky="w")
        ttk.Spinbox(f, from_=0, to=365, textvariable=self.vars["retain_days"], width=6).grid(row=0, column=1, sticky="w")
        ttk.Label(f, text="(Applies to emails, logs, and input copies)").grid(row=1, column=0, columnspan=3, sticky="w", pady=(6,0))
        return f

    # ----- actions -----
    def on_test_server(self):
        ok, err = test_server(self.vars["host"].get(), self.vars["port"].get())
        if ok:
            messagebox.showinfo(APP_TITLE, "Connection OK – server is reachable.")
        else:
            messagebox.showwarning(APP_TITLE, f"Connection failed.\n{self.vars['host'].get()}:{self.vars['port'].get()}\n\n{err}")

    def pick_base_dir(self):
        cur = self.vars["base_dir"].get()
        path = filedialog.askdirectory(initialdir=cur or app_dir(), title="Select Base Directory")
        if path:
            self.vars["base_dir"].set(path)

    def on_save(self):
        try:
            int(self.vars["port"].get())
            int(self.vars["retain_days"].get())
        except ValueError:
            messagebox.showerror(APP_TITLE, "Port and Retention must be integers.")
            return

        emails = self.email_list.get_emails()
        bad = [e for e in emails if not EMAIL_RE.match(e)]
        if bad:
            messagebox.showerror(APP_TITLE, "Invalid email(s):\n- " + "\n- ".join(bad))
            return

        values = {k: v.get() for k, v in self.vars.items()}
        values["recipients_list"] = emails

        try:
            save_config(values)
            messagebox.showinfo(APP_TITLE, "Saved to config.ini.")
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Failed to save config: {e}")

    def on_reload(self):
        self.cfg = load_config()
        self.vars["from_address"].set(self.cfg.get("EMAIL", "from_address", fallback=self.vars["from_address"].get()))
        self.vars["recipients"].set(self.cfg.get("EMAIL", "recipients", fallback=self.vars["recipients"].get()))
        self.vars["base_dir"].set(self.cfg.get("PATHS", "base_dir", fallback=self.vars["base_dir"].get()))
        self.vars["host"].set(self.cfg.get("SERVER", "host", fallback=self.vars["host"].get()))
        self.vars["port"].set(self.cfg.get("SERVER", "port", fallback=self.vars["port"].get()))
        self.vars["fa_user"].set(self.cfg.get("UPLOAD", "fadataloader_user", fallback=self.vars["fa_user"].get()))
        self.vars["fa_pass"].set(self.cfg.get("UPLOAD", "fadataloader_pass", fallback=self.vars["fa_pass"].get()))
        self.vars["retain_days"].set(self.cfg.get("RETENTION", "days", fallback=self.vars["retain_days"].get()))
        # rebuild recipients list
        if hasattr(self, "email_list"):
            self.email_list.destroy()
        emails = [e.strip() for e in re.split(r"[;,]", self.vars["recipients"].get()) if e.strip()]
        self.email_list = EmailList(self._email_tab_parent, emails=emails)
        self.email_list.pack(fill="x", pady=(8,0))

# ----------------------- main -----------------------
def main():
    root = tk.Tk()
    try:
        if sys.platform.startswith("win"):
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    ManagerConsole(root)
    root.mainloop()

if __name__ == "__main__":
    main()

from __future__ import annotations

"""Windows desktop launcher for the private/offline research stack."""

import hashlib
import json
import os
from pathlib import Path
import queue
import re
import socket
import subprocess
import sys
import threading
import time
import tkinter as tk
from tkinter import messagebox, ttk
import webbrowser

ROOT = Path(__file__).resolve().parent
IR = ROOT / "institutional_research"
STATE_FILE = ROOT / ".desktop_launcher_state.json"
PORTFOLIO_URL = "http://localhost:8501"
TICKER_RE = re.compile(r"^[A-Z0-9.\-]{1,10}$")


def python_executable() -> str:
    exe = Path(sys.executable)
    if exe.name.lower() == "pythonw.exe":
        candidate = exe.with_name("python.exe")
        if candidate.exists():
            return str(candidate)
    return str(exe)


def normalize_ticker(raw: str) -> str:
    ticker = str(raw or "").strip().upper()
    if not TICKER_RE.fullmatch(ticker):
        raise ValueError("Enter a ticker such as GOOGL, MSFT, JPM, TSM, or SIE.DE.")
    return ticker


def requirement_fingerprint(paths: list[Path] | None = None) -> str:
    paths = paths or [ROOT / "requirements.txt", IR / "requirements.txt"]
    h = hashlib.sha256()
    for path in paths:
        try:
            name = str(path.relative_to(ROOT))
        except ValueError:
            name = str(path)
        h.update(name.encode("utf-8"))
        if path.exists():
            h.update(path.read_bytes())
    return h.hexdigest()


def portfolio_server_running(host: str = "127.0.0.1", port: int = 8501) -> bool:
    try:
        with socket.create_connection((host, port), timeout=0.35):
            return True
    except OSError:
        return False


def latest_workbook(ticker: str) -> Path | None:
    candidates = sorted(
        (ROOT / "updated_models").glob(f"{ticker}_Equity_Research_*.xlsx"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return candidates[0] if candidates else None


class ResearchHub(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Antzaz Research Hub")
        self.geometry("920x680")
        self.minsize(780, 580)
        self.configure(bg="#0b1220")
        self.log_queue: queue.Queue[str] = queue.Queue()
        self.busy = False
        self.task_buttons: list[ttk.Button] = []
        self._build_styles()
        self._build_ui()
        self.after(100, self._drain_log_queue)
        self.after(250, lambda: self.run_task("Checking GitHub for the latest version", self._update_only))

    def _build_styles(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Root.TFrame", background="#0b1220")
        style.configure("Card.TFrame", background="#121b2e")
        style.configure("Title.TLabel", background="#0b1220", foreground="#f3f7ff", font=("Segoe UI", 21, "bold"))
        style.configure("Sub.TLabel", background="#0b1220", foreground="#9eabc0", font=("Segoe UI", 10))
        style.configure("CardTitle.TLabel", background="#121b2e", foreground="#f3f7ff", font=("Segoe UI", 12, "bold"))
        style.configure("CardText.TLabel", background="#121b2e", foreground="#9eabc0", font=("Segoe UI", 9))
        style.configure("Status.TLabel", background="#121b2e", foreground="#8bd3a8", font=("Segoe UI", 9, "bold"))
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), padding=(14, 10))
        style.configure("Normal.TButton", font=("Segoe UI", 9), padding=(12, 8))
        style.configure("Ticker.TEntry", fieldbackground="#f5e6a8", foreground="#1b1b1b", padding=8)

    def _build_ui(self):
        outer = ttk.Frame(self, style="Root.TFrame", padding=22)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="Antzaz Research Hub", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            outer,
            text="Private/offline launcher • automatically updates from GitHub main before running research",
            style="Sub.TLabel",
        ).pack(anchor="w", pady=(2, 16))

        top = ttk.Frame(outer, style="Root.TFrame")
        top.pack(fill="x")

        portfolio = ttk.Frame(top, style="Card.TFrame", padding=18)
        portfolio.pack(side="left", fill="both", expand=True, padx=(0, 8))
        ttk.Label(portfolio, text="Portfolio Research", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(
            portfolio,
            text="Refresh institutional analytics, then open the latest local portfolio dashboard.",
            style="CardText.TLabel",
            wraplength=360,
        ).pack(anchor="w", pady=(4, 12))
        b = ttk.Button(portfolio, text="Open Portfolio Dashboard", style="Accent.TButton", command=self.open_portfolio)
        b.pack(fill="x")
        self.task_buttons.append(b)

        equity = ttk.Frame(top, style="Card.TFrame", padding=18)
        equity.pack(side="left", fill="both", expand=True, padx=(8, 0))
        ttk.Label(equity, text="Equity Research", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(
            equity,
            text="Full Research runs the guarded workbook plus the normal agent/ML research pipeline.",
            style="CardText.TLabel",
            wraplength=360,
        ).pack(anchor="w", pady=(4, 9))

        ticker_row = ttk.Frame(equity, style="Card.TFrame")
        ticker_row.pack(fill="x", pady=(0, 8))
        ttk.Label(ticker_row, text="Ticker", style="CardText.TLabel").pack(side="left", padx=(0, 8))
        self.ticker = tk.StringVar(value="GOOGL")
        entry = ttk.Entry(ticker_row, textvariable=self.ticker, width=14, style="Ticker.TEntry")
        entry.pack(side="left", fill="x", expand=True)
        entry.bind("<Return>", lambda _e: self.run_full_equity())

        row = ttk.Frame(equity, style="Card.TFrame")
        row.pack(fill="x")
        b1 = ttk.Button(row, text="Run Full Research", style="Accent.TButton", command=self.run_full_equity)
        b1.pack(side="left", fill="x", expand=True, padx=(0, 5))
        b2 = ttk.Button(row, text="Workbook Only", style="Normal.TButton", command=self.run_workbook_only)
        b2.pack(side="left", fill="x", expand=True, padx=(5, 0))
        self.task_buttons.extend([b1, b2])

        controls = ttk.Frame(outer, style="Root.TFrame")
        controls.pack(fill="x", pady=(14, 10))
        refresh = ttk.Button(
            controls, text="Update from GitHub", style="Normal.TButton",
            command=lambda: self.run_task("Updating project", self._update_only),
        )
        refresh.pack(side="left")
        ttk.Button(controls, text="Open Project Folder", style="Normal.TButton", command=lambda: os.startfile(ROOT)).pack(side="left", padx=8)
        ttk.Button(controls, text="Open Latest Workbook", style="Normal.TButton", command=self.open_latest_workbook).pack(side="left")
        self.task_buttons.append(refresh)

        status_card = ttk.Frame(outer, style="Card.TFrame", padding=12)
        status_card.pack(fill="x", pady=(0, 10))
        self.status = tk.StringVar(value="Starting…")
        ttk.Label(status_card, textvariable=self.status, style="Status.TLabel").pack(anchor="w")

        log_frame = ttk.Frame(outer, style="Card.TFrame", padding=10)
        log_frame.pack(fill="both", expand=True)
        self.log = tk.Text(
            log_frame, bg="#08101d", fg="#dbe7f5", insertbackground="#ffffff",
            relief="flat", font=("Cascadia Mono", 9), wrap="word", height=18,
        )
        scroll = ttk.Scrollbar(log_frame, command=self.log.yview)
        self.log.configure(yscrollcommand=scroll.set)
        self.log.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self._write_log("Research Hub started.\n")

    def _write_log(self, text: str):
        self.log.insert("end", text)
        self.log.see("end")

    def _drain_log_queue(self):
        try:
            while True:
                self._write_log(self.log_queue.get_nowait())
        except queue.Empty:
            pass
        self.after(100, self._drain_log_queue)

    def emit(self, text: str):
        self.log_queue.put(text if text.endswith("\n") else text + "\n")

    def set_busy(self, busy: bool, status: str):
        self.busy = busy
        self.status.set(status)
        state = "disabled" if busy else "normal"
        for button in self.task_buttons:
            button.configure(state=state)

    def run_task(self, label: str, fn):
        if self.busy:
            messagebox.showinfo("Research Hub", "A task is already running.")
            return
        self.set_busy(True, label + "…")
        self.emit(f"\n=== {label} ===")

        def worker():
            ok = False
            try:
                fn()
                ok = True
            except Exception as exc:
                message = str(exc)
                self.emit(f"ERROR: {message}")
                self.after(0, lambda msg=message: messagebox.showerror("Research Hub", msg))
            finally:
                self.after(0, lambda: self.set_busy(False, "Ready" if ok else "Task failed — see log"))

        threading.Thread(target=worker, daemon=True).start()

    def _run(self, args: list[str], cwd: Path = ROOT, check: bool = True):
        display = " ".join(str(x) for x in args)
        self.emit(f"> {display}")
        proc = subprocess.Popen(
            args,
            cwd=str(cwd),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        assert proc.stdout is not None
        for line in proc.stdout:
            self.emit(line.rstrip())
        code = proc.wait()
        if check and code != 0:
            raise RuntimeError(f"Command failed with exit code {code}: {display}")
        return code

    def _git_update(self):
        current = subprocess.run(
            ["git", "branch", "--show-current"],
            cwd=ROOT,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        if current.returncode != 0:
            raise RuntimeError("This folder is not a usable Git repository.")
        branch = current.stdout.strip()
        if branch != "main":
            self.emit(f"Current branch is {branch or '(detached)'}. Switching to main…")
            self._run(["git", "switch", "main"])
        self._run(["git", "pull", "--ff-only", "origin", "main"])
        sha = subprocess.run(
            ["git", "rev-parse", "--short", "HEAD"],
            cwd=ROOT,
            capture_output=True,
            text=True,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        ).stdout.strip()
        self.emit(f"Latest GitHub version ready: {sha}")

    def _sync_dependencies(self):
        fingerprint = requirement_fingerprint()
        prior = {}
        if STATE_FILE.exists():
            try:
                prior = json.loads(STATE_FILE.read_text(encoding="utf-8"))
            except Exception:
                prior = {}
        if prior.get("requirements_sha256") == fingerprint:
            self.emit("Dependencies already match the current requirements.")
            return
        self.emit("Requirements changed (or first launch). Updating Python dependencies…")
        py = python_executable()
        self._run([py, "-m", "pip", "install", "-r", str(ROOT / "requirements.txt")])
        ir_req = IR / "requirements.txt"
        if ir_req.exists():
            self._run([py, "-m", "pip", "install", "-r", str(ir_req)])
        STATE_FILE.write_text(json.dumps({"requirements_sha256": fingerprint}, indent=2), encoding="utf-8")
        self.emit("Dependencies are current.")

    def _prepare_latest(self):
        self._git_update()
        self._sync_dependencies()

    def _update_only(self):
        self._prepare_latest()

    def open_portfolio(self):
        self.run_task("Opening latest portfolio dashboard", self._portfolio_task)

    def _portfolio_task(self):
        self._prepare_latest()
        py = python_executable()
        self.emit("Refreshing institutional portfolio research…")
        self._run([py, "run_research.py"], cwd=IR)
        if portfolio_server_running():
            self.emit("Portfolio server is already running. Opening it in your browser.")
            webbrowser.open(PORTFOLIO_URL)
            return
        self.emit("Starting Streamlit portfolio dashboard…")
        subprocess.Popen(
            [py, "-m", "streamlit", "run", "app.py"],
            cwd=str(IR),
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        for _ in range(40):
            if portfolio_server_running():
                webbrowser.open(PORTFOLIO_URL)
                self.emit("Portfolio dashboard opened.")
                return
            time.sleep(0.25)
        webbrowser.open(PORTFOLIO_URL)
        self.emit("Streamlit was started. The browser may need a few more seconds to connect.")

    def _get_ticker(self) -> str:
        return normalize_ticker(self.ticker.get())

    def run_full_equity(self):
        try:
            ticker = self._get_ticker()
        except ValueError as exc:
            messagebox.showerror("Ticker", str(exc))
            return
        self.run_task(f"Full equity research — {ticker}", lambda: self._equity_task(ticker, full=True))

    def run_workbook_only(self):
        try:
            ticker = self._get_ticker()
        except ValueError as exc:
            messagebox.showerror("Ticker", str(exc))
            return
        self.run_task(f"Guarded workbook — {ticker}", lambda: self._equity_task(ticker, full=False))

    def _equity_task(self, ticker: str, full: bool):
        self._prepare_latest()
        py = python_executable()
        script = "research.py" if full else "commodity_safe_runner.py"
        self._run([py, script, ticker], cwd=ROOT)
        wb = latest_workbook(ticker)
        if wb:
            self.emit(f"Opening workbook: {wb.name}")
            os.startfile(wb)
        else:
            self.emit("Research completed, but no matching workbook was found to open automatically.")

    def open_latest_workbook(self):
        try:
            ticker = self._get_ticker()
        except ValueError as exc:
            messagebox.showerror("Ticker", str(exc))
            return
        wb = latest_workbook(ticker)
        if not wb:
            messagebox.showinfo("Latest Workbook", f"No generated workbook found yet for {ticker}.")
            return
        os.startfile(wb)


def main():
    ResearchHub().mainloop()


if __name__ == "__main__":
    main()

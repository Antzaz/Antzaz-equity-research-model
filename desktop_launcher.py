from __future__ import annotations

"""Windows desktop launcher for the private investment-research project."""

from pathlib import Path
import os
import queue
import re
import subprocess
import sys
import threading

ROOT = Path(__file__).resolve().parent
IR = ROOT / "institutional_research"
UPDATED_MODELS = ROOT / "updated_models"


def normalize_ticker(raw: str) -> str:
    ticker = str(raw or "").upper().strip()
    if not re.fullmatch(r"[A-Z0-9.\-]{1,10}", ticker):
        raise ValueError("Enter a ticker such as GOOGL, MSFT, NVDA, TSM, or SIE.DE.")
    return ticker


def git_update_command() -> list[str]:
    return ["git", "pull", "--ff-only", "origin", "main"]


def full_research_command(ticker: str) -> list[str]:
    return [sys.executable, str(ROOT / "research.py"), normalize_ticker(ticker)]


def model_only_command(ticker: str) -> list[str]:
    return [sys.executable, str(ROOT / "commodity_safe_runner.py"), normalize_ticker(ticker)]


def portfolio_dashboard_command() -> list[str]:
    if os.name == "nt":
        return ["cmd.exe", "/c", str(IR / "start_dashboard.bat")]
    return [sys.executable, "-m", "streamlit", "run", str(IR / "app.py")]


def _open_path(path: Path) -> None:
    if os.name == "nt":
        os.startfile(str(path))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


def latest_workbook(ticker: str) -> Path | None:
    ticker = normalize_ticker(ticker)
    files = sorted(
        UPDATED_MODELS.glob(f"{ticker}_Equity_Research_*.xlsx"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return files[0] if files else None


def main() -> None:
    import tkinter as tk
    from tkinter import messagebox, ttk

    class InvestmentLauncher:
        def __init__(self, root: tk.Tk):
            self.root = root
            self.root.title("Antzaz Investment Research")
            self.root.geometry("920x650")
            self.root.minsize(780, 560)
            self.root.configure(bg="#0b1220")

            self.events: queue.Queue = queue.Queue()
            self.current_process: subprocess.Popen | None = None
            self.busy = False

            style = ttk.Style()
            try:
                style.theme_use("clam")
            except Exception:
                pass
            style.configure("TFrame", background="#0b1220")
            style.configure("TLabel", background="#0b1220", foreground="#e8eef8", font=("Segoe UI", 10))
            style.configure("Title.TLabel", background="#0b1220", foreground="#ffffff", font=("Segoe UI Semibold", 20))
            style.configure("Sub.TLabel", background="#0b1220", foreground="#9fb0c8", font=("Segoe UI", 10))
            style.configure("TCheckbutton", background="#0b1220", foreground="#dce6f4")
            style.configure("TButton", font=("Segoe UI Semibold", 10), padding=(12, 9))
            style.configure("Accent.TButton", font=("Segoe UI Semibold", 10), padding=(12, 10))

            outer = ttk.Frame(root, padding=24)
            outer.pack(fill="both", expand=True)

            ttk.Label(outer, text="Antzaz Investment Research", style="Title.TLabel").pack(anchor="w")
            ttk.Label(
                outer,
                text="Private desktop launcher · GitHub stays the source of truth",
                style="Sub.TLabel",
            ).pack(anchor="w", pady=(2, 18))

            ticker_row = ttk.Frame(outer)
            ticker_row.pack(fill="x", pady=(0, 10))
            ttk.Label(ticker_row, text="Equity ticker").pack(side="left")
            self.ticker = tk.StringVar(value="GOOGL")
            entry = ttk.Entry(ticker_row, textvariable=self.ticker, width=18, font=("Segoe UI", 11))
            entry.pack(side="left", padx=(12, 18))
            entry.bind("<Return>", lambda _e: self.run_equity(full=True))

            self.auto_update = tk.BooleanVar(value=True)
            ttk.Checkbutton(
                ticker_row,
                text="Update from GitHub before launch",
                variable=self.auto_update,
            ).pack(side="left")

            buttons = ttk.Frame(outer)
            buttons.pack(fill="x", pady=(2, 14))
            ttk.Button(buttons, text="Update from GitHub", command=self.update_only).grid(row=0, column=0, padx=(0, 8), pady=5, sticky="ew")
            ttk.Button(buttons, text="Run Equity Research", style="Accent.TButton", command=lambda: self.run_equity(full=True)).grid(row=0, column=1, padx=8, pady=5, sticky="ew")
            ttk.Button(buttons, text="Build Model Only", command=lambda: self.run_equity(full=False)).grid(row=0, column=2, padx=8, pady=5, sticky="ew")
            ttk.Button(buttons, text="Open Portfolio Dashboard", style="Accent.TButton", command=self.run_portfolio).grid(row=1, column=0, padx=(0, 8), pady=5, sticky="ew")
            ttk.Button(buttons, text="Open Latest Workbook", command=self.open_latest_workbook).grid(row=1, column=1, padx=8, pady=5, sticky="ew")
            ttk.Button(buttons, text="Open Project Folder", command=lambda: _open_path(ROOT)).grid(row=1, column=2, padx=8, pady=5, sticky="ew")
            for col in range(3):
                buttons.grid_columnconfigure(col, weight=1)

            status_row = ttk.Frame(outer)
            status_row.pack(fill="x", pady=(0, 8))
            self.status = tk.StringVar(value="Ready")
            ttk.Label(status_row, textvariable=self.status, style="Sub.TLabel").pack(side="left")
            ttk.Button(status_row, text="Stop Current Process", command=self.stop_process).pack(side="right")

            self.log = tk.Text(
                outer,
                height=24,
                bg="#07101d",
                fg="#d8e4f3",
                insertbackground="#ffffff",
                relief="flat",
                borderwidth=0,
                font=("Consolas", 9),
                wrap="word",
            )
            self.log.pack(fill="both", expand=True)
            self.log.configure(state="disabled")

            self._write("Launcher ready. The repository is the source of truth.\n")
            self._write(f"Project: {ROOT}\n")
            self.root.after(100, self._drain_events)

        def _write(self, value: str) -> None:
            self.log.configure(state="normal")
            self.log.insert("end", value)
            self.log.see("end")
            self.log.configure(state="disabled")

        def _drain_events(self) -> None:
            try:
                while True:
                    kind, payload = self.events.get_nowait()
                    if kind == "log":
                        self._write(payload)
                    elif kind == "status":
                        self.status.set(payload)
                    elif kind == "done":
                        self.busy = False
                        self.current_process = None
                        self.status.set(payload)
            except queue.Empty:
                pass
            self.root.after(100, self._drain_events)

        def _run_process(self, command: list[str], cwd: Path, label: str) -> int:
            self.events.put(("log", f"\n> {label}\n"))
            self.events.put(("log", "  " + " ".join(command) + "\n"))
            try:
                creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
                self.current_process = subprocess.Popen(
                    command,
                    cwd=str(cwd),
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    bufsize=1,
                    creationflags=creationflags,
                )
                assert self.current_process.stdout is not None
                for line in self.current_process.stdout:
                    self.events.put(("log", line))
                return self.current_process.wait()
            except FileNotFoundError as exc:
                self.events.put(("log", f"ERROR: {exc}\n"))
                return 127
            except Exception as exc:
                self.events.put(("log", f"ERROR: {exc}\n"))
                return 1

        def _task(self, label: str, command: list[str], cwd: Path, update_first: bool) -> None:
            if self.busy:
                messagebox.showinfo("Already running", "Another research process is already running.")
                return
            self.busy = True
            self.status.set(label + "...")
            def worker():
                if update_first:
                    rc = self._run_process(git_update_command(), ROOT, "Updating latest version from GitHub")
                    if rc != 0:
                        self.events.put(("done", "Update failed — existing local version was not changed."))
                        return
                rc = self._run_process(command, cwd, label)
                self.events.put(("done", f"{label} finished." if rc == 0 else f"{label} failed with exit code {rc}."))
            threading.Thread(target=worker, daemon=True).start()

        def update_only(self) -> None:
            self._task("GitHub update", git_update_command(), ROOT, False)

        def run_equity(self, full: bool) -> None:
            try:
                ticker = normalize_ticker(self.ticker.get())
            except ValueError as exc:
                messagebox.showerror("Invalid ticker", str(exc))
                return
            command = full_research_command(ticker) if full else model_only_command(ticker)
            label = ("Full equity research" if full else "Equity model") + f" — {ticker}"
            self._task(label, command, ROOT, bool(self.auto_update.get()))

        def run_portfolio(self) -> None:
            self._task(
                "Portfolio dashboard",
                portfolio_dashboard_command(),
                IR,
                bool(self.auto_update.get()),
            )

        def open_latest_workbook(self) -> None:
            try:
                workbook = latest_workbook(self.ticker.get())
            except ValueError as exc:
                messagebox.showerror("Invalid ticker", str(exc))
                return
            if workbook is None:
                messagebox.showinfo("No workbook yet", "No workbook for this ticker has been generated yet.")
                return
            _open_path(workbook)

        def stop_process(self) -> None:
            proc = self.current_process
            if proc is None or proc.poll() is not None:
                self.status.set("No active process")
                return
            try:
                proc.terminate()
                self.events.put(("log", "\nStop requested.\n"))
                self.events.put(("done", "Process stopped."))
            except Exception as exc:
                messagebox.showerror("Could not stop process", str(exc))

    app = tk.Tk()
    InvestmentLauncher(app)
    app.mainloop()


if __name__ == "__main__":
    main()

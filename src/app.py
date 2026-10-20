"""rumps menu bar application for Excel AI Solver.

The solving work runs on a background thread so the menu bar never blocks.
State machine: IDLE -> RUNNING -> PAUSED/COMPLETED/FAILED.
"""

from __future__ import annotations

import threading
from enum import Enum
from typing import Optional

try:
    import rumps
except Exception:  # pragma: no cover - rumps only present on macOS with deps
    rumps = None  # type: ignore[assignment]

from .ai_client import AIClient
from .config import Config
from .settings_window import open_settings_window
from .job_manager import JobManager, JobState
from .solver import GradualSolver, ImmediateSolver
from . import workbook as wb_mod


class AppState(str, Enum):
    IDLE = "IDLE"
    RUNNING = "RUNNING"
    PAUSED = "PAUSED"
    COMPLETED = "COMPLETED"
    FAILED = "FAILED"


# Menu labels (Estonian).
MENU_SOLVE_NOW = "Lahenda kohe"
MENU_SOLVE_GRADUAL = "Lahenda järk-järgult"
MENU_PAUSE = "Peata"
MENU_RESUME = "Jätka"
MENU_CANCEL = "Tühista"
MENU_STATUS = "Olek: jõude"
MENU_SETTINGS = "Seaded..."
MENU_QUIT = "Välju"


def _friendly_error(exc: Exception) -> str:
    """Map technical exceptions to short, actionable Estonian messages."""
    msg = str(exc)
    low = msg.lower()
    # Anthropic SDK error types (imported lazily to avoid hard dependency here).
    try:
        import anthropic as _ant
        if isinstance(exc, _ant.AuthenticationError):
            return "API võti on vigane. Ava Seaded → kontrolli võtit."
        if isinstance(exc, _ant.RateLimitError):
            return "Liiga palju päringuid. Oota hetk ja proovi uuesti."
        if isinstance(exc, _ant.APIConnectionError):
            return "Ühendus Anthropic API-ga ebaõnnestus. Kontrolli internetiühendust."
        if isinstance(exc, _ant.BadRequestError):
            return f"Vigane API päring: {msg[:160]}"
        if isinstance(exc, _ant.APIStatusError):
            return f"API viga {exc.status_code}: {msg[:120]}"
    except Exception:
        pass
    # Generic keyword matching as fallback.
    if "api key" in low or "api_key" in low or "unavailable" in low or "missing sdk" in low:
        return "API võti puudub või on vigane. Ava Seaded → lisa võti."
    if "budget" in low:
        return f"Eelarve piir ületatud ({msg[:80]}). Tõsta eelarvet Seadetes."
    if "excel" in low or "xlwings" in low:
        return f"Exceli viga: {msg[:120]}"
    if "workbook" in low or "töövihik" in low:
        return f"Töövihiku viga: {msg[:120]}"
    return f"Viga: {msg[:180]}"


def _require_rumps():
    if rumps is None:
        raise RuntimeError("rumps is required to run the menu bar app (macOS only).")


class ExcelAISolverApp:
    def __init__(self):
        _require_rumps()
        self.config = Config()
        self.jobs = JobManager()
        self.state = AppState.IDLE
        self._pause_flag = threading.Event()
        self._cancel_flag = threading.Event()
        self._worker: Optional[threading.Thread] = None
        self._current_job_id: Optional[int] = None

        self.app = rumps.App("Excel AI Solver", title="\U0001F4CA")  # bar chart emoji
        self.status_item = rumps.MenuItem(MENU_STATUS)
        self.app.menu = [
            rumps.MenuItem(MENU_SOLVE_NOW, callback=self.on_solve_now),
            rumps.MenuItem(MENU_SOLVE_GRADUAL, callback=self.on_solve_gradual),
            None,
            rumps.MenuItem(MENU_PAUSE, callback=self.on_pause),
            rumps.MenuItem(MENU_RESUME, callback=self.on_resume),
            rumps.MenuItem(MENU_CANCEL, callback=self.on_cancel),
            None,
            self.status_item,
            None,
            rumps.MenuItem(MENU_SETTINGS, callback=self.on_settings),
            rumps.MenuItem(MENU_QUIT, callback=self.on_quit),
        ]

    # ------------------------------------------------------------------ helpers
    def _set_status(self, text: str) -> None:
        self.status_item.title = f"Olek: {text}"

    def _notify(self, title: str, message: str) -> None:
        try:
            rumps.notification(title, "", message)
        except Exception:
            pass

    def _resolve_target_workbook(self) -> Optional[str]:
        active = wb_mod.detect_active_workbook()
        if active:
            return active
        # tkinter crashes with PyObjC's NSApplication — use NSOpenPanel instead.
        try:
            from AppKit import NSApplication, NSOpenPanel
            panel = NSOpenPanel.openPanel()
            panel.setCanChooseFiles_(True)
            panel.setCanChooseDirectories_(False)
            panel.setAllowsMultipleSelection_(False)
            panel.setTitle_("Vali Exceli töövihik")
            panel.setAllowedFileTypes_(["xlsx", "xls"])
            NSApplication.sharedApplication().activateIgnoringOtherApps_(True)
            if panel.runModal() == 1:  # NSModalResponseOK
                return panel.URL().path()
        except Exception:
            pass
        return None

    def _progress(self, message: str, done: int, total: int) -> None:
        if total:
            self._set_status(f"{message} ({done}/{total})")
        else:
            self._set_status(message)
        if self._current_job_id is not None:
            self.jobs.update_job_progress(self._current_job_id, done, total)

    # ------------------------------------------------------------------ actions
    def on_solve_now(self, _sender) -> None:
        self._start_job(mode="immediate")

    def on_solve_gradual(self, _sender) -> None:
        self._start_job(mode="gradual")

    def _start_job(self, mode: str) -> None:
        if self.state == AppState.RUNNING:
            self._notify("Excel AI Solver", "Töö juba käib.")
            return

        # Preflight: API key must be present before anything else starts.
        if not self.config.get_api_key():
            self._set_status("API võti puudub")
            self._notify(
                "Excel AI Solver",
                "API võti puudub. Klõpsa Seaded ja lisa oma Anthropic API võti.",
            )
            return

        path = self._resolve_target_workbook()
        if not path:
            self._set_status("töövihikut ei leitud")
            self._notify(
                "Excel AI Solver",
                "Aktiivset Exceli töövihikut ei leitud. Ava töövihik Excelis ja proovi uuesti.",
            )
            return

        self._pause_flag.clear()
        self._cancel_flag.clear()
        self.state = AppState.RUNNING
        self._current_job_id = self.jobs.create_job(path, mode)
        self.jobs.resume_job(self._current_job_id)

        self._worker = threading.Thread(
            target=self._run_pipeline, args=(path, mode), daemon=True
        )
        self._worker.start()

    def _run_pipeline(self, path: str, mode: str) -> None:
        try:
            ai = AIClient(self.config)
            if mode == "gradual":
                solver = GradualSolver(ai, self.config, progress=self._progress)
                result = solver.run(path, should_continue=self._should_continue)
            else:
                solver = ImmediateSolver(ai, self.config, progress=self._progress)
                result = solver.run(path)

            if self._cancel_flag.is_set():
                self.state = AppState.PAUSED
                if self._current_job_id is not None:
                    self.jobs.cancel_job(self._current_job_id)
                return

            self.state = AppState.COMPLETED
            if self._current_job_id is not None:
                self.jobs.complete_job(self._current_job_id)
            out = result.get("output_path", "?")
            sheet_report = result.get("sheet_report", {})
            written = sheet_report.get("written", [])
            qa_issues = result.get("qa_issues", [])
            status_msg = f"valmis — {len(written)} lehte lisatud"
            notif_msg = f"Lahendus kirjutatud töövihikusse. Lehed: {', '.join(written[:3])}{'...' if len(written)>3 else ''}"
            if qa_issues:
                notif_msg += f"\n⚠ {len(qa_issues)} QA probleem(i)."
            self._set_status(status_msg)
            self._notify("Excel AI Solver", notif_msg)
        except wb_mod.XlwingsWriteError as exc:
            self.state = AppState.PAUSED
            if self._current_job_id is not None:
                self.jobs.pause_job(self._current_job_id)
            self._set_status("peatatud — Exceli kirjutamisviga")
            self._notify("Excel AI Solver", f"Töö peatatud: {exc}")
        except Exception as exc:
            self.state = AppState.FAILED
            if self._current_job_id is not None:
                self.jobs.fail_job(self._current_job_id)
            msg = _friendly_error(exc)
            self._set_status(f"viga — {msg[:60]}")
            self._notify("Excel AI Solver", msg)

    def _should_continue(self) -> bool:
        if self._cancel_flag.is_set():
            return False
        # Block while paused (but bail if cancelled meanwhile).
        while self._pause_flag.is_set():
            if self._cancel_flag.is_set():
                return False
            self._pause_flag.wait(timeout=0.5)
        return True

    def on_pause(self, _sender) -> None:
        if self.state != AppState.RUNNING:
            return
        self._pause_flag.set()
        self.state = AppState.PAUSED
        if self._current_job_id is not None:
            self.jobs.pause_job(self._current_job_id)
        self._set_status("peatatud")

    def on_resume(self, _sender) -> None:
        if self.state != AppState.PAUSED:
            return
        self._pause_flag.clear()
        self.state = AppState.RUNNING
        if self._current_job_id is not None:
            self.jobs.resume_job(self._current_job_id)
        self._set_status("jätkub")

    def on_cancel(self, _sender) -> None:
        self._cancel_flag.set()
        self._pause_flag.clear()
        if self._current_job_id is not None:
            self.jobs.cancel_job(self._current_job_id)
        self.state = AppState.IDLE
        self._set_status("tühistatud")

    def on_settings(self, _sender) -> None:
        open_settings_window(self.config)

    def on_quit(self, _sender) -> None:
        self.jobs.close()
        rumps.quit_application()

    def run(self) -> None:
        self.app.run()


def main() -> None:
    ExcelAISolverApp().run()


if __name__ == "__main__":
    main()

"""Settings, API-key storage (macOS Keychain), and a small tkinter settings window."""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

try:
    import keyring
except Exception:  # pragma: no cover - keyring may be unavailable in some envs
    keyring = None  # type: ignore[assignment]

KEYRING_SERVICE = "ExcelAISolver"
KEYRING_USERNAME = "anthropic_api_key"

APP_SUPPORT_DIR = Path.home() / "Library" / "Application Support" / "ExcelAISolver"
CONFIG_PATH = APP_SUPPORT_DIR / "config.json"

DEFAULT_MODEL = "claude-sonnet-5-5"

# USD per 1M tokens — verified against docs.anthropic.com/en/docs/about-claude/pricing (Oct 2026).
DEFAULT_PRICING: dict[str, dict[str, float]] = {
    "claude-sonnet-5-5": {"input": 2.0, "output": 10.0},
    "claude-opus-5-5": {"input": 8.0, "output": 40.0},
    "claude-haiku-5-5": {"input": 0.8, "output": 4.0},
    # Legacy models kept so existing config files don't lose their price entry.
    "claude-sonnet-4-6": {"input": 3.0, "output": 15.0},
    "claude-sonnet-4-5": {"input": 3.0, "output": 15.0},
    "claude-opus-4-8": {"input": 5.0, "output": 25.0},
    "claude-haiku-4-5": {"input": 1.0, "output": 5.0},
}

DEFAULT_CONFIG: dict[str, Any] = {
    "model": DEFAULT_MODEL,
    "budget_per_workbook": 1.0,
    "pricing": DEFAULT_PRICING,
    "gradual_interval_seconds": 300,
}


class Config:
    """Loads/saves JSON config and brokers the API key via Keychain + env fallback."""

    def __init__(self, path: Path = CONFIG_PATH):
        self.path = path
        self._data: dict[str, Any] = dict(DEFAULT_CONFIG)
        self.load()

    # ------------------------------------------------------------------ load/save
    def load(self) -> None:
        if self.path.exists():
            try:
                stored = json.loads(self.path.read_text())
                merged = dict(DEFAULT_CONFIG)
                merged.update(stored)
                # Merge pricing so new default models survive an old config file.
                pricing = dict(DEFAULT_PRICING)
                pricing.update(stored.get("pricing", {}))
                merged["pricing"] = pricing
                self._data = merged
            except (json.JSONDecodeError, OSError):
                self._data = dict(DEFAULT_CONFIG)
        else:
            self._data = dict(DEFAULT_CONFIG)

    def save(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.path.write_text(json.dumps(self._data, indent=2))

    # ------------------------------------------------------------------ accessors
    @property
    def model(self) -> str:
        return self._data.get("model", DEFAULT_MODEL)

    @model.setter
    def model(self, value: str) -> None:
        self._data["model"] = value

    @property
    def budget_per_workbook(self) -> float:
        return float(self._data.get("budget_per_workbook", 1.0))

    @budget_per_workbook.setter
    def budget_per_workbook(self, value: float) -> None:
        self._data["budget_per_workbook"] = float(value)

    @property
    def pricing(self) -> dict[str, dict[str, float]]:
        return self._data.get("pricing", DEFAULT_PRICING)

    @property
    def gradual_interval_seconds(self) -> int:
        return int(self._data.get("gradual_interval_seconds", 300))

    def price_for(self, model: str | None = None) -> dict[str, float]:
        model = model or self.model
        return self.pricing.get(model, DEFAULT_PRICING.get(DEFAULT_MODEL, {"input": 3.0, "output": 15.0}))

    # ------------------------------------------------------------------ API key
    def get_api_key(self) -> str | None:
        """Keychain first, then ANTHROPIC_API_KEY environment variable."""
        if keyring is not None:
            try:
                key = keyring.get_password(KEYRING_SERVICE, KEYRING_USERNAME)
                if key:
                    return key
            except Exception:
                pass
        return os.environ.get("ANTHROPIC_API_KEY")

    def set_api_key(self, key: str) -> bool:
        """Store the key in Keychain. Returns True on success."""
        if keyring is None:
            return False
        try:
            keyring.set_password(KEYRING_SERVICE, KEYRING_USERNAME, key)
            return True
        except Exception:
            return False

    def clear_api_key(self) -> None:
        if keyring is None:
            return
        try:
            keyring.delete_password(KEYRING_SERVICE, KEYRING_USERNAME)
        except Exception:
            pass


def open_settings_window(config: Config) -> None:
    """Open a simple tkinter settings window (API key, model, budget)."""
    import tkinter as tk
    from tkinter import messagebox, ttk

    root = tk.Tk()
    root.title("Excel AI Solver - Settings")
    root.geometry("460x300")

    frame = ttk.Frame(root, padding=16)
    frame.pack(fill="both", expand=True)

    ttk.Label(frame, text="Anthropic API Key:").grid(row=0, column=0, sticky="w", pady=6)
    api_var = tk.StringVar(value=config.get_api_key() or "")
    api_entry = ttk.Entry(frame, textvariable=api_var, width=40, show="*")
    api_entry.grid(row=0, column=1, sticky="ew", pady=6)

    ttk.Label(frame, text="Model:").grid(row=1, column=0, sticky="w", pady=6)
    model_var = tk.StringVar(value=config.model)
    model_combo = ttk.Combobox(
        frame, textvariable=model_var, values=list(config.pricing.keys()), width=37
    )
    model_combo.grid(row=1, column=1, sticky="ew", pady=6)

    ttk.Label(frame, text="Budget per workbook ($):").grid(row=2, column=0, sticky="w", pady=6)
    budget_var = tk.StringVar(value=str(config.budget_per_workbook))
    ttk.Entry(frame, textvariable=budget_var, width=40).grid(row=2, column=1, sticky="ew", pady=6)

    frame.columnconfigure(1, weight=1)

    def save() -> None:
        key = api_var.get().strip()
        if key:
            if not config.set_api_key(key):
                messagebox.showwarning(
                    "Keychain",
                    "Could not store the key in Keychain. Set ANTHROPIC_API_KEY instead.",
                )
        config.model = model_var.get().strip() or DEFAULT_MODEL
        try:
            config.budget_per_workbook = float(budget_var.get())
        except ValueError:
            messagebox.showerror("Invalid value", "Budget must be a number.")
            return
        config.save()
        messagebox.showinfo("Saved", "Settings saved.")
        root.destroy()

    btns = ttk.Frame(frame)
    btns.grid(row=3, column=0, columnspan=2, pady=16)
    ttk.Button(btns, text="Save", command=save).pack(side="left", padx=6)
    ttk.Button(btns, text="Cancel", command=root.destroy).pack(side="left", padx=6)

    root.mainloop()

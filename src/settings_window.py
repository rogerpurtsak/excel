"""Native AppKit settings panel for Excel AI Solver.

tkinter cannot be used here: when PyObjC owns NSApplication (which rumps
requires), Tk's internal TkSetMacColor calls -[NSApplication macOSVersion]
which is not a recognized selector on PyObjC's NSApplication subclass,
causing an NSInvalidArgumentException crash on Tk.init().

This module builds the panel entirely with AppKit/PyObjC, which already
owns the run loop, so no secondary event loop is needed.
"""
from __future__ import annotations

import objc
from AppKit import (
    NSApplication,
    NSBackingStoreBuffered,
    NSBezelStyleRounded,
    NSButton,
    NSMakeRect,
    NSMomentaryLightButton,
    NSPanel,
    NSSecureTextField,
    NSTextField,
    NSWindowStyleMaskClosable,
    NSWindowStyleMaskTitled,
)
from Foundation import NSObject

# Keep a reference alive at module level; the GC must not collect the
# controller while its window is still on screen.
_active: _SettingsController | None = None  # type: ignore[name-defined]


class _SettingsController(NSObject):
    """Minimal AppKit controller for the settings panel."""

    # ------------------------------------------------------------------ init
    def initWithConfig_(self, config):
        self = objc.super(_SettingsController, self).init()
        if self is None:
            return None
        self._config = config
        self._window = None
        self._key_field = None
        self._model_field = None
        self._budget_field = None
        self._status_label = None
        return self

    # ------------------------------------------------------------------ helpers
    @staticmethod
    def _static_label(text: str, x: float, y: float, w: float, h: float = 20):
        f = NSTextField.alloc().initWithFrame_(NSMakeRect(x, y, w, h))
        f.setStringValue_(text)
        f.setBezeled_(False)
        f.setDrawsBackground_(False)
        f.setEditable_(False)
        f.setSelectable_(False)
        return f

    # ------------------------------------------------------------------ build
    def buildAndShow(self) -> None:
        W, H = 460, 290
        style = NSWindowStyleMaskTitled | NSWindowStyleMaskClosable
        panel = NSPanel.alloc().initWithContentRect_styleMask_backing_defer_(
            NSMakeRect(0, 0, W, H), style, NSBackingStoreBuffered, False
        )
        panel.setTitle_("Excel AI Solver – Seaded")
        # Prevent the panel from being released when the close button is clicked;
        # we manage lifetime through _active.
        panel.setReleasedWhenClosed_(False)
        self._window = panel

        v = panel.contentView()
        px = 20
        y = H - 48

        # ---- API key (secure) ----
        v.addSubview_(self._static_label("Anthropic API võti:", px, y, 220))
        y -= 28
        self._key_field = NSSecureTextField.alloc().initWithFrame_(
            NSMakeRect(px, y, W - 40, 24)
        )
        existing = self._config.get_api_key() or ""
        self._key_field.setStringValue_(existing)
        self._key_field.setPlaceholderString_("sk-ant-…")
        v.addSubview_(self._key_field)
        y -= 40

        # ---- Model ----
        v.addSubview_(self._static_label("Mudel:", px, y, 220))
        y -= 28
        self._model_field = NSTextField.alloc().initWithFrame_(
            NSMakeRect(px, y, W - 40, 24)
        )
        self._model_field.setStringValue_(self._config.model)
        self._model_field.setPlaceholderString_("claude-sonnet-5-5")
        v.addSubview_(self._model_field)
        y -= 40

        # ---- Budget ----
        v.addSubview_(
            self._static_label("Eelarve töövihiku kohta ($):", px, y, 240)
        )
        y -= 28
        self._budget_field = NSTextField.alloc().initWithFrame_(
            NSMakeRect(px, y, 120, 24)
        )
        self._budget_field.setStringValue_(str(self._config.budget_per_workbook))
        v.addSubview_(self._budget_field)
        y -= 36

        # ---- Status feedback label ----
        self._status_label = self._static_label("", px, y, W - 40)
        v.addSubview_(self._status_label)

        # ---- Buttons ----
        btn_y = 14
        for title, x, action in [
            ("Tühista", W - 190, "cancel:"),
            ("Salvesta", W - 100, "save:"),
        ]:
            btn = NSButton.alloc().initWithFrame_(NSMakeRect(x, btn_y, 82, 28))
            btn.setTitle_(title)
            btn.setBezelStyle_(NSBezelStyleRounded)
            btn.setButtonType_(NSMomentaryLightButton)
            btn.setTarget_(self)
            btn.setAction_(action)
            v.addSubview_(btn)

        # Centre and show — activateIgnoringOtherApps_ is required for
        # LSUIElement apps (no Dock icon) to bring the window to the front.
        panel.center()
        panel.makeKeyAndOrderFront_(None)
        NSApplication.sharedApplication().activateIgnoringOtherApps_(True)

    # ------------------------------------------------------------------ actions
    def save_(self, sender) -> None:
        key = (self._key_field.stringValue() or "").strip()
        model = (self._model_field.stringValue() or "").strip()
        budget_raw = (self._budget_field.stringValue() or "").strip()

        if key:
            if not self._config.set_api_key(key):
                self._status_label.setStringValue_(
                    "⚠ Keychain salvestamine ebaõnnestus. Kasuta ANTHROPIC_API_KEY muutujat."
                )
                return

        if model:
            self._config.model = model

        try:
            self._config.budget_per_workbook = float(budget_raw)
        except ValueError:
            self._status_label.setStringValue_("⚠ Eelarve peab olema arv (nt 1.0).")
            return

        self._config.save()
        self._window.close()

    def cancel_(self, sender) -> None:
        self._window.close()


# ------------------------------------------------------------------ public API

def open_settings_window(config) -> None:
    """Open the native AppKit settings panel.

    Safe to call directly on the main thread from a rumps callback.
    If the window is already open, brings it to the front instead of
    opening a second copy.
    """
    global _active

    if _active is not None and _active._window is not None:
        try:
            if _active._window.isVisible():
                _active._window.makeKeyAndOrderFront_(None)
                NSApplication.sharedApplication().activateIgnoringOtherApps_(True)
                return
        except Exception:
            pass

    _active = _SettingsController.alloc().initWithConfig_(config)
    _active.buildAndShow()

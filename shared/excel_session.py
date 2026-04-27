"""Excel session helpers for fetch scripts.

CapIQ requires SAML+MFA login at the start of a session, and that auth lives
in the *currently running* Excel app. Spawning a fresh xw.App() means no
auth, which means CapIQ formulas return #NAME? and the fetch silently
populates garbage. So we attach to a running instance whenever one exists.
"""
from __future__ import annotations

from pathlib import Path


def get_or_create_app(headless: bool = False):
    """Return (app, owns_app).

    If an Excel instance is already running, attach to the active app and
    return owns_app=False. Otherwise spawn a new app and return owns_app=True.

    Caller MUST only quit the app when owns_app is True, and MUST only close
    workbooks the script itself opened.
    """
    import xlwings as xw

    if xw.apps.count > 0:
        return xw.apps.active, False
    app = xw.App(visible=not headless, add_book=False)
    return app, True


def workbook_already_open(app, path) -> bool:
    """True iff a workbook with this filename is already open in `app`."""
    target = Path(path).resolve().name.lower()
    for wb in app.books:
        try:
            if Path(wb.fullname).name.lower() == target:
                return True
        except Exception:
            # Some books (unsaved) may not have a fullname; compare on name only.
            if (wb.name or "").lower() == target:
                return True
    return False


class AppPrefs:
    """Save and restore screen_updating + display_alerts across a fetch.

    Use as a context manager:

        with AppPrefs(app):
            ...  # screen_updating=False, display_alerts=False inside
        # prior values restored on exit, even if an exception was raised.
    """

    def __init__(self, app):
        self.app = app
        self._prior_screen_updating = None
        self._prior_display_alerts = None

    def __enter__(self):
        try:
            self._prior_screen_updating = self.app.screen_updating
        except Exception:
            self._prior_screen_updating = None
        try:
            self._prior_display_alerts = self.app.display_alerts
        except Exception:
            self._prior_display_alerts = None
        try:
            self.app.screen_updating = False
        except Exception:
            pass
        try:
            self.app.display_alerts = False
        except Exception:
            pass
        return self

    def __exit__(self, exc_type, exc, tb):
        if self._prior_screen_updating is not None:
            try:
                self.app.screen_updating = self._prior_screen_updating
            except Exception:
                pass
        if self._prior_display_alerts is not None:
            try:
                self.app.display_alerts = self._prior_display_alerts
            except Exception:
                pass
        return False  # don't swallow exceptions

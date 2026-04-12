"""Workbook session lifecycle management for Graph API."""

from __future__ import annotations

from .client import GraphClient
from .locator import GraphWorkbookLocator


class WorkbookSession:
    """Manages lazy open / close of a Graph API workbook session.

    The session is opened on the first call to ``ensure_open()`` and closed
    explicitly via ``close()``.  ``persist_changes`` controls whether edits
    are committed to the workbook on close.
    """

    def __init__(
        self,
        client: GraphClient,
        locator: GraphWorkbookLocator,
        *,
        persist_changes: bool = False,
    ) -> None:
        self._client = client
        self._locator = locator
        self._persist_changes = persist_changes
        self._open = False

    @property
    def is_open(self) -> bool:
        return self._open

    def ensure_open(self) -> None:
        """Open a workbook session if not already open."""
        if self._open:
            return
        path = f"{self._locator.item_path}/workbook/createSession"
        resp = self._client.post(path, json={"persistChanges": self._persist_changes})
        session_id = resp.json()["id"]
        self._client.session_id = session_id
        self._open = True

    def reopen(self) -> None:
        """Close (if open) and open a fresh session.

        Used for stale-session recovery when the server has expired a session.
        """
        if self._open:
            # Best-effort close of the stale session
            try:
                self._close_remote()
            except Exception:  # noqa: BLE001
                pass
            self._client.session_id = None
            self._open = False
        self.ensure_open()

    def close(self) -> None:
        """Close the current session (no-op if already closed)."""
        if not self._open:
            return
        try:
            self._close_remote()
        except Exception:  # noqa: BLE001 — best-effort close
            pass
        self._client.session_id = None
        self._open = False

    def _close_remote(self) -> None:
        path = f"{self._locator.item_path}/workbook/closeSession"
        self._client.post(path, json={})

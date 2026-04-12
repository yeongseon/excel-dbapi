"""Parse ``msgraph://`` DSNs into drive/item coordinates."""

from __future__ import annotations

from dataclasses import dataclass
from urllib.parse import urlparse


@dataclass(frozen=True)
class GraphWorkbookLocator:
    """Immutable reference to a workbook on OneDrive/SharePoint."""

    drive_id: str
    item_id: str

    @property
    def item_path(self) -> str:
        """Graph API path segment for this workbook."""
        return f"/drives/{self.drive_id}/items/{self.item_id}"


def parse_msgraph_dsn(dsn: str) -> GraphWorkbookLocator:
    """Parse ``msgraph://drives/{drive_id}/items/{item_id}`` into a locator.

    Raises:
        ValueError: If the DSN is malformed.
    """
    parsed = urlparse(dsn)
    if parsed.scheme != "msgraph":
        raise ValueError(f"Expected 'msgraph' scheme, got {parsed.scheme!r}")

    # netloc + path: urlparse puts "drives" in netloc for msgraph://drives/...
    raw_path = parsed.netloc + parsed.path  # e.g. "drives/abc/items/xyz"
    parts = [p for p in raw_path.split("/") if p]

    if len(parts) != 4 or parts[0] != "drives" or parts[2] != "items":
        raise ValueError(
            f"Invalid msgraph DSN: expected "
            f"'msgraph://drives/{{drive_id}}/items/{{item_id}}', got {dsn!r}"
        )

    return GraphWorkbookLocator(drive_id=parts[1], item_id=parts[3])

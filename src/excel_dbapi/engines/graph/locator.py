"""Parse Graph locator DSNs into drive/item coordinates."""

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
        if self.drive_id == "me":
            return f"/me/drive/items/{self.item_id}"
        return f"/drives/{self.drive_id}/items/{self.item_id}"


def parse_msgraph_dsn(dsn: str) -> GraphWorkbookLocator:
    """Parse Graph DSNs into a workbook locator."""
    parsed = urlparse(dsn)
    scheme = parsed.scheme
    if scheme not in {"msgraph", "sharepoint", "onedrive"}:
        raise ValueError(
            "Expected 'msgraph' scheme. "
            f"Expected one of 'msgraph', 'sharepoint', or 'onedrive' scheme, got {scheme!r}"
        )

    raw_path = parsed.netloc + parsed.path
    parts = [p for p in raw_path.split("/") if p]

    if len(parts) == 4 and parts[0] == "drives" and parts[2] == "items":
        return GraphWorkbookLocator(drive_id=parts[1], item_id=parts[3])

    if (
        len(parts) == 6
        and parts[0] == "sites"
        and parts[2] == "drives"
        and parts[4] == "items"
    ):
        return GraphWorkbookLocator(drive_id=parts[3], item_id=parts[5])

    if (
        len(parts) == 4
        and parts[0] == "me"
        and parts[1] == "drive"
        and parts[2] == "items"
    ):
        return GraphWorkbookLocator(drive_id="me", item_id=parts[3])

    if scheme == "msgraph":
        raise ValueError(
            f"Invalid msgraph DSN: expected 'msgraph://drives/{{drive_id}}/items/{{item_id}}', got {dsn!r}"
        )

    if scheme == "sharepoint":
        raise ValueError(
            f"Invalid sharepoint DSN: expected 'sharepoint://sites/{{site_name}}/drives/{{drive_id}}/items/{{item_id}}', got {dsn!r}"
        )

    raise ValueError(
        f"Invalid onedrive DSN: expected 'onedrive://me/drive/items/{{item_id}}', got {dsn!r}"
    )

import pytest

from excel_dbapi.engines.graph.locator import parse_msgraph_dsn


def test_parse_msgraph_drive_item_locator() -> None:
    locator = parse_msgraph_dsn("msgraph://drives/drv-1/items/itm-1")
    assert locator.drive_id == "drv-1"
    assert locator.item_id == "itm-1"
    assert locator.item_path == "/drives/drv-1/items/itm-1"


def test_parse_sharepoint_site_locator() -> None:
    locator = parse_msgraph_dsn(
        "sharepoint://sites/finance-team/drives/drv-sp/items/itm-sp"
    )
    assert locator.drive_id == "drv-sp"
    assert locator.item_id == "itm-sp"
    assert locator.item_path == "/drives/drv-sp/items/itm-sp"


def test_parse_onedrive_me_locator() -> None:
    locator = parse_msgraph_dsn("onedrive://me/drive/items/itm-me")
    assert locator.drive_id == "me"
    assert locator.item_id == "itm-me"
    assert locator.item_path == "/me/drive/items/itm-me"


@pytest.mark.parametrize(
    "dsn",
    [
        "sharepoint://sites/finance-team/drives/drv-sp/files/itm-sp",
        "sharepoint://sites/finance-team/drives/drv-sp/items",
        "sharepoint://contoso.sharepoint.com/sites/finance/Shared Documents/workbook.xlsx",
        "onedrive://me/drives/items/itm-me",
        "onedrive://Documents/workbook.xlsx",
        "onedrive://users/abc/drive/items/itm-me",
    ],
)
def test_parse_extended_locators_reject_invalid_shapes(dsn: str) -> None:
    with pytest.raises(ValueError):
        _ = parse_msgraph_dsn(dsn)

"""Tests for GraphWorkbookLocator and DSN parsing."""

import pytest

from excel_dbapi.engines.graph.locator import GraphWorkbookLocator, parse_msgraph_dsn


class TestGraphWorkbookLocator:
    def test_item_path(self):
        loc = GraphWorkbookLocator(drive_id="drv-1", item_id="itm-2")
        assert loc.item_path == "/drives/drv-1/items/itm-2"

    def test_item_path_for_onedrive_me(self):
        loc = GraphWorkbookLocator(drive_id="me", item_id="itm-2")
        assert loc.item_path == "/me/drive/items/itm-2"

    def test_frozen(self):
        loc = GraphWorkbookLocator(drive_id="d", item_id="i")
        with pytest.raises(AttributeError):
            setattr(loc, "drive_id", "x")


class TestParseMsgraphDsn:
    def test_valid_dsn(self):
        dsn = "msgraph://drives/abc123/items/xyz789"
        loc = parse_msgraph_dsn(dsn)
        assert loc.drive_id == "abc123"
        assert loc.item_id == "xyz789"

    def test_wrong_scheme(self):
        with pytest.raises(ValueError, match=r"Expected 'msgraph' scheme"):
            parse_msgraph_dsn("https://drives/abc/items/xyz")

    def test_missing_items_segment(self):
        with pytest.raises(ValueError, match="Invalid msgraph DSN"):
            parse_msgraph_dsn("msgraph://drives/abc/files/xyz")

    def test_too_few_segments(self):
        with pytest.raises(ValueError, match="Invalid msgraph DSN"):
            parse_msgraph_dsn("msgraph://drives/abc")

    def test_too_many_segments(self):
        with pytest.raises(ValueError, match="Invalid msgraph DSN"):
            parse_msgraph_dsn("msgraph://drives/abc/items/xyz/extra/stuff")

    def test_trailing_slash(self):
        dsn = "msgraph://drives/d1/items/i1/"
        loc = parse_msgraph_dsn(dsn)
        assert loc.drive_id == "d1"
        assert loc.item_id == "i1"

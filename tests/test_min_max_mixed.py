"""Tests for MIN/MAX on mixed-type columns (issue #90)."""

from excel_dbapi.connection import ExcelConnection


class TestMinMaxMixedTypes:
    """MIN/MAX must not crash on columns with mixed types."""

    def test_min_mixed_int_string(self, tmp_path):
        file = tmp_path / "test.xlsx"
        with ExcelConnection(str(file), autocommit=True, create=True) as conn:
            cur = conn.cursor()
            cur.execute("CREATE TABLE t (val)")
            cur.execute("INSERT INTO t VALUES (10)")
            cur.execute("INSERT INTO t VALUES ('hello')")
            cur.execute("INSERT INTO t VALUES (5)")

            cur.execute("SELECT MIN(val) FROM t")
            result = cur.fetchone()
            # Numeric sorts before strings via _sort_key
            assert result is not None
            assert result[0] == 5

    def test_max_mixed_int_string(self, tmp_path):
        file = tmp_path / "test.xlsx"
        with ExcelConnection(str(file), autocommit=True, create=True) as conn:
            cur = conn.cursor()
            cur.execute("CREATE TABLE t (val)")
            cur.execute("INSERT INTO t VALUES (10)")
            cur.execute("INSERT INTO t VALUES ('hello')")
            cur.execute("INSERT INTO t VALUES (5)")

            cur.execute("SELECT MAX(val) FROM t")
            result = cur.fetchone()
            # Strings sort after numbers via _sort_key
            assert result is not None
            assert result[0] == "hello"

    def test_min_all_nulls(self, tmp_path):
        file = tmp_path / "test.xlsx"
        with ExcelConnection(str(file), autocommit=True, create=True) as conn:
            cur = conn.cursor()
            cur.execute("CREATE TABLE t (val)")
            cur.execute("INSERT INTO t VALUES (NULL)")
            cur.execute("INSERT INTO t VALUES (NULL)")

            cur.execute("SELECT MIN(val) FROM t")
            result = cur.fetchone()
            assert result[0] is None

    def test_min_max_homogeneous_still_works(self, tmp_path):
        """Ensure homogeneous columns still work correctly."""
        file = tmp_path / "test.xlsx"
        with ExcelConnection(str(file), autocommit=True, create=True) as conn:
            cur = conn.cursor()
            cur.execute("CREATE TABLE t (val)")
            cur.execute("INSERT INTO t VALUES (30)")
            cur.execute("INSERT INTO t VALUES (10)")
            cur.execute("INSERT INTO t VALUES (20)")

            cur.execute("SELECT MIN(val), MAX(val) FROM t")
            result = cur.fetchone()
            assert result == (10, 30)

    def test_min_max_strings_only(self, tmp_path):
        file = tmp_path / "test.xlsx"
        with ExcelConnection(str(file), autocommit=True, create=True) as conn:
            cur = conn.cursor()
            cur.execute("CREATE TABLE t (val)")
            cur.execute("INSERT INTO t VALUES ('banana')")
            cur.execute("INSERT INTO t VALUES ('apple')")
            cur.execute("INSERT INTO t VALUES ('cherry')")

            cur.execute("SELECT MIN(val), MAX(val) FROM t")
            result = cur.fetchone()
            assert result == ("apple", "cherry")

"""
tests/test_embed_extractor.py
==============================
Unit tests for embed_extractor.py.

These tests do NOT require a browser or network access; they cover the
pure-Python helpers that can be exercised offline.
"""

import csv
import os
import sys
import tempfile
from unittest.mock import patch

import pytest

# Ensure the project root is on the path so we can import the script directly
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import embed_extractor as ee


# ---------------------------------------------------------------------------
# extract_youtube_id
# ---------------------------------------------------------------------------


class TestExtractYoutubeId:
    def test_standard_watch_url(self):
        assert ee.extract_youtube_id("https://www.youtube.com/watch?v=dQw4w9WgXcQ") == "dQw4w9WgXcQ"

    def test_short_url(self):
        assert ee.extract_youtube_id("https://youtu.be/dQw4w9WgXcQ") == "dQw4w9WgXcQ"

    def test_shorts_url(self):
        assert ee.extract_youtube_id("https://www.youtube.com/shorts/abc12345678") == "abc12345678"

    def test_embed_url(self):
        assert ee.extract_youtube_id("https://www.youtube.com/embed/dQw4w9WgXcQ") == "dQw4w9WgXcQ"

    def test_url_with_extra_params(self):
        assert (
            ee.extract_youtube_id(
                "https://www.youtube.com/watch?v=dQw4w9WgXcQ&t=42s&list=PLxxx"
            )
            == "dQw4w9WgXcQ"
        )

    def test_non_youtube_url_returns_none(self):
        assert ee.extract_youtube_id("https://www.example.com/video/12345") is None

    def test_empty_string_returns_none(self):
        assert ee.extract_youtube_id("") is None


# ---------------------------------------------------------------------------
# build_youtube_embed
# ---------------------------------------------------------------------------


class TestBuildYoutubeEmbed:
    def test_contains_iframe(self):
        html = ee.build_youtube_embed("dQw4w9WgXcQ")
        assert html.startswith("<iframe")
        assert html.endswith("></iframe>")

    def test_contains_video_id(self):
        html = ee.build_youtube_embed("dQw4w9WgXcQ")
        assert "dQw4w9WgXcQ" in html

    def test_contains_embed_src(self):
        html = ee.build_youtube_embed("dQw4w9WgXcQ")
        assert 'src="https://www.youtube.com/embed/dQw4w9WgXcQ"' in html

    def test_allowfullscreen(self):
        html = ee.build_youtube_embed("dQw4w9WgXcQ")
        assert "allowfullscreen" in html


# ---------------------------------------------------------------------------
# _get_embed  (no driver → programmatic fallback)
# ---------------------------------------------------------------------------


class TestGetEmbedNoDriver:
    def test_youtube_url_returns_iframe(self):
        code = ee._get_embed(None, "https://www.youtube.com/watch?v=dQw4w9WgXcQ")
        assert code is not None
        assert "<iframe" in code
        assert "dQw4w9WgXcQ" in code

    def test_non_video_url_returns_none(self):
        code = ee._get_embed(None, "https://www.example.com/page")
        assert code is None


# ---------------------------------------------------------------------------
# process_csv  (end-to-end with a temp file, no browser)
# ---------------------------------------------------------------------------


@pytest.fixture(autouse=True)
def no_browser(monkeypatch):
    """Prevent any test from spawning a real browser / downloading ChromeDriver."""
    monkeypatch.setattr(ee, "_create_driver", lambda headless=True: None)


class TestProcessCsv:
    def _make_csv(self, rows, headers=None):
        """Write rows to a temp CSV and return its path."""
        fd, path = tempfile.mkstemp(suffix=".csv")
        os.close(fd)
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            if headers:
                writer.writerow(headers)
            writer.writerows(rows)
        return path

    def test_embed_code_written_to_new_column(self):
        path = self._make_csv(
            [["https://www.youtube.com/watch?v=dQw4w9WgXcQ"]],
            headers=["Video URL"],
        )
        try:
            # Column 1 = URL, column 2 = embed output
            ee.process_csv(path, "1", "2", headless=True)

            with open(path, newline="", encoding="utf-8-sig") as f:
                rows = list(csv.reader(f))

            assert len(rows) == 2  # header + 1 data row
            assert "<iframe" in rows[1][1]
            assert "dQw4w9WgXcQ" in rows[1][1]
        finally:
            os.unlink(path)

    def test_embed_code_by_header_name(self):
        path = self._make_csv(
            [["https://youtu.be/dQw4w9WgXcQ", ""]],
            headers=["URL", "Embed"],
        )
        try:
            ee.process_csv(path, "URL", "Embed", headless=True)

            with open(path, newline="", encoding="utf-8-sig") as f:
                rows = list(csv.reader(f))

            assert "<iframe" in rows[1][1]
        finally:
            os.unlink(path)

    def test_empty_url_rows_skipped(self):
        path = self._make_csv(
            [[""], ["https://www.youtube.com/watch?v=dQw4w9WgXcQ"]],
            headers=["URL"],
        )
        try:
            ee.process_csv(path, "1", "2", headless=True)

            with open(path, newline="", encoding="utf-8-sig") as f:
                rows = list(csv.reader(f))

            # Row with empty URL should have no embed code
            assert rows[1][1] == ""
            # Row with YouTube URL should have embed code
            assert "<iframe" in rows[2][1]
        finally:
            os.unlink(path)

    def test_non_youtube_url_no_embed(self):
        path = self._make_csv(
            [["https://www.example.com/not-a-video"]],
            headers=["URL"],
        )
        try:
            ee.process_csv(path, "1", "2", headless=True)

            with open(path, newline="", encoding="utf-8-sig") as f:
                rows = list(csv.reader(f))

            assert rows[1][1] == ""
        finally:
            os.unlink(path)


# ---------------------------------------------------------------------------
# process_excel  (end-to-end with a temp .xlsx, no browser)
# ---------------------------------------------------------------------------


class TestProcessExcel:
    def _make_xlsx(self, rows, headers=None):
        import openpyxl

        fd, path = tempfile.mkstemp(suffix=".xlsx")
        os.close(fd)
        wb = openpyxl.Workbook()
        ws = wb.active
        if headers:
            ws.append(headers)
        for row in rows:
            ws.append(row)
        wb.save(path)
        return path

    def test_embed_code_written_to_column(self):
        import openpyxl

        path = self._make_xlsx(
            [["https://www.youtube.com/watch?v=dQw4w9WgXcQ"]],
            headers=["Video URL"],
        )
        try:
            ee.process_excel(path, "1", "2", headless=True)

            wb = openpyxl.load_workbook(path)
            ws = wb.active
            embed = ws.cell(row=2, column=2).value
            assert embed is not None
            assert "<iframe" in embed
            assert "dQw4w9WgXcQ" in embed
        finally:
            os.unlink(path)

    def test_embed_code_by_header_name(self):
        import openpyxl

        path = self._make_xlsx(
            [["https://youtu.be/dQw4w9WgXcQ", ""]],
            headers=["URL", "Embed Code"],
        )
        try:
            ee.process_excel(path, "URL", "Embed Code", headless=True)

            wb = openpyxl.load_workbook(path)
            ws = wb.active
            embed = ws.cell(row=2, column=2).value
            assert embed is not None
            assert "<iframe" in embed
        finally:
            os.unlink(path)

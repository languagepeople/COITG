#!/usr/bin/env python3
"""
embed_extractor.py
==================
Reads a spreadsheet (Excel .xlsx/.xls or CSV), follows URLs found in a
specified column, navigates each URL in a headless browser, clicks the
Share → Embed buttons, and writes the resulting <iframe> HTML into a
configurable output column on the same row.

YouTube is the primary target.  For YouTube URLs a fast programmatic
fallback is also included so the script still works even when browser
automation is blocked or the page layout has changed.

Usage
-----
    python embed_extractor.py <spreadsheet> [options]

Examples
--------
    # URL in column B (2), embed code written to column C (3)
    python embed_extractor.py videos.xlsx --url-col 2 --embed-col 3

    # URL column by header name, output to a named column
    python embed_extractor.py videos.csv --url-col "Video URL" --embed-col "Embed Code"

    # Keep the browser window visible (non-headless) for debugging
    python embed_extractor.py videos.xlsx --no-headless
"""
from __future__ import annotations

import argparse
import csv
import os
import re
import sys
import time
from urllib.parse import urlparse

# ---------------------------------------------------------------------------
# Optional dependency: selenium  (installed via requirements.txt)
# ---------------------------------------------------------------------------
try:
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options as ChromeOptions
    from selenium.webdriver.chrome.service import Service as ChromeService
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.ui import WebDriverWait

    try:
        from webdriver_manager.chrome import ChromeDriverManager

        _WDM_AVAILABLE = True
    except ImportError:
        _WDM_AVAILABLE = False

    _SELENIUM_AVAILABLE = True
except ImportError:
    _SELENIUM_AVAILABLE = False

# ---------------------------------------------------------------------------
# Optional dependency: openpyxl  (installed via requirements.txt)
# ---------------------------------------------------------------------------
try:
    import openpyxl

    _OPENPYXL_AVAILABLE = True
except ImportError:
    _OPENPYXL_AVAILABLE = False


# ---------------------------------------------------------------------------
# YouTube helpers
# ---------------------------------------------------------------------------

_YT_PATTERNS = [
    r"(?:youtube\.com/watch\?(?:.*&)?v=|youtu\.be/)([A-Za-z0-9_-]{11})",
    r"youtube\.com/embed/([A-Za-z0-9_-]{11})",
    r"youtube\.com/shorts/([A-Za-z0-9_-]{11})",
]

_YT_HOSTNAMES = frozenset(
    {"www.youtube.com", "youtube.com", "m.youtube.com", "youtu.be"}
)


def _is_youtube_url(url: str) -> bool:
    """Return True only when *url*'s hostname is a known YouTube domain."""
    try:
        hostname = urlparse(url).hostname or ""
        return hostname in _YT_HOSTNAMES
    except Exception:
        return False


def extract_youtube_id(url: str) -> str | None:
    """Return the 11-character YouTube video ID from *url*, or None."""
    for pattern in _YT_PATTERNS:
        m = re.search(pattern, url)
        if m:
            return m.group(1)
    return None


def build_youtube_embed(video_id: str) -> str:
    """Return the standard YouTube <iframe> embed HTML for *video_id*."""
    return (
        f'<iframe width="560" height="315" '
        f'src="https://www.youtube.com/embed/{video_id}" '
        f'title="YouTube video player" '
        f'frameborder="0" '
        f'allow="accelerometer; autoplay; clipboard-write; encrypted-media; '
        f'gyroscope; picture-in-picture; web-share" '
        f'referrerpolicy="strict-origin-when-cross-origin" '
        f'allowfullscreen></iframe>'
    )


# ---------------------------------------------------------------------------
# Browser automation
# ---------------------------------------------------------------------------


def _create_driver(headless: bool = True) -> "webdriver.Chrome | None":
    """
    Create and return a Chrome WebDriver instance, or **None** if Chrome /
    ChromeDriver is not available in this environment.

    A *None* return value lets the callers fall back to programmatic embed
    code generation (e.g. for YouTube URLs) without crashing.
    """
    if not _SELENIUM_AVAILABLE:
        return None

    options = ChromeOptions()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument(
        "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )

    try:
        if _WDM_AVAILABLE:
            service = ChromeService(ChromeDriverManager().install())
            return webdriver.Chrome(service=service, options=options)
        return webdriver.Chrome(options=options)
    except Exception as exc:
        print(
            f"Warning: could not start Chrome WebDriver ({exc}). "
            "Browser automation is disabled; falling back to programmatic "
            "embed code generation where supported.",
            file=sys.stderr,
        )
        return None


def _dismiss_consent_dialogs(driver: "webdriver.Chrome") -> None:
    """Try to dismiss common cookie/consent overlays."""
    consent_selectors = [
        'button[aria-label*="Accept"]',
        'button[aria-label*="Agree"]',
        "#accept-button",
        ".accept-cookies",
        'form[action*="consent"] button',
    ]
    for sel in consent_selectors:
        try:
            btn = driver.find_element(By.CSS_SELECTOR, sel)
            if btn.is_displayed():
                btn.click()
                time.sleep(0.5)
                return
        except Exception:
            pass


def get_embed_via_browser(driver: "webdriver.Chrome", url: str) -> str | None:
    """
    Navigate to *url* in *driver*, click Share → Embed, and return the
    <iframe> HTML string, or None on failure.

    Currently implements the YouTube share flow.  Other sites can be
    added by extending the dispatch block below.
    """
    try:
        driver.get(url)
        _dismiss_consent_dialogs(driver)

        if _is_youtube_url(url):
            return _youtube_share_embed(driver)

        # Generic fallback: look for a Share button anywhere on the page
        return _generic_share_embed(driver)

    except Exception as exc:
        print(f"  [browser] Error fetching embed for {url}: {exc}", file=sys.stderr)
        return None


def _youtube_share_embed(driver: "webdriver.Chrome") -> str | None:
    """Automate the YouTube Share → Embed dialog and return the embed HTML."""
    wait = WebDriverWait(driver, 20)

    # Wait until the player area is present so buttons are loaded
    try:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#player")))
    except Exception:
        pass  # continue anyway

    # --- Click the Share button ---
    share_btn = None
    share_selectors = [
        (By.CSS_SELECTOR, 'button[aria-label="Share"]'),
        (By.XPATH, '//button[.//yt-formatted-string[text()="Share"]]'),
        (By.XPATH, '//ytd-button-renderer[.//yt-formatted-string[text()="Share"]]//button'),
    ]
    for by, selector in share_selectors:
        try:
            share_btn = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((by, selector))
            )
            break
        except Exception:
            pass

    if share_btn is None:
        print("  [browser] Could not find YouTube Share button.", file=sys.stderr)
        return None

    driver.execute_script("arguments[0].scrollIntoView(true);", share_btn)
    time.sleep(0.5)
    share_btn.click()

    # --- Wait for the share dialog, then click the Embed tab ---
    embed_btn = None
    embed_selectors = [
        (By.XPATH, '//yt-share-panel//button[normalize-space()="Embed"]'),
        (By.XPATH, '//ytd-unified-share-panel-renderer//button[normalize-space()="Embed"]'),
        (By.CSS_SELECTOR, 'tp-yt-paper-tab[aria-label="Embed"]'),
        (By.XPATH, '//tp-yt-paper-tab[.//div[normalize-space()="Embed"]]'),
    ]
    for by, selector in embed_selectors:
        try:
            embed_btn = WebDriverWait(driver, 8).until(
                EC.element_to_be_clickable((by, selector))
            )
            break
        except Exception:
            pass

    if embed_btn is None:
        print("  [browser] Could not find Embed tab in share dialog.", file=sys.stderr)
        return None

    embed_btn.click()
    time.sleep(1)

    # --- Extract embed code from the textarea / code panel ---
    code_selectors = [
        (By.CSS_SELECTOR, "textarea.ytd-embed-code-panel"),
        (By.CSS_SELECTOR, "#embed-code"),
        (By.XPATH, '//yt-share-panel//textarea'),
        (By.XPATH, '//div[@id="embed-code"]//textarea'),
        (By.CSS_SELECTOR, "yt-copy-code-panel textarea"),
    ]
    for by, selector in code_selectors:
        try:
            el = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((by, selector))
            )
            code = el.get_attribute("value") or el.text
            if code and "<iframe" in code:
                return code.strip()
        except Exception:
            pass

    return None


def _generic_share_embed(driver: "webdriver.Chrome") -> str | None:
    """Attempt a generic Share → Embed flow for non-YouTube pages."""
    wait = WebDriverWait(driver, 10)

    # Look for any button whose text contains "Share"
    try:
        share_btn = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, '//button[contains(translate(., "SHARE", "share"), "share")]')
            )
        )
        share_btn.click()
        time.sleep(1)
    except Exception:
        return None

    # Look for an "Embed" option
    try:
        embed_btn = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, '//button[contains(translate(., "EMBED", "embed"), "embed")]')
            )
        )
        embed_btn.click()
        time.sleep(1)
    except Exception:
        return None

    # Return any <iframe> found on the page at this point
    try:
        textarea = driver.find_element(By.CSS_SELECTOR, "textarea")
        code = textarea.get_attribute("value") or textarea.text
        if "<iframe" in code:
            return code.strip()
    except Exception:
        pass

    return None


# ---------------------------------------------------------------------------
# Spreadsheet I/O  –  Excel
# ---------------------------------------------------------------------------


def _col_index(wb_ws, column_ref: str) -> int:
    """
    Resolve *column_ref* to a 1-based column index for an openpyxl worksheet.
    Accepts an integer (as string), a letter like 'B', or a header name.

    Resolution order:
    1. Plain integer  →  used as-is (1-based)
    2. Header name match in row 1  →  that column
    3. Spreadsheet column letter(s)  (A, B, … Z, AA, …)
    """
    if column_ref.isdigit():
        return int(column_ref)

    # Header name – search row 1 first so named columns are never confused
    # with column letters (e.g. "URL" or "Embed" are valid header names).
    for cell in wb_ws[1]:
        if str(cell.value or "").strip().lower() == column_ref.strip().lower():
            return cell.column

    # Single / multi-letter column (A, B, AA …)
    if re.fullmatch(r"[A-Za-z]+", column_ref):
        return openpyxl.utils.column_index_from_string(column_ref.upper())

    raise ValueError(
        f"Cannot find column '{column_ref}' in the spreadsheet. "
        "Specify a column number, letter, or exact header name."
    )


def _ensure_header(ws, col_idx: int, header: str = "Embed Code") -> None:
    """Write a header into row 1 of *col_idx* if the cell is empty."""
    if not ws.cell(row=1, column=col_idx).value:
        ws.cell(row=1, column=col_idx, value=header)


def process_excel(
    path: str,
    url_col: str,
    embed_col: str,
    headless: bool,
) -> None:
    if not _OPENPYXL_AVAILABLE:
        sys.exit("openpyxl is not installed.  Run: pip install -r requirements.txt")

    wb = openpyxl.load_workbook(path)
    ws = wb.active

    url_idx = _col_index(ws, url_col)
    embed_idx = _col_index(ws, embed_col)
    _ensure_header(ws, embed_idx)

    driver = _create_driver(headless)
    max_row = ws.max_row
    try:
        for row_num in range(2, max_row + 1):
            url = ws.cell(row=row_num, column=url_idx).value
            if not url:
                continue
            url = str(url).strip()
            print(f"Row {row_num}: {url}")
            embed = _get_embed(driver, url)
            if embed:
                ws.cell(row=row_num, column=embed_idx, value=embed)
                print(f"  → embed code written ({len(embed)} chars)")
            else:
                print(f"  → no embed code found", file=sys.stderr)
    finally:
        if driver:
            driver.quit()

    wb.save(path)
    print(f"\nSaved: {path}")


# ---------------------------------------------------------------------------
# Spreadsheet I/O  –  CSV
# ---------------------------------------------------------------------------


def _col_index_csv(headers: list[str], column_ref: str) -> int:
    """
    Resolve *column_ref* to a 0-based index for a CSV row list.
    Accepts an integer (1-based), letter, or header name.

    Resolution order:
    1. Plain integer  →  used as 1-based, converted to 0-based
    2. Header name match  →  that index
    3. Spreadsheet column letter(s)  (A, B, … AA, …)
    """
    if column_ref.isdigit():
        return int(column_ref) - 1

    # Header name – check first so that words like "URL" or "Embed" are
    # matched against actual column headers before being treated as letters.
    for i, h in enumerate(headers):
        if h.strip().lower() == column_ref.strip().lower():
            return i

    # Column letter(s) fallback
    if re.fullmatch(r"[A-Za-z]+", column_ref):
        val = 0
        for ch in column_ref.upper():
            val = val * 26 + (ord(ch) - ord("A") + 1)
        return val - 1

    raise ValueError(
        f"Cannot find column '{column_ref}' in the CSV headers. "
        "Specify a column number, letter, or exact header name."
    )


def process_csv(
    path: str,
    url_col: str,
    embed_col: str,
    headless: bool,
) -> None:
    # Read all rows
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        rows = list(reader)

    if not rows:
        print("CSV file is empty.")
        return

    headers = rows[0]
    url_idx = _col_index_csv(headers, url_col)

    # Resolve embed column; add header if needed
    try:
        embed_idx = _col_index_csv(headers, embed_col)
    except ValueError:
        embed_idx = None

    if embed_idx is None or embed_idx >= len(headers):
        # Append a new column
        embed_idx = len(headers)
        headers.append(embed_col if not embed_col.isdigit() else "Embed Code")
        for row in rows[1:]:
            while len(row) <= embed_idx:
                row.append("")

    driver = _create_driver(headless)

    try:
        for i, row in enumerate(rows[1:], start=2):
            while len(row) <= max(url_idx, embed_idx):
                row.append("")
            url = row[url_idx].strip()
            if not url:
                continue
            print(f"Row {i}: {url}")
            embed = _get_embed(driver, url)
            if embed:
                row[embed_idx] = embed
                print(f"  → embed code written ({len(embed)} chars)")
            else:
                print(f"  → no embed code found", file=sys.stderr)
    finally:
        if driver:
            driver.quit()

    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows(rows[1:])

    print(f"\nSaved: {path}")


# ---------------------------------------------------------------------------
# Central embed-fetching logic
# ---------------------------------------------------------------------------


def _get_embed(driver, url: str) -> str | None:
    """
    Return an embed HTML string for *url* using:
    1. Browser automation (if selenium is available and driver is provided)
    2. Programmatic YouTube embed code as a reliable fallback
    """
    # Try browser first
    if driver:
        code = get_embed_via_browser(driver, url)
        if code:
            return code

    # Programmatic fallback for YouTube
    video_id = extract_youtube_id(url)
    if video_id:
        print("  [fallback] Using programmatic YouTube embed code.")
        return build_youtube_embed(video_id)

    return None


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def _parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description=(
            "Extract embed codes from video URLs in a spreadsheet and write "
            "them back to a configurable column."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("spreadsheet", help="Path to the .xlsx, .xls, or .csv file.")
    parser.add_argument(
        "--url-col",
        default="1",
        metavar="COL",
        help=(
            "Column containing video URLs. "
            "Accepts a 1-based number ('1'), a letter ('A'), or a header name. "
            "Default: 1"
        ),
    )
    parser.add_argument(
        "--embed-col",
        default="2",
        metavar="COL",
        help=(
            "Column where embed code will be written. "
            "Same format as --url-col. "
            "Default: 2"
        ),
    )
    parser.add_argument(
        "--no-headless",
        action="store_true",
        help="Run browser in visible (non-headless) mode. Useful for debugging.",
    )
    return parser.parse_args(argv)


def main(argv=None) -> int:
    args = _parse_args(argv)
    path = args.spreadsheet

    if not os.path.isfile(path):
        print(f"Error: file not found: {path}", file=sys.stderr)
        return 1

    headless = not args.no_headless
    ext = os.path.splitext(path)[1].lower()

    if ext in (".xlsx", ".xls"):
        if not _OPENPYXL_AVAILABLE:
            print(
                "Error: openpyxl is required for Excel files.\n"
                "Run: pip install -r requirements.txt",
                file=sys.stderr,
            )
            return 1
        process_excel(path, args.url_col, args.embed_col, headless)

    elif ext == ".csv":
        process_csv(path, args.url_col, args.embed_col, headless)

    else:
        print(
            f"Error: unsupported file type '{ext}'. Use .xlsx, .xls, or .csv.",
            file=sys.stderr,
        )
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())

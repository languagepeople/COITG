# COITG
Useful Tools for COITG

---

## embed_extractor.py — Video Embed Code Extractor

`embed_extractor.py` is an offline script that reads a spreadsheet (Excel or
CSV), follows each video URL it finds, retrieves the `<iframe>` embed HTML by
automating the **Share → Embed** flow in a headless browser, and writes the
result into a separate column on the same row.

YouTube is the primary target.  When browser automation is unavailable (e.g.
Chrome is not installed), the script falls back to constructing the standard
YouTube embed code programmatically from the video ID.

### Supported file formats

| Format | Extension |
|--------|-----------|
| Excel (xlsx / xlsb) | `.xlsx`, `.xls` |
| CSV | `.csv` |

### Requirements

Python 3.10+ and the packages listed in `requirements.txt`:

```bash
pip install -r requirements.txt
```

Chrome (or Chromium) must also be installed for full browser-automation
support.  On most systems `webdriver-manager` (included in `requirements.txt`)
will download the correct ChromeDriver automatically.

### Usage

```
python embed_extractor.py <spreadsheet> [options]
```

| Option | Default | Description |
|--------|---------|-------------|
| `--url-col COL` | `1` | Column that contains the video URLs. |
| `--embed-col COL` | `2` | Column where the embed code will be written. |
| `--no-headless` | *(headless)* | Open a visible browser window (useful for debugging). |

`COL` can be:
* a **1-based column number** — `1`, `3`, …
* a **spreadsheet letter** — `A`, `B`, `C`, …
* an **exact column header name** — `"Video URL"`, `"Embed Code"`, …

### Examples

```bash
# URLs in column B (2), embed codes written to column C (3)
python embed_extractor.py videos.xlsx --url-col 2 --embed-col 3

# Use column letters
python embed_extractor.py videos.xlsx --url-col B --embed-col C

# Use header names
python embed_extractor.py videos.csv --url-col "Video URL" --embed-col "Embed Code"

# Visible browser window for debugging
python embed_extractor.py videos.xlsx --no-headless
```

The spreadsheet is **modified in-place**: the embed code is written directly
into the specified column and the file is saved.

### Spreadsheet layout example

Before running the script:

| Title | Video URL | Notes |
|-------|-----------|-------|
| My Tutorial | https://www.youtube.com/watch?v=dQw4w9WgXcQ | |

After running `python embed_extractor.py sheet.xlsx --url-col 2 --embed-col 4`:

| Title | Video URL | Notes | Embed Code |
|-------|-----------|-------|------------|
| My Tutorial | https://www.youtube.com/watch?v=dQw4w9WgXcQ | | `<iframe width="560" …></iframe>` |

### Running the tests

```bash
python -m pytest tests/ -v
```

The unit tests do not require a browser or network connection.

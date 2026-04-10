# COITG
Useful Tools for COITG

---

## embed_extractor.py — Video Embed Code Extractor

`embed_extractor.py` is an offline script that reads a spreadsheet (Excel or
CSV), follows each video URL it finds, retrieves the `<iframe>` embed HTML by
automating the **Share → Embed** flow in a headless browser, and writes the
result into a separate column on the same row.  It can also fetch the video
**duration** and write it to an additional column.

YouTube is the primary target.  When browser automation is unavailable (e.g.
Chrome is not installed), the script falls back to constructing the standard
YouTube embed code programmatically from the video ID.

The default column layout matches the COITG course-content spreadsheet:

| Purpose | Default column |
|---------|---------------|
| Video URL (input) | **F** |
| Embed code (output) | **O** |
| Video duration (output) | **P** |

### Supported file formats

| Format | Extension |
|--------|-----------|
| Excel | `.xlsx`, `.xls` |
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
| `--url-col COL` | `F` | Column that contains the video URLs. |
| `--embed-col COL` | `O` | Column where the embed code will be written. |
| `--duration-col COL` | `P` | Column where the video duration will be written (e.g. `"4:33"`). Pass `""` to disable. |
| `--no-headless` | *(headless)* | Open a visible browser window (useful for debugging). |

`COL` can be:
* a **1-based column number** — `1`, `6`, …
* a **spreadsheet letter** — `A`, `F`, `O`, …
* an **exact column header name** — `"Video URL"`, `"Embed Code"`, …

### Examples

```bash
# Run against the COITG course-content spreadsheet (default columns)
python embed_extractor.py "Excel Course Content for Moodle.xlsx"

# Override columns
python embed_extractor.py videos.xlsx --url-col 2 --embed-col 3 --duration-col 4

# Use header names
python embed_extractor.py videos.csv --url-col "Video URL" \
    --embed-col "Embed Code" --duration-col "Duration"

# Visible browser window for debugging
python embed_extractor.py videos.xlsx --no-headless

# Disable duration extraction
python embed_extractor.py videos.xlsx --duration-col ""
```

The spreadsheet is **modified in-place**: the embed code and duration are
written directly into the specified columns and the file is saved.

### Spreadsheet layout example

Before running the script (using default columns):

| … | F (Video URL) | … | O (Embed Code) | P (Duration) |
|---|--------------|---|----------------|--------------|
| … | https://www.youtube.com/watch?v=dQw4w9WgXcQ | … | | |

After running `python embed_extractor.py sheet.xlsx`:

| … | F (Video URL) | … | O (Embed Code) | P (Duration) |
|---|--------------|---|----------------|--------------|
| … | https://www.youtube.com/watch?v=dQw4w9WgXcQ | … | `<iframe …></iframe>` | `3:33` |

### Running the tests

```bash
python -m pytest tests/ -v
```

The unit tests do not require a browser or network connection.

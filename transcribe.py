import os
import subprocess
import argparse
import re
import json
import nltk
import whisper
from datetime import datetime
from docx import Document
from openpyxl import load_workbook
from pathlib import Path

nltk.download("punkt", quiet=True)
nltk.download("punkt_tab", quiet=True)
from nltk.tokenize import sent_tokenize


# ── Helpers ───────────────────────────────────────────────────────────────────

def format_timestamp(seconds: float) -> str:
    seconds = int(seconds)
    h = seconds // 3600
    m = (seconds % 3600) // 60
    s = seconds % 60
    return f"{h:02}:{m:02}:{s:02}"

def format_duration(seconds: float) -> str:
    """Convert raw seconds into a human-readable duration string."""
    seconds = int(seconds)
    h = seconds // 3600
    m = (seconds % 3600) // 60
    s = seconds % 60
    parts = []
    if h:
        parts.append(f"{h}h")
    if m:
        parts.append(f"{m}m")
    parts.append(f"{s}s")
    return " ".join(parts)

def sanitize_filename(name: str) -> str:
    return re.sub(r'[\\/*?:"<>|]', "", name).strip()

def get_video_metadata(url: str) -> dict | None:
    """
    Fetch title and duration from YouTube without downloading the video.
    Returns { "title": str, "duration": float } or None on failure.
    """
    result = subprocess.run(
        [
            "yt-dlp",
            "--print", "%{title}\t%{duration}s",
            "--no-playlist",
            url,
        ],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0 or not result.stdout.strip():
        return None
    parts = result.stdout.strip().split("\t")
    if len(parts) != 2:
        return None
    try:
        return {
            "title": sanitize_filename(parts[0].strip()),
            "duration": float(parts[1].strip()),
        }
    except ValueError:
        return None

def download_audio(url: str, output_path: str) -> bool:
    print(f"    Downloading audio...")
    result = subprocess.run(
        [
            "yt-dlp",
            "-x",
            "--audio-format", "mp3",
            "--audio-quality", "0",
            "--no-playlist",
            "-o", output_path,
            url,
        ],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        print(f"    ERROR downloading:\n{result.stderr.strip()}")
        return False
    return True

def transcribe_audio(audio_path: str, model) -> list[dict]:
    print(f"    Transcribing...")
    result = model.transcribe(audio_path, word_timestamps=True)

    sentences_with_times = []
    for segment in result["segments"]:
        text = segment["text"].strip()
        if not text:
            continue

        sentences = sent_tokenize(text)
        words = segment.get("words", [])
        word_idx = 0

        for sentence in sentences:
            word_count = len(sentence.split())
            sentence_words = words[word_idx : word_idx + word_count]

            start = sentence_words[0]["start"] if sentence_words else segment["start"]
            end = sentence_words[-1]["end"] if sentence_words else segment["end"]

            sentences_with_times.append({
                "sentence": sentence,
                "start": start,
                "end": end,
            })
            word_idx += word_count

    return sentences_with_times

def save_docx(sentences: list[dict], output_path: str, source_name: str):
    doc = Document()
    doc.add_heading(f"Transcription: {source_name}", level=1)

    for item in sentences:
        start = format_timestamp(item["start"])
        end = format_timestamp(item["end"])
        paragraph = doc.add_paragraph()
        run_ts = paragraph.add_run(f"[{start} → {end}]  ")
        run_ts.bold = True
        paragraph.add_run(item["sentence"])

    doc.save(output_path)
    print(f"    Saved: {output_path}")


# ── Logging ───────────────────────────────────────────────────────────────────

def write_log(log_path: str, entry: dict):
    """
    Append a single job entry to the log file.

    Each entry is a JSON object on its own line (newline-delimited JSON / NDJSON),
    making it easy to parse programmatically while still being human-readable.

    A human-readable text block is written above each JSON line as a comment.
    """
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    entry["logged_at"] = timestamp

    with open(log_path, "a", encoding="utf-8") as f:
        f.write("\n")
        f.write("─" * 60 + "\n")
        f.write(f"  Logged at  : {timestamp}\n")
        f.write(f"  Status     : {entry.get('status', 'unknown')}\n")
        f.write(f"  Title      : {entry.get('title', 'N/A')}\n")
        f.write(f"  URL        : {entry.get('url', 'N/A')}\n")
        f.write(f"  Duration   : {entry.get('duration_readable', 'N/A')}\n")
        f.write(f"  Audio file : {entry.get('audio_path', 'N/A')}\n")
        f.write(f"  Transcript : {entry.get('transcript_path', 'N/A')}\n")
        if entry.get("error"):
            f.write(f"  Error      : {entry['error']}\n")
        f.write(f"  JSON       : {json.dumps(entry)}\n")


# ── Core Processing ───────────────────────────────────────────────────────────

def process_audio_file(audio_path: str, output_dir: str, model, skip_existing: bool, log_path: str, url: str = "N/A"):
    """Transcribe a single local audio file and save as .docx."""
    stem = Path(audio_path).stem
    output_path = os.path.join(output_dir, f"{stem}.docx")

    log_entry = {
        "url": url,
        "title": stem,
        "audio_path": str(Path(audio_path).resolve()),
        "transcript_path": str(Path(output_path).resolve()),
        "duration_raw": None,
        "duration_readable": "N/A",
        "status": None,
        "error": None,
    }

    if skip_existing and os.path.exists(output_path):
        print(f"    Skipping — transcript already exists: {output_path}")
        log_entry["status"] = "skipped"
        write_log(log_path, log_entry)
        return True

    sentences = transcribe_audio(audio_path, model)
    if not sentences:
        print(f"    WARNING: Empty transcription.")
        log_entry["status"] = "failed"
        log_entry["error"] = "Empty transcription result"
        write_log(log_path, log_entry)
        return False

    save_docx(sentences, output_path, stem)
    log_entry["status"] = "success"
    write_log(log_path, log_entry)
    return True

def process_spreadsheet(xlsx_path: str, audio_dir: str, output_dir: str, model, skip_existing: bool, log_path: str):
    wb = load_workbook(xlsx_path)
    ws = wb.active

    # Column F is index 5 (zero-based)
    URL_COL = 5

    rows = [
        (i + 1, str(row[URL_COL].value).strip())
        for i, row in enumerate(ws.iter_rows())
        if row[URL_COL].value and str(row[URL_COL].value).strip().startswith("http")
    ]

    if not rows:
        print("No URLs found in the spreadsheet.")
        return

    total = len(rows)
    print(f"\nFound {total} URL(s) in spreadsheet. Starting...\n")

    failed = []

    for row_num, url in rows:
        print(f"[Row {row_num}/{total}] {url}")

        log_entry = {
            "url": url,
            "title": None,
            "audio_path": None,
            "transcript_path": None,
            "duration_raw": None,
            "duration_readable": None,
            "status": None,
            "error": None,
        }

        # Step 1: Fetch metadata (title + duration)
        print(f"    Fetching metadata...")
        meta = get_video_metadata(url)
        if not meta:
            print(f"    ERROR: Could not fetch metadata. Skipping.")
            log_entry["status"] = "failed"
            log_entry["error"] = "Could not fetch video metadata"
            write_log(log_path, log_entry)
            failed.append((row_num, url, "Could not fetch metadata"))
            continue

        title = meta["title"]
        duration = meta["duration"]
        audio_path = os.path.join(audio_dir, f"{title}.mp3")
        output_path = os.path.join(output_dir, f"{title}.docx")

        log_entry["title"] = title
        log_entry["duration_raw"] = duration
        log_entry["duration_readable"] = format_duration(duration)
        log_entry["audio_path"] = str(Path(audio_path).resolve())
        log_entry["transcript_path"] = str(Path(output_path).resolve())

        print(f"    Title    : {title}")
        print(f"    Duration : {format_duration(duration)}")

        # Step 2: Skip if transcript already exists
        if skip_existing and os.path.exists(output_path):
            print(f"    Skipping — transcript already exists.\n")
            log_entry["status"] = "skipped"
            write_log(log_path, log_entry)
            continue

        # Step 3: Download audio if not already on disk
        if not os.path.exists(audio_path):
            success = download_audio(url, audio_path)
            if not success:
                log_entry["status"] = "failed"
                log_entry["error"] = "Audio download failed"
                write_log(log_path, log_entry)
                failed.append((row_num, url, "Download failed"))
                continue
        else:
            print(f"    Audio already on disk, skipping download.")

        # Step 4: Transcribe and save
        sentences = transcribe_audio(audio_path, model)
        if not sentences:
            print(f"    WARNING: Empty transcription.")
            log_entry["status"] = "failed"
            log_entry["error"] = "Empty transcription result"
            write_log(log_path, log_entry)
            failed.append((row_num, url, "Empty transcription"))
            continue

        save_docx(sentences, output_path, title)
        log_entry["status"] = "success"
        write_log(log_path, log_entry)
        print(f"    ✓ Done.\n")

    # Summary
    print("\n" + "=" * 50)
    print(f"Completed: {total - len(failed)}/{total} rows")
    if failed:
        print(f"\nFailed rows:")
        for row_num, url, reason in failed:
            print(f"  Row {row_num}: {reason} — {url}")
    else:
        print("All rows completed successfully.")
    print("=" * 50)
    print(f"\nLog written to: {log_path}")

def process_directory(audio_dir: str, output_dir: str, model, skip_existing: bool, log_path: str):
    extensions = {".mp3", ".mp4", ".wav", ".m4a", ".ogg", ".flac", ".webm"}
    files = sorted([
        f for f in Path(audio_dir).iterdir()
        if f.is_file() and f.suffix.lower() in extensions
    ])

    if not files:
        print(f"No audio files found in: {audio_dir}")
        return

    total = len(files)
    print(f"\nFound {total} audio file(s). Starting...\n")

    failed = []

    for i, audio_file in enumerate(files, 1):
        print(f"[{i}/{total}] {audio_file.name}")
        success = process_audio_file(
            str(audio_file), output_dir, model, skip_existing, log_path
        )
        if not success:
            failed.append(audio_file.name)
        print()

    print("=" * 50)
    print(f"Completed: {total - len(failed)}/{total} files")
    if failed:
        print(f"\nFailed files:")
        for name in failed:
            print(f"  {name}")
    else:
        print("All files completed successfully.")
    print("=" * 50)
    print(f"\nLog written to: {log_path}")


# ── CLI Entry Point ───────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="YouTube Audio Transcriber → DOCX with timestamps"
    )

    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--spreadsheet", "-s", help="Path to .xlsx file with YouTube URLs in column F.")
    group.add_argument("--directory", "-d", help="Path to a directory of audio files to transcribe.")

    parser.add_argument("--audio-dir", "-a", default="./audio", help="Where to save downloaded audio. (default: ./audio)")
    parser.add_argument("--output-dir", "-o", default="./transcripts", help="Where to save .docx files. (default: ./transcripts)")
    parser.add_argument("--log-file", "-l", default="./transcription_log.txt", help="Path to the log file. (default: ./transcription_log.txt)")
    parser.add_argument("--model", "-m", default="medium", choices=["tiny", "base", "small", "medium", "large"], help="Whisper model size. (default: medium)")
    parser.add_argument("--no-skip", action="store_true", help="Re-transcribe even if a .docx already exists.")

    args = parser.parse_args()

    os.makedirs(args.audio_dir, exist_ok=True)
    os.makedirs(args.output_dir, exist_ok=True)

    # Write a session header to the log
    with open(args.log_file, "a", encoding="utf-8") as f:
        f.write("\n" + "═" * 60 + "\n")
        f.write(f"  SESSION START: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"  Mode         : {'spreadsheet' if args.spreadsheet else 'directory'}\n")
        f.write(f"  Input        : {args.spreadsheet or args.directory}\n")
        f.write(f"  Audio dir    : {Path(args.audio_dir).resolve()}\n")
        f.write(f"  Output dir   : {Path(args.output_dir).resolve()}\n")
        f.write(f"  Whisper model: {args.model}\n")
        f.write(f"  Skip existing: {not args.no_skip}\n")
        f.write("═" * 60 + "\n")

    print(f"Loading Whisper model '{args.model}'...")
    model = whisper.load_model(args.model)

    if args.spreadsheet:
        process_spreadsheet(args.spreadsheet, args.audio_dir, args.output_dir, model, not args.no_skip, args.log_file)
    elif args.directory:
        process_directory(args.directory, args.output_dir, model, not args.no_skip, args.log_file)


if __name__ == "__main__":
    main()
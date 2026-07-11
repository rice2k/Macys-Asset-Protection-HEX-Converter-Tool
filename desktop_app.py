from __future__ import annotations

import csv
import json
import os
import re
import sys
import webbrowser
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from datetime import datetime
from html.parser import HTMLParser
from pathlib import Path
from typing import Any

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Side, Border
from PIL import Image, ImageTk

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter, landscape
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
except Exception:  # pragma: no cover - optional dependency in dev
    colors = None
    letter = landscape = getSampleStyleSheet = None
    SimpleDocTemplate = Paragraph = Spacer = Table = TableStyle = None

try:
    import xlrd
except Exception:  # pragma: no cover - optional dependency in dev
    xlrd = None

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except Exception:  # pragma: no cover - drag/drop is optional
    DND_FILES = None
    TkinterDnD = None


BLUEWAVE_URL = "http://ma000xsblw1001/"
APP_DISPLAY_NAME = "Macy's Asset Protection - China Grove Hex Converter Utility"
APP_SHORT_NAME = "Macy's AP China Grove Hex Utility"
APP_STATE_DIR = Path(os.environ.get("APPDATA", str(Path.home()))) / "AP_Access_Control_Converter"
SETTINGS_FILE = APP_STATE_DIR / "settings.json"
TkRoot = TkinterDnD.Tk if TkinterDnD else tk.Tk


class ToolTip:
    def __init__(self, widget: tk.Widget, text: str, delay: int = 650) -> None:
        self.widget = widget
        self.text = text
        self.delay = delay
        self.after_id: str | None = None
        self.window: tk.Toplevel | None = None
        widget.bind("<Enter>", self.schedule, add="+")
        widget.bind("<Leave>", self.hide, add="+")
        widget.bind("<ButtonPress>", self.hide, add="+")

    def schedule(self, _event: Any | None = None) -> None:
        self.cancel()
        self.after_id = self.widget.after(self.delay, self.show)

    def cancel(self) -> None:
        if self.after_id:
            self.widget.after_cancel(self.after_id)
            self.after_id = None

    def show(self) -> None:
        if self.window or not self.text:
            return
        x = self.widget.winfo_rootx() + 14
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 8
        self.window = tk.Toplevel(self.widget)
        self.window.wm_overrideredirect(True)
        self.window.wm_geometry(f"+{x}+{y}")
        frame = tk.Frame(self.window, bg="#05070b", highlightthickness=1, highlightbackground="#3c4656")
        frame.pack()
        tk.Label(
            frame,
            text=self.text,
            bg="#05070b",
            fg="#f5f7fb",
            justify="left",
            wraplength=280,
            padx=10,
            pady=7,
            font=("Segoe UI", 9),
        ).pack()

    def hide(self, _event: Any | None = None) -> None:
        self.cancel()
        if self.window:
            self.window.destroy()
            self.window = None


@dataclass
class ConvertedRow:
    line: int
    raw: str
    hex_value: str
    facility: int
    card: int
    converted_at: str
    suggestions: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


@dataclass
class InvalidRow:
    line: int
    raw: str
    reason: str
    converted_at: str


@dataclass
class UnconvertRow:
    line: int
    raw: str
    facility: int
    card: int
    hex_value: str
    converted_at: str
    warnings: list[str] = field(default_factory=list)


@dataclass
class HistoryEntry:
    timestamp: str
    action: str
    total: int
    valid: int
    invalid: int
    warnings: int


def asset_path(name: str) -> Path:
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / "assets" / name
    return Path(__file__).resolve().parent / "src" / "assets" / name


def valid_hex8(value: str) -> bool:
    return bool(re.fullmatch(r"[0-9A-Fa-f]{8}", value or ""))


def hex_to_fc_cn(hex_value: str) -> tuple[int, int]:
    clean = hex_value.strip().upper()
    if not valid_hex8(clean):
        raise ValueError("Hex value must be exactly 8 characters.")
    number = int(clean, 16)
    return (number >> 16) & 0xFFFF, number & 0xFFFF


def strict_int(value: str) -> int:
    text = str(value).strip()
    if not re.fullmatch(r"\d+", text):
        raise ValueError("Facility Code and Card Number must be whole numbers.")
    return int(text)


def fc_cn_to_hex(fc: str | int, cn: str | int) -> str:
    facility = strict_int(str(fc))
    card = strict_int(str(cn))
    if facility < 0 or facility > 65535 or card < 0 or card > 65535:
        raise ValueError("FC and CN must be between 0 and 65535.")
    return f"{((facility << 16) | card):08X}"


def clean_candidate_line(line: str) -> dict[str, Any]:
    original = str(line or "")
    cleaned = original.strip()
    suggestions: list[str] = []
    if not cleaned:
        return {"original": original, "cleaned": "", "extracted": "", "suggestions": suggestions}

    token = re.search(r"\b([0-9A-Fa-f]{8})\b", cleaned)
    split_token = re.search(r"\b([0-9A-Fa-f]{4})[\s-]+([0-9A-Fa-f]{4})\b", cleaned)
    extracted = ""
    if token:
        extracted = token.group(1).upper()
        if cleaned != token.group(1):
            suggestions.append("Extracted 8-character ID from full text.")
    elif split_token:
        extracted = f"{split_token.group(1)}{split_token.group(2)}".upper()
        suggestions.append("Joined a split 8-character ID.")

    fallback = re.sub(r"\s+", " ", re.sub(r"[\t,;\"']", " ", cleaned)).strip()
    compact = re.sub(r"[^0-9A-Fa-f]", "", fallback).upper()
    if not extracted and len(compact) == 8:
        extracted = compact
        if compact != cleaned.upper():
            suggestions.append("Removed spaces or punctuation around the ID.")

    return {"original": original, "cleaned": fallback, "extracted": extracted, "suggestions": suggestions}


def invalid_reason(original: str) -> str:
    raw = str(original or "").strip()
    if not raw:
        return "Blank line"
    hex_like = re.sub(r"[^0-9A-Fa-f]", "", raw)
    if len(hex_like) == 0:
        return "No hex characters found"
    if len(hex_like) < 8:
        return "Too short; needs exactly 8 characters"
    if len(hex_like) > 8 and not re.search(r"\b[0-9A-Fa-f]{8}\b", raw):
        return "Too long or mixed with extra characters"
    if re.search(r"[^0-9A-Fa-f\s,;\"'\-]", raw):
        return "Contains invalid characters"
    return "Must be exactly 8 characters using 0-9 and A-F"


def unusual_warnings(hex_value: str, facility: int, card: int) -> list[str]:
    warnings: list[str] = []
    if re.fullmatch(r"0{8}", hex_value):
        warnings.append("All zeros is unusual.")
    if re.fullmatch(r"F{8}", hex_value, re.I):
        warnings.append("All Fs is unusual.")
    if re.fullmatch(r"\d{8}", hex_value) and not re.fullmatch(r"88\d{6}", hex_value):
        warnings.append("Numeric ID does not start with the common 88 prefix.")
    if facility == 0:
        warnings.append("Facility Code is 0.")
    if card == 0:
        warnings.append("Card Number is 0.")
    if facility > 60000 or card > 60000:
        warnings.append("Very high FC or CN value.")
    return warnings


def convert_lines(text: str, converted_at: str) -> tuple[list[ConvertedRow], list[InvalidRow]]:
    converted: list[ConvertedRow] = []
    invalid: list[InvalidRow] = []
    seen_hex: dict[str, int] = {}
    for idx, line in enumerate(str(text or "").splitlines(), start=1):
        prepared = clean_candidate_line(line)
        if not prepared["original"].strip():
            continue
        hex_value = prepared["extracted"].strip().upper()
        if not valid_hex8(hex_value):
            invalid.append(InvalidRow(idx, prepared["original"], invalid_reason(prepared["original"]), converted_at))
            continue
        facility, card = hex_to_fc_cn(hex_value)
        warnings = unusual_warnings(hex_value, facility, card)
        if hex_value in seen_hex:
            warnings.append(f"Duplicate of line {seen_hex[hex_value]}.")
        else:
            seen_hex[hex_value] = idx
        converted.append(
            ConvertedRow(
                line=idx,
                raw=prepared["original"],
                hex_value=hex_value,
                facility=facility,
                card=card,
                converted_at=converted_at,
                suggestions=list(prepared["suggestions"]),
                warnings=warnings,
            )
        )
    return converted, invalid


def parse_fc_cn_line(line: str) -> tuple[int, int]:
    raw = str(line or "")
    labeled = re.search(r"\bfc\b\D*(\d{1,5})\D+\bcn\b\D*(\d{1,5})", raw, re.I)
    if labeled:
        return strict_int(labeled.group(1)), strict_int(labeled.group(2))
    numbers = re.findall(r"\b\d{1,5}\b", raw)
    if len(numbers) < 2:
        raise ValueError("Needs Facility Code and Card Number.")
    return strict_int(numbers[0]), strict_int(numbers[1])


def unconvert_lines(text: str, converted_at: str) -> tuple[list[UnconvertRow], list[InvalidRow]]:
    converted: list[UnconvertRow] = []
    invalid: list[InvalidRow] = []
    seen_hex: dict[str, int] = {}
    for idx, line in enumerate(str(text or "").splitlines(), start=1):
        raw = str(line or "").strip()
        if not raw:
            continue
        try:
            facility, card = parse_fc_cn_line(raw)
            hex_value = fc_cn_to_hex(facility, card)
        except ValueError as exc:
            invalid.append(InvalidRow(idx, raw, str(exc), converted_at))
            continue
        warnings = unusual_warnings(hex_value, facility, card)
        if hex_value in seen_hex:
            warnings.append(f"Duplicate of line {seen_hex[hex_value]}.")
        else:
            seen_hex[hex_value] = idx
        converted.append(UnconvertRow(idx, raw, facility, card, hex_value, converted_at, warnings))
    return converted, invalid


def cell_text(value: Any) -> str:
    return re.sub(r"\s+", " ", "" if value is None else str(value)).strip()


def normalize_header(value: Any) -> str:
    text = cell_text(value).lower().replace("&nbsp;", " ")
    text = re.sub(r"[\r\n\t]+", " ", text)
    text = re.sub(r"[^a-z0-9#]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


NAME_RULES = [
    ("candidate name", 100),
    ("colleague name", 92),
    ("employee name", 90),
    ("associate name", 88),
    ("full name", 80),
    ("name", 60),
]

ID_RULES = [
    ("colleague #", 100),
    ("colleague number", 98),
    ("colleague no", 96),
    ("colleague id", 94),
    ("employee id", 88),
    ("employee number", 84),
    ("associate id", 80),
    ("badge id", 72),
]

BAD_NAME_PATTERNS = [
    re.compile(r"\bpl name\b"),
    re.compile(r"\bmanager\b"),
    re.compile(r"\blead\b"),
    re.compile(r"\bsupervisor\b"),
    re.compile(r"\brecruiter\b"),
    re.compile(r"\btrainer\b"),
    re.compile(r"\binterviewer\b"),
]

BAD_ID_PATTERNS = [
    re.compile(r"\bcandidate id\b"),
    re.compile(r"\breq id\b"),
    re.compile(r"\bjob id\b"),
    re.compile(r"\bposition id\b"),
]


def header_score(normalized: str, rules: list[tuple[str, int]], bad_patterns: list[re.Pattern[str]]) -> int:
    if not normalized:
        return 0
    score = 0
    for label, points in rules:
        if normalized == label:
            score = max(score, points + 20)
        elif label in normalized:
            score = max(score, points)
    for pattern in bad_patterns:
        if pattern.search(normalized):
            score -= 120
    return score


def find_header_pair(rows: list[list[str]]) -> dict[str, Any] | None:
    candidates: list[dict[str, Any]] = []
    for row_index, row in enumerate(rows[:8]):
        for col_index, cell in enumerate(row):
            normalized = normalize_header(cell)
            name_score = header_score(normalized, NAME_RULES, BAD_NAME_PATTERNS)
            id_score = header_score(normalized, ID_RULES, BAD_ID_PATTERNS)
            if name_score > 0 or id_score > 0:
                candidates.append(
                    {
                        "row": row_index,
                        "col": col_index,
                        "normalized": normalized,
                        "name_score": name_score,
                        "id_score": id_score,
                    }
                )
    best = None
    for name in [c for c in candidates if c["name_score"] > 0]:
        for emp_id in [c for c in candidates if c["id_score"] > 0]:
            score = name["name_score"] + emp_id["id_score"]
            score -= abs(name["row"] - emp_id["row"]) * 12
            score -= max(0, abs(name["col"] - emp_id["col"]) - 8) * 2
            if name["normalized"] == "candidate name":
                score += 60
            if emp_id["normalized"] == "colleague #":
                score += 60
            if name["normalized"] == "candidate name" and emp_id["normalized"] == "colleague #":
                score += 200
            if not best or score > best["score"]:
                best = {
                    "score": score,
                    "name_index": name["col"],
                    "id_index": emp_id["col"],
                    "header_row": max(name["row"], emp_id["row"]),
                    "name_header": name["normalized"],
                    "id_header": emp_id["normalized"],
                    "exact": name["normalized"] == "candidate name" and emp_id["normalized"] == "colleague #",
                }
    return best


def extract_eight_digit_id(value: Any) -> str:
    raw = cell_text(value)
    preferred = re.search(r"\b(88\d{6})\b", raw)
    if preferred:
        return preferred.group(1)
    generic = re.search(r"\b(\d{8})\b", raw)
    return generic.group(1) if generic else ""


def is_noise_name(value: Any) -> bool:
    text = cell_text(value)
    normalized = normalize_header(text)
    if not normalized:
        return True
    if any(pattern.search(normalized) for pattern in BAD_NAME_PATTERNS):
        return True
    if re.search(r"\b(candidate id|colleague #|employee id|employee number|requisition|job title|position|department|location|status|email|phone)\b", normalized):
        return True
    return bool(re.fullmatch(r"\d+", text) or "@" in text)


def is_likely_person_name(value: Any) -> bool:
    text = cell_text(value)
    if not text or is_noise_name(text):
        return False
    if len(text) < 3 or len(text) > 80 or re.match(r"^\d", text):
        return False
    if not re.search(r"[A-Za-z]", text):
        return False
    if re.search(r"\b(warehouse|associate|operator|shift|full time|part time|china grove|salisbury|north carolina)\b", text, re.I):
        return False
    return bool(re.fullmatch(r"[A-Za-z][A-Za-z\s.\-'\u2019]+", text))


def clean_imported_pair(name: Any, emp_id: Any) -> tuple[str, str] | None:
    clean_name = cell_text(name)
    clean_id = extract_eight_digit_id(emp_id)
    if not is_likely_person_name(clean_name) or not clean_id:
        return None
    return clean_name, clean_id


def dedupe_pairs(pairs: list[tuple[str, str]]) -> list[tuple[str, str]]:
    seen = set()
    out = []
    for name, emp_id in pairs:
        key = emp_id.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append((name, emp_id))
    return out


def fallback_extract_pairs(rows: list[list[str]]) -> list[tuple[str, str]]:
    pairs = []
    for row in rows:
        ids = [(idx, extract_eight_digit_id(cell)) for idx, cell in enumerate(row)]
        ids = [(idx, emp_id) for idx, emp_id in ids if emp_id]
        if not ids:
            continue
        names = [(idx, cell_text(cell)) for idx, cell in enumerate(row) if is_likely_person_name(cell)]
        if not names:
            continue
        best_id = ids[0]
        names.sort(key=lambda item: abs(item[0] - best_id[0]))
        pair = clean_imported_pair(names[0][1], best_id[1])
        if pair:
            pairs.append(pair)
    return dedupe_pairs(pairs)


def extract_name_id_lines_from_tables(tables: list[list[list[str]]], source_label: str) -> dict[str, Any]:
    pairs: list[tuple[str, str]] = []
    strategy = "fallback"
    match_label = ""
    for rows in tables:
        header_pair = find_header_pair(rows)
        if header_pair:
            extracted = []
            for row in rows[header_pair["header_row"] + 1 :]:
                pair = clean_imported_pair(
                    row[header_pair["name_index"]] if header_pair["name_index"] < len(row) else "",
                    row[header_pair["id_index"]] if header_pair["id_index"] < len(row) else "",
                )
                if pair:
                    extracted.append(pair)
            if extracted:
                pairs.extend(dedupe_pairs(extracted))
                if header_pair["exact"]:
                    strategy = "exact headers"
                    match_label = "Candidate Name + Colleague #"
                elif strategy != "exact headers":
                    strategy = "alternate headers"
                    match_label = f"{header_pair['name_header']} + {header_pair['id_header']}"
                continue
        pairs.extend(fallback_extract_pairs(rows))

    unique = dedupe_pairs(pairs)
    mode = f"{strategy} ({match_label})" if match_label else strategy
    return {
        "lines": [f"{name}, {emp_id}" for name, emp_id in unique],
        "found_rows": len(unique),
        "strategy": strategy,
        "message": f"Imported {len(unique)} row(s) from {source_label} using {mode}." if unique else f"No matching name and employee ID pairs could be pulled from {source_label}.",
    }


def parse_delimited_text(text: str) -> list[list[str]]:
    lines = [line for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n") if line.strip()]
    sample = "\n".join(lines[:10])
    delimiter = max(["\t", ",", ";"], key=lambda item: sample.count(item))
    if sample.count(delimiter) == 0:
        delimiter = ","
    return [[cell_text(cell) for cell in row] for row in csv.reader(lines, delimiter=delimiter)]


class SimpleTableParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.tables: list[list[list[str]]] = []
        self._table: list[list[str]] | None = None
        self._row: list[str] | None = None
        self._cell: list[str] | None = None

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        tag = tag.lower()
        if tag == "table":
            self._table = []
        elif tag == "tr" and self._table is not None:
            self._row = []
        elif tag in {"td", "th"} and self._row is not None:
            self._cell = []

    def handle_data(self, data: str) -> None:
        if self._cell is not None:
            self._cell.append(data)

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if tag in {"td", "th"} and self._row is not None and self._cell is not None:
            self._row.append(cell_text("".join(self._cell)))
            self._cell = None
        elif tag == "tr" and self._table is not None and self._row is not None:
            if any(self._row):
                self._table.append(self._row)
            self._row = None
        elif tag == "table" and self._table is not None:
            if self._table:
                self.tables.append(self._table)
            self._table = None


def parse_html_tables(text: str) -> list[list[list[str]]]:
    parser = SimpleTableParser()
    parser.feed(text or "")
    return parser.tables


def local_name(tag: str) -> str:
    return tag.split("}", 1)[-1] if "}" in tag else tag


def parse_spreadsheet_xml(text: str) -> list[list[list[str]]]:
    root = ET.fromstring(text)
    rows = []
    for row_node in root.iter():
        if local_name(row_node.tag) != "Row":
            continue
        cells = []
        expected_index = 1
        for cell_node in list(row_node):
            if local_name(cell_node.tag) != "Cell":
                continue
            idx_attr = None
            for key, value in cell_node.attrib.items():
                if local_name(key) == "Index":
                    idx_attr = value
                    break
            if idx_attr:
                while expected_index < int(idx_attr):
                    cells.append("")
                    expected_index += 1
            data_text = "".join(cell_node.itertext())
            cells.append(cell_text(data_text))
            expected_index += 1
        if any(cells):
            rows.append(cells)
    return [rows] if rows else []


def parse_xlsx(path: Path) -> list[list[list[str]]]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    tables: list[list[list[str]]] = []
    try:
        for sheet in workbook.worksheets:
            rows = []
            for row in sheet.iter_rows(values_only=True):
                values = [cell_text(cell) for cell in row]
                if any(values):
                    rows.append(values)
            if rows:
                tables.append(rows)
    finally:
        workbook.close()
    return tables


def parse_xls(path: Path) -> list[list[list[str]]]:
    if xlrd is None:
        raise RuntimeError("XLS support requires xlrd.")
    book = xlrd.open_workbook(str(path))
    tables = []
    for sheet in book.sheets():
        rows = []
        for row_index in range(sheet.nrows):
            values = [cell_text(sheet.cell_value(row_index, col_index)) for col_index in range(sheet.ncols)]
            if any(values):
                rows.append(values)
        if rows:
            tables.append(rows)
    return tables


def import_structured_file(path: Path) -> dict[str, Any]:
    lower = path.name.lower()
    if lower.endswith((".xlsx", ".xlsm")):
        return extract_name_id_lines_from_tables(parse_xlsx(path), path.name)
    if lower.endswith(".xls"):
        return extract_name_id_lines_from_tables(parse_xls(path), path.name)

    text = path.read_text(encoding="utf-8-sig", errors="replace")
    if lower.endswith(".xml") or "<Workbook" in text:
        return extract_name_id_lines_from_tables(parse_spreadsheet_xml(text), path.name)
    if lower.endswith((".html", ".htm")) or re.search(r"<table[\s>]", text, re.I):
        return extract_name_id_lines_from_tables(parse_html_tables(text), path.name)
    if lower.endswith((".txt", ".csv", ".tsv")):
        structured = extract_name_id_lines_from_tables([parse_delimited_text(text)], path.name)
        if structured["found_rows"]:
            return structured
        return {
            "lines": [line for line in text.replace("\r\n", "\n").split("\n") if line.strip()],
            "found_rows": 0,
            "strategy": "raw text",
            "message": f"Imported raw text from {path.name}.",
        }
    raise ValueError("Unsupported file type.")


class ConverterApp(TkRoot):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_DISPLAY_NAME)
        self.geometry("1200x780")
        self.minsize(980, 640)
        self.configure(bg="#0b0d12")
        self.settings = self.load_settings()
        self.results: list[ConvertedRow] = []
        self.invalid: list[InvalidRow] = []
        self.unconvert_results: list[UnconvertRow] = []
        self.unconvert_invalid: list[InvalidRow] = []
        self.last_converted_at = ""
        self.row_lookup: dict[str, ConvertedRow | InvalidRow] = {}
        self.unconvert_row_lookup: dict[str, UnconvertRow | InvalidRow] = {}
        self.sort_column = "Line"
        self.sort_reverse = False
        self.logo_photo: ImageTk.PhotoImage | None = None
        self.header_accent_photo: ImageTk.PhotoImage | None = None
        self.nav_icon_photos: list[ImageTk.PhotoImage] = []
        self.icon_photos: dict[tuple[str, int, int], ImageTk.PhotoImage] = {}
        self.empty_results_photo: ImageTk.PhotoImage | None = None
        self._setup_icon()
        self._setup_styles()
        self._build_shell()
        self._setup_shortcuts()
        self.render_results()
        self.render_unconvert_results()

    def load_settings(self) -> dict[str, Any]:
        defaults = {"default_export_dir": str(Path.home() / "Desktop"), "recent_files": [], "history": []}
        try:
            if SETTINGS_FILE.exists():
                data = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
                defaults.update({key: data.get(key, value) for key, value in defaults.items()})
        except Exception:
            pass
        return defaults

    def save_settings(self) -> None:
        APP_STATE_DIR.mkdir(parents=True, exist_ok=True)
        SETTINGS_FILE.write_text(json.dumps(self.settings, indent=2), encoding="utf-8")

    def add_recent_file(self, path: Path) -> None:
        recent = [item for item in self.settings.get("recent_files", []) if item != str(path)]
        recent.insert(0, str(path))
        self.settings["recent_files"] = recent[:8]
        self.save_settings()
        if hasattr(self, "recent_menu"):
            self.refresh_recent_menu()

    def add_history(self, action: str, total: int, valid: int, invalid: int, warnings: int) -> None:
        entry = HistoryEntry(datetime.now().strftime("%m/%d/%Y %I:%M:%S %p"), action, total, valid, invalid, warnings)
        history = self.settings.get("history", [])
        history.insert(0, entry.__dict__)
        self.settings["history"] = history[:50]
        self.save_settings()
        if hasattr(self, "history_text"):
            self.render_history()

    def _setup_icon(self) -> None:
        icon = asset_path("app-icon.ico")
        if icon.exists():
            try:
                self.iconbitmap(str(icon))
            except tk.TclError:
                pass

    def _setup_styles(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(".", background="#0f1218", foreground="#f5f7fb", fieldbackground="#151922")
        style.configure("TNotebook", background="#0b0d12", borderwidth=0)
        style.configure("TNotebook.Tab", padding=(16, 9), background="#151922", foreground="#a8b2c2")
        style.map("TNotebook.Tab", background=[("selected", "#251017")], foreground=[("selected", "#ffffff")])
        style.configure("Treeview", background="#11151d", foreground="#f5f7fb", fieldbackground="#11151d", rowheight=30, borderwidth=0)
        style.configure("Treeview.Heading", background="#1d222c", foreground="#a8b2c2", font=("Segoe UI", 9, "bold"))
        style.map("Treeview", background=[("selected", "#981624")], foreground=[("selected", "#ffffff")])
        style.configure("TCombobox", padding=6)
        style.configure("Dark.Vertical.TScrollbar", background="#202633", troughcolor="#0b0d12", bordercolor="#323b49", arrowcolor="#f5f7fb")
        style.configure("Dark.Horizontal.TScrollbar", background="#202633", troughcolor="#0b0d12", bordercolor="#323b49", arrowcolor="#f5f7fb")

    def _setup_shortcuts(self) -> None:
        self.bind_all("<Control-i>", lambda _event: self.import_file())
        self.bind_all("<Control-r>", lambda _event: self.convert_batch())
        self.bind_all("<Control-e>", lambda _event: self.export_excel())
        self.bind_all("<Control-Shift-E>", lambda _event: self.export_csv())
        self.bind_all("<Control-p>", lambda _event: self.export_pdf())
        self.bind_all("<Control-l>", lambda _event: self.clear_workspace())
        self.bind_all("<Control-f>", lambda _event: self.focus_search())

    def focus_search(self) -> None:
        self.select_tab(0)
        if hasattr(self, "search_entry"):
            self.search_var.set("")
            self.search_entry.focus_set()
        self.set_status("Search is ready.")

    def _build_menu(self) -> None:
        menubar = tk.Menu(self)
        import_menu = tk.Menu(menubar, tearoff=False)
        import_menu.add_command(label="Browse Files", command=self.import_file)
        import_menu.add_command(label="Paste Clipboard To Queue", command=self.paste_clipboard_to_queue)
        import_menu.add_command(label="Load Sample IDs", command=self.load_sample)
        menubar.add_cascade(label="Import", menu=import_menu)

        file_menu = tk.Menu(menubar, tearoff=False)
        file_menu.add_command(label="Choose Default Export Folder", command=self.choose_export_folder)
        file_menu.add_command(label="Open Export Folder", command=self.open_export_folder)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.destroy)
        menubar.add_cascade(label="File", menu=file_menu)

        export_menu = tk.Menu(menubar, tearoff=False)
        export_menu.add_command(label="Excel Workbook", command=self.export_excel)
        export_menu.add_command(label="CSV Report", command=self.export_csv)
        export_menu.add_command(label="TXT Report", command=self.export_txt)
        export_menu.add_command(label="PDF Report", command=self.export_pdf)
        menubar.add_cascade(label="Export", menu=export_menu)

        help_menu = tk.Menu(menubar, tearoff=False)
        help_menu.add_command(label="How To Use", command=self.show_help)
        help_menu.add_command(label="About", command=self.show_about)
        help_menu.add_separator()
        help_menu.add_command(label="Open BlueWave", command=self.open_bluewave)
        help_menu.add_command(label="Open GitHub Profile", command=lambda: webbrowser.open("https://github.com/rice2k"))
        menubar.add_cascade(label="Help", menu=help_menu)
        self.config(menu=menubar)

    def refresh_recent_menu(self) -> None:
        self.recent_menu.delete(0, "end")
        recent = [item for item in self.settings.get("recent_files", []) if Path(item).exists()]
        if not recent:
            self.recent_menu.add_command(label="No recent files", state="disabled")
            return
        for item in recent:
            self.recent_menu.add_command(label=Path(item).name, command=lambda p=item: self.import_file(Path(p)))

    def _load_icon(self, name: str, size: int = 18) -> ImageTk.PhotoImage | None:
        key = (name, size, size)
        if key in self.icon_photos:
            return self.icon_photos[key]
        path = asset_path(name)
        if not path.exists():
            return None
        image = Image.open(path).resize((size, size), Image.LANCZOS)
        photo = ImageTk.PhotoImage(image)
        self.icon_photos[key] = photo
        return photo

    def _button(
        self,
        parent: tk.Widget,
        text: str,
        command: Any,
        primary: bool = False,
        icon: str | None = None,
        tooltip: str | None = None,
    ) -> tk.Button:
        bg = "#e51b2d" if primary else "#202633"
        hover = "#b91525" if primary else "#2c3444"
        icon_photo = self._load_icon(icon, 18) if icon else None
        button = tk.Button(
            parent,
            text=text,
            image=icon_photo,
            compound="left" if icon_photo else "none",
            command=command,
            bg=bg,
            fg="#ffffff",
            activebackground=hover,
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground="#4a5261" if not primary else "#ff6d78",
            font=("Segoe UI", 9, "bold"),
            padx=13,
            pady=8,
            cursor="hand2",
        )
        button.bind("<Enter>", lambda _event: button.configure(bg=hover))
        button.bind("<Leave>", lambda _event: button.configure(bg=bg))
        if tooltip:
            ToolTip(button, tooltip)
        return button

    def _menu_button(
        self,
        parent: tk.Widget,
        text: str,
        items: list[tuple[str, Any] | None],
        icon: str | None = None,
        tooltip: str | None = None,
    ) -> tk.Menubutton:
        icon_photo = self._load_icon(icon, 18) if icon else None
        bg = "#202633"
        hover = "#2c3444"
        button = tk.Menubutton(
            parent,
            text=f"{text} v",
            image=icon_photo,
            compound="left" if icon_photo else "none",
            bg=bg,
            fg="#ffffff",
            activebackground=hover,
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground="#4a5261",
            font=("Segoe UI", 9, "bold"),
            padx=13,
            pady=8,
            cursor="hand2",
        )
        menu = tk.Menu(
            button,
            tearoff=False,
            bg="#151922",
            fg="#ffffff",
            activebackground="#e51b2d",
            activeforeground="#ffffff",
            disabledforeground="#6f7a8b",
            bd=0,
            relief="flat",
        )
        for item in items:
            if item is None:
                menu.add_separator()
            else:
                label, command = item
                menu.add_command(label=label, command=command)
        button.configure(menu=menu)
        button.bind("<Enter>", lambda _event: button.configure(bg=hover))
        button.bind("<Leave>", lambda _event: button.configure(bg=bg))
        if tooltip:
            ToolTip(button, tooltip)
        return button

    def _panel(self, parent: tk.Widget) -> tk.Frame:
        return tk.Frame(parent, bg="#10141b", highlightthickness=1, highlightbackground="#323b49")

    def _link_label(self, parent: tk.Widget, text: str, url: str, bg: str, tooltip: str | None = None) -> tk.Label:
        label = tk.Label(
            parent,
            text=text,
            bg=bg,
            fg="#46d9ff",
            cursor="hand2",
            padx=10,
            font=("Segoe UI", 9, "underline"),
        )
        label.bind("<Button-1>", lambda _event: webbrowser.open(url))
        label.bind("<Enter>", lambda _event: label.configure(fg="#8beaff"))
        label.bind("<Leave>", lambda _event: label.configure(fg="#46d9ff"))
        if tooltip:
            ToolTip(label, tooltip)
        return label

    def _nav_section(self, parent: tk.Widget, text: str) -> None:
        tk.Label(
            parent,
            text=text.upper(),
            bg="#111721",
            fg="#6f7a8b",
            font=("Segoe UI", 8, "bold"),
            anchor="w",
        ).pack(fill="x", padx=14, pady=(14, 6))

    def _nav_button(
        self,
        parent: tk.Widget,
        text: str,
        index: int,
        image: ImageTk.PhotoImage | None = None,
        tooltip: str | None = None,
    ) -> tk.Button:
        button = tk.Button(
            parent,
            text=text,
            image=image,
            compound="left" if image else "none",
            command=lambda: self.select_tab(index),
            anchor="w",
            bg="#111721",
            fg="#a8b2c2",
            activebackground="#251017",
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground="#202734",
            font=("Segoe UI", 10, "bold"),
            padx=14,
            pady=11,
            cursor="hand2",
        )
        if tooltip:
            ToolTip(button, tooltip)
        return button

    def _load_nav_icon(self, name: str) -> ImageTk.PhotoImage | None:
        path = asset_path(name)
        if not path.exists():
            return None
        image = Image.open(path).resize((24, 24), Image.LANCZOS)
        photo = ImageTk.PhotoImage(image)
        self.nav_icon_photos.append(photo)
        return photo

    def _build_shell(self) -> None:
        header = tk.Frame(self, bg="#090b10", height=72)
        header.pack(fill="x")
        header.pack_propagate(False)

        logo_path = asset_path("macys-ap-icon.png")
        if not logo_path.exists():
            logo_path = asset_path("macys-ap-orb.png")
        if logo_path.exists():
            image = Image.open(logo_path).resize((50, 50), Image.LANCZOS)
            self.logo_photo = ImageTk.PhotoImage(image)
            tk.Label(header, image=self.logo_photo, bg="#090b10").pack(side="left", padx=(18, 12))

        title_box = tk.Frame(header, bg="#090b10")
        title_box.pack(side="left")
        tk.Label(title_box, text=APP_SHORT_NAME, bg="#090b10", fg="#ffffff", font=("Segoe UI", 14, "bold")).pack(anchor="w")
        tk.Label(title_box, text="Asset Protection access-control workstation", bg="#090b10", fg="#a8b2c2", font=("Segoe UI", 9)).pack(anchor="w")

        self.mode_var = tk.StringVar(value="Batch Converter")

        accent_path = asset_path("ap-window-accent.png")
        if accent_path.exists():
            accent_image = Image.open(accent_path).resize((420, 58), Image.LANCZOS)
            self.header_accent_photo = ImageTk.PhotoImage(accent_image)
            tk.Label(header, image=self.header_accent_photo, bg="#090b10").pack(side="right", padx=(0, 18))

        toolbar = tk.Frame(self, bg="#111721", height=50, highlightthickness=1, highlightbackground="#252d3a")
        toolbar.pack(fill="x")
        toolbar.pack_propagate(False)
        tk.Label(toolbar, text="AP COMMANDS", bg="#111721", fg="#ff6d78", font=("Segoe UI", 9, "bold")).pack(side="left", padx=(16, 10))
        self._menu_button(toolbar, "File", [
            ("Default Export Folder", self.choose_export_folder),
            ("Open Export Folder", self.open_export_folder),
            None,
            ("Exit", self.destroy),
        ], icon="icon-folder.png", tooltip="Set or open the folder where reports are saved.").pack(side="left", padx=(0, 8), pady=8)
        self._menu_button(toolbar, "Import", [
            ("Browse Files", self.import_file),
            ("Paste Clipboard To Queue", self.paste_clipboard_to_queue),
            ("Load Sample IDs", self.load_sample),
        ], icon="icon-import.png", tooltip="Import files, paste clipboard text, or drag files onto the Input Queue box.").pack(side="left", padx=(0, 8), pady=8)
        self._menu_button(toolbar, "Export", [
            ("Excel Workbook", self.export_excel),
            ("CSV Report", self.export_csv),
            ("TXT Report", self.export_txt),
            ("PDF Report", self.export_pdf),
        ], icon="icon-excel.png", tooltip="Save the current conversion results as a report.").pack(side="left", padx=(0, 8), pady=8)
        self._menu_button(toolbar, "Help", [
            ("How To Use", self.show_help),
            ("About This Utility", self.show_about),
            ("Open GitHub Profile", lambda: webbrowser.open("https://github.com/rice2k")),
        ], icon="icon-help.png", tooltip="Open usage help, app details, and support links.").pack(side="right", padx=(0, 16), pady=8)
        self._button(
            toolbar,
            "BlueWave",
            self.open_bluewave,
            icon="icon-bluewave.png",
            tooltip="Open the BlueWave access-control site in your browser.",
        ).pack(side="right", padx=(0, 8), pady=8)

        body = tk.Frame(self, bg="#0b0d12")
        body.pack(fill="both", expand=True, padx=16, pady=16)
        body.grid_rowconfigure(0, weight=1)
        body.grid_columnconfigure(1, weight=1)

        nav = self._panel(body)
        nav.grid(row=0, column=0, sticky="ns", padx=(0, 14))
        nav.configure(width=220, bg="#111721")
        nav.grid_propagate(False)
        tk.Label(nav, text="Workspace", bg="#111721", fg="#ffffff", font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=14, pady=(14, 2))
        tk.Label(nav, text="Choose one work area", bg="#111721", fg="#6f7a8b", font=("Segoe UI", 8)).pack(anchor="w", padx=14, pady=(0, 4))
        self.nav_buttons: list[tk.Button] = []
        nav_specs = [
            ("Convert", "Batch Converter", 0, "nav-batch.png"),
            ("Convert", "Single Lookup", 1, "nav-single.png"),
            ("Reverse", "FC/CN to Hex", 2, "nav-reverse.png"),
            ("Reverse", "Unconvert Batch", 3, "nav-unconvert.png"),
            ("Review", "History", 4, "nav-history.png"),
        ]
        nav_tips = {
            "Batch Converter": "Convert many HEX IDs into Facility Code and Card Number rows.",
            "Single Lookup": "Quickly convert one HEX ID and copy the FC/CN pair.",
            "FC/CN to Hex": "Create one HEX ID from a Facility Code and Card Number.",
            "Unconvert Batch": "Convert many FC/CN pairs back into HEX IDs.",
            "History": "Review recent conversion activity saved by this utility.",
        }
        last_section = ""
        for section, label, index, icon_name in nav_specs:
            if section != last_section:
                self._nav_section(nav, section)
                last_section = section
            button = self._nav_button(nav, label, index, self._load_nav_icon(icon_name), nav_tips.get(label))
            self.nav_buttons.append(button)
            button.pack(fill="x", padx=10, pady=(0, 8))
        tk.Frame(nav, bg="#1f2632", height=1).pack(fill="x", padx=12, pady=(8, 12))
        self.nav_status = tk.StringVar(value="Ready")
        tk.Label(nav, textvariable=self.nav_status, bg="#111721", fg="#46d9ff", wraplength=188, justify="left", font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=14, pady=(0, 8))

        content = tk.Frame(body, bg="#0b0d12")
        content.grid(row=0, column=1, sticky="nsew")
        content.grid_rowconfigure(0, weight=1)
        content.grid_columnconfigure(0, weight=1)

        self.batch_tab = tk.Frame(content, bg="#0b0d12")
        self.single_tab = tk.Frame(content, bg="#0b0d12")
        self.reverse_tab = tk.Frame(content, bg="#0b0d12")
        self.unconvert_tab = tk.Frame(content, bg="#0b0d12")
        self.history_tab = tk.Frame(content, bg="#0b0d12")
        self.tab_frames = [self.batch_tab, self.single_tab, self.reverse_tab, self.unconvert_tab, self.history_tab]
        for frame in self.tab_frames:
            frame.grid(row=0, column=0, sticky="nsew")

        self._build_batch_tab()
        self._build_single_tab()
        self._build_reverse_tab()
        self._build_unconvert_tab()
        self._build_history_tab()
        self.select_tab(0)

        footer = tk.Frame(self, bg="#090b10", height=34)
        footer.pack(fill="x")
        footer.pack_propagate(False)
        self.status_var = tk.StringVar(value="Ready")
        tk.Label(footer, textvariable=self.status_var, bg="#090b10", fg="#a8b2c2", anchor="w", padx=12).pack(side="left", fill="x", expand=True)
        github = self._link_label(footer, "github.com/rice2k", "https://github.com/rice2k", "#090b10", "Open Christopher Schumacher's GitHub profile.")
        github.pack(side="right")
        tk.Label(footer, text="Made by Christopher Schumacher, Asset Protection FLO", bg="#090b10", fg="#f5f7fb", padx=10).pack(side="right")

    def _build_batch_tab(self) -> None:
        top = tk.Frame(self.batch_tab, bg="#0b0d12")
        top.pack(fill="x", pady=(0, 12))

        input_panel = self._panel(top)
        input_panel.pack(side="left", fill="both", expand=True, padx=(0, 12))
        tk.Label(input_panel, text="Input Queue", bg="#10141b", fg="#ffffff", font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=12, pady=(12, 8))
        tk.Label(
            input_panel,
            text="Paste HEX IDs, import files, or drag TXT/CSV/Excel/XML/HTML files onto this box. The app extracts clean 8-character IDs and flags duplicates.",
            bg="#10141b",
            fg="#a8b2c2",
            font=("Segoe UI", 9),
            wraplength=650,
            justify="left",
        ).pack(anchor="w", padx=12, pady=(0, 8))
        multi_frame = tk.Frame(input_panel, bg="#10141b")
        multi_frame.pack(fill="both", expand=True, padx=12)
        multi_frame.rowconfigure(0, weight=1)
        multi_frame.columnconfigure(0, weight=1)
        self.multi_text = tk.Text(multi_frame, height=8, bg="#080a0f", fg="#f5f7fb", insertbackground="#46d9ff", relief="flat", padx=10, pady=10, font=("Cascadia Mono", 10), wrap="none")
        multi_scroll_y = ttk.Scrollbar(multi_frame, orient="vertical", command=self.multi_text.yview, style="Dark.Vertical.TScrollbar")
        multi_scroll_x = ttk.Scrollbar(multi_frame, orient="horizontal", command=self.multi_text.xview, style="Dark.Horizontal.TScrollbar")
        self.multi_text.configure(yscrollcommand=multi_scroll_y.set, xscrollcommand=multi_scroll_x.set)
        self.multi_text.grid(row=0, column=0, sticky="nsew")
        multi_scroll_y.grid(row=0, column=1, sticky="ns")
        multi_scroll_x.grid(row=1, column=0, sticky="ew")
        self.multi_text.insert("1.0", "")
        self.multi_text.bind("<<Modified>>", self.handle_batch_input_changed)
        self._enable_drop_target(self.multi_text)
        self._enable_drop_target(multi_frame)
        btns = tk.Frame(input_panel, bg="#10141b")
        btns.pack(fill="x", padx=12, pady=12)
        self._button(btns, "Import", self.import_file, icon="icon-import.png", tooltip="Browse for one or more supported files and add extracted IDs to the queue.").pack(side="left", padx=(0, 8))
        self._button(btns, "Sample", self.load_sample, icon="icon-sample.png", tooltip="Load sample data so you can see how conversion results look.").pack(side="left", padx=(0, 8))
        self._button(btns, "Convert All", self.convert_batch, True, icon="icon-convert.png", tooltip="Convert every queued HEX ID into Facility Code and Card Number.").pack(side="left", padx=(0, 8))
        self._button(btns, "Clear", self.clear_workspace, icon="icon-clear.png", tooltip="Clear input, results, and single-lookup fields.").pack(side="left")

        summary = self._panel(top)
        summary.pack(side="right", fill="y")
        summary.configure(width=280)
        summary.pack_propagate(False)
        tk.Label(summary, text="Run Summary", bg="#10141b", fg="#ffffff", font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=12, pady=(12, 8))
        tk.Label(summary, text="Reports save from the Export menu.", bg="#10141b", fg="#a8b2c2", font=("Segoe UI", 9), wraplength=230, justify="left").pack(anchor="w", padx=12, pady=(0, 10))
        self.stat_vars = {name: tk.StringVar(value="0") for name in ["Input", "Valid", "Invalid", "Warnings"]}
        grid = tk.Frame(summary, bg="#10141b")
        grid.pack(padx=12, pady=(4, 8))
        for idx, name in enumerate(self.stat_vars):
            card = tk.Frame(grid, bg="#151922", width=112, height=66, highlightthickness=1, highlightbackground="#303845")
            card.grid(row=idx // 2, column=idx % 2, padx=5, pady=5)
            card.grid_propagate(False)
            tk.Label(card, text=name, bg="#151922", fg="#a8b2c2", font=("Segoe UI", 8, "bold")).pack(pady=(8, 0))
            tk.Label(card, textvariable=self.stat_vars[name], bg="#151922", fg="#46d9ff", font=("Segoe UI", 16, "bold")).pack()
        self.time_var = tk.StringVar(value="No run yet")
        tk.Label(summary, textvariable=self.time_var, bg="#151922", fg="#f5f7fb", wraplength=230, justify="left", padx=10, pady=8).pack(fill="x", padx=12, pady=(4, 8))
        self.notice_var = tk.StringVar(value="")
        tk.Label(summary, textvariable=self.notice_var, bg="#10141b", fg="#ffca63", wraplength=230, justify="left").pack(anchor="w", padx=12, pady=(0, 12))
        results_panel = self._panel(self.batch_tab)
        results_panel.pack(fill="both", expand=True)
        header = tk.Frame(results_panel, bg="#10141b")
        header.pack(fill="x", padx=12, pady=(12, 8))
        tk.Label(header, text="Results", bg="#10141b", fg="#ffffff", font=("Segoe UI", 12, "bold")).pack(side="left")
        copy_row = tk.Frame(header, bg="#10141b")
        copy_row.pack(side="right")
        self._button(copy_row, "Copy All", self.copy_all_pairs, icon="icon-copy.png", tooltip="Copy all valid FC,CN pairs to the clipboard.").pack(side="left", padx=(0, 6))
        self._button(copy_row, "Clear Invalid", self.clear_invalid_rows, icon="icon-clear.png", tooltip="Remove invalid rows from the review table.").pack(side="left", padx=(0, 6))
        self._button(copy_row, "Copy FC", lambda: self.copy_selected("fc"), icon="icon-copy.png", tooltip="Copy the selected row's Facility Code.").pack(side="left", padx=(0, 6))
        self._button(copy_row, "Copy CN", lambda: self.copy_selected("cn"), icon="icon-copy.png", tooltip="Copy the selected row's Card Number.").pack(side="left", padx=(0, 6))
        self._button(copy_row, "Copy Pair", lambda: self.copy_selected("pair"), icon="icon-copy.png", tooltip="Copy the selected row as FC,CN.").pack(side="left")

        filter_row = tk.Frame(results_panel, bg="#10141b")
        filter_row.pack(fill="x", padx=12, pady=(0, 8))
        tk.Label(filter_row, text="Search", bg="#10141b", fg="#a8b2c2", font=("Segoe UI", 9, "bold")).pack(side="left", padx=(0, 8))
        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(filter_row, textvariable=self.search_var, bg="#080a0f", fg="#f5f7fb", insertbackground="#46d9ff", relief="flat", width=32)
        self.search_entry.pack(side="left", ipady=7, padx=(0, 10))
        self.search_var.trace_add("write", lambda *_args: self.render_results())
        tk.Label(filter_row, text="Status", bg="#10141b", fg="#a8b2c2", font=("Segoe UI", 9, "bold")).pack(side="left", padx=(0, 8))
        self.status_filter_var = tk.StringVar(value="All")
        status_menu = tk.OptionMenu(filter_row, self.status_filter_var, "All", "Valid", "Warning", "Invalid", command=lambda _value: self.render_results())
        status_menu.configure(bg="#202633", fg="#ffffff", activebackground="#2c3444", activeforeground="#ffffff", relief="flat", highlightthickness=1, highlightbackground="#4a5261")
        status_menu["menu"].configure(bg="#151922", fg="#ffffff", activebackground="#e51b2d", activeforeground="#ffffff")
        status_menu.pack(side="left")

        self.results_empty_frame = tk.Frame(results_panel, bg="#10141b", highlightthickness=1, highlightbackground="#263141")
        self.results_empty_frame.pack(fill="x", padx=12, pady=(0, 8))
        empty_path = asset_path("empty-results.png")
        if empty_path.exists():
            empty_image = Image.open(empty_path).resize((246, 106), Image.LANCZOS)
            self.empty_results_photo = ImageTk.PhotoImage(empty_image)
            tk.Label(self.results_empty_frame, image=self.empty_results_photo, bg="#10141b").pack(side="left", padx=12, pady=10)
        empty_copy = tk.Frame(self.results_empty_frame, bg="#10141b")
        empty_copy.pack(side="left", fill="both", expand=True, padx=(4, 12), pady=12)
        tk.Label(empty_copy, text="No results yet", bg="#10141b", fg="#ffffff", font=("Segoe UI", 12, "bold")).pack(anchor="w")
        tk.Label(empty_copy, text="Paste or import access-control IDs, then use Convert All. Results, warnings, and exports will appear here.", bg="#10141b", fg="#a8b2c2", wraplength=520, justify="left", font=("Segoe UI", 9)).pack(anchor="w", pady=(5, 0))

        table_frame = tk.Frame(results_panel, bg="#10141b")
        self.results_table_frame = table_frame
        table_frame.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        columns = ("Line", "Hex ID", "FC", "CN", "Status", "Notes")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings")
        for col in columns:
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_results_by(c))
        widths = {"Line": 60, "Hex ID": 150, "FC": 90, "CN": 90, "Status": 120, "Notes": 420}
        for col, width in widths.items():
            self.tree.column(col, width=width, minwidth=60, stretch=col == "Notes")
        yscroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview, style="Dark.Vertical.TScrollbar")
        xscroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview, style="Dark.Horizontal.TScrollbar")
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)
        self.tree.tag_configure("invalid", background="#2a2119", foreground="#ffd9a0")
        self.tree.tag_configure("warning", background="#1d1a12", foreground="#ffe1a3")
        self.tree.tag_configure("empty", foreground="#a8b2c2")
        self.tree.bind("<Double-1>", lambda _event: self.copy_selected("pair"))

    def _build_single_tab(self) -> None:
        panel = self._panel(self.single_tab)
        panel.pack(fill="x", padx=2, pady=2)
        tk.Label(panel, text="Single Hex Lookup", bg="#10141b", fg="#ffffff", font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=12, pady=(12, 8))
        self.single_var = tk.StringVar()
        entry = tk.Entry(panel, textvariable=self.single_var, bg="#080a0f", fg="#f5f7fb", insertbackground="#46d9ff", relief="flat", font=("Cascadia Mono", 12))
        entry.pack(fill="x", padx=12, ipady=9)
        entry.bind("<Return>", lambda _event: self.convert_single())
        row = tk.Frame(panel, bg="#10141b")
        row.pack(fill="x", padx=12, pady=12)
        self._button(row, "Convert", self.convert_single, True, icon="icon-convert.png", tooltip="Convert one HEX ID and copy the FC,CN pair.").pack(side="left", padx=(0, 8))
        self._button(row, "Clear", self.clear_single, icon="icon-clear.png", tooltip="Clear the single lookup field.").pack(side="left")
        self.single_result = tk.StringVar(value="Waiting for a hex ID.")
        tk.Label(panel, textvariable=self.single_result, bg="#151922", fg="#46d9ff", anchor="w", justify="left", padx=12, pady=16, font=("Cascadia Mono", 13), wraplength=900).pack(fill="x", padx=12, pady=(0, 12))

    def _build_reverse_tab(self) -> None:
        panel = self._panel(self.reverse_tab)
        panel.pack(fill="x", padx=2, pady=2)
        tk.Label(panel, text="FC/CN to Hex", bg="#10141b", fg="#ffffff", font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=12, pady=(12, 8))
        self.fc_var = tk.StringVar()
        self.cn_var = tk.StringVar()
        for label, var in [("Facility Code", self.fc_var), ("Card Number", self.cn_var)]:
            tk.Label(panel, text=label, bg="#10141b", fg="#a8b2c2").pack(anchor="w", padx=12, pady=(8, 4))
            tk.Entry(panel, textvariable=var, bg="#080a0f", fg="#f5f7fb", insertbackground="#46d9ff", relief="flat", font=("Cascadia Mono", 12)).pack(fill="x", padx=12, ipady=9)
        row = tk.Frame(panel, bg="#10141b")
        row.pack(fill="x", padx=12, pady=12)
        self._button(row, "Convert", self.convert_reverse, True, icon="icon-convert.png", tooltip="Build one 8-character HEX ID from FC and CN.").pack(side="left", padx=(0, 8))
        self._button(row, "Clear", self.clear_reverse, icon="icon-clear.png", tooltip="Clear the FC and CN fields.").pack(side="left")
        self.reverse_result = tk.StringVar(value="Waiting for FC and CN values.")
        tk.Label(panel, textvariable=self.reverse_result, bg="#151922", fg="#46d9ff", anchor="w", justify="left", padx=12, pady=16, font=("Cascadia Mono", 13), wraplength=900).pack(fill="x", padx=12, pady=(0, 12))

    def _build_unconvert_tab(self) -> None:
        top = tk.Frame(self.unconvert_tab, bg="#0b0d12")
        top.pack(fill="x", pady=(0, 12))

        input_panel = self._panel(top)
        input_panel.pack(side="left", fill="both", expand=True, padx=(0, 12))
        tk.Label(input_panel, text="Unconvert FC/CN Batch", bg="#10141b", fg="#ffffff", font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=12, pady=(12, 8))
        tk.Label(input_panel, text="Paste one pair per line. Examples: 34968,18199 or FC 34968 CN 18199.", bg="#10141b", fg="#a8b2c2", font=("Segoe UI", 9)).pack(anchor="w", padx=12, pady=(0, 8))
        unconvert_input_frame = tk.Frame(input_panel, bg="#10141b")
        unconvert_input_frame.pack(fill="both", expand=True, padx=12)
        unconvert_input_frame.rowconfigure(0, weight=1)
        unconvert_input_frame.columnconfigure(0, weight=1)
        self.unconvert_text = tk.Text(unconvert_input_frame, height=8, bg="#080a0f", fg="#f5f7fb", insertbackground="#46d9ff", relief="flat", padx=10, pady=10, font=("Cascadia Mono", 10), wrap="none")
        unconvert_scroll_y = ttk.Scrollbar(unconvert_input_frame, orient="vertical", command=self.unconvert_text.yview, style="Dark.Vertical.TScrollbar")
        unconvert_scroll_x = ttk.Scrollbar(unconvert_input_frame, orient="horizontal", command=self.unconvert_text.xview, style="Dark.Horizontal.TScrollbar")
        self.unconvert_text.configure(yscrollcommand=unconvert_scroll_y.set, xscrollcommand=unconvert_scroll_x.set)
        self.unconvert_text.grid(row=0, column=0, sticky="nsew")
        unconvert_scroll_y.grid(row=0, column=1, sticky="ns")
        unconvert_scroll_x.grid(row=1, column=0, sticky="ew")
        self.unconvert_text.bind("<<Modified>>", self.handle_unconvert_input_changed)
        row = tk.Frame(input_panel, bg="#10141b")
        row.pack(fill="x", padx=12, pady=12)
        self._button(row, "Sample", self.load_unconvert_sample, icon="icon-sample.png", tooltip="Load sample FC/CN pairs for the unconvert workflow.").pack(side="left", padx=(0, 8))
        self._button(row, "Unconvert All", self.convert_unconvert_batch, True, icon="icon-convert.png", tooltip="Convert every queued FC/CN pair back into HEX.").pack(side="left", padx=(0, 8))
        self._button(row, "Copy All HEX", self.copy_all_unconverted_hex, icon="icon-copy.png", tooltip="Copy all valid unconverted HEX IDs.").pack(side="left", padx=(0, 8))
        self._button(row, "Clear", self.clear_unconvert, icon="icon-clear.png", tooltip="Clear the unconvert input and results.").pack(side="left")

        summary = self._panel(top)
        summary.pack(side="right", fill="y")
        summary.configure(width=280)
        summary.pack_propagate(False)
        tk.Label(summary, text="Unconvert Summary", bg="#10141b", fg="#ffffff", font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=12, pady=(12, 8))
        self.unconvert_summary_var = tk.StringVar(value="No unconvert run yet")
        tk.Label(summary, textvariable=self.unconvert_summary_var, bg="#10141b", fg="#46d9ff", wraplength=235, justify="left", font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=12, pady=(8, 12))
        tk.Label(summary, text="This tab reverses FC/CN pairs back into 8-character HEX IDs.", bg="#10141b", fg="#a8b2c2", wraplength=235, justify="left").pack(anchor="w", padx=12, pady=(0, 12))

        results_panel = self._panel(self.unconvert_tab)
        results_panel.pack(fill="both", expand=True)
        header = tk.Frame(results_panel, bg="#10141b")
        header.pack(fill="x", padx=12, pady=(12, 8))
        tk.Label(header, text="Unconvert Results", bg="#10141b", fg="#ffffff", font=("Segoe UI", 12, "bold")).pack(side="left")
        self._button(header, "Copy Selected HEX", self.copy_selected_unconvert_hex, icon="icon-copy.png", tooltip="Copy the selected unconverted HEX ID.").pack(side="right")

        self.unconvert_empty_frame = tk.Frame(results_panel, bg="#10141b", highlightthickness=1, highlightbackground="#263141")
        self.unconvert_empty_frame.pack(fill="x", padx=12, pady=(0, 8))
        if self.empty_results_photo:
            tk.Label(self.unconvert_empty_frame, image=self.empty_results_photo, bg="#10141b").pack(side="left", padx=12, pady=10)
        empty_copy = tk.Frame(self.unconvert_empty_frame, bg="#10141b")
        empty_copy.pack(side="left", fill="both", expand=True, padx=(4, 12), pady=12)
        tk.Label(empty_copy, text="Ready to rebuild HEX IDs", bg="#10141b", fg="#ffffff", font=("Segoe UI", 12, "bold")).pack(anchor="w")
        tk.Label(empty_copy, text="Paste FC/CN pairs, run Unconvert All, then copy the returned HEX IDs from this review area.", bg="#10141b", fg="#a8b2c2", wraplength=520, justify="left", font=("Segoe UI", 9)).pack(anchor="w", pady=(5, 0))

        table_frame = tk.Frame(results_panel, bg="#10141b")
        self.unconvert_table_frame = table_frame
        table_frame.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        columns = ("Line", "FC", "CN", "Hex ID", "Status", "Notes")
        self.unconvert_tree = ttk.Treeview(table_frame, columns=columns, show="headings")
        for col in columns:
            self.unconvert_tree.heading(col, text=col)
        widths = {"Line": 60, "FC": 120, "CN": 120, "Hex ID": 150, "Status": 120, "Notes": 430}
        for col, width in widths.items():
            self.unconvert_tree.column(col, width=width, minwidth=60, stretch=col == "Notes")
        yscroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.unconvert_tree.yview, style="Dark.Vertical.TScrollbar")
        xscroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.unconvert_tree.xview, style="Dark.Horizontal.TScrollbar")
        self.unconvert_tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.unconvert_tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)
        self.unconvert_tree.tag_configure("invalid", background="#2a2119", foreground="#ffd9a0")
        self.unconvert_tree.tag_configure("warning", background="#1d1a12", foreground="#ffe1a3")
        self.unconvert_tree.tag_configure("empty", foreground="#a8b2c2")

    def _build_history_tab(self) -> None:
        panel = self._panel(self.history_tab)
        panel.pack(fill="both", expand=True, padx=2, pady=2)
        header = tk.Frame(panel, bg="#10141b")
        header.pack(fill="x", padx=12, pady=(12, 8))
        tk.Label(header, text="Conversion History", bg="#10141b", fg="#ffffff", font=("Segoe UI", 12, "bold")).pack(side="left")
        self._button(header, "Clear History", self.clear_history, icon="icon-clear.png", tooltip="Delete saved conversion history from this app.").pack(side="right")
        history_frame = tk.Frame(panel, bg="#10141b")
        history_frame.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        history_frame.rowconfigure(0, weight=1)
        history_frame.columnconfigure(0, weight=1)
        self.history_text = tk.Text(history_frame, bg="#080a0f", fg="#f5f7fb", relief="flat", padx=14, pady=14, wrap="word", font=("Cascadia Mono", 10))
        history_scroll = ttk.Scrollbar(history_frame, orient="vertical", command=self.history_text.yview, style="Dark.Vertical.TScrollbar")
        self.history_text.configure(yscrollcommand=history_scroll.set)
        self.history_text.grid(row=0, column=0, sticky="nsew")
        history_scroll.grid(row=0, column=1, sticky="ns")
        self.history_text.configure(state="disabled")
        self.render_history()

    def _mode_changed(self, _event: Any) -> None:
        self.select_tab(["Batch Converter", "Single Hex Lookup", "FC/CN to Hex", "Unconvert Batch", "History"].index(self.mode_var.get()))

    def _tab_changed(self, _event: Any) -> None:
        return

    def select_tab(self, index: int) -> None:
        labels = ["Batch Converter", "Single Hex Lookup", "FC/CN to Hex", "Unconvert Batch", "History"]
        self.tab_frames[index].tkraise()
        self.mode_var.set(labels[index])
        for i, button in enumerate(self.nav_buttons):
            if i == index:
                button.configure(bg="#251017", fg="#ffffff", highlightbackground="#e51b2d")
            else:
                button.configure(bg="#111721", fg="#a8b2c2", highlightbackground="#202734")
        if index == 4:
            self.render_history()

    def set_status(self, message: str) -> None:
        self.status_var.set(message)
        if hasattr(self, "nav_status"):
            self.nav_status.set(message)

    def _enable_drop_target(self, widget: tk.Widget) -> None:
        if DND_FILES is None or not hasattr(widget, "drop_target_register"):
            return
        try:
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", self.handle_file_drop)
            widget.dnd_bind("<<DragEnter>>", self.handle_drag_enter)
            widget.dnd_bind("<<DragLeave>>", self.handle_drag_leave)
        except Exception:
            return

    def handle_drag_enter(self, event: Any) -> str | None:
        if hasattr(self, "multi_text"):
            self.multi_text.configure(bg="#0d1722")
        self.set_status("Drop supported files onto the Input Queue to import them.")
        return getattr(event, "action", None)

    def handle_drag_leave(self, event: Any) -> str | None:
        if hasattr(self, "multi_text"):
            self.multi_text.configure(bg="#080a0f")
        return getattr(event, "action", None)

    def handle_file_drop(self, event: Any) -> str | None:
        if hasattr(self, "multi_text"):
            self.multi_text.configure(bg="#080a0f")
        try:
            paths = [Path(item) for item in self.tk.splitlist(event.data)]
        except Exception:
            paths = [Path(str(event.data))]
        self.import_paths_to_queue(paths, preview=False)
        return getattr(event, "action", None)

    def append_text_to_queue(self, text: str) -> None:
        incoming = text.strip()
        if not incoming:
            return
        current = self.multi_text.get("1.0", "end").strip()
        self.multi_text.delete("1.0", "end")
        self.multi_text.insert("1.0", f"{current}\n{incoming}".strip() if current else incoming)
        self.multi_text.edit_modified(True)
        self.handle_batch_input_changed()

    def paste_clipboard_to_queue(self) -> None:
        try:
            text = self.clipboard_get()
        except tk.TclError:
            messagebox.showinfo("Clipboard empty", "There is no text on the clipboard to paste.")
            return
        if not text.strip():
            messagebox.showinfo("Clipboard empty", "There is no text on the clipboard to paste.")
            return
        self.append_text_to_queue(text)
        self.set_status("Clipboard text added to the Input Queue.")

    def _count_nonempty_text_lines(self, widget: tk.Text) -> int:
        text = widget.get("1.0", "end").strip()
        return len([line for line in text.splitlines() if line.strip()]) if text else 0

    def handle_batch_input_changed(self, _event: Any | None = None) -> None:
        if not hasattr(self, "multi_text") or not self.multi_text.edit_modified():
            return
        self.multi_text.edit_modified(False)
        queued = self._count_nonempty_text_lines(self.multi_text)
        if hasattr(self, "stat_vars"):
            self.stat_vars["Input"].set(str(queued))
        if not hasattr(self, "time_var"):
            return
        if queued == 0:
            if not self.results and not self.invalid:
                self.time_var.set("No run yet")
                self.notice_var.set("")
            return
        if self.results or self.invalid:
            self.time_var.set(f"{queued} line(s) in queue")
            self.notice_var.set("Input changed. Run Convert All to refresh results.")
        else:
            self.time_var.set(f"{queued} line(s) queued")
            self.notice_var.set("")

    def handle_unconvert_input_changed(self, _event: Any | None = None) -> None:
        if not hasattr(self, "unconvert_text") or not self.unconvert_text.edit_modified():
            return
        self.unconvert_text.edit_modified(False)
        queued = self._count_nonempty_text_lines(self.unconvert_text)
        if not hasattr(self, "unconvert_summary_var"):
            return
        if queued:
            self.unconvert_summary_var.set(f"{queued} line(s) queued")
        elif not self.unconvert_results and not self.unconvert_invalid:
            self.unconvert_summary_var.set("No unconvert run yet")

    def load_sample(self) -> None:
        self.multi_text.delete("1.0", "end")
        self.multi_text.insert("1.0", "88984717\n88984130\nActive Christopher Benson, 88984765\nBAD-LINE")
        self.multi_text.edit_modified(True)
        self.handle_batch_input_changed()
        self.set_status("Sample IDs loaded.")

    def load_unconvert_sample(self) -> None:
        self.unconvert_text.delete("1.0", "end")
        self.unconvert_text.insert("1.0", "34968,18199\nFC 34968 CN 18192\n34968\t18277\nBAD-LINE")
        self.unconvert_text.edit_modified(True)
        self.handle_unconvert_input_changed()
        self.set_status("Unconvert sample loaded.")

    def clear_workspace(self) -> None:
        self.multi_text.delete("1.0", "end")
        self.results.clear()
        self.invalid.clear()
        self.unconvert_results.clear()
        self.unconvert_invalid.clear()
        self.last_converted_at = ""
        self.notice_var.set("")
        self.time_var.set("No run yet")
        self.single_var.set("")
        self.single_result.set("Waiting for a hex ID.")
        self.reverse_result.set("Waiting for FC and CN values.")
        self.fc_var.set("")
        self.cn_var.set("")
        if hasattr(self, "unconvert_text"):
            self.unconvert_text.delete("1.0", "end")
            self.unconvert_summary_var.set("No unconvert run yet")
            self.render_unconvert_results()
        self.render_results()
        self.set_status("Workspace cleared.")

    def convert_batch(self) -> None:
        converted_at = datetime.now().strftime("%m/%d/%Y %I:%M:%S %p")
        self.results, self.invalid = convert_lines(self.multi_text.get("1.0", "end"), converted_at)
        self.last_converted_at = converted_at
        self.render_results()
        total = len(self.results) + len(self.invalid)
        warn_count = sum(len(row.warnings) for row in self.results)
        self.add_history("HEX to FC/CN", total, len(self.results), len(self.invalid), warn_count)
        self.set_status(f"Converted {len(self.results)} valid line(s); {len(self.invalid)} invalid line(s).")
        if total == 0:
            messagebox.showinfo("No input", "Add at least one ID before converting.")

    def convert_unconvert_batch(self) -> None:
        converted_at = datetime.now().strftime("%m/%d/%Y %I:%M:%S %p")
        self.unconvert_results, self.unconvert_invalid = unconvert_lines(self.unconvert_text.get("1.0", "end"), converted_at)
        self.render_unconvert_results()
        total = len(self.unconvert_results) + len(self.unconvert_invalid)
        warn_count = sum(len(row.warnings) for row in self.unconvert_results)
        self.add_history("FC/CN to HEX", total, len(self.unconvert_results), len(self.unconvert_invalid), warn_count)
        self.set_status(f"Unconverted {len(self.unconvert_results)} valid pair(s); {len(self.unconvert_invalid)} invalid line(s).")
        if total == 0:
            messagebox.showinfo("No input", "Add at least one FC/CN pair before unconverting.")

    def clear_unconvert(self) -> None:
        self.unconvert_text.delete("1.0", "end")
        self.unconvert_results.clear()
        self.unconvert_invalid.clear()
        self.unconvert_summary_var.set("No unconvert run yet")
        self.render_unconvert_results()
        self.set_status("Unconvert workspace cleared.")

    def sort_results_by(self, column: str) -> None:
        if self.sort_column == column:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column = column
            self.sort_reverse = False
        self.render_results()

    def render_results(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.row_lookup.clear()
        combined: list[tuple[int, str, ConvertedRow | InvalidRow]] = [(row.line, "valid", row) for row in self.results] + [(row.line, "invalid", row) for row in self.invalid]
        query = self.search_var.get().strip().lower() if hasattr(self, "search_var") else ""
        status_filter = self.status_filter_var.get() if hasattr(self, "status_filter_var") else "All"
        filtered: list[tuple[int, str, ConvertedRow | InvalidRow]] = []
        if hasattr(self, "results_empty_frame"):
            if combined:
                self.results_empty_frame.pack_forget()
            elif not self.results_empty_frame.winfo_ismapped():
                self.results_empty_frame.pack(fill="x", padx=12, pady=(0, 8), before=self.results_table_frame)
        for line, kind, row in combined:
            values = [str(line)]
            if isinstance(row, InvalidRow):
                status = "Invalid"
                values.extend([row.raw, row.reason])
            else:
                status = "Warning" if row.warnings else "Valid"
                values.extend([row.hex_value, str(row.facility), str(row.card), " ".join([*row.suggestions, *row.warnings])])
            if status_filter != "All" and status != status_filter:
                continue
            if query and query not in " ".join(values).lower():
                continue
            filtered.append((line, kind, row))
        def sort_key(item: tuple[int, str, ConvertedRow | InvalidRow]) -> Any:
            _line, kind, row = item
            if self.sort_column == "Line":
                return row.line
            if isinstance(row, InvalidRow):
                mapping = {"Hex ID": row.raw, "FC": "", "CN": "", "Status": "Invalid", "Notes": row.reason}
            else:
                mapping = {"Hex ID": row.hex_value, "FC": row.facility, "CN": row.card, "Status": "Warning" if row.warnings else "Valid", "Notes": " ".join([*row.suggestions, *row.warnings])}
            return mapping.get(self.sort_column, row.line)
        filtered.sort(key=sort_key, reverse=self.sort_reverse)
        if not filtered and combined:
            self.tree.insert("", "end", values=("", "No matching results.", "", "", "", ""), tags=("empty",))
        for _line, kind, row in filtered:
            if kind == "invalid":
                item_id = self.tree.insert("", "end", values=(row.line, row.raw, "", "", "Invalid", row.reason), tags=("invalid",))
            else:
                notes = " | ".join([*row.suggestions, *row.warnings])
                status = "Warning" if row.warnings else "Valid"
                tags = ("warning",) if row.warnings else ()
                item_id = self.tree.insert("", "end", values=(row.line, row.hex_value, row.facility, row.card, status, notes), tags=tags)
            self.row_lookup[item_id] = row
        warn_count = sum(len(row.warnings) for row in self.results)
        self.stat_vars["Input"].set(str(len(self.results) + len(self.invalid)))
        self.stat_vars["Valid"].set(str(len(self.results)))
        self.stat_vars["Invalid"].set(str(len(self.invalid)))
        self.stat_vars["Warnings"].set(str(warn_count))
        self.time_var.set(f"Converted {self.last_converted_at}" if self.last_converted_at else "No run yet")
        self.notice_var.set(f"{warn_count} warning(s) | {len(self.invalid)} invalid line(s)" if self.results or self.invalid else "")

    def render_unconvert_results(self) -> None:
        for item in self.unconvert_tree.get_children():
            self.unconvert_tree.delete(item)
        self.unconvert_row_lookup.clear()
        combined: list[tuple[int, str, UnconvertRow | InvalidRow]] = [(row.line, "valid", row) for row in self.unconvert_results] + [(row.line, "invalid", row) for row in self.unconvert_invalid]
        if hasattr(self, "unconvert_empty_frame"):
            if combined:
                self.unconvert_empty_frame.pack_forget()
            elif not self.unconvert_empty_frame.winfo_ismapped():
                self.unconvert_empty_frame.pack(fill="x", padx=12, pady=(0, 8), before=self.unconvert_table_frame)
        if not combined:
            return
        for _line, kind, row in sorted(combined, key=lambda item: item[0]):
            if kind == "invalid":
                item_id = self.unconvert_tree.insert("", "end", values=(row.line, "", "", row.raw, "Invalid", row.reason), tags=("invalid",))
            else:
                notes = " | ".join(row.warnings)
                status = "Warning" if row.warnings else "Valid"
                tags = ("warning",) if row.warnings else ()
                item_id = self.unconvert_tree.insert("", "end", values=(row.line, row.facility, row.card, row.hex_value, status, notes), tags=tags)
            self.unconvert_row_lookup[item_id] = row
        warn_count = sum(len(row.warnings) for row in self.unconvert_results)
        total = len(self.unconvert_results) + len(self.unconvert_invalid)
        self.unconvert_summary_var.set(f"Input: {total}\nValid: {len(self.unconvert_results)}\nInvalid: {len(self.unconvert_invalid)}\nWarnings: {warn_count}")

    def copy_selected(self, mode: str) -> None:
        selected = self.tree.selection()
        if not selected:
            self.set_status("Select a result row first.")
            return
        row = self.row_lookup.get(selected[0])
        if row is None:
            self.set_status("Select a converted result row first.")
            return
        if isinstance(row, InvalidRow):
            text = row.raw
        elif mode == "fc":
            text = str(row.facility)
        elif mode == "cn":
            text = str(row.card)
        else:
            text = f"{row.facility},{row.card}"
        self.clipboard_clear()
        self.clipboard_append(text)
        self.set_status("Copied to clipboard.")

    def copy_all_pairs(self) -> None:
        if not self.results:
            self.set_status("No valid FC/CN pairs to copy.")
            return
        text = "\n".join(f"{row.facility},{row.card}" for row in self.results)
        self.clipboard_clear()
        self.clipboard_append(text)
        self.set_status(f"Copied {len(self.results)} FC/CN pair(s).")

    def clear_invalid_rows(self) -> None:
        if not self.invalid:
            self.set_status("No invalid rows to clear.")
            return
        count = len(self.invalid)
        self.invalid.clear()
        self.render_results()
        self.set_status(f"Cleared {count} invalid row(s).")

    def copy_selected_unconvert_hex(self) -> None:
        selected = self.unconvert_tree.selection()
        if not selected:
            self.set_status("Select an unconvert result first.")
            return
        row = self.unconvert_row_lookup.get(selected[0])
        if not isinstance(row, UnconvertRow):
            self.set_status("Select a valid unconvert result first.")
            return
        self.clipboard_clear()
        self.clipboard_append(row.hex_value)
        self.set_status("Copied HEX to clipboard.")

    def copy_all_unconverted_hex(self) -> None:
        if not self.unconvert_results:
            self.set_status("No unconverted HEX values to copy.")
            return
        self.clipboard_clear()
        self.clipboard_append("\n".join(row.hex_value for row in self.unconvert_results))
        self.set_status(f"Copied {len(self.unconvert_results)} HEX value(s).")

    def choose_export_folder(self) -> None:
        folder = filedialog.askdirectory(
            title="Choose default export folder",
            initialdir=self.settings.get("default_export_dir") or str(Path.home() / "Desktop"),
        )
        if not folder:
            return
        self.settings["default_export_dir"] = folder
        self.save_settings()
        self.set_status(f"Default export folder set: {folder}")

    def open_export_folder(self) -> None:
        folder = Path(self.default_export_dir())
        folder.mkdir(parents=True, exist_ok=True)
        try:
            os.startfile(folder)  # type: ignore[attr-defined]
            self.set_status(f"Opened export folder: {folder}")
        except OSError as exc:
            messagebox.showerror("Folder unavailable", f"Could not open the export folder.\n\n{exc}")

    def default_export_dir(self) -> str:
        folder = Path(self.settings.get("default_export_dir") or Path.home() / "Desktop")
        return str(folder if folder.exists() else Path.home())

    def open_bluewave(self) -> None:
        webbrowser.open(BLUEWAVE_URL)
        self.set_status("Opened BlueWave in your browser.")

    def render_history(self) -> None:
        if not hasattr(self, "history_text"):
            return
        history = self.settings.get("history", [])
        lines = [f"{APP_SHORT_NAME} Conversion History", ""]
        if not history:
            lines.append("No conversion history yet.")
        for item in history:
            lines.append(
                f"{item.get('timestamp', '')} | {item.get('action', '')} | "
                f"input {item.get('total', 0)} | valid {item.get('valid', 0)} | "
                f"invalid {item.get('invalid', 0)} | warnings {item.get('warnings', 0)}"
            )
        self.history_text.configure(state="normal")
        self.history_text.delete("1.0", "end")
        self.history_text.insert("1.0", "\n".join(lines))
        self.history_text.configure(state="disabled")

    def clear_history(self) -> None:
        self.settings["history"] = []
        self.save_settings()
        self.render_history()
        self.set_status("Conversion history cleared.")

    def convert_single(self) -> None:
        prepared = clean_candidate_line(self.single_var.get())
        hex_value = prepared["extracted"].strip().upper()
        if not valid_hex8(hex_value):
            messagebox.showerror("Invalid hex", f"Single hex is invalid: {invalid_reason(prepared['original'])}.")
            return
        facility, card = hex_to_fc_cn(hex_value)
        notes = [*prepared["suggestions"], *unusual_warnings(hex_value, facility, card)]
        self.single_result.set(f"HEX: {hex_value}    FC: {facility}    CN: {card}" + (f"\n{' '.join(notes)}" if notes else ""))
        self.clipboard_clear()
        self.clipboard_append(f"{facility},{card}")
        self.set_status("Single ID converted and FC,CN copied.")

    def clear_single(self) -> None:
        self.single_var.set("")
        self.single_result.set("Waiting for a hex ID.")
        self.set_status("Single lookup cleared.")

    def clear_reverse(self) -> None:
        self.fc_var.set("")
        self.cn_var.set("")
        self.reverse_result.set("Waiting for FC and CN values.")

    def convert_reverse(self) -> None:
        try:
            hex_value = fc_cn_to_hex(self.fc_var.get(), self.cn_var.get())
            warnings = unusual_warnings(hex_value, int(self.fc_var.get()), int(self.cn_var.get()))
        except ValueError as exc:
            messagebox.showerror("Invalid FC/CN", str(exc))
            return
        self.reverse_result.set(f"HEX: {hex_value}" + (f"\n{' '.join(warnings)}" if warnings else ""))
        self.clipboard_clear()
        self.clipboard_append(hex_value)
        self.set_status("Hex value created and copied.")

    def confirm_import_preview(self, path: Path, result: dict[str, Any]) -> bool:
        preview = "\n".join(result.get("lines", [])[:30])
        if not preview:
            return messagebox.askyesno("Import preview", f"{result.get('message', '')}\n\nNo rows were extracted. Continue?")
        dialog = tk.Toplevel(self)
        dialog.title("Import Preview")
        dialog.configure(bg="#10141b")
        dialog.geometry("680x420")
        dialog.transient(self)
        dialog.grab_set()
        approved = tk.BooleanVar(value=False)
        tk.Label(dialog, text=f"Preview: {path.name}", bg="#10141b", fg="#ffffff", font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=14, pady=(14, 6))
        tk.Label(dialog, text=result.get("message", ""), bg="#10141b", fg="#a8b2c2", wraplength=640, justify="left").pack(anchor="w", padx=14, pady=(0, 8))
        text = tk.Text(dialog, bg="#080a0f", fg="#f5f7fb", relief="flat", padx=10, pady=10, font=("Cascadia Mono", 10))
        text.pack(fill="both", expand=True, padx=14, pady=(0, 12))
        text.insert("1.0", preview)
        text.configure(state="disabled")
        row = tk.Frame(dialog, bg="#10141b")
        row.pack(fill="x", padx=14, pady=(0, 14))
        self._button(row, "Add To Queue", lambda: (approved.set(True), dialog.destroy()), True, icon="icon-import.png").pack(side="right", padx=(8, 0))
        self._button(row, "Cancel", dialog.destroy, icon="icon-clear.png").pack(side="right")
        self.wait_window(dialog)
        return bool(approved.get())

    def import_paths_to_queue(self, paths: list[Path], preview: bool = True) -> None:
        files = [path for path in paths if path.is_file()]
        if not files:
            messagebox.showinfo("No files", "Drop or choose one or more supported files.")
            return
        all_lines: list[str] = []
        imported = 0
        errors: list[str] = []
        messages: list[str] = []
        for path in files:
            try:
                result = import_structured_file(path)
            except Exception as exc:
                errors.append(f"{path.name}: {exc}")
                continue
            if preview and len(files) == 1 and not self.confirm_import_preview(path, result):
                self.set_status("Import canceled.")
                return
            lines = result.get("lines", [])
            if lines:
                all_lines.extend(lines)
                imported += 1
            messages.append(f"{path.name}: {result.get('message', '')}")
            self.add_recent_file(path)
        if all_lines:
            self.append_text_to_queue("\n".join(all_lines))
        if errors:
            messagebox.showwarning("Some imports failed", "\n".join(errors[:8]))
        if len(files) == 1 and messages:
            self.notice_var.set(messages[0])
            self.set_status(messages[0])
        else:
            status = f"Imported {len(all_lines)} row(s) from {imported} file(s)."
            if errors:
                status += f" {len(errors)} file(s) failed."
            self.notice_var.set(status)
            self.set_status(status)

    def import_file(self, path: Path | None = None) -> None:
        if path is None:
            path_texts = filedialog.askopenfilenames(
                title="Import access control data",
                filetypes=[
                    ("Supported files", "*.txt *.csv *.tsv *.xlsx *.xlsm *.xls *.xml *.html *.htm"),
                    ("Excel files", "*.xlsx *.xlsm *.xls"),
                    ("Text files", "*.txt *.csv *.tsv"),
                    ("Table files", "*.xml *.html *.htm"),
                    ("All files", "*.*"),
                ],
            )
            if not path_texts:
                return
            paths = [Path(item) for item in path_texts]
            self.import_paths_to_queue(paths, preview=len(paths) == 1)
            return
        self.import_paths_to_queue([path], preview=True)

    def export_excel(self) -> None:
        if not self.results and not self.invalid:
            messagebox.showinfo("No report", "Run a conversion before exporting.")
            return
        stamp = datetime.now().strftime("%Y-%m-%d")
        path_text = filedialog.asksaveasfilename(
            title="Save Excel report",
            defaultextension=".xlsx",
            initialdir=self.default_export_dir(),
            initialfile=f"Macys_AP_China_Grove_Hex_Utility_{stamp}.xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if not path_text:
            return
        self._write_excel(Path(path_text))
        self.set_status(f"Saved Excel report: {path_text}")

    def export_csv(self) -> None:
        if not self.results and not self.invalid:
            messagebox.showinfo("No report", "Run a conversion before exporting.")
            return
        stamp = datetime.now().strftime("%Y-%m-%d")
        path_text = filedialog.asksaveasfilename(
            title="Save CSV report",
            defaultextension=".csv",
            initialdir=self.default_export_dir(),
            initialfile=f"Macys_AP_China_Grove_Hex_Utility_{stamp}.csv",
            filetypes=[("CSV Report", "*.csv")],
        )
        if not path_text:
            return
        with Path(path_text).open("w", newline="", encoding="utf-8-sig") as handle:
            writer = csv.writer(handle)
            writer.writerow(["Line", "Hex ID", "Facility Code", "Card Number", "Status", "Warnings / Cleanup", "Converted At"])
            for item in self.results:
                writer.writerow([item.line, item.hex_value, item.facility, item.card, "Warning" if item.warnings else "Valid", " | ".join([*item.suggestions, *item.warnings]), item.converted_at])
            for item in self.invalid:
                writer.writerow([item.line, item.raw, "", "", "Invalid", item.reason, item.converted_at])
        self.set_status(f"Saved CSV report: {path_text}")

    def export_pdf(self) -> None:
        if not self.results and not self.invalid:
            messagebox.showinfo("No report", "Run a conversion before exporting.")
            return
        stamp = datetime.now().strftime("%Y-%m-%d")
        path_text = filedialog.asksaveasfilename(
            title="Save PDF report",
            defaultextension=".pdf",
            initialdir=self.default_export_dir(),
            initialfile=f"Macys_AP_China_Grove_Hex_Utility_{stamp}.pdf",
            filetypes=[("PDF Report", "*.pdf")],
        )
        if not path_text:
            return
        if SimpleDocTemplate is None:
            messagebox.showerror("PDF export unavailable", "PDF export needs reportlab installed.")
            return
        doc = SimpleDocTemplate(path_text, pagesize=landscape(letter), rightMargin=24, leftMargin=24, topMargin=24, bottomMargin=24)
        styles = getSampleStyleSheet()
        story = [
            Paragraph(f"{APP_SHORT_NAME} Conversion Report", styles["Title"]),
            Paragraph(f"Generated: {datetime.now().strftime('%m/%d/%Y %I:%M:%S %p')}", styles["Normal"]),
            Paragraph(f"Valid: {len(self.results)} | Invalid: {len(self.invalid)} | Warnings: {sum(len(row.warnings) for row in self.results)}", styles["Normal"]),
            Spacer(1, 12),
        ]
        data = [["Line", "Hex ID", "FC", "CN", "Status", "Notes"]]
        for item in self.results:
            data.append([item.line, item.hex_value, item.facility, item.card, "Warning" if item.warnings else "Valid", " | ".join([*item.suggestions, *item.warnings])])
        for item in self.invalid:
            data.append([item.line, item.raw, "", "", "Invalid", item.reason])
        table = Table(data, repeatRows=1, colWidths=[44, 110, 70, 70, 80, 410])
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E51B2D")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#B8B8B8")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
        ]))
        story.append(table)
        doc.build(story)
        self.set_status(f"Saved PDF report: {path_text}")

    def _write_excel(self, path: Path) -> None:
        wb = Workbook()
        ws = wb.active
        ws.title = "Hex Converter"
        widths = {"A": 8, "B": 18, "C": 14, "D": 14, "E": 32, "F": 24}
        for column, width in widths.items():
            ws.column_dimensions[column].width = width
        header_fill = PatternFill("solid", fgColor="E51B2D")
        table_fill = PatternFill("solid", fgColor="E6E6E6")
        thin = Side(style="thin", color="B8B8B8")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        ws.merge_cells("A1:F1")
        ws["A1"] = f"{APP_SHORT_NAME} Conversion Report"
        ws["A1"].font = Font(bold=True, size=16, color="FFFFFF")
        ws["A1"].alignment = Alignment(horizontal="center")
        ws["A1"].fill = header_fill
        meta = [
            ("Generated", datetime.now().strftime("%m/%d/%Y %I:%M:%S %p")),
            ("Last Conversion", self.last_converted_at),
            ("Total Input Lines", len(self.results) + len(self.invalid)),
            ("Valid Lines", len(self.results)),
            ("Invalid Lines", len(self.invalid)),
            ("Warnings", sum(len(row.warnings) for row in self.results)),
            ("Note", "Facility Code = high 16 bits; Card Number = low 16 bits"),
        ]
        row_index = 3
        for key, value in meta:
            ws.cell(row_index, 1, key).font = Font(bold=True)
            ws.cell(row_index, 2, value)
            row_index += 1
        row_index += 1
        headers = ["Line", "Hex ID", "Facility Code", "Card Number", "Warnings / Cleanup", "Converted At"]
        for col, label in enumerate(headers, start=1):
            cell = ws.cell(row_index, col, label)
            cell.font = Font(bold=True)
            cell.fill = table_fill
            cell.border = border
            cell.alignment = Alignment(horizontal="center")
        row_index += 1
        for item in self.results:
            values = [item.line, item.hex_value, item.facility, item.card, " | ".join([*item.suggestions, *item.warnings]), item.converted_at]
            for col, value in enumerate(values, start=1):
                cell = ws.cell(row_index, col, value)
                cell.border = border
                cell.alignment = Alignment(horizontal="center" if col in {1, 3, 4, 6} else "left")
            row_index += 1
        if self.invalid:
            row_index += 2
            ws.cell(row_index, 1, "Invalid Inputs").font = Font(bold=True)
            row_index += 1
            for col, label in enumerate(["Line", "Raw Input", "Reason", "Converted At"], start=1):
                cell = ws.cell(row_index, col, label)
                cell.font = Font(bold=True)
                cell.fill = table_fill
                cell.border = border
            row_index += 1
            for item in self.invalid:
                for col, value in enumerate([item.line, item.raw, item.reason, item.converted_at], start=1):
                    cell = ws.cell(row_index, col, value)
                    cell.border = border
                row_index += 1
        wb.save(path)

    def export_txt(self) -> None:
        if not self.results and not self.invalid:
            messagebox.showinfo("No report", "Run a conversion before exporting.")
            return
        stamp = datetime.now().strftime("%Y-%m-%d")
        path_text = filedialog.asksaveasfilename(
            title="Save TXT report",
            defaultextension=".txt",
            initialdir=self.default_export_dir(),
            initialfile=f"Macys_AP_China_Grove_Hex_Utility_{stamp}.txt",
            filetypes=[("Text Report", "*.txt")],
        )
        if not path_text:
            return
        Path(path_text).write_text(self._build_text_report(), encoding="utf-8")
        self.set_status(f"Saved TXT report: {path_text}")

    def _build_text_report(self) -> str:
        lines = [
            f"{APP_SHORT_NAME} Export",
            f"Generated: {datetime.now().strftime('%m/%d/%Y %I:%M:%S %p')}",
            f"Last Conversion: {self.last_converted_at}",
            f"Valid Records: {len(self.results)}   Invalid Lines: {len(self.invalid)}",
            "-" * 60,
            "",
        ]
        if self.results:
            lines.append(f"{'LINE':<6} {'HEX ID':<12} {'FC':<8} {'CN':<8} WARNINGS/CLEANUP")
            lines.append("-" * 80)
            for item in self.results:
                notes = " | ".join([*item.suggestions, *item.warnings])
                lines.append(f"{item.line:<6} {item.hex_value:<12} {item.facility:<8} {item.card:<8} {notes}")
        if self.invalid:
            lines.extend(["", "INVALID LINES", "-" * 30])
            for item in self.invalid:
                lines.append(f"Line {item.line}: {item.raw} | {item.reason} | {item.converted_at}")
        return "\n".join(lines) + "\n"

    def show_help(self) -> None:
        dialog = tk.Toplevel(self)
        dialog.title("How To Use")
        dialog.configure(bg="#10141b")
        dialog.geometry("700x560")
        dialog.transient(self)
        dialog.grab_set()
        tk.Label(dialog, text=f"How To Use {APP_SHORT_NAME}", bg="#10141b", fg="#ffffff", font=("Segoe UI", 14, "bold")).pack(anchor="w", padx=16, pady=(16, 8))
        help_text = (
            "Batch Converter\n"
            "Paste HEX IDs, one per line, then press Convert All. You can also paste full employee lines; the app will pull out clean 8-character IDs when it can.\n\n"
            "Results\n"
            "Valid rows show HEX, Facility Code, and Card Number. Warning rows are still converted, but the app noticed something unusual such as a duplicate, all zeros, or a high value.\n\n"
            "Export\n"
            "Use the Export menu to save reports as Excel, CSV, TXT, or PDF. The Run Summary panel only shows the current counts so the screen stays easy to read.\n\n"
            "Navigation\n"
            "Use the left Workspace list to move between Convert, Reverse, and Review areas. Hover over buttons and menus for quick tips.\n\n"
            "Import Options\n"
            "Use the Import menu to browse files, paste clipboard text, or load sample IDs. You can also drag supported files directly onto the Input Queue text box; the app imports them into that same queue without creating a separate drop area.\n\n"
            "Single Lookup\n"
            "Use Single Lookup when you only need to convert one HEX ID quickly. The FC,CN pair is copied to the clipboard after conversion.\n\n"
            "FC/CN to Hex\n"
            "Use FC/CN to Hex for one Facility Code and Card Number pair.\n\n"
            "Unconvert Batch\n"
            "Use Unconvert Batch for many FC/CN pairs at once. Accepted examples include 34968,18199, tab-separated columns, or FC 34968 CN 18199.\n\n"
            "Import\n"
            "Import supports TXT, CSV, TSV, XLS, XLSX, XLSM, XML, and HTML table files. Browse imports show a preview for one file. Drag/drop imports add supported files directly to the Input Queue.\n\n"
            "BlueWave\n"
            "Use the BlueWave button in the top bar to open the access-control site in your browser.\n\n"
            "Helpful Shortcuts\n"
            "Ctrl+I imports, Ctrl+R converts, Ctrl+E exports Excel, Ctrl+P exports PDF, Ctrl+F jumps to search, and Ctrl+L clears the workspace."
        )
        help_frame = tk.Frame(dialog, bg="#10141b")
        help_frame.pack(fill="both", expand=True, padx=16, pady=(0, 12))
        help_frame.rowconfigure(0, weight=1)
        help_frame.columnconfigure(0, weight=1)
        text = tk.Text(help_frame, bg="#080a0f", fg="#f5f7fb", relief="flat", padx=14, pady=14, wrap="word", font=("Segoe UI", 10))
        help_scroll = ttk.Scrollbar(help_frame, orient="vertical", command=text.yview, style="Dark.Vertical.TScrollbar")
        text.configure(yscrollcommand=help_scroll.set)
        text.grid(row=0, column=0, sticky="nsew")
        help_scroll.grid(row=0, column=1, sticky="ns")
        text.insert("1.0", help_text)
        text.configure(state="disabled")
        row = tk.Frame(dialog, bg="#10141b")
        row.pack(fill="x", padx=16, pady=(0, 16))
        self._button(row, "Close", dialog.destroy, True, icon="icon-clear.png").pack(side="right")

    def show_about(self) -> None:
        dialog = tk.Toplevel(self)
        dialog.title("About")
        dialog.configure(bg="#10141b")
        dialog.geometry("560x360")
        dialog.transient(self)
        dialog.grab_set()

        header = tk.Frame(dialog, bg="#090b10")
        header.pack(fill="x")
        logo = self._load_icon("macys-ap-icon.png", 46)
        if logo:
            tk.Label(header, image=logo, bg="#090b10").pack(side="left", padx=16, pady=14)
        title = tk.Frame(header, bg="#090b10")
        title.pack(side="left", fill="x", expand=True, pady=14)
        tk.Label(title, text=APP_SHORT_NAME, bg="#090b10", fg="#ffffff", font=("Segoe UI", 13, "bold")).pack(anchor="w")
        tk.Label(title, text="Asset Protection access-control utility", bg="#090b10", fg="#a8b2c2").pack(anchor="w")

        body = tk.Frame(dialog, bg="#10141b")
        body.pack(fill="both", expand=True, padx=18, pady=16)
        copy = (
            "Made for Macy's Asset Protection operations at the China Grove, North Carolina facility.\n\n"
            "Built by Christopher Schumacher, Asset Protection FLO.\n\n"
            "This utility converts access-control HEX IDs, Facility Codes, and Card Numbers for review and export."
        )
        tk.Label(body, text=copy, bg="#10141b", fg="#f5f7fb", justify="left", wraplength=500, font=("Segoe UI", 10)).pack(anchor="w")
        links = tk.Frame(body, bg="#10141b")
        links.pack(fill="x", pady=(16, 0))
        self._link_label(links, "GitHub: github.com/rice2k", "https://github.com/rice2k", "#10141b", "Open the project profile link.").pack(anchor="w", pady=(0, 6))
        self._link_label(links, "BlueWave access-control site", BLUEWAVE_URL, "#10141b", "Open BlueWave in your browser.").pack(anchor="w")

        row = tk.Frame(dialog, bg="#10141b")
        row.pack(fill="x", padx=18, pady=(0, 16))
        self._button(row, "Close", dialog.destroy, True, icon="icon-clear.png").pack(side="right")


def main() -> None:
    if "--self-test" in sys.argv:
        assert hex_to_fc_cn("88984717") == (34968, 18199)
        assert fc_cn_to_hex(34968, 18199) == "88984717"
        assert clean_candidate_line("Active Christopher Benson, 88984765")["extracted"] == "88984765"
        results, invalid = convert_lines("88984717\nBAD-LINE\n8898-4765", "TEST")
        assert len(results) == 2
        assert len(invalid) == 1
        unconverted, unconvert_invalid = unconvert_lines("34968,18199\nBAD-LINE", "TEST")
        assert len(unconverted) == 1
        assert len(unconvert_invalid) == 1
        assert unconverted[0].hex_value == "88984717"
        imported = extract_name_id_lines_from_tables(
            [[["Candidate Name", "Colleague #"], ["Christopher Benson", "88984765"], ["Supervisor", "88123456"]]],
            "sample.xlsx",
        )
        assert imported["lines"] == ["Christopher Benson, 88984765"]
        return
    app = ConverterApp()
    app.mainloop()


if __name__ == "__main__":
    main()

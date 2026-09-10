
import csv
import os
import re
import sys
import tempfile
import subprocess
from pathlib import Path
from collections import defaultdict

import pandas as pd
from openpyxl import load_workbook
import openpyxl
import pdfplumber

try:
    from docx import Document
except ImportError:
    Document = None


DAY_NAMES = {
    "monday": "Monday", "mon": "Monday",
    "tuesday": "Tuesday", "tue": "Tuesday", "tues": "Tuesday",
    "wednesday": "Wednesday", "wed": "Wednesday",
    "thursday": "Thursday", "thu": "Thursday", "thur": "Thursday", "thurs": "Thursday",
    "friday": "Friday", "fri": "Friday",
    "saturday": "Saturday", "sat": "Saturday",
    "sunday": "Sunday", "sun": "Sunday",
}

DAY_ORDER = {
    "Monday": 1, "Tuesday": 2, "Wednesday": 3, "Thursday": 4,
    "Friday": 5, "Saturday": 6, "Sunday": 7
}

ROMAN = {
    "I": 1, "II": 2, "III": 3, "IV": 4, "V": 5, "VI": 6,
    "VII": 7, "VIII": 8
}


def clean(x):
    if x is None:
        return ""
    return re.sub(r"\s+", " ", str(x).replace("\xa0", " ").strip())


def clean_multiline(x):
    """Normalize text while preserving meaningful line breaks inside timetable cells."""
    if x is None:
        return ""
    lines = [re.sub(r"[ \t]+", " ", str(line).replace("\xa0", " ").strip()) for line in str(x).replace("\r", "\n").split("\n")]
    return "\n".join(line for line in lines if line)


def norm(x):
    return re.sub(r"[^a-z0-9]+", "", clean(x).lower())


def normalize_roll(x):
    return re.sub(r"[\s-]+", "", clean(x).upper())


def is_real_student_roll(x):
    s = normalize_roll(x)
    return bool(re.fullmatch(r"[A-Z]{1,4}\d+[A-Z]?\d*", s))


def find_column(columns, candidates):
    mapping = {norm(c): c for c in columns}
    for c in candidates:
        if norm(c) in mapping:
            return mapping[norm(c)]
    return None


def looks_like_subject_file(path):
    return path.suffix.lower() in {".xlsx", ".xls", ".csv"}


def is_timetable_name(path):
    n = norm(path.stem)
    return "timetable" in n or "timetable" in n or "schedule" in n or "timeslot" in n


def infer_subject_name(path):
    stem = re.sub(r"\s*\(\d+\)\s*$", "", path.stem.strip())
    # Keep the user's filename as the default subject label.
    # If the filename is "Customer Relationship Management - CRM", prefer CRM.
    parts = re.split(r"\s*[-_]\s*", stem)
    if len(parts) > 1:
        last = clean(parts[-1])
        if 1 <= len(last.split()) <= 4:
            return last
    return stem


def sheet_has_student_structure(df):
    cols = list(df.columns)
    roll = find_column(cols, ["Roll No", "Roll Number", "Roll", "Registration No", "Reg No", "ID"])
    name = find_column(cols, ["Name", "Student Name", "Student"])
    return roll is not None or name is not None


def parse_subject_file(path):
    records = []
    subject = infer_subject_name(path)

    if path.suffix.lower() == ".csv":
        sheets = {"CSV": pd.read_csv(path)}
    else:
        sheets = pd.read_excel(path, sheet_name=None)

    for sheet_name, df in sheets.items():
        if df is None or df.empty:
            continue

        df = df.copy()
        df.columns = [clean(c) for c in df.columns]

        # Some subject workbooks contain a duplicate administrative EMAIL sheet.
        # It has the same students as the real enrolment sheet, but it is not a
        # separate section/course list. Skip it entirely.
        sheet_key = norm(sheet_name)
        has_email_columns = any(
            "email" in norm(c) for c in df.columns
        )
        if sheet_key in {"email", "emails"} or has_email_columns:
            continue

        roll_col = find_column(
            df.columns,
            ["Roll No", "Roll Number", "Roll", "Registration No", "Reg No", "Student ID", "ID"]
        )
        name_col = find_column(
            df.columns,
            ["Name", "Student Name", "Student", "Candidate Name"]
        )
        section_col = find_column(
            df.columns,
            ["Section", "Sec", "Batch", "Group", "Class", "Division"]
        )

        if not roll_col and not name_col:
            # Ignore non-student sheets, but allow a simple one-column student list.
            if len(df.columns) == 1:
                only_col = df.columns[0]
                values = [clean(v) for v in df[only_col].tolist() if clean(v)]
                if any(is_real_student_roll(v) for v in values):
                    roll_col = only_col
                else:
                    name_col = only_col
            else:
                continue

        multi_sheet = len(sheets) > 1
        sheet_section = "" if str(sheet_name).lower() in {"sheet1", "sheet2", "csv"} else clean(sheet_name)

        for _, row in df.iterrows():
            roll = normalize_roll(row[roll_col]) if roll_col else ""
            name = clean(row[name_col]) if name_col else ""
            section = clean(row[section_col]) if section_col else ""
            if not section and multi_sheet:
                # Treat only section-like sheet names as sections.
                # "LIST", "EMAIL", "SHEET1", etc. are administrative/data sheets.
                if re.fullmatch(
                    r"(?:[A-D]|[A-D]\d{1,2}|\d{1,2}[A-D]|PRE|POST)",
                    sheet_section or "",
                    flags=re.I,
                ):
                    section = sheet_section.upper()


            if roll or name:
                if roll and not is_real_student_roll(roll):
                    # Don't accidentally treat "Roll No" as data.
                    if norm(roll) in {"rollno", "rollnumber", "studentid", "id"}:
                        continue
                records.append({
                    "roll": roll,
                    "name": name,
                    "subject": subject,
                    "section": section,
                    "source_file": path.name,
                    "source_sheet": clean(sheet_name),
                })

    return records


def parse_all_subject_files(folder, timetable_path):
    records = []
    for path in sorted(folder.iterdir()):
        if path.is_file() and looks_like_subject_file(path) and path.resolve() != timetable_path.resolve():
            # Avoid treating likely master timetable workbooks as subject files.
            if is_timetable_name(path):
                continue
            # Common administrative files should not be treated as subjects.
            if norm(path.stem) in {
                "studentlist", "studentlists", "masterlist", "allstudents",
                "students", "timetable", "schedule"
            }:
                continue
            try:
                records.extend(parse_subject_file(path))
            except Exception:
                # Let the UI report ignored files later; one malformed file shouldn't crash everything.
                continue

    # De-duplicate.
    seen = set()
    out = []
    for r in records:
        key = (r["roll"], norm(r["name"]), norm(r["subject"]), norm(r["section"]))
        if key not in seen and (r["roll"] or r["name"]):
            seen.add(key)
            out.append(r)
    return out


def extract_time(text):
    text = clean(text).replace("–", "-").replace("—", "-")
    m = re.search(
        r"(\d{1,2}:\d{2}\s*(?:AM|PM)?)\s*-\s*(\d{1,2}:\d{2}\s*(?:AM|PM)?)",
        text,
        flags=re.I
    )
    if not m:
        return ""
    return f"{m.group(1).upper().replace(' ', '')}-{m.group(2).upper().replace(' ', '')}"


def period_label_from_header(text, index):
    text = clean(text)
    # Prefer Roman numeral when present.
    m = re.search(r"\b(I{1,3}|IV|V|VI|VII|VIII)\b", text.upper())
    if m:
        return m.group(1).upper()
    t = extract_time(text)
    if t:
        return t
    return f"Period {index}"


def weekday(text):
    k = clean(text).lower()
    return DAY_NAMES.get(k, "")


def convert_legacy_doc_to_docx(doc_path):
    # Windows + Microsoft Word: best support for old .doc files.
    if os.name == "nt":
        try:
            import win32com.client  # type: ignore
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(str(doc_path))
            tmp = Path(tempfile.gettempdir()) / (doc_path.stem + "_converted.docx")
            doc.SaveAs2(str(tmp), FileFormat=16)  # wdFormatXMLDocument
            doc.Close(False)
            word.Quit()
            return tmp
        except Exception:
            pass

    # Linux/macOS fallback if LibreOffice is installed.
    for command in ("soffice", "libreoffice"):
        try:
            out_dir = Path(tempfile.mkdtemp(prefix="tt_"))
            subprocess.run(
                [command, "--headless", "--convert-to", "docx", "--outdir", str(out_dir), str(doc_path)],
                check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True
            )
            converted = out_dir / (doc_path.stem + ".docx")
            if converted.exists():
                return converted
        except Exception:
            continue

    return None


def load_docx_table(path):
    actual = Path(path)
    temp_converted = None

    if actual.suffix.lower() == ".doc":
        converted = convert_legacy_doc_to_docx(actual)
        if converted is None:
            raise RuntimeError(
                "This is an old .doc timetable. On Windows, Microsoft Word must be installed "
                "for the program to read .doc files. Another option is to save the timetable as .docx."
            )
        actual = converted
        temp_converted = converted

    if Document is None:
        raise RuntimeError("python-docx is not installed.")

    try:
        doc = Document(str(actual))
        return doc.tables
    finally:
        if temp_converted and temp_converted.exists():
            try:
                temp_converted.unlink()
            except Exception:
                pass


def find_timetable_word_table(path):
    tables = load_docx_table(path)
    best = None
    best_score = -1

    for table in tables:
        if not table.rows:
            continue
        score = 0
        first = " ".join(clean(c.text).lower() for c in table.rows[0].cells)
        if "day" in first:
            score += 3
        if extract_time(first):
            score += 2
        if len(table.rows[0].cells) >= 4:
            score += 1
        if score > best_score:
            best_score = score
            best = table

    if best is None:
        raise RuntimeError("Could not find the timetable table in the Word file.")
    return best


def parse_matrix_timetable_from_doc(path):
    table = find_timetable_word_table(path)
    rows = table.rows
    header = [clean(c.text) for c in rows[0].cells]

    periods = []
    for i, h in enumerate(header[1:], start=1):
        periods.append({
            "index": i,
            "label": period_label_from_header(h, i),
            "time": extract_time(h),
        })

    matrix = []
    for row in rows[1:]:
        raw_cells = [str(c.text or "") for c in row.cells]
        if not raw_cells:
            continue
        day = weekday(raw_cells[0])
        if not day:
            continue
        for i in range(1, min(len(raw_cells), len(periods) + 1)):
            cell = raw_cells[i]
            if not clean(cell):
                continue
            matrix.append({
                "day": day,
                "period": periods[i - 1]["label"],
                "time": periods[i - 1]["time"],
                "cell": cell
            })
    return matrix


def find_timetable_excel_table(path):
    sheets = pd.read_excel(path, sheet_name=None)
    # First try long format.
    for sheet_name, df in sheets.items():
        if df is None or df.empty:
            continue
        df = df.copy()
        df.columns = [clean(c) for c in df.columns]
        day_col = find_column(df.columns, ["Day", "Weekday", "Week Day"])
        subject_col = find_column(df.columns, ["Subject", "Course", "Course Name", "Paper", "Class"])
        period_col = find_column(df.columns, ["Period", "Period No", "Slot", "Time", "Time Slot"])
        if day_col and subject_col and period_col:
            rows = []
            start_col = find_column(df.columns, ["Start Time", "Start", "From"])
            end_col = find_column(df.columns, ["End Time", "End", "To"])
            section_col = find_column(df.columns, ["Section", "Sec", "Batch", "Group"])
            room_col = find_column(df.columns, ["Room", "Classroom", "Venue", "Location"])
            for _, r in df.iterrows():
                day = weekday(r[day_col])
                subject = clean(r[subject_col])
                if not day or not subject:
                    continue
                rows.append({
                    "day": day,
                    "period": clean(r[period_col]),
                    "time": (
                        f"{clean(r[start_col])}-{clean(r[end_col])}"
                        if start_col and end_col else clean(r[period_col])
                    ),
                    "subject": subject,
                    "section": clean(r[section_col]) if section_col else "",
                    "room": clean(r[room_col]) if room_col else "",
                    "cell": subject,
                    "long_format": True,
                })
            if rows:
                return rows

    # Otherwise treat the sheet as a matrix like the Word timetable.
    for sheet_name, df in sheets.items():
        if df is None or df.empty or df.shape[1] < 3:
            continue
        df = df.copy()
        df.columns = [clean(c) for c in df.columns]
        header = [clean(c) for c in df.columns]
        if norm(header[0]) != "day" and "day" not in norm(header[0]):
            continue
        periods = [
            {"label": period_label_from_header(h, i), "time": extract_time(h)}
            for i, h in enumerate(header[1:], start=1)
        ]
        rows = []
        for _, r in df.iterrows():
            day = weekday(r.iloc[0])
            if not day:
                continue
            for i, p in enumerate(periods, start=1):
                cell = clean(r.iloc[i])
                if cell:
                    rows.append({
                        "day": day,
                        "period": p["label"],
                        "time": p["time"],
                        "cell": cell
                    })
        if rows:
            return rows

    raise RuntimeError("Could not detect the timetable structure in the Excel file.")



def parse_matrix_timetable_from_pdf(path):
    """Parse a text-based timetable PDF using pdfplumber table extraction."""
    candidate_tables = []

    with pdfplumber.open(str(path)) as pdf:
        for page_no, page in enumerate(pdf.pages, start=1):
            settings_list = [
                {"vertical_strategy": "lines", "horizontal_strategy": "lines"},
                {},
                {"vertical_strategy": "lines", "horizontal_strategy": "text"},
            ]
            for settings in settings_list:
                try:
                    tables = page.extract_tables(settings)
                except Exception:
                    tables = []
                for table in tables:
                    if not table:
                        continue
                    rows = [[clean_multiline(c) for c in (row or [])] for row in table]
                    if not rows:
                        continue

                    # Score tables rather than blindly selecting the widest one.
                    # The desired timetable table normally starts with DAY and has
                    # 6+ time-bearing period columns.
                    first = rows[0]
                    first_text = " ".join(first).replace("\n", " ")
                    day_first = bool(first and norm(first[0]) == "day")
                    time_cells = sum(1 for c in first if extract_time(c))
                    day_rows = sum(1 for r in rows if r and weekday(r[0]))
                    score = (
                        1000 if day_first else 0
                    ) + time_cells * 100 + min(len(first), 10) * 5 + min(day_rows, 7) * 20

                    if day_first or ("day" in norm(first_text) and extract_time(first_text)):
                        candidate_tables.append((score, page_no, rows))

    if not candidate_tables:
        raise RuntimeError(
            "Could not identify a timetable table in the PDF. "
            "Use a text-based PDF with days on the left and periods/times across the top."
        )

    _, page_no, rows = max(candidate_tables, key=lambda x: x[0])

    # The selected standard table should have the header in its first row. If not,
    # locate the best row containing DAY + time information.
    header_idx = 0
    if not (rows[0] and norm(rows[0][0]) == "day"):
        for i, row in enumerate(rows[:12]):
            row_text = " ".join(row)
            if row and norm(row[0]) == "day" and extract_time(row_text):
                header_idx = i
                break

    header = rows[header_idx]
    if len(header) < 3:
        raise RuntimeError("The PDF timetable does not contain enough period columns.")

    ncols = len(header)

    # The supplied master timetable PDF uses two header rows:
    # row 1 contains I/II/III/... and row 2 contains the actual times.
    # Combine them so the generated personal timetable gets both period and time.
    time_row = None
    if header_idx + 1 < len(rows):
        candidate = rows[header_idx + 1]
        if len(candidate) >= ncols and any(extract_time(c or "") for c in candidate):
            time_row = candidate

    periods = []
    for idx, h in enumerate(header[1:], start=1):
        h = clean(h)
        time_text = clean(time_row[idx] if time_row is not None and idx < len(time_row) else "")
        time_value = extract_time(time_text) or extract_time(h)
        period_value = re.sub(r"[^IVX]+", "", h.upper()) if re.fullmatch(r"[IVX]+", h.upper()) else h
        periods.append({
            "label": period_value or f"Period {idx}",
            "time": time_value,
        })

    # Coalesce continuation rows created by merged PDF cells.
    timetable_rows = []
    current = None
    for raw_row in rows[header_idx + 1:]:
        row = list(raw_row or []) + [""] * max(0, ncols - len(raw_row or []))
        row = [clean_multiline(c) for c in row[:ncols]]
        if not any(row):
            continue

        row_day = weekday(row[0]) if row else ""
        if row_day:
            if current is not None:
                timetable_rows.append(current)
            current = {"day": row_day, "cells": row[1:]}
        elif current is not None:
            for j in range(1, ncols):
                extra = row[j]
                if not extra:
                    continue
                if current["cells"][j - 1]:
                    current["cells"][j - 1] += "\n" + extra
                else:
                    current["cells"][j - 1] = extra

    if current is not None:
        timetable_rows.append(current)

    output = []
    for tr in timetable_rows:
        day = tr["day"]
        for j, period in enumerate(periods):
            cell = clean_multiline(tr["cells"][j] if j < len(tr["cells"]) else "")
            if not cell:
                continue
            output.append({
                "day": day,
                "period": period["label"],
                "time": period["time"],
                "cell": cell,
                "pdf_page": page_no,
            })

    if not output:
        raise RuntimeError("The PDF timetable table was found, but no class cells were extracted.")

    return output



def parse_timetable(path):
    path = Path(path)
    if not path.exists():
        raise RuntimeError(
            f"Timetable file was not found at:\n{path}\n\n"
            "Check that the file still exists and that the folder was selected correctly."
        )
    ext = path.suffix.lower()
    if ext in {".doc", ".docx"}:
        return parse_matrix_timetable_from_doc(path)
    if ext == ".pdf":
        return parse_matrix_timetable_from_pdf(path)
    if ext in {".xlsx", ".xls"}:
        return find_timetable_excel_table(path)
    if ext == ".csv":
        tmp = pd.read_csv(path)
        # Reuse Excel logic by writing a temporary xlsx.
        temp = Path(tempfile.gettempdir()) / "_tt_temp.xlsx"
        tmp.to_excel(temp, index=False)
        try:
            return find_timetable_excel_table(temp)
        finally:
            try:
                temp.unlink()
            except Exception:
                pass
    raise RuntimeError("Unsupported timetable format.")


def subject_matches_cell(subject, section, cell):
    """Match a subject entry without accidentally matching another section."""
    text = clean(cell)
    s = clean(subject)
    sec = clean(section)
    if not s:
        return False

    subj_re = re.escape(s)

    # The subject itself must occur as a token.
    if not re.search(rf"(?<![A-Za-z0-9]){subj_re}(?![A-Za-z0-9])", text, flags=re.I):
        return False

    if not sec:
        return True

    sec_re = re.escape(sec)

    # Style 1: BrM(D), BrM (D)
    if re.search(
        rf"(?<![A-Za-z0-9]){subj_re}\s*\(\s*{sec_re}\s*\)",
        text, flags=re.I
    ):
        return True

    # Style 2: SO (AP) B L-3 / SO B L-3
    # Permit one parenthesized faculty token between subject and section,
    # but require the section to be the next meaningful token. This prevents
    # BrM(B) from incorrectly matching BrM(C).
    if re.search(
        rf"(?<![A-Za-z0-9]){subj_re}\s+(?:\([^)]*\)\s+)?{sec_re}(?![A-Za-z0-9])",
        text, flags=re.I
    ):
        return True

    return False

def matching_line(cell, subject, section):
    # Word tables often place multiple classes in one cell using line breaks.
    # Keep the line boundaries so the room belongs to the matched subject.
    raw = str(cell or "").replace("\r", "\n")
    lines = [clean(x) for x in re.split(r"\n|\|", raw) if clean(x)]

    for line in lines:
        if subject_matches_cell(subject, section, line):
            return line

    if subject_matches_cell(subject, section, clean(raw)):
        return clean(raw)

    return ""


def extract_room(cell):
    text = str(cell or "")
    # Locations in the timetable include G-12, G-14, L-B and LB.
    m = re.search(
        r"\b(?:G|L)\s*-\s*\d+\b|\bL\s*-\s*B\b|\bLB\b",
        text,
        flags=re.I
    )
    return clean(m.group(0)) if m else ""


def extract_faculty(cell, subject="", section=""):
    """
    Extract faculty initials from a timetable line such as:
        CRM (ANK/NB) G-12
        BrM (D) (AK)
        BrM(D) (AK) G-12
        SO (AP) B L-3

    When a section is present in parentheses, ignore that parenthesis and
    return the faculty initials.
    """
    raw = str(cell or "")
    candidates = re.findall(
        r"\(\s*([A-Za-z]+(?:\s*/\s*[A-Za-z]+)*)\s*\)",
        raw
    )
    candidates = [clean(x).replace(" ", "") for x in candidates]

    sec = clean(section).upper()
    for candidate in candidates:
        if sec and candidate.upper() == sec:
            continue

        # Faculty initials in the supplied timetable are alphabetic tokens.
        # Avoid treating words such as PRE/POST as faculty.
        if candidate.upper() in {"PRE", "POST"}:
            continue
        return candidate

    return ""



def build_student_timetable(records, timetable, roll):
    roll = normalize_roll(roll)
    student_rows = [r for r in records if normalize_roll(r["roll"]) == roll]

    if not student_rows:
        return None, [], ["Roll number not found in the subject files."]

    name = next((r["name"] for r in student_rows if r["name"]), "")

    schedule = []
    missing = []

    # Preserve the subject/section combination from the enrollment files.
    seen_courses = set()
    for r in student_rows:
        course_key = (norm(r["subject"]), norm(r["section"]))
        if course_key in seen_courses:
            continue
        seen_courses.add(course_key)

        hits = []
        for slot in timetable:
            if "long_format" in slot and slot["long_format"]:
                if norm(slot.get("subject")) != norm(r["subject"]):
                    # Allow filename code to appear in long-format subject.
                    if norm(r["subject"]) not in norm(slot.get("subject")):
                        continue
                if r["section"] and clean(slot.get("section")) and norm(slot.get("section")) != norm(r["section"]):
                    continue
                if r["section"] and not clean(slot.get("section")):
                    # Long format without section cannot safely decide.
                    continue
                hit = dict(slot)
                hit["subject"] = r["subject"]
                hit["section"] = r["section"]
                hit["room"] = slot.get("room", "")
                hits.append(hit)
            else:
                line = matching_line(slot["cell"], r["subject"], r["section"])
                if line:
                    hit = dict(slot)
                    hit["subject"] = r["subject"]
                    hit["section"] = r["section"]
                    hit["room"] = extract_room(line)
                    hit["faculty"] = extract_faculty(
                        line, subject=r["subject"], section=r["section"]
                    )
                    hits.append(hit)

        if not hits:
            missing.append(
                f'{r["subject"]}' + (f' ({r["section"]})' if r["section"] else "")
            )

        schedule.extend(hits)

    # Remove duplicate matches and sort.
    unique = []
    seen = set()
    for r in schedule:
        key = (
            r["day"], r["period"], norm(r["subject"]), norm(r.get("section", "")),
            norm(r.get("room", ""))
        )
        if key not in seen:
            seen.add(key)
            unique.append(r)

    unique.sort(key=lambda x: (
        DAY_ORDER.get(x["day"], 99),
        str(x["period"]),
        norm(x["subject"])
    ))

    return {
        "roll": roll,
        "name": name,
        "courses": student_rows,
        "schedule": unique,
        "missing": sorted(set(missing), key=str.lower),
    }, student_rows, []


def make_text(student):
    lines = [
        f'{student["name"]} — Roll No: {student["roll"]}',
        "",
        "PERSONAL TIMETABLE",
        "",
    ]

    grouped = defaultdict(list)
    for x in student["schedule"]:
        grouped[x["day"]].append(x)

    for day_num in sorted(DAY_ORDER.values()):
        day_name = next(
            (name for name, number in DAY_ORDER.items() if number == day_num),
            None
        )
        if not day_name or day_name not in grouped:
            continue

        lines.append(day_name.upper())
        lines.append(
            "Time                  Subject              Section   Faculty     Location"
        )
        lines.append(
            "--------------------  -------------------  -------   ----------  --------"
        )

        for x in grouped[day_name]:
            time = x.get("time", "") or x.get("period", "")
            subject = clean(x.get("subject", ""))
            section = clean(x.get("section", "")) or "—"
            faculty = clean(x.get("faculty", "")) or "—"
            location = clean(x.get("room", "")) or "—"

            lines.append(
                f'{time:<20}  {subject:<19}  {section:<7}   '
                f'{faculty:<10}  {location:<8}'
            )
        lines.append("")

    if student["missing"]:
        lines.append("SUBJECTS WITH NO MATCHING TIMETABLE ENTRY")
        for s in student["missing"]:
            lines.append(f"- {s}")
        lines.append("")

    if not student["schedule"]:
        lines.append("No timetable entries were found for this student.")
        lines.append("")

    return "\n".join(lines).rstrip() + "\n"



def make_master_grid_text(student):
    master_periods = [
        ("I", "8:30-10:00 AM"),
        ("II", "10:15-11:45 AM"),
        ("III", "12:00-1:30 PM"),
        ("IV", "2:15-3:45 PM"),
        ("V", "4:00-5:30 PM"),
        ("VI", "5:45-7:15 PM"),
    ]

    lookup = defaultdict(list)
    for x in student["schedule"]:
        period = clean(x.get("period", "")).upper()
        if period not in {p for p, _ in master_periods}:
            tkey = re.sub(r"[^0-9apm]", "", clean(x.get("time", "")).lower())
            for p, t in master_periods:
                if re.sub(r"[^0-9apm]", "", t.lower()) == tkey:
                    period = p
                    break

        details = [clean(x["subject"])]
        meta = []
        if clean(x.get("section", "")):
            meta.append(f"Sec {clean(x['section'])}")
        if clean(x.get("faculty", "")):
            meta.append(clean(x["faculty"]))
        if clean(x.get("room", "")):
            meta.append(clean(x["room"]))
        if meta:
            details.append(" | ".join(meta))

        lookup[(x["day"], period)].append(" — ".join(details))

    lines = [
        f"PERSONAL TIME TABLE — {student['name']}",
        f"Roll No: {student['roll']}",
        "",
        "DAY | " + " | ".join(f"{p}\n{t}" for p, t in master_periods),
    ]

    for day in ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]:
        cells = [
            " / ".join(lookup.get((day, p), []))
            for p, _ in master_periods
        ]
        lines.append(day + " | " + " | ".join(cells))

    return "\n".join(lines) + "\n\n"


def save_outputs(student, folder):
    out_dir = folder / "Student Timetables"
    out_dir.mkdir(parents=True, exist_ok=True)

    safe_roll = re.sub(r"[^A-Za-z0-9_-]+", "_", student["roll"])
    safe_name = re.sub(r"[^A-Za-z0-9._ -]+", "_", student["name"]).strip() or "Student"

    # The supplied master timetable has a fixed six-period structure.
    # Keep these columns exactly, even when a student has no class in a period.
    master_periods = [
        ("I", "8:30-10:00 AM"),
        ("II", "10:15-11:45 AM"),
        ("III", "12:00-1:30 PM"),
        ("IV", "2:15-3:45 PM"),
        ("V", "4:00-5:30 PM"),
        ("VI", "5:45-7:15 PM"),
    ]

    def time_key(value):
        return re.sub(r"[^0-9apm]", "", clean(value).lower().replace("–", "-").replace("—", "-"))

    master_by_time = {time_key(t): p for p, t in master_periods}
    master_by_period = {p.upper(): p for p, _ in master_periods}

    rows = []
    for x in student["schedule"]:
        source_time = clean(x.get("time", ""))
        source_period = clean(x.get("period", "")).upper()

        # Prefer time because the PDF header has a separate period row and time row.
        canonical_period = master_by_time.get(time_key(source_time), "")
        if not canonical_period:
            canonical_period = master_by_period.get(source_period, "")

        # If an unusual source format has a time written inside the period field,
        # match that too.
        if not canonical_period:
            canonical_period = master_by_time.get(time_key(source_period), "")

        # Fall back gracefully; regular master timetable inputs will map to I-VI.
        if not canonical_period:
            canonical_period = source_period or "UNKNOWN"

        rows.append({
            "Day": x["day"],
            "Time": source_time,
            "Period": canonical_period,
            "Subject": x["subject"],
            "Section": x.get("section", ""),
            "Faculty": x.get("faculty", ""),
            "Location": x.get("room", ""),
        })

    df = pd.DataFrame(
        rows,
        columns=["Day", "Time", "Period", "Subject", "Section", "Faculty", "Location"]
    )

    xlsx_path = out_dir / f"{safe_roll}_{safe_name}_Timetable.xlsx"
    txt_path = out_dir / f"{safe_roll}_{safe_name}_Timetable.txt"

    # Build class-cell text.
    cell_lookup = defaultdict(list)
    for _, r in df.iterrows():
        details = [
            clean(r["Subject"]),
            " | ".join(
                p for p in [
                    f"Section: {clean(r['Section'])}" if clean(r["Section"]) else "",
                    f"Faculty: {clean(r['Faculty'])}" if clean(r["Faculty"]) else "",
                    f"Location: {clean(r['Location'])}" if clean(r["Location"]) else "",
                ]
                if p
            ),
        ]
        cell_lookup[(clean(r["Day"]), clean(r["Period"]))].append(
            "\n".join(p for p in details if p)
        )

    # Full master timetable day structure.
    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Timetable"

    # Match the master timetable's hierarchy.
    ws.merge_cells("A1:G1")
    ws["A1"] = "MANAGEMENT DEVELOPMENT INSTITUTE GURGAON"
    ws.merge_cells("A2:G2")
    ws["A2"] = f"PERSONAL TIME TABLE — {student['name']}"
    ws.merge_cells("A3:G3")
    ws["A3"] = f"Roll No: {student['roll']}"

    ws["A5"] = "DAY"
    for c, (period, time) in enumerate(master_periods, start=2):
        ws.cell(5, c, period)
        ws.cell(6, c, time)
    ws["A6"] = ""

    # Insert exact five weekday rows.
    for row_idx, day in enumerate(days, start=7):
        ws.cell(row_idx, 1, day)
        for col_idx, (period, _) in enumerate(master_periods, start=2):
            entries = cell_lookup.get((day, period), [])
            ws.cell(row_idx, col_idx, "\n---\n".join(entries))

    # Details sheet.
    details_ws = wb.create_sheet("Details")
    headers = ["Day", "Time", "Period", "Subject", "Section", "Faculty", "Location"]
    for c, h in enumerate(headers, start=1):
        details_ws.cell(1, c, h)
    for r_idx, (_, r) in enumerate(df.iterrows(), start=2):
        for c_idx, h in enumerate(headers, start=1):
            details_ws.cell(r_idx, c_idx, r[h])

    # Formatting to resemble the master timetable grid.
    thin = openpyxl.styles.Side(style="thin")
    border = openpyxl.styles.Border(left=thin, right=thin, top=thin, bottom=thin)

    for row in ws.iter_rows(min_row=5, max_row=11, min_col=1, max_col=7):
        for cell in row:
            cell.border = border
            cell.alignment = openpyxl.styles.Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

    for row in ws.iter_rows(min_row=1, max_row=3, min_col=1, max_col=7):
        for cell in row:
            cell.alignment = openpyxl.styles.Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

    ws["A1"].font = openpyxl.styles.Font(bold=True, size=14)
    ws["A2"].font = openpyxl.styles.Font(bold=True, size=13)
    ws["A3"].font = openpyxl.styles.Font(bold=True, size=11)

    ws.column_dimensions["A"].width = 17
    for col in range(2, 8):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 23

    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 24
    ws.row_dimensions[3].height = 20
    ws.row_dimensions[5].height = 23
    ws.row_dimensions[6].height = 32
    for r in range(7, 12):
        ws.row_dimensions[r].height = 82

    ws.freeze_panes = "B7"
    ws.sheet_view.showGridLines = False
    ws.print_area = "A1:G11"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    details_ws.freeze_panes = "A2"
    details_ws.auto_filter.ref = details_ws.dimensions
    details_ws.sheet_view.showGridLines = False
    for c in range(1, 8):
        details_ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 22

    wb.save(xlsx_path)

    # Copyable text uses the same master-grid representation.
    master_text = make_master_grid_text(student)

    txt_path.write_text(master_text, encoding="utf-8")
    return xlsx_path, txt_path





import streamlit as st
import html
import tempfile
import shutil
from pathlib import Path
from collections import defaultdict

st.set_page_config(page_title="Student Timetable Generator", page_icon="📅", layout="wide")

DATA_DIR = Path(__file__).parent / "data"
MASTER_PERIODS = [
    ("I","8:30–10:00 AM"),
    ("II","10:15–11:45 AM"),
    ("III","12:00–1:30 PM"),
    ("IV","2:15–3:45 PM"),
    ("V","4:00–5:30 PM"),
    ("VI","5:45–7:15 PM"),
]

def find_master_timetable():
    pdfs=sorted(DATA_DIR.glob("*.pdf"))
    if pdfs:
        named=[p for p in pdfs if "timetable" in p.stem.lower()]
        return (named or pdfs)[0]
    return None

def canonical_period(period,time):
    p=clean(period).upper()
    valid={x[0] for x in MASTER_PERIODS}
    if p in valid: return p
    key=re.sub(r"[^0-9apm]","",clean(time).lower())
    for mp,mt in MASTER_PERIODS:
        if key and key==re.sub(r"[^0-9apm]","",mt.lower()):
            return mp
    return p

def render_master_grid(student):
    lookup=defaultdict(list)
    for x in student["schedule"]:
        p=canonical_period(x.get("period",""),x.get("time",""))
        parts=[clean(x.get("subject",""))]
        if clean(x.get("section","")): parts.append(f"Section: {clean(x['section'])}")
        if clean(x.get("faculty","")): parts.append(f"Faculty: {clean(x['faculty'])}")
        if clean(x.get("room","")): parts.append(f"Location: {clean(x['room'])}")
        lookup[(x["day"],p)].append(parts)

    rows=[
        '<tr><th rowspan="2" class="day">DAY</th>'+''.join(
            f'<th class="period">{html.escape(p)}</th>' for p,_ in MASTER_PERIODS
        )+'</tr>',
        '<tr>'+''.join(
            f'<th class="time">{html.escape(t)}</th>' for _,t in MASTER_PERIODS
        )+'</tr>'
    ]
    for day in ["Monday","Tuesday","Wednesday","Thursday","Friday"]:
        cells=[f'<td class="day">{day}</td>']
        for p,_ in MASTER_PERIODS:
            entries=lookup.get((day,p),[])
            if not entries:
                cells.append('<td></td>'); continue
            blocks=[]
            for parts in entries:
                blocks.append(
                    '<div class="subject">'+html.escape(parts[0])+'</div>'+
                    ''.join('<div class="detail">'+html.escape(v)+'</div>' for v in parts[1:])
                )
            cells.append('<td>'+'<hr class="sep">'.join(blocks)+'</td>')
        rows.append('<tr>'+''.join(cells)+'</tr>')
    return """
<style>
.tt{width:100%;border-collapse:collapse;table-layout:fixed;background:#0e1117;color:#f0f2f6}
.tt th,.tt td{border:1px solid #3b414b;text-align:center;vertical-align:middle;padding:7px}
.tt th{color:#f0f2f6}
.tt .day{width:11%;background:#161b22;color:#f0f2f6;font-weight:700}
.tt .period{background:#1b222c;color:#f0f2f6;font-size:14px}
.tt .time{background:#11161d;color:#c9d1d9;font-size:12px}
.tt tr td:not(.day){background:#0e1117;color:#f0f2f6}
.tt td{height:92px;font-size:11px}
.subject{font-size:14px;font-weight:700;color:#ffffff;margin-bottom:4px}
.detail{font-size:11px;line-height:1.3;color:#e6edf3}
.sep{border:0;border-top:1px solid #30363d;margin:5px 0}
</style>
""" + f"""
<div style="text-align:center;margin-bottom:8px">
<div style="font-size:20px;font-weight:700">MANAGEMENT DEVELOPMENT INSTITUTE GURGAON</div>
<div style="font-size:16px;font-weight:600">PERSONAL TIME TABLE — {html.escape(student["name"])}</div>
<div style="font-size:13px">Roll No: {html.escape(student["roll"])}</div>
</div>
<table class="tt"><tbody>{''.join(rows)}</tbody></table>
"""

def make_downloads(student):
    tmp=Path(tempfile.mkdtemp(prefix="ttweb_"))
    try:
        xlsx,txt=save_outputs(student,tmp)
        return xlsx.read_bytes(),txt.read_bytes()
    finally:
        shutil.rmtree(tmp,ignore_errors=True)

tt_path=find_master_timetable()
if not tt_path:
    st.error("The administrator has not installed a timetable.")
    st.stop()

st.markdown(
    "<h1 style='text-align:center;margin-bottom:0'>📅 Student Timetable Generator</h1>"
    "<p style='text-align:center;margin-top:4px'>Enter your roll number to view your personal timetable.</p>",
    unsafe_allow_html=True
)

with st.form("roll_form"):
    roll=st.text_input("Roll Number",placeholder="e.g. 25P060")
    submitted=st.form_submit_button("Generate My Timetable")

if submitted:
    roll=normalize_roll(roll)
    if not roll:
        st.warning("Please enter your roll number.")
        st.stop()
    try:
        records=parse_all_subject_files(DATA_DIR,tt_path)
        timetable=parse_timetable(tt_path)
        student,_,errors=build_student_timetable(records,timetable,roll)
    except Exception as e:
        st.error(f"Could not generate the timetable: {e}")
        st.stop()

    if student is None:
        st.error(f"Roll number {roll} was not found.")
        st.stop()

    if student["missing"]:
        st.warning("Some enrolled subjects could not be matched: "+", ".join(student["missing"]))

    st.success(f'Generated timetable for {student["name"]} ({student["roll"]}) — {len(student["schedule"])} class entries.')
    st.markdown(render_master_grid(student),unsafe_allow_html=True)

    xb,tb=make_downloads(student)
    c1,c2=st.columns(2)
    with c1:
        st.download_button("Download Excel Timetable",xb,
            file_name=f"{student['roll']}_timetable.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
    with c2:
        st.download_button("Download Copyable Timetable",tb,
            file_name=f"{student['roll']}_timetable.txt",
            mime="text/plain",use_container_width=True)
else:
    st.info("Enter your roll number and press Enter or click Generate My Timetable.")

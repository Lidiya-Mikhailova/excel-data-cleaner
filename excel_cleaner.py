"""
Excel / CSV Data Cleaner and Normalizer
----------------------------------------
This script processes raw text data files (TXT, CSV) and outputs:
- Cleaned and deduplicated data
- Excel report with auto-fitted columns
- CSV report for universal compatibility
- Audit flags for incomplete or suspicious data

Features:
- Normalizes Russian FIO (Last, First, Middle) with proper capitalization
- Corrects initials (e.g., Петров П. П.)
- Normalizes city names with proper title case
- Preserves regular text capitalization (first letter of sentences)
- Deduplicates data
- Generates logs of processing

Author: Your Name
"""

import re
import os
import platform
from pathlib import Path
from typing import List
import pandas as pd
from datetime import datetime

# PATH SETTINGS
RAW_FILE_DIR: Path = Path("raw_data")  # Directory with input files
OUTPUT_DIR: Path = Path("output")      # Directory for output files
OUTPUT_DIR.mkdir(exist_ok=True)

OUTPUT_EXCEL: Path = OUTPUT_DIR / "cleaned_data_report.xlsx"
OUTPUT_CSV: Path = OUTPUT_DIR / "cleaned_data_report.csv"
LOG_FILE: Path = OUTPUT_DIR / "processing_log.txt"

# LOGGING FUNCTION
def write_log(message: str) -> None:
    """Append a timestamped message to the log file."""
    timestamp: str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] {message}\n")

# FILE OPENING FUNCTION
def open_file(path: Path) -> None:
    """Open a file using the default application depending on OS."""
    if platform.system() == "Windows":
        os.startfile(path)
    elif platform.system() == "Darwin":  # macOS
        os.system(f'open "{path}"')
    else:  # Linux
        os.system(f'xdg-open "{path}"')

# NORMALIZATION FUNCTIONS
def normalize_fio(value: str) -> str:
    """
    Normalize Russian FIO:
    - Иванов иван иванович -> Иванов Иван Иванович
    - Инициалы: Петров П. П.
    """
    if not value:
        return value
    words = value.split()
    normalized_words = []
    for word in words:
        # Initials like "П."
        if len(word) == 2 and word[1] == ".":
            normalized_words.append(word[0].upper() + ".")
        else:
            normalized_words.append(word.capitalize())
    return " ".join(normalized_words)

def normalize_location(value: str) -> str:
    """
    Normalize city or location names:
    - санкт-петербург -> Санкт-Петербург
    """
    if not value:
        return value
    return value.title()

def normalize_sentence(value: str) -> str:
    """
    Normalize a normal sentence:
    - Only the first letter of the sentence capitalized
    """
    value = value.strip()
    if not value:
        return value
    return value[0].upper() + value[1:]

def normalize_text(value: str, column_name: str) -> str:
    """
    Main normalization dispatcher based on column semantics.
    """
    if not value:
        return value
    if "сотрудник" in column_name.lower():
        return normalize_fio(value)
    elif "город" in column_name.lower() or "место" in column_name.lower():
        return normalize_location(value)
    else:
        return normalize_sentence(value)

# MAIN PROCESSING ENGINE
def run_data_processing() -> None:
    write_log("--- DATA PROCESSING STARTED ---")
    rows: List[List[str]] = []

    try:
        # Automatically select the first available input file
        input_files: List[Path] = list(RAW_FILE_DIR.glob("*.*"))
        if not input_files:
            raise FileNotFoundError("No files found in raw_data directory.")

        target_file: Path = input_files[0]
        write_log(f"Target file acquired: {target_file.name}")

        # Read raw lines and strip empty lines
        with open(target_file, "r", encoding="utf-8") as f:
            lines: List[str] = [line.strip() for line in f if line.strip()]

        if not lines:
            raise ValueError("Input file contains no data.")

        # 1. Parsing and normalization
        for line in lines:
            parts: List[str] = [p.strip() for p in re.split(r"[;,]", line)]
            # For first row (header) just strip spaces
            if len(rows) == 0:
                clean_parts = [p for p in parts]
            else:
                # Normalize based on column
                clean_parts = [normalize_text(p, col_name) for p, col_name in zip(parts, rows[0])]
            rows.append(clean_parts)

        # 2. Structural alignment
        header: List[str] = rows[0]
        data_rows: List[List[str]] = rows[1:]
        max_cols_found: int = max(len(r) for r in rows)

        if len(header) < max_cols_found:
            write_log(f"Detected {max_cols_found} columns. Extending header.")
            for i in range(len(header), max_cols_found):
                header.append(f"Additional_Field_{i + 1}")

        normalized_data: List[List[str]] = [
            r + [""] * (len(header) - len(r)) if len(r) < len(header) else r[:len(header)]
            for r in data_rows
        ]

        df: pd.DataFrame = pd.DataFrame(normalized_data, columns=header)

        # 3. Deduplication
        initial_count: int = len(df)
        df = df.drop_duplicates().reset_index(drop=True)
        write_log(f"Duplicates removed: {initial_count - len(df)}")

        # 4. Integrity audit
        def audit_integrity(row: pd.Series) -> str:
            flags: List[str] = []
            if "" in row.values:
                flags.append("Incomplete Entry")
            for val in row.values:
                s = str(val)
                if s.isdigit():
                    continue
                if 0 < len(s) < 2:
                    flags.append("Validation Required")
                    break
            return ", ".join(flags)

        df["Audit_Flag"] = df.apply(audit_integrity, axis=1)

        # 5. Export
        df.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig", sep=";")

        with pd.ExcelWriter(OUTPUT_EXCEL, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Cleaned_Data")
            worksheet = writer.sheets["Cleaned_Data"]
            for i, col_name in enumerate(df.columns):
                max_len = max(df[col_name].astype(str).map(len).max(), len(col_name)) + 2
                worksheet.set_column(i, i, min(max_len, 60))

        write_log("--- PROCESSING COMPLETED SUCCESSFULLY ---")
        print(f"Data processing complete: {len(df)} records finalized.")
        open_file(OUTPUT_EXCEL)

    except Exception as e:
        import traceback
        write_log(traceback.format_exc())
        print(f"Critical Error: {e}")


if __name__ == "__main__":
    run_data_processing()

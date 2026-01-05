# DataCleansing_Tool (Pro Version)

An automated Python-based solution designed to transform raw, unstructured text batches into clean, reporting-ready Excel and CSV datasets.

## Key Features
- Dynamic Structural Repair: Automatically detects and aligns rows with inconsistent column counts or misplaced delimiters.
- Linguistic Standardization: Enforces consistent Title Case capitalization and strips whitespace to ensure data uniformity.
- Smart Deduplication: Identifies and removes redundant records, maintaining a detailed audit trail in the log file.
- Integrity Auditing: Generates an Audit_Flag column to highlight records that require manual verification (e.g., missing fields or corrupted entries).
- Pro-Grade Export: Delivers polished Excel files with auto-fitted column widths and universal CSV files (UTF-8-SIG).

## Setup & Usage
1. Place your raw input files in the `raw_data/` directory.
2. Run `data_processor.py`.
3. Retrieve processed files and the change log from the `output/` directory.

## Technical Specifications
- Language: Python 3.13
- Libraries: Pandas, XlsxWriter, Re
- Logging: Comprehensive processing_log.txt generated per batch.

# TTP Sanitizer

TTP Sanitizer is a lightweight utility designed to simplify the process of extracting, sanitizing, and transforming raw Excel spreadsheet data into multiple formats required by administrative systems.

The tool provides an interactive interface for reviewing and editing extracted data, ensuring that outputs are formatted correctly and consistently for downstream use, specifically:

- **TTP Enterable Data** — For external systems requiring sanitized, structured input.  
- **Internal Tracker Data** — For local administrative tracking and record-keeping.  
- **One45 Data** — For systems such as One45 requiring customized data structures.  

---

## Project Purpose

Many academic and administrative environments rely on semi-structured spreadsheets for scheduling and attendance tracking. Manually cleaning and transforming these files is repetitive, error-prone, and inefficient.

TTP Sanitizer provides:

✔ Automated extraction and initial sanitization of raw Excel data.  
✔ A user-friendly web interface for manual review and edits.  
✔ Export-ready output for multiple target systems.  
✔ Built-in tools for dynamic row insertion and inline editing.  
✔ Consistent formatting across all outputs.  

---

## Core Technologies

- **Python**  
- **Pandas** — for data manipulation  
- **Streamlit** — for the interactive web interface  
- **Custom Extraction & Cleaning Functions** (`extractionFunctions.py`)  

---

## Basic Workflow

1. **File Selection**  
   Enter the path to your raw Excel spreadsheet and basic rotation details.

2. **Automated Extraction**  
   The tool extracts relevant data blocks and applies initial sanitization logic.

3. **Data Review & Editing**  
   Through the Streamlit interface:  
   - Review session information  
   - Insert rows dynamically if needed  
   - Edit individual fields (dates, times, session types, etc.)

4. **Format Selection**  
   Process the cleaned data into one of the supported output formats:  
   - TTPS Data  
   - Internal Tracker  
   - One45  

5. **Output Review**  
   Final processed data is displayed for verification before export.

---

## Design Philosophy

This utility prioritizes:

- **Simplicity** — Minimal UI, focused on efficiency, not over-engineering  
- **Transparency** — Edits are clearly visible and applied only when confirmed  
- **Separation of Logic** — Data processing and UI interaction remain modular and maintainable  
- **Practicality** — Built to solve a specific, repetitive administrative task, not for generalized data pipelines  

---

## Disclaimer

This project is a small internal-facing utility. It is designed with pragmatic functionality in mind, not as a polished commercial software product.

Viewers and contributors should understand:

- This is not intended as a code review exercise  
- Documentation focuses on practical functionality, not deep code theory  
- Assumes basic familiarity with Python and Pandas  

---

## Future Improvements (Optional)

Potential features under consideration:

- Enhanced input validation for common editing mistakes  
- Export-to-Google Sheets integration  
- File download buttons for processed outputs  
- More granular user prompts for unsaved changes  

---

## Contact

For questions or suggestions, contact the wmaxxing on github.


# AI Agent Instructions for Commercialista

## Project overview
- This repository is a small Python/Flask application for running an Excel processing pipeline on Windows.
- `app.py` is the web frontend and orchestrator: it exposes a Flask UI, selects folders, runs `run_pipeline.py`, and manages Excel template copy/paste and output file handling.
- `run_pipeline.py` orchestrates `latest_zip.py` and `process_zip_excel.py` to process the most recent ZIP archive.
- `process_zip_excel.py` performs Excel automation via `pywin32`/`win32com` and rewrites the ZIP archive.
- `README_APP.md` describes installation and basic app usage.

## Key files
- `app.py`: main Flask app, async pipeline execution, Windows folder picker, Excel output file creation, clipboard handling, and open file support.
- `run_pipeline.py`: pipeline entrypoint invoked by `app.py`; finds latest ZIP then processes it.
- `process_zip_excel.py`: ZIP extraction, Excel column move, and ZIP reconstruction.
- `latest_zip.py`: utility to find the newest `.zip` file in a given folder.
- `templates/index.html`: frontend UI template.

## Environment & dependencies
- Target platform: Windows.
- Python 3.7+ is expected.
- Required packages include at least:
  - `flask`
  - `openpyxl`
  - `pandas`
  - `pyperclip`
  - `pywin32` (for Excel COM automation and clipboard use)
- `tkinter` is used for the Windows folder picker; it is usually bundled with Python on Windows.
- Excel must be installed on the host machine for `pywin32` automation to work.

## Running the app
- Launch from the repository root:
  - `python app.py`
- Access the UI at `http://localhost:5000`.
- The app uses a background thread and global state (`execution_result`) for async status polling.

## Important conventions
- Preserve existing Windows-specific behavior unless asked to make it cross-platform.
- `app.py` uses `Path(__file__).parent` as the working directory when calling `run_pipeline.py`; keep this relative execution behavior intact.
- Excel sheet names are hardcoded in the pipeline:
  - `Fatture Attive` for the main pipeline output
  - `Fatture Passive` for the passive invoice flow
- `process_zip_excel.py` expects the ZIP to contain at least one Excel file and relies on Excel COM to move the source column.

## Guidance for agents
- Focus fixes on robust folder validation, error handling, and Windows Excel automation reliability.
- Avoid introducing non-Windows platform behavior unless the user explicitly requests cross-platform support.
- Keep the UI/data flow simple: folder selection → run pipeline → poll status → open file.
- Do not assume a package manager file exists; use `README_APP.md` and the source imports to infer dependencies.

## Suggested next customization
- Add a `.github/copilot-instructions.md` if you want to provide user-facing behavior guidance for future AI tools, especially around Windows Excel automation and pipeline execution.

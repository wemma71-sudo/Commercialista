import os
import zipfile
import tempfile
import shutil
from pathlib import Path
import sys
from typing import Optional


def find_first_excel(root: Path) -> Optional[Path]:
    for p in root.rglob('*'):
        if p.suffix.lower() in ('.xlsx', '.xlsm', '.xls') and p.is_file():
            return p
    return None


def copy_data_excluding_header_to_clipboard(sheet) -> None:
    used = sheet.UsedRange
    last_row = used.Rows.Count
    last_col = used.Columns.Count
    if last_row <= 1:
        return
    start = sheet.Cells(2, 1)
    end = sheet.Cells(last_row, last_col)
    rng = sheet.Range(start, end)
    rng.Copy()


def move_excel_column_after(excel_path: Path, src_col: str = 'Y', after_col: str = 'E') -> None:
    try:
        import win32com.client
    except ImportError as exc:
        raise ImportError('pywin32 is required. Install with: pip install pywin32') from exc

    excel = win32com.client.Dispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False

    workbook = None
    try:
        workbook = excel.Workbooks.Open(str(excel_path), False, False)
        if workbook is None:
            raise RuntimeError(f'Excel failed to open workbook: {excel_path}')
        sheet = workbook.Worksheets(1)
        src_range = f'{src_col}:{src_col}'
        insert_before_col = chr(ord(after_col.upper()) + 1)
        insert_range = f'{insert_before_col}:{insert_before_col}'
        xlToRight = -4161

        sheet.Columns(src_range).Cut()
        sheet.Columns(insert_range).Insert(Shift=xlToRight)

        workbook.Save()
        copy_data_excluding_header_to_clipboard(sheet)
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=True)
        excel.Quit()


def recreate_zip_from_folder(zip_path: Path, folder: Path) -> None:
    backup = zip_path.with_suffix(zip_path.suffix + '.bak')
    shutil.copy2(zip_path, backup)

    tmpzip = zip_path.with_suffix('.tmp.zip')
    with zipfile.ZipFile(tmpzip, 'w', compression=zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(folder):
            for f in files:
                full = Path(root) / f
                arc = full.relative_to(folder)
                z.write(full, arc.as_posix())

    shutil.move(str(tmpzip), str(zip_path))


def process_zip(zip_path: Path, col_from: str, col_to: str) -> int:
    if not zip_path.exists():
        alt = zip_path.with_suffix('.zip')
        if alt.exists():
            zip_path = alt
        else:
            print(f'Zip not found: {zip_path}', file=sys.stderr)
            return 2

    with tempfile.TemporaryDirectory() as td:
        td_path = Path(td)
        with zipfile.ZipFile(zip_path, 'r') as z:
            z.extractall(td)

        excel = find_first_excel(td_path)
        if excel is None:
            print('No Excel file found inside the archive', file=sys.stderr)
            return 3

        try:
            move_excel_column_after(excel, src_col=col_from, after_col=col_to)
        except ImportError as exc:
            print(exc, file=sys.stderr)
            return 4
        except Exception as exc:
            print(f'Excel automation failed: {exc}', file=sys.stderr)
            return 5

        recreate_zip_from_folder(zip_path, td_path)

    return 0


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Usage: process_zip_excel.py <zip-path>', file=sys.stderr)
        raise SystemExit(1)
    zp = Path(sys.argv[1])
    rc = process_zip(zp, sys.argv[2], sys.argv[3])
    raise SystemExit(rc)

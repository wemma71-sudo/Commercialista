#!/usr/bin/env python3
"""
Web application per eseguire run_pipeline.py con gestione template Excel
Fornisce un'interfaccia web con selezione cartella e esecuzione pipeline
"""

import subprocess
import sys
from pathlib import Path
from flask import Flask, render_template, request, jsonify
import json
import os
from threading import Thread
import queue
from datetime import datetime
from shutil import copy2
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import pyperclip
import webbrowser
import zipfile
import tempfile
import pandas as pd
import pythoncom

app = Flask(__name__)
app.config['TEMPLATES_AUTO_RELOAD'] = True

# Queue per gestire output dei processi
process_queue = queue.Queue()
# Variabili globali per il risultato
execution_result = {}


@app.route('/')
def index():
    """Pagina principale"""
    return render_template('index.html')


@app.route('/api/browse-folder', methods=['POST'])
def browse_folder():
    """
    Apre il dialog di selezione cartella di Windows
    Ritorna il percorso selezionato
    """
    try:
        import tkinter as tk
        from tkinter import filedialog
        
        folder_type = request.json.get('type', 'input')
        title_map = {
            'input': 'Seleziona cartella per l\'elaborazione',
            'template': 'Seleziona cartella con template'
        }
        
        # Crea una finestra Tkinter nascosta
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        
        # Apri dialog di selezione cartella
        folder_path = filedialog.askdirectory(
            title=title_map.get(folder_type, 'Seleziona cartella'),
            initialdir=os.path.expanduser("~\\Documents")
        )
        
        root.destroy()
        
        if folder_path:
            return jsonify({
                'success': True,
                'folder': folder_path
            })
        else:
            return jsonify({
                'success': False,
                'error': 'Nessuna cartella selezionata'
            })
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 400


@app.route('/api/run-pipeline', methods=['POST'])
def run_pipeline():
    """
    Esegue il workflow completo:
    1. Copia template Excel e lo rinomina con data
    2. Esegue run_pipeline.py
    3. Apre file Excel e incolla clipboard
    4. Salva il file
    """
    global execution_result
    execution_result = {}
    
    try:
        data = request.json
        input_folder = data.get('folder', '').strip()
        template_folder = data.get('template_folder', '').strip()
        
        if not input_folder:
            return jsonify({
                'success': False,
                'error': 'Cartella di input non specificata'
            }), 400
        
        if not template_folder:
            return jsonify({
                'success': False,
                'error': 'Cartella con template non specificata'
            }), 400
        
        # Verifica che le cartelle esistano
        if not os.path.isdir(input_folder):
            return jsonify({
                'success': False,
                'error': f'Cartella di input non trovata: {input_folder}'
            }), 400
        
        if not os.path.isdir(template_folder):
            return jsonify({
                'success': False,
                'error': f'Cartella template non trovata: {template_folder}'
            }), 400
        
        # Avvia il processo in background
        def run_process():
            global execution_result
            try:
                # Step 1: Trova e copia il template
                template_path = Path(template_folder) / 'template_ReportIncassiPagamenti.xlsx'
                
                if not template_path.exists():
                    execution_result = {
                        'status': 'error',
                        'error': f'Template non trovato: {template_path}',
                        'success': False
                    }
                    return
                
                # Genera nome file con data corrente
                today = datetime.now().strftime('%Y%m%d')
                output_filename = f'{today}_ReportIncassiPagamenti.xlsx'
                output_path = Path(template_folder) / output_filename
                
                # Copia il template
                copy2(str(template_path), str(output_path))
                
                # Step 2: Esegui run_pipeline.py
                result = subprocess.run(
                    [sys.executable, 'run_pipeline.py', input_folder],
                    cwd=Path(__file__).parent,
                    capture_output=True,
                    text=True
                )
                
                if result.returncode != 0:
                    execution_result = {
                        'status': 'error',
                        'error': f'Pipeline fallita: {result.stderr}',
                        'success': False,
                        'output_file': str(output_path)
                    }
                    return
                
                # Step 3: Apri file Excel e incolla clipboard
                try:
                    wb = load_workbook(str(output_path))
                    
                    # Accedi al foglio "Fatture Attive"
                    if 'Fatture Attive' not in wb.sheetnames:
                        execution_result = {
                            'status': 'error',
                            'error': f'Foglio "Fatture Attive" non trovato. Fogli disponibili: {wb.sheetnames}',
                            'success': False,
                            'output_file': str(output_path)
                        }
                        wb.close()
                        return
                    
                    ws = wb['Fatture Attive']
                    
                    # Copia il contenuto della clipboard
                    clipboard_content = pyperclip.paste()
                    
                    if not clipboard_content:
                        execution_result = {
                            'status': 'warning',
                            'error': 'Clipboard è vuoto, file creato senza dati',
                            'success': False,
                            'output_file': str(output_path)
                        }
                        wb.close()
                        return
                    
                    # Incolla il contenuto nella cella A2
                    # Se è una lista di righe, inseriamo i dati
                    lines = clipboard_content.strip().split('\n')
                    for row_idx, line in enumerate(lines, start=2):
                        # Se la riga contiene tab, è una tabella
                        if '\t' in line:
                            cells = line.split('\t')
                            for col_idx, cell_value in enumerate(cells, start=1):
                                ws.cell(row=row_idx, column=col_idx).value = cell_value
                        else:
                            ws.cell(row=row_idx, column=1).value = line
                    
                    # Step 4: Salva il file
                    wb.save(str(output_path))
                    wb.close()
                    
                    execution_result = {
                        'status': 'completed',
                        'success': True,
                        'message': f'Pipeline completata con successo',
                        'output_file': str(output_path),
                        'stdout': result.stdout,
                        'stderr': result.stderr
                    }
                    
                except Exception as e:
                    execution_result = {
                        'status': 'error',
                        'error': f'Errore durante l\'elaborazione Excel: {str(e)}',
                        'success': False,
                        'output_file': str(output_path)
                    }
                
            except Exception as e:
                execution_result = {
                    'status': 'error',
                    'error': str(e),
                    'success': False
                }
        
        # Avvia il processo in un thread separato
        thread = Thread(target=run_process, daemon=True)
        thread.start()
        
        return jsonify({
            'success': True,
            'message': 'Pipeline in esecuzione...',
            'input_folder': input_folder,
            'template_folder': template_folder
        })
    
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 400


@app.route('/api/check-status', methods=['GET'])
def check_status():
    """
    Controlla lo stato della pipeline in esecuzione
    """
    global execution_result
    
    if execution_result and execution_result.get('status') in ['completed', 'error', 'warning']:
        result = execution_result.copy()
        execution_result = {}
        return jsonify(result)
    
    return jsonify({'status': 'running'})


@app.route('/api/open-file', methods=['POST'])
def open_file():
    """
    Apre il file Excel generato
    """
    try:
        data = request.json
        file_path = data.get('file_path', '').strip()
        
        if not file_path or not os.path.isfile(file_path):
            return jsonify({
                'success': False,
                'error': f'File non trovato: {file_path}'
            }), 400
        
        # Apri il file con l'applicazione predefinita (Excel)
        os.startfile(file_path)
        
        return jsonify({
            'success': True,
            'message': f'File aperto: {file_path}'
        })
    
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 400


def find_latest_zip(folder: Path) -> Path | None:
    """Trova l'ultimo file .zip nella cartella"""
    zips = list(folder.glob("*.zip"))
    if not zips:
        return None
    return max(zips, key=lambda p: p.stat().st_mtime)


def find_first_excel(root: Path) -> Path | None:
    """Trova il primo file Excel nella cartella (ricorsivo)"""
    for p in root.rglob('*'):
        if p.suffix.lower() in ('.xlsx', '.xlsm', '.xls') and p.is_file():
            return p
    return None


def move_excel_column(excel_path: Path, src_col: str = 'W', after_col: str = 'E') -> None:
    """Sposta una colonna Excel dopo un'altra usando win32com"""
    try:
        import win32com.client
    except ImportError as exc:
        raise ImportError('pywin32 is required. Install with: pip install pywin32') from exc

    pythoncom.CoInitialize()
    try:
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
        finally:
            if workbook is not None:
                workbook.Close(SaveChanges=True)
            excel.Quit()
    finally:
        pythoncom.CoUninitialize()


def copy_data_excluding_header_to_clipboard(excel_path: Path) -> None:
    """Copia i dati (escluso header) del primo foglio Excel in clipboard"""
    try:
        import win32com.client
    except ImportError as exc:
        raise ImportError('pywin32 is required. Install with: pip install pywin32') from exc

    pythoncom.CoInitialize()
    try:
        excel = win32com.client.Dispatch('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False

        workbook = None
        try:
            workbook = excel.Workbooks.Open(str(excel_path), False, False)
            sheet = workbook.Worksheets(1)
            used = sheet.UsedRange
            last_row = used.Rows.Count
            last_col = used.Columns.Count
            if last_row <= 1:
                return
            start = sheet.Cells(2, 1)
            end = sheet.Cells(last_row, last_col)
            rng = sheet.Range(start, end)
            rng.Copy()
        finally:
            if workbook is not None:
                workbook.Close(SaveChanges=False)
            excel.Quit()
    finally:
        pythoncom.CoUninitialize()


@app.route('/api/process-passive', methods=['POST'])
def process_passive():
    """
    Elabora le fatture passive:
    1. Trova l'ultimo .zip in input folder
    2. Unzippa e estrae Excel
    3. Sposta colonna W dopo E
    4. Copia dati (escluso header) in clipboard
    5. Apre il file output
    6. Incolla nel foglio "Fatture Passive" a partire da A2
    7. Salva il file
    """
    global execution_result
    execution_result = {}
    
    try:
        data = request.json
        input_folder = data.get('input_folder', '').strip()
        output_file = data.get('output_file', '').strip()
        
        if not input_folder:
            return jsonify({
                'success': False,
                'error': 'Cartella di input non specificata'
            }), 400
        
        if not output_file:
            return jsonify({
                'success': False,
                'error': 'File di output non specificato'
            }), 400
        
        # Verifica che le cartelle/file esistano
        if not os.path.isdir(input_folder):
            return jsonify({
                'success': False,
                'error': f'Cartella di input non trovata: {input_folder}'
            }), 400
        
        if not os.path.isfile(output_file):
            return jsonify({
                'success': False,
                'error': f'File di output non trovato: {output_file}'
            }), 400
        
        # Avvia il processo in background
        def run_process():
            global execution_result
            try:
                # Step 1: Trova l'ultimo file .zip
                input_path = Path(input_folder)
                zip_file = find_latest_zip(input_path)
                
                if not zip_file:
                    execution_result = {
                        'status': 'error',
                        'error': f'Nessun file .zip trovato in: {input_folder}',
                        'success': False
                    }
                    return
                
                # Step 2: Unzippa il file
                with tempfile.TemporaryDirectory() as temp_dir:
                    temp_path = Path(temp_dir)
                    with zipfile.ZipFile(zip_file, 'r') as z:
                        z.extractall(temp_path)
                    
                    # Step 3: Trova il file Excel estratto
                    excel_file = find_first_excel(temp_path)
                    if not excel_file:
                        execution_result = {
                            'status': 'error',
                            'error': f'Nessun file Excel trovato nel .zip: {zip_file}',
                            'success': False
                        }
                        return
                    
                    # Step 4: Sposta colonna W dopo E
                    try:
                        move_excel_column(excel_file, src_col='W', after_col='E')
                    except Exception as e:
                        execution_result = {
                            'status': 'error',
                            'error': f'Errore durante lo spostamento colonna: {str(e)}',
                            'success': False
                        }
                        return
                    
                    # Step 5: Copia dati in clipboard
                    try:
                        copy_data_excluding_header_to_clipboard(excel_file)
                    except Exception as e:
                        execution_result = {
                            'status': 'error',
                            'error': f'Errore durante la copia in clipboard: {str(e)}',
                            'success': False
                        }
                        return
                    
                    # Step 6: Apri il file output e incolla nel foglio "Fatture Passive"
                    try:
                        output_wb = load_workbook(output_file)
                        
                        if 'Fatture Passive' not in output_wb.sheetnames:
                            execution_result = {
                                'status': 'error',
                                'error': f'Foglio "Fatture Passive" non trovato. Fogli disponibili: {output_wb.sheetnames}',
                                'success': False
                            }
                            output_wb.close()
                            return
                        
                        ws = output_wb['Fatture Passive']
                        
                        # Leggi il file Excel estratto usando pandas (supporta sia .xls che .xlsx)
                        try:
                            df = pd.read_excel(excel_file, sheet_name=0)
                        except Exception as e:
                            execution_result = {
                                'status': 'error',
                                'error': f'Errore nella lettura del file Excel: {str(e)}',
                                'success': False
                            }
                            output_wb.close()
                            return
                        
                        # Scrivi i dati nel foglio "Fatture Passive" a partire da A2
                        for row_idx, (_, row) in enumerate(df.iterrows(), start=2):
                            for col_idx, (_, value) in enumerate(row.items(), start=1):
                                ws.cell(row=row_idx, column=col_idx).value = value
                        
                        # Step 7: Salva il file
                        output_wb.save(output_file)
                        output_wb.close()
                        
                        execution_result = {
                            'status': 'completed',
                            'success': True,
                            'message': f'Fatture passive elaborate con successo',
                            'output_file': output_file
                        }
                        
                    except Exception as e:
                        execution_result = {
                            'status': 'error',
                            'error': f'Errore durante l\'elaborazione del file output: {str(e)}',
                            'success': False
                        }
                
            except Exception as e:
                execution_result = {
                    'status': 'error',
                    'error': str(e),
                    'success': False
                }
        
        # Avvia il processo in un thread separato
        thread = Thread(target=run_process, daemon=True)
        thread.start()
        
        return jsonify({
            'success': True,
            'message': 'Elaborazione fatture passive in corso...'
        })
    
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 400


if __name__ == '__main__':
    print("Applicazione avviata su http://localhost:5000")
    print("Premi Ctrl+C per fermare l'applicazione")
    app.run(debug=False, host='localhost', port=5000)

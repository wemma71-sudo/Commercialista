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
import re
import time
import urllib.request
import urllib.error
from dotenv import load_dotenv

# Moduli per l'invio al Sistema Tessera Sanitaria
import genera_invio_ts
import invia_ts
import ricevute_ts

# Carica le variabili d'ambiente dal file .env
load_dotenv()

app = Flask(__name__)
app.config['TEMPLATES_AUTO_RELOAD'] = True

# Queue per gestire output dei processi
process_queue = queue.Queue()
# Variabili globali per il risultato
execution_result = {}


@app.route('/')
def index():
    """Pagina principale"""
    ts_ambiente = os.getenv('TS_AMBIENTE', 'test').strip().lower()
    return render_template('index.html', ts_ambiente=ts_ambiente)


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
            'template': 'Seleziona cartella con template',
            'private': 'Seleziona cartella di input fatture private'
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


import re
import json
import urllib.request


def estrai_contesto_intorno_s(testo: str, righe_contorno: int = 4) -> str:
    """Individua il contesto intorno alla riga con 'S' isolata."""
    righe = testo.splitlines()
    indici_s = []

    for i, riga in enumerate(righe):
        if re.match(r'^\s*S\s*$', riga):
            indici_s.append(i)

    if not indici_s:
        return testo

    idx = indici_s[0]
    inizio = max(0, idx - righe_contorno)
    fine = min(len(righe), idx + righe_contorno + 1)
    blocco = righe[inizio:fine]

    return "\n".join(blocco)


RIGHE_DA_ESCLUDERE = [
    "isabella eusebio",
    "specialista in medicina",
    "ritenuta d'acconto",
    "p. iva",
]

def filtra_contesto(contesto: str) -> str:
    """Rimuove dal contesto le righe che contengono stringhe da escludere (case-insensitive)."""
    righe_filtrate = [
        riga for riga in contesto.splitlines()
        if not any(esclusa in riga.lower() for esclusa in RIGHE_DA_ESCLUDERE)
    ]
    return "\n".join(righe_filtrate)


def chiama_claude_api_per_nome(contesto: str) -> dict:
    """Chiama Claude API per estrarre nome e cognome dal contesto."""
    contesto = filtra_contesto(contesto)
    api_key = os.getenv('ANTHROPIC_API_KEY')
    if not api_key:
        return {'nome': '', 'cognome': '', 'nome_completo': '', 'errore': 'API key non configurata'}

    system_prompt = """Sei un assistente che estrae nome e cognome del paziente da testi di fatture/ricevute mediche italiane.

Il testo che ricevi è il contesto intorno a una riga chiave del documento.
Devi trovare il nome e cognome del PAZIENTE (non del medico).

Il medico è sempre "Dott.ssa Isabella EUSEBIO" o simile — ignoralo.
Il paziente è la persona che ha ricevuto la prestazione sanitaria.

Rispondi SOLO con un oggetto JSON nel formato:
{"nome": "...", "cognome": "...", "nome_completo": "..."}

Se non riesci a determinare con certezza nome o cognome, usa null per quel campo.
Non aggiungere nulla prima o dopo il JSON."""

    payload = {
        "model": "claude-sonnet-4-6",
        "max_tokens": 200,
        "system": system_prompt,
        "messages": [
            {
                "role": "user",
                "content": f"Ecco il contesto estratto dalla fattura:\n\n{contesto}\n\nEstrai nome e cognome del paziente."
            }
        ]
    }

    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=data,
        headers={
            "Content-Type": "application/json",
            "anthropic-version": "2023-06-01",
            "x-api-key": api_key,
        },
        method="POST"
    )

    try:
        with urllib.request.urlopen(req) as resp:
            risposta = json.loads(resp.read().decode("utf-8"))

        testo_risposta = risposta["content"][0]["text"].strip()
        testo_risposta = re.sub(r"```(?:json)?", "", testo_risposta).strip()

        try:
            result = json.loads(testo_risposta)
            nome = result.get('nome') or ''
            cognome = result.get('cognome') or ''
            nome_completo = result.get('nome_completo') or ''
            if not nome_completo and (nome or cognome):
                nome_completo = f"{nome} {cognome}".strip()
            return {
                'nome': nome,
                'cognome': cognome,
                'nome_completo': nome_completo
            }
        except json.JSONDecodeError:
            return {'nome': '', 'cognome': '', 'nome_completo': testo_risposta, 'errore_parsing': True}
    except urllib.error.HTTPError as e:
        # Mostra il corpo della risposta dell'API (es. modello inesistente, auth) invece del solo codice HTTP
        try:
            dettaglio = e.read().decode("utf-8")
        except Exception:
            dettaglio = ""
        return {'nome': '', 'cognome': '', 'nome_completo': '', 'errore': f"HTTP {e.code}: {dettaglio}"}
    except Exception as e:
        return {'nome': '', 'cognome': '', 'nome_completo': '', 'errore': str(e)}


def parse_private_invoice_text(text: str) -> dict:
    """Estrae i campi richiesti da un file .txt di fattura privata."""
    import re

    data = {
        'numero': '',
        'data': '',
        'importo': '',
        'nome_cliente': ''
    }

    lines = [line.strip() for line in text.splitlines() if line.strip()]
    joined = '\n'.join(lines)

    # Il numero è nella forma nn/yyyy: nn = 1 o 2 cifre, yyyy = un anno.
    def numero_valido(valore: str) -> bool:
        m = re.fullmatch(r'(\d{1,2})/(\d{4})', valore)
        return bool(m) and 1900 <= int(m.group(2)) <= 2100

    # Prima cerca tramite il marker corrente "RICEVUTA n."
    numero_match = re.search(r'RICEVUTA\s*n\.\s*(\S+)', joined, re.IGNORECASE)
    if numero_match and numero_valido(numero_match.group(1).strip()):
        data['numero'] = numero_match.group(1).strip()
    else:
        # Se il marker non dà un valore nella forma attesa, cerca tutte le
        # stringhe nel formato nn/yyyy nell'intero testo e prendi la prima valida.
        for candidato in re.findall(r'\b\d{1,2}/\d{4}\b', joined):
            if numero_valido(candidato):
                data['numero'] = candidato
                break

    # Abbrevia l'anno alle ultime due cifre: es. 16/2026 -> 16/26
    if data['numero']:
        nn, anno = data['numero'].split('/')
        data['numero'] = f'{nn}/{anno[2:]}'

    # Analizza tutte e sole le stringhe di 8 cifre: la prima che corrisponde
    # a una data valida nel formato ddmmyyyy è la data della fattura.
    for candidato in re.findall(r'\b\d{8}\b', joined):
        try:
            d = datetime.strptime(candidato, '%d%m%Y')
        except ValueError:
            continue
        data['data'] = d.strftime('%d/%m/%Y')
        break

    def converti_importo(valore: str):
        """Converte una stringa in numero intero, gestendo i separatori. None se non convertibile."""
        normalized = valore.replace(' ', '')
        if ',' in normalized and '.' in normalized:
            normalized = normalized.replace('.', '').replace(',', '.')
        elif ',' in normalized:
            normalized = normalized.replace(',', '.')
        elif normalized.count('.') > 1:
            normalized = normalized.replace('.', '')
        try:
            return int(float(normalized))
        except ValueError:
            return None

    def importo_valido(n) -> bool:
        """L'importo è valido se è compreso tra 50 e 150 ed è un multiplo di 10."""
        return n is not None and 50 <= n <= 180 and n % 10 == 0

    # Logica corrente: cerca l'importo nella riga dopo il marker "Riepilogo...".
    importo_trovato = None
    for idx, line in enumerate(lines):
        if 'Riepilogo degli onorari e delle fatture'.lower() in line.lower() and idx + 1 < len(lines):
            next_line = lines[idx + 1]
            import_match = re.search(r'([0-9]+[\d\.,]*)', next_line)
            if import_match:
                candidato = converti_importo(import_match.group(1).strip())
                if importo_valido(candidato):
                    importo_trovato = candidato
                break

    # Se il candidato del marker non è un importo valido, continua a scorrere il
    # testo finché trova una stringa convertibile in un numero valido.
    if importo_trovato is None:
        for candidato_str in re.findall(r'[0-9]+[\d\.,]*', joined):
            candidato = converti_importo(candidato_str)
            if importo_valido(candidato):
                importo_trovato = candidato
                break

    data['importo'] = str(importo_trovato) if importo_trovato is not None else ''

    def is_person_name(line: str) -> bool:
        if re.search(r'\d', line):
            return False
        if len(line) < 5 or len(line) > 80:
            return False
        lower = line.lower()
        blocked = [
            'corso', 'via', 'alba', 'cn', 'rata', 'importo', 'totale', 'ricevuta',
            'riepilogo', 'fattura', 'numero', 'cliente', 'spett', 'p.iva', 'cf',
            'telefono', 'fax', 'iban', 'cognome', 'nome', 'isabella eusebio', 'p. iva',
            'ritenuta d''acconto', 'specialista in medicina'
        ]
        if any(token in lower for token in blocked):
            return False
        words = [w for w in re.findall(r"[A-Za-zÀ-ÖØ-öø-ÿ'\-]+", line) if w]
        if len(words) < 2 or len(words) > 3:
            return False
        capitalized = sum(1 for w in words if w[0].isupper() or w.isupper())
        return capitalized >= 2

    # Prova prima con Claude API
    contesto = estrai_contesto_intorno_s(text)
    claude_result = chiama_claude_api_per_nome(contesto)
    if claude_result.get('nome_completo') and not claude_result.get('errore'):
        data['nome_cliente'] = claude_result.get('nome_completo', '')
    else:
        # Fallback alla ricerca locale se Claude non ha trovato nulla
        for idx, line in enumerate(lines):
            if line.strip().upper() in ('S', 'S.'):
                for offset in range(1, 6):
                    next_index = idx + offset
                    if next_index >= len(lines):
                        break
                    candidate = lines[next_index]
                    if is_person_name(candidate):
                        data['nome_cliente'] = candidate
                        break
                if data['nome_cliente']:
                    break

        if not data['nome_cliente']:
            for idx, line in enumerate(lines):
                if 'S' in line and len(line.strip()) == 1:
                    for offset in range(1, 6):
                        next_index = idx + offset
                        if next_index >= len(lines):
                            break
                        candidate = lines[next_index]
                        if is_person_name(candidate):
                            data['nome_cliente'] = candidate
                            break
                    if data['nome_cliente']:
                        break

        if not data['nome_cliente']:
            for line in lines:
                if is_person_name(line):
                    data['nome_cliente'] = line
                    break

    # Normalizza il nome: iniziale maiuscola di ogni parola, resto minuscolo.
    if data['nome_cliente']:
        data['nome_cliente'] = data['nome_cliente'].title()

    return data


@app.route('/api/process-private', methods=['POST'])
def process_private():
    """
    Elabora le fatture private dai file .txt e le scrive nel foglio "Fatture Privati" del file di output.
    """
    global execution_result
    execution_result = {}

    try:
        data = request.json
        private_folder = data.get('private_folder', '').strip()
        output_file = data.get('output_file', '').strip()

        if not private_folder:
            return jsonify({'success': False, 'error': 'Cartella di input fatture private non specificata'}), 400
        if not output_file:
            return jsonify({'success': False, 'error': 'File di output non specificato'}), 400
        if not os.path.isdir(private_folder):
            return jsonify({'success': False, 'error': f'Cartella di input fatture private non trovata: {private_folder}'}), 400
        if not os.path.isfile(output_file):
            return jsonify({'success': False, 'error': f'File di output non trovato: {output_file}'}), 400

        def run_process():
            global execution_result
            try:
                txt_files = sorted([p for p in Path(private_folder).glob('*.txt') if p.is_file()])
                if not txt_files:
                    execution_result = {'status': 'error', 'error': f'Nessun file .txt trovato in: {private_folder}', 'success': False}
                    return

                try:
                    output_wb = load_workbook(output_file)
                except Exception as e:
                    execution_result = {'status': 'error', 'error': f'Impossibile aprire file di output: {str(e)}', 'success': False}
                    return

                if 'Fatture Privati' not in output_wb.sheetnames:
                    output_wb.close()
                    execution_result = {
                        'status': 'error',
                        'error': f'Foglio "Fatture Privati" non trovato. Fogli disponibili: {output_wb.sheetnames}',
                        'success': False
                    }
                    return

                ws = output_wb['Fatture Privati']
                row = 2
                for txt_file in txt_files:
                    try:
                        text = txt_file.read_text(encoding='utf-8', errors='ignore')
                    except Exception as e:
                        output_wb.close()
                        execution_result = {'status': 'error', 'error': f'Errore leggendo {txt_file.name}: {str(e)}', 'success': False}
                        return

                    parsed = parse_private_invoice_text(text)
                    # L'importo va scritto come numero: così la cella conserva il
                    # formato presente sul template invece di essere trattato come testo.
                    importo_raw = parsed.get('importo', '')
                    importo_cell = int(importo_raw) if importo_raw else None
                    ws.cell(row=row, column=1).value = 'FATTURA'
                    ws.cell(row=row, column=2).value = parsed.get('numero', '')
                    ws.cell(row=row, column=3).value = parsed.get('data', '')
                    ws.cell(row=row, column=4).value = importo_cell
                    ws.cell(row=row, column=5).value = parsed.get('data', '')
                    ws.cell(row=row, column=6).value = parsed.get('nome_cliente', '')
                    row += 1

                try:
                    output_wb.save(output_file)
                    output_wb.close()
                except Exception as e:
                    output_wb.close()
                    execution_result = {'status': 'error', 'error': f'Errore salvando il file di output: {str(e)}', 'success': False}
                    return

                execution_result = {'status': 'completed', 'success': True, 'message': 'Fatture private elaborate con successo', 'output_file': output_file}
            except Exception as e:
                execution_result = {'status': 'error', 'error': str(e), 'success': False}

        thread = Thread(target=run_process, daemon=True)
        thread.start()

        return jsonify({'success': True, 'message': 'Elaborazione fatture private in corso...'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 400


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


# Partita IVA fissa dello studio, usata in tutti i record del foglio documentoSpesa
DOCUMENTO_SPESA_PIVA = 'IT03972130045'
CODICE_FISCALE_PROPRIETARIO = 'SBESLL71B46B111Q'


@app.route('/api/output-ts', methods=['POST'])
def output_ts():
    """
    Genera i record del foglio "documentoSpesa" a partire dalle fatture
    presenti nel foglio "Fatture Privati" del file di output.

    Per ogni fattura privata (a partire dalla riga 2) crea un record nel
    foglio "documentoSpesa" a partire dalla riga 2 con la struttura prevista
    dal tracciato Sistema TS.
    """
    global execution_result
    execution_result = {}

    try:
        data = request.json
        output_file = data.get('output_file', '').strip()

        if not output_file:
            return jsonify({'success': False, 'error': 'File di output non specificato'}), 400
        if not os.path.isfile(output_file):
            return jsonify({'success': False, 'error': f'File di output non trovato: {output_file}'}), 400

        def run_process():
            global execution_result
            try:
                try:
                    output_wb = load_workbook(output_file)
                except Exception as e:
                    execution_result = {'status': 'error', 'error': f'Impossibile aprire file di output: {str(e)}', 'success': False}
                    return

                if 'Fatture Privati' not in output_wb.sheetnames:
                    output_wb.close()
                    execution_result = {
                        'status': 'error',
                        'error': f'Foglio "Fatture Privati" non trovato. Fogli disponibili: {output_wb.sheetnames}',
                        'success': False
                    }
                    return

                if 'documentoSpesa' not in output_wb.sheetnames:
                    output_wb.close()
                    execution_result = {
                        'status': 'error',
                        'error': f'Foglio "documentoSpesa" non trovato. Fogli disponibili: {output_wb.sheetnames}',
                        'success': False
                    }
                    return

                src = output_wb['Fatture Privati']
                dst = output_wb['documentoSpesa']

                dst_row = 2
                count = 0
                # Le fatture private partono dalla riga 2; la colonna A contiene
                # sempre il marker 'FATTURA'. Si scorre finché si trovano dati.
                src_row = 2
                while True:
                    marker = src.cell(row=src_row, column=1).value          # A
                    numero = src.cell(row=src_row, column=2).value          # B -> NumDocumento
                    data_emissione = src.cell(row=src_row, column=3).value  # C -> dataEmissione
                    importo = src.cell(row=src_row, column=4).value        # D -> importo
                    data_pagamento = src.cell(row=src_row, column=5).value  # E -> dataPagamento
                    cf_cittadino = src.cell(row=src_row, column=7).value    # G -> cfCittadino

                    # Riga vuota -> fine dei dati
                    if not marker and not numero and not data_emissione:
                        break
                    dst.cell(row=dst_row, column=1).value = CODICE_FISCALE_PROPRIETARIO  # codice fiscale    
                    dst.cell(row=dst_row, column=2).value = DOCUMENTO_SPESA_PIVA  # pIVA
                    dst.cell(row=dst_row, column=3).value = data_emissione        # dataEmissione
                    dst.cell(row=dst_row, column=4).value = 1                     # dispositivo
                    dst.cell(row=dst_row, column=5).value = numero                # NumDocumento
                    dst.cell(row=dst_row, column=6).value = data_pagamento        # dataPagamento
                    dst.cell(row=dst_row, column=7).value = None                  # flagPagamentoAnticipato
                    dst.cell(row=dst_row, column=8).value = 'I'                   # flagOperazione
                    dst.cell(row=dst_row, column=9).value = cf_cittadino          # cfCittadino
                    dst.cell(row=dst_row, column=10).value = 'SI'                  # pagamentoTracciato
                    dst.cell(row=dst_row, column=11).value = 'F'                  # tipoDocumento
                    dst.cell(row=dst_row, column=12).value = 0                    # flagOpposizione
                    dst.cell(row=dst_row, column=13).value = 'SR'                 # tipoSpesa
                    dst.cell(row=dst_row, column=14).value = None                 # flagTipoSpesa
                    dst.cell(row=dst_row, column=15).value = None                 # AliquotaIva
                    dst.cell(row=dst_row, column=16).value = 'N4'                 # naturaIVA
                    dst.cell(row=dst_row, column=17).value = None                 # idRimborso
                    dst.cell(row=dst_row, column=18).value = importo              # R -> importo (da Fatture Privati col. D)

                    dst_row += 1
                    src_row += 1
                    count += 1

                try:
                    output_wb.save(output_file)
                    output_wb.close()
                except Exception as e:
                    output_wb.close()
                    execution_result = {'status': 'error', 'error': f'Errore salvando il file di output: {str(e)}', 'success': False}
                    return

                execution_result = {
                    'status': 'completed',
                    'success': True,
                    'message': f'Output TS generato: {count} record creati nel foglio "documentoSpesa"',
                    'output_file': output_file
                }
            except Exception as e:
                execution_result = {'status': 'error', 'error': str(e), 'success': False}

        thread = Thread(target=run_process, daemon=True)
        thread.start()

        return jsonify({'success': True, 'message': 'Generazione Output TS in corso...'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 400


@app.route('/api/invia-ts', methods=['POST'])
def invia_ts_route():
    """
    Invio completo al Sistema Tessera Sanitaria:
    1. Genera l'XML <precompilata> + .zip dal foglio "documentoSpesa" del file di output
    2. Trasmette il file (SOAP MTOM) all'ambiente configurato (TS_AMBIENTE)
    3. Recupera l'esito e scarica la ricevuta PDF

    Le credenziali di produzione sono lette dal file .env:
    TS_AMBIENTE (test|produzione), TS_USERNAME, TS_PASSWORD, TS_PINCODE, TS_CF_PROPRIETARIO
    """
    global execution_result
    execution_result = {}

    try:
        data = request.json
        output_file = data.get('output_file', '').strip()
        richiesta_ambiente = (data.get('ambiente') or '').strip()

        if not output_file:
            return jsonify({'success': False, 'error': 'File di output non specificato'}), 400
        if not os.path.isfile(output_file):
            return jsonify({'success': False, 'error': f'File di output non trovato: {output_file}'}), 400

        def run_process():
            global execution_result
            try:
                # L'ambiente può essere scelto dal frontend; in mancanza usa quello del .env
                ambiente = (richiesta_ambiente or os.getenv('TS_AMBIENTE', 'test')).strip().lower()
                prod = ambiente.startswith('prod')

                # Step 1: genera XML + ZIP dal foglio documentoSpesa
                try:
                    gen = genera_invio_ts.genera_file(output_file)
                except Exception as e:
                    execution_result = {'status': 'error', 'success': False,
                                        'error': f'Errore nella generazione del file: {str(e)}',
                                        'output_file': output_file}
                    return
                zip_path = gen['zip_path']

                # Credenziali e endpoint in base all'ambiente
                if prod:
                    user = os.getenv('TS_USERNAME', '').strip()
                    password = os.getenv('TS_PASSWORD', '').strip()
                    pincode = os.getenv('TS_PINCODE', '').strip()
                    cf_prop = os.getenv('TS_CF_PROPRIETARIO', '').strip()
                    mancanti = [n for n, v in (('TS_USERNAME', user), ('TS_PASSWORD', password),
                                               ('TS_PINCODE', pincode), ('TS_CF_PROPRIETARIO', cf_prop)) if not v]
                    if mancanti:
                        execution_result = {'status': 'error', 'success': False,
                                            'error': f'Credenziali di produzione mancanti nel file .env: {", ".join(mancanti)}',
                                            'output_file': output_file}
                        return
                    endpoint = invia_ts.ENDPOINT_PROD
                    verify = True
                else:
                    user, password = invia_ts.TEST_USER, invia_ts.TEST_PASSWORD
                    pincode, cf_prop = invia_ts.TEST_PINCODE, invia_ts.TEST_CF_PROPRIETARIO
                    endpoint = invia_ts.ENDPOINT_TEST
                    verify = False

                # Step 2: invio del file
                try:
                    campi = invia_ts.invia(endpoint, user, password, pincode, cf_prop, zip_path, verify=verify)
                except Exception as e:
                    execution_result = {'status': 'error', 'success': False,
                                        'error': f'Errore durante l\'invio: {str(e)}',
                                        'output_file': output_file}
                    return

                codice = campi.get('codiceEsito')
                protocollo = campi.get('protocollo')
                descrizione_invio = campi.get('descrizioneEsito', '')

                if codice != '000' or not protocollo:
                    execution_result = {'status': 'error', 'success': False,
                                        'error': f'Invio non accolto (codice {codice}): {descrizione_invio}',
                                        'ambiente': ambiente, 'documenti': gen['n'],
                                        'output_file': output_file}
                    return

                # Step 3: attende l'elaborazione e recupera l'esito
                esito = {}
                for _ in range(6):
                    time.sleep(5)
                    try:
                        esito = ricevute_ts.recupera_esito(protocollo, user, password, pincode, prod=prod, verify=verify)
                    except Exception:
                        esito = {}
                    if esito.get('stato'):  # esito di merito disponibile
                        break

                # Scarica la ricevuta PDF (disponibile anche in caso di scarto)
                ricevuta_file = None
                try:
                    out_pdf = Path(output_file).parent / f'ricevuta_{protocollo}.pdf'
                    got = ricevute_ts.scarica_ricevuta_pdf(protocollo, out_pdf, user, password, pincode, prod=prod, verify=verify)
                    if got:
                        ricevuta_file = str(got)
                except Exception:
                    ricevuta_file = None

                stato = esito.get('stato')
                accolto = stato == '2'
                if accolto:
                    msg = (f'Invio ACCOLTO ({ambiente}). Protocollo {protocollo}. '
                           f'Documenti accolti: {esito.get("nAccolti", "?")} su {esito.get("nInviati", "?")}, '
                           f'errori: {esito.get("nErrori", "0")}.')
                else:
                    descr = esito.get('descrizione') or descrizione_invio or 'in elaborazione'
                    msg = (f'Invio trasmesso ({ambiente}), protocollo {protocollo}. '
                           f'Esito: {descr}. Apri la ricevuta per i dettagli.')

                execution_result = {
                    'status': 'completed',
                    'success': bool(accolto),
                    'message': msg,
                    'protocollo': protocollo,
                    'ambiente': ambiente,
                    'documenti': gen['n'],
                    'esito': esito,
                    'ricevuta_file': ricevuta_file,
                    'output_file': output_file,
                }
            except Exception as e:
                execution_result = {'status': 'error', 'success': False, 'error': str(e),
                                    'output_file': output_file}

        thread = Thread(target=run_process, daemon=True)
        thread.start()

        return jsonify({'success': True, 'message': 'Invio al Sistema TS in corso...'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 400


if __name__ == '__main__':
    print("Applicazione avviata su http://localhost:5000")
    print("Premi Ctrl+C per fermare l'applicazione")
    app.run(debug=False, host='localhost', port=5000)

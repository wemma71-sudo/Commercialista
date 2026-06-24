
    #!/usr/bin/env python3
"""
    Estrazione automatica di nome e cognome, data, numero e importo da PDF di ricevute/fatture.

    Requisiti:
      - Python 3.9+
      - PyPDF2 (`pip install PyPDF2`)

    Utilizzo:
      python extract_invoice_fields.py documento.pdf
      python extract_invoice_fields.py cartella_con_pdf --csv output.csv
      python extract_invoice_fields.py *.pdf --json output.json

    Note:
      - Lo script funziona bene con PDF testuali o PDF già OCRizzati.
      - Per PDF immagine puri serve una fase OCR separata.
    """

from __future__ import annotations

import argparse
import csv
import glob
import json
import re
import sys
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Iterable, List, Optional

try:
        from PyPDF2 import PdfReader
except Exception:
        print("Errore: libreria PyPDF2 non disponibile. Installa con: pip install PyPDF2", file=sys.stderr)
        raise

@dataclass
class ExtractedInvoice:
        file_name: str
        nome_cognome: Optional[str] = None
        data: Optional[str] = None
        numero: Optional[str] = None
        importo: Optional[str] = None
        raw_text_preview: Optional[str] = None


def extract_text_from_pdf(pdf_path: Path) -> str:
        reader = PdfReader(str(pdf_path))
        chunks = []
        for page in reader.pages:
            text = page.extract_text() or ""
            chunks.append(text)
        return "".join(chunks)


def normalize_spaces(text: str) -> str:
        text = text.replace(" ", " ")
        text = re.sub(r"[ 	]+", " ", text)
        text = re.sub(r"{2,}", "", text)
        return text.strip()


def normalize_amount(value: str) -> str:
        value = value.strip().replace(" ", "")
        # Normalizza formati tipo 120,00 / 120.00 / 1.234,56
        if "," in value and "." in value:
            # assume formato europeo 1.234,56
            value = value.replace(".", "")
        if "." in value and "," not in value:
            # assume formato 120.00 -> 120,00
            value = value.replace(".", ",")
        return value


def normalize_date(raw: str) -> str:
        raw = raw.strip()
        digits = re.sub(r"\D", "", raw)
        if len(digits) == 8:
            return f"{digits[0:2]}/{digits[2:4]}/{digits[4:8]}"
        # già in formato con separatori
        parts = re.split(r"[\./\-]", raw)
        if len(parts) == 3:
            d, m, y = parts
            d = d.zfill(2)
            m = m.zfill(2)
            if len(y) == 2:
                y = ("20" + y) if int(y) < 50 else ("19" + y)
            return f"{d}/{m}/{y}"
        return raw


def find_numero(text: str) -> Optional[str]:
        patterns = [
            r"RICEVUTA\s*n\.?\s*([0-9]+\s*/\s*[0-9]{4})",
            r"n\.?\s*([0-9]+\s*/\s*[0-9]{4})",
            r"numero\s*([0-9]+\s*/\s*[0-9]{4})",
        ]
        for pat in patterns:
            m = re.search(pat, text, flags=re.IGNORECASE | re.DOTALL)
            if m:
                return re.sub(r"\s+", "", m.group(1))
        return None


def find_data(text: str) -> Optional[str]:
        patterns = [
            r"Data\s*([0-9]{8})",
            r"Data\s*([0-9]{1,2}[\./\-][0-9]{1,2}[\./\-][0-9]{2,4})",
            r"data\s*fattura\s*[:\-]?\s*([0-9]{1,2}[\./\-][0-9]{1,2}[\./\-][0-9]{2,4})",
        ]
        for pat in patterns:
            m = re.search(pat, text, flags=re.IGNORECASE)
            if m:
                return normalize_date(m.group(1))
        return None


def find_nome_cognome(text: str) -> Optional[str]:
        # Caso tipico: riga con NOME COGNOME seguito dalla data di nascita
        patterns = [
            r"([A-ZÀ-ÖØ-Ý'`\.\-]+\s+[A-ZÀ-ÖØ-Ý'`\.\-]+(?:\s+[A-ZÀ-ÖØ-Ý'`\.\-]+)?)\s+\d{1,2}[\./\-]\d{1,2}[\./\-]\d{2,4}",
            r"Cliente\s*[:\-]?\s*([A-ZÀ-ÖØ-Ý'`\.\-]+\s+[A-ZÀ-ÖØ-Ý'`\.\-]+(?:\s+[A-ZÀ-ÖØ-Ý'`\.\-]+)?)",
            r"Paziente\s*[:\-]?\s*([A-ZÀ-ÖØ-Ý'`\.\-]+\s+[A-ZÀ-ÖØ-Ý'`\.\-]+(?:\s+[A-ZÀ-ÖØ-Ý'`\.\-]+)?)",
        ]
        # Evita di prendere il medico/fornitore nelle intestazioni: cerca dopo Data/n.
        search_zone = text
        anchor = re.search(r"RICEVUTA|FATTURA|Data", text, flags=re.IGNORECASE)
        if anchor:
            search_zone = text[anchor.start():]
        for pat in patterns:
            m = re.search(pat, search_zone, flags=re.IGNORECASE)
            if m:
                candidate = re.sub(r"\s+", " ", m.group(1)).strip()
                banned = {"MEDICO CHIRURGO", "IMPORTO", "RIEPILOGO RIPORTI"}
                if candidate.upper() not in banned:
                    return candidate.title()
        return None


def find_importo(text: str) -> Optional[str]:
        patterns = [
            r"TOTALE\s*([0-9\., ]{2,})",
            r"da\s+Pagare\s*€?\s*([0-9\., ]{2,})",
            r"IMPORTI\s+RICEVUTA.*?([0-9\., ]{2,})",
            r"Visita\s+fisiatrica\s*([0-9\., ]{2,})",
        ]
        for pat in patterns:
            m = re.search(pat, text, flags=re.IGNORECASE | re.DOTALL)
            if m:
                return normalize_amount(m.group(1))
        # fallback: prende l'ultimo importo nel documento, spesso coincide col totale
        all_amounts = re.findall(r"\d{1,3}(?:[\. ]\d{3})*[\.,]\d{2}", text)
        if all_amounts:
            return normalize_amount(all_amounts[-1])
        return None


def extract_fields_from_text(text: str, file_name: str) -> ExtractedInvoice:
        txt = normalize_spaces(text)
        result = ExtractedInvoice(
            file_name=file_name,
            nome_cognome=find_nome_cognome(txt),
            data=find_data(txt),
            numero=find_numero(txt),
            importo=find_importo(txt),
            raw_text_preview=txt[:250],
        )
        return result


def expand_inputs(inputs: List[str]) -> List[Path]:
        paths: List[Path] = []
        for item in inputs:
            p = Path(item)
            if p.is_dir():
                paths.extend(sorted(p.glob("*.pdf")))
            elif any(ch in item for ch in "*?[]"):
                paths.extend(sorted(Path(x) for x in glob.glob(item)))
            else:
                paths.append(p)
        # de-dup e solo pdf esistenti
        unique = []
        seen = set()
        for p in paths:
            if p.suffix.lower() == ".pdf" and p.exists() and p.resolve() not in seen:
                unique.append(p)
                seen.add(p.resolve())
        return unique


def write_csv(results: List[ExtractedInvoice], out_path: Path) -> None:
        with out_path.open("w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(
                f,
                fieldnames=["file_name", "nome_cognome", "data", "numero", "importo", "raw_text_preview"],
            )
            writer.writeheader()
            for r in results:
                writer.writerow(asdict(r))


def write_json(results: List[ExtractedInvoice], out_path: Path) -> None:
        with out_path.open("w", encoding="utf-8") as f:
            json.dump([asdict(r) for r in results], f, ensure_ascii=False, indent=2)


def main() -> int:
        parser = argparse.ArgumentParser(description="Estrae nome/cognome, data, numero e importo da PDF di fatture/ricevute")
        parser.add_argument("inputs", nargs="+", help="Uno o più PDF, wildcard (*.pdf), oppure una cartella")
        parser.add_argument("--csv", dest="csv_path", help="Salva i risultati anche in CSV")
        parser.add_argument("--json", dest="json_path", help="Salva i risultati anche in JSON")
        args = parser.parse_args()

        pdf_files = expand_inputs(args.inputs)
        if not pdf_files:
            print("Nessun PDF trovato.", file=sys.stderr)
            return 1

        results: List[ExtractedInvoice] = []
        for pdf in pdf_files:
            try:
                text = extract_text_from_pdf(pdf)
                results.append(extract_fields_from_text(text, pdf.name))
            except Exception as exc:
                results.append(ExtractedInvoice(file_name=pdf.name, raw_text_preview=f"ERRORE: {exc}"))

        # Output a video
        for r in results:
            print("-" * 80)
            print(f"File        : {r.file_name}")
            print(f"Nome        : {r.nome_cognome or ''}")
            print(f"Data        : {r.data or ''}")
            print(f"Numero      : {r.numero or ''}")
            print(f"Importo     : {r.importo or ''}")

        if args.csv_path:
            write_csv(results, Path(args.csv_path))
            print(f"CSV salvato in: {args.csv_path}")

        if args.json_path:
            write_json(results, Path(args.json_path))
            print(f"JSON salvato in: {args.json_path}")

        return 0


if __name__ == "__main__":
    raise SystemExit(main())
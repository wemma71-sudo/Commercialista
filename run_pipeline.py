#!/usr/bin/env python3
"""
Pipeline che connette latest_zip.py e process_zip_excel.py
1. Trova lo zip più recente
2. Lo elabora con process_zip_excel
"""

import sys
import subprocess
from pathlib import Path


def main(argv: list[str]) -> int:
    # Cartella in input o cartella corrente
    folder = argv[1] if len(argv) > 1 else r"C:\Users\francesco.emma\Downloads"
    
    # Step 1: Trova l'ultimo zip
    print(f"Step 1: Trovando lo zip più recente in {folder}...", file=sys.stderr)
    result = subprocess.run(
        [sys.executable, "latest_zip.py", folder],
        capture_output=True,
        text=True,
        cwd=Path(__file__).parent
    )
    
    if result.returncode != 0:
        print(result.stderr, file=sys.stderr)
        return result.returncode
    
    zip_path = result.stdout.strip()
    print(f"✓ Trovato: {zip_path}", file=sys.stderr)
    
    # Step 2: Elabora lo zip
    print(f"Step 2: Elaborando {Path(zip_path).name}...", file=sys.stderr)
    result = subprocess.run(
        [sys.executable, "process_zip_excel.py", zip_path, 'Y', 'E'],
        cwd=Path(__file__).parent
    )
    
    if result.returncode == 0:
        print("✓ Pipeline completato con successo", file=sys.stderr)
    else:
        print(f"✗ Errore durante l'elaborazione (exit code: {result.returncode})", file=sys.stderr)
    
    return result.returncode


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))

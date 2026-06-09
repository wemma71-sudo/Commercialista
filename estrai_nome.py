#!/usr/bin/env python3
"""
Estrae nome e cognome del paziente da file di testo di fatture mediche.
Strategia: individua il contesto intorno alla riga con "S" isolata,
poi usa Claude API per estrarre nome e cognome in modo robusto.
"""

import os
import re
import json
import glob
import argparse
import urllib.request


# ── Configurazione ────────────────────────────────────────────────────────────

ANTHROPIC_API_URL = "https://api.anthropic.com/v1/messages"
MODEL = "claude-sonnet-4-20250514"

SYSTEM_PROMPT = """Sei un assistente che estrae nome e cognome del paziente da testi di fatture/ricevute mediche italiane.

Il testo che ricevi è il contesto intorno a una riga chiave del documento.
Devi trovare il nome e cognome del PAZIENTE (non del medico).

Il medico è sempre "Dott.ssa Isabella EUSEBIO" o simile — ignoralo.
Il paziente è la persona che ha ricevuto la prestazione sanitaria.

Rispondi SOLO con un oggetto JSON nel formato:
{"nome": "...", "cognome": "...", "nome_completo": "..."}

Se non riesci a determinare con certezza nome o cognome, usa null per quel campo.
Non aggiungere nulla prima o dopo il JSON."""


def estrai_contesto_intorno_s(testo: str, righe_contorno: int = 4) -> str:
    """
    Individua le righe vicine alla riga che contiene 'S' isolata
    (il separatore caratteristico dei moduli Buffetti).
    Restituisce un blocco di testo con quelle righe.
    """
    righe = testo.splitlines()
    indici_s = []

    for i, riga in enumerate(righe):
        # "S" isolata: riga che contiene solo S, o S con spazi, o S seguita da poco
        if re.match(r'^\s*S\s*$', riga):
            indici_s.append(i)

    if not indici_s:
        # Fallback: prendi l'intero testo (file corto)
        return testo

    # Prendi il contesto attorno al primo "S" trovato
    idx = indici_s[0]
    inizio = max(0, idx - righe_contorno)
    fine = min(len(righe), idx + righe_contorno + 1)
    blocco = righe[inizio:fine]

    return "\n".join(blocco)


def chiama_claude_api(contesto: str) -> dict:
    """
    Chiama l'API di Claude per estrarre nome e cognome dal contesto.
    """
    payload = {
        "model": MODEL,
        "max_tokens": 200,
        "system": SYSTEM_PROMPT,
        "messages": [
            {
                "role": "user",
                "content": f"Ecco il contesto estratto dalla fattura:\n\n{contesto}\n\nEstrai nome e cognome del paziente."
            }
        ]
    }

    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        raise ValueError("Variabile d'ambiente ANTHROPIC_API_KEY non impostata.")

    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        ANTHROPIC_API_URL,
        data=data,
        headers={
            "Content-Type": "application/json",
            "anthropic-version": "2023-06-01",
            "x-api-key": api_key,
        },
        method="POST"
    )

    with urllib.request.urlopen(req) as resp:
        risposta = json.loads(resp.read().decode("utf-8"))

    testo_risposta = risposta["content"][0]["text"].strip()

    # Rimuovi eventuali backtick markdown
    testo_risposta = re.sub(r"```(?:json)?", "", testo_risposta).strip()

    try:
        return json.loads(testo_risposta)
    except json.JSONDecodeError:
        return {"nome": None, "cognome": None, "nome_completo": testo_risposta, "errore_parsing": True}


def processa_file(percorso: str) -> dict:
    """
    Legge un file .txt, estrae il contesto e chiama l'API.
    """
    with open(percorso, "r", encoding="utf-8", errors="replace") as f:
        testo = f.read()

    contesto = estrai_contesto_intorno_s(testo)
    risultato = chiama_claude_api(contesto)
    risultato["file"] = os.path.basename(percorso)
    risultato["contesto_usato"] = contesto
    return risultato


def main():
    parser = argparse.ArgumentParser(
        description="Estrae nome e cognome del paziente da fatture mediche in formato .txt"
    )
    parser.add_argument(
        "files",
        nargs="+",
        help="Uno o più file .txt (accetta glob, es: fatture/*.txt)"
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Output in formato JSON invece di tabella testuale"
    )
    args = parser.parse_args()

    # Espandi glob
    file_list = []
    for pattern in args.files:
        espansi = glob.glob(pattern)
        file_list.extend(espansi if espansi else [pattern])

    if not file_list:
        print("Nessun file trovato.")
        return

    risultati = []
    for percorso in sorted(file_list):
        if not os.path.exists(percorso):
            print(f"  [!] File non trovato: {percorso}")
            continue

        print(f"  Elaborando: {os.path.basename(percorso)} ...", end=" ", flush=True)
        try:
            res = processa_file(percorso)
            risultati.append(res)
            print(f"→ {res.get('nome_completo', 'N/D')}")
        except Exception as e:
            print(f"ERRORE: {e}")
            risultati.append({"file": os.path.basename(percorso), "errore": str(e)})

    print()

    if args.json:
        # Output JSON pulito (senza contesto_usato)
        output = []
        for r in risultati:
            output.append({
                "file": r.get("file"),
                "nome": r.get("nome"),
                "cognome": r.get("cognome"),
                "nome_completo": r.get("nome_completo"),
            })
        print(json.dumps(output, ensure_ascii=False, indent=2))
    else:
        # Tabella testuale
        print(f"{'FILE':<45} {'NOME':<15} {'COGNOME':<20} {'NOME COMPLETO'}")
        print("-" * 100)
        for r in risultati:
            if "errore" in r:
                print(f"{r['file']:<45} ERRORE: {r['errore']}")
            else:
                print(
                    f"{r.get('file',''):<45} "
                    f"{r.get('nome') or '':<15} "
                    f"{r.get('cognome') or '':<20} "
                    f"{r.get('nome_completo') or ''}"
                )


if __name__ == "__main__":
    main()
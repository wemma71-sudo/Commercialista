#!/usr/bin/env python3
"""
Invio al Sistema Tessera Sanitaria (730 spese sanitarie, modalità asincrona).

Esegue la chiamata SOAP MTOM all'operazione `inviaFileMtom`, allegando lo .zip
contenente il file XML <precompilata> e restituendo l'esito (codiceEsito,
descrizioneEsito, protocollo).

Per impostazione predefinita usa l'AMBIENTE DI TEST con l'utente fittizio del
Development Kit (PROVAX00X00X000Y) e l'allegato di esempio ts730/provaMedico.zip.
Per inviare in PRODUZIONE servono le credenziali reali (vedi sezione --prod).

Uso:
    # Test (utente fittizio, allegato di esempio):
    python invia_ts.py

    # Test con uno zip specifico:
    python invia_ts.py --zip "C:\\...\\20260609_invio730.zip"

    # Produzione (richiede credenziali reali):
    python invia_ts.py --prod --user <CF> --password <PWD> --pincode <PIN> \\
                       --cf-proprietario <CF> --zip "C:\\...\\20260609_invio730.zip"
"""

import argparse
import re
import sys
import uuid
from base64 import b64encode
from datetime import datetime
from pathlib import Path

import requests
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.x509 import load_der_x509_certificate, load_pem_x509_certificate

# Usa lo store dei certificati di Windows (include le CA della PA) per la verifica TLS,
# evitando l'errore "unable to get local issuer certificate" di certifi.
try:
    import truststore
    truststore.inject_into_ssl()
except ImportError:
    pass

BASE_DIR = Path(__file__).parent
CERT_PATH = BASE_DIR / 'ts730' / 'SanitelCF.cer'

ENDPOINT_TEST = 'https://invioSS730pTest.sanita.finanze.it/InvioTelematicoSS730pMtomWeb/InvioTelematicoSS730pMtomPort'
ENDPOINT_PROD = 'https://invioSS730p.sanita.finanze.it/InvioTelematicoSS730pMtomWeb/InvioTelematicoSS730pMtomPort'

# Credenziali dell'utente di test fornite dal Development Kit (SoggettoMedico)
TEST_USER = 'PROVAX00X00X000Y'
TEST_PASSWORD = 'Salve123'
TEST_PINCODE = '1234567890'
TEST_CF_PROPRIETARIO = 'PROVAX00X00X000Y'  # nel SOAP datiProprietario va in chiaro

SOAP_NS = 'http://ejb.invioTelematicoSS730p.sanita.finanze.it/'


def cifra(text: str) -> str:
    """Cifra una stringa col certificato pubblico SanitelCF (RSA PKCS#1 v1.5) + base64."""
    raw = CERT_PATH.read_bytes()
    try:
        cert = load_der_x509_certificate(raw)
    except Exception:
        cert = load_pem_x509_certificate(raw)
    cifrato = cert.public_key().encrypt(text.encode('utf-8'), padding.PKCS1v15())
    return b64encode(cifrato).decode('ascii')


def costruisci_envelope(nome_file: str, pincode_cifrato: str, cf_proprietario: str, attach_cid: str) -> str:
    """Costruisce l'envelope SOAP con il riferimento MTOM all'allegato."""
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" '
        f'xmlns:ejb="{SOAP_NS}">'
        '<soapenv:Header/>'
        '<soapenv:Body>'
        '<ejb:inviaFileMtom>'
        f'<nomeFileAllegato>{nome_file}</nomeFileAllegato>'
        f'<pincodeInvianteCifrato>{pincode_cifrato}</pincodeInvianteCifrato>'
        '<datiProprietario>'
        f'<cfProprietario>{cf_proprietario}</cfProprietario>'
        '</datiProprietario>'
        '<documento>'
        f'<xop:Include xmlns:xop="http://www.w3.org/2004/08/xop/include" href="cid:{attach_cid}"/>'
        '</documento>'
        '</ejb:inviaFileMtom>'
        '</soapenv:Body>'
        '</soapenv:Envelope>'
    )


def invia(endpoint, user, password, pincode, cf_proprietario, zip_path: Path, verify=True):
    nome_file = zip_path.name
    if not (6 <= len(nome_file) <= 60):
        raise SystemExit(f'nomeFileAllegato deve essere 6-60 caratteri: {nome_file!r}')

    pincode_cifrato = cifra(pincode)
    root_cid = f'root-{uuid.uuid4().hex}@sts'
    attach_cid = f'doc-{uuid.uuid4().hex}@sts'
    boundary = f'----=_Part_{uuid.uuid4().hex}'

    envelope = costruisci_envelope(nome_file, pincode_cifrato, cf_proprietario, attach_cid)
    zip_bytes = zip_path.read_bytes()

    crlf = b'\r\n'
    parts = []
    # Parte 1: root XOP (SOAP envelope)
    parts.append(b'--' + boundary.encode())
    parts.append(b'Content-Type: application/xop+xml; charset=UTF-8; type="text/xml"')
    parts.append(b'Content-Transfer-Encoding: 8bit')
    parts.append(f'Content-ID: <{root_cid}>'.encode())
    parts.append(b'')
    parts.append(envelope.encode('utf-8'))
    # Parte 2: allegato binario (lo zip)
    parts.append(b'--' + boundary.encode())
    parts.append(b'Content-Type: application/zip')
    parts.append(b'Content-Transfer-Encoding: binary')
    parts.append(f'Content-ID: <{attach_cid}>'.encode())
    parts.append(b'')
    parts.append(zip_bytes)
    parts.append(b'--' + boundary.encode() + b'--')
    body = crlf.join(parts)

    content_type = (
        f'multipart/related; type="application/xop+xml"; '
        f'start="<{root_cid}>"; start-info="text/xml"; boundary="{boundary}"'
    )
    headers = {
        'Content-Type': content_type,
        'SOAPAction': '""',
        'MIME-Version': '1.0',
    }

    print(f'Invio a: {endpoint}')
    print(f'Allegato: {zip_path} ({len(zip_bytes)} byte) come "{nome_file}"')
    if not verify:
        print('ATTENZIONE: verifica del certificato TLS disabilitata (CA di test non pubblica).')
        import urllib3
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    resp = requests.post(endpoint, data=body, headers=headers,
                         auth=(user, password), timeout=120, verify=verify)
    print(f'HTTP {resp.status_code}')
    testo = resp.text

    # Estrae i campi della ricevuta dalla response (SOAP o MTOM)
    campi = {}
    for k in ('codiceEsito', 'descrizioneEsito', 'protocollo', 'dataAccoglienza',
              'nomeFileAllegato', 'dimensioneFileAllegato', 'idErrore'):
        m = re.search(rf'<{k}>(.*?)</{k}>', testo, re.S)
        if m:
            campi[k] = m.group(1).strip()

    if campi:
        print('\n--- ESITO ---')
        for k, v in campi.items():
            print(f'  {k}: {v}')
        # codiceEsito 000 = file accolto e in carico (asincrono); l'esito di merito
        # arriva con la ricevuta. Codici 0xx/1xx = file NON accolto (vedi PDF par. 4.7).
        if campi.get('codiceEsito') == '000':
            print('\nFile ACCOLTO (in attesa di elaborazione). Conserva il protocollo: '
                  'scarica poi la ricevuta PDF e l\'esito (servizi WS_Ricevute) '
                  'per i controlli di merito sui dati.')
        else:
            print('\nFile NON accolto o con errore: vedi codiceEsito/descrizioneEsito '
                  'e la tabella codici nel PDF (par. 4.7).')
    else:
        print('\nRisposta non interpretata. Corpo grezzo (primi 2000 caratteri):')
        print(testo[:2000])
    return campi


def main():
    p = argparse.ArgumentParser(description='Invio file spese sanitarie 730 al Sistema TS (MTOM).')
    p.add_argument('--prod', action='store_true', help='Usa l\'ambiente di produzione (invio reale).')
    p.add_argument('--zip', dest='zip_path', help='Percorso dello .zip da inviare.')
    p.add_argument('--user', help='User (CF) — solo produzione.')
    p.add_argument('--password', help='Password — solo produzione.')
    p.add_argument('--pincode', help='Pincode in chiaro — solo produzione.')
    p.add_argument('--cf-proprietario', dest='cf_proprietario', help='CF proprietario — solo produzione.')
    p.add_argument('--insecure', action='store_true', help='Disabilita la verifica del certificato TLS.')
    args = p.parse_args()

    if not CERT_PATH.is_file():
        raise SystemExit(f'Certificato non trovato: {CERT_PATH}')

    if args.prod:
        mancanti = [n for n in ('user', 'password', 'pincode', 'cf_proprietario', 'zip_path')
                    if not getattr(args, n)]
        if mancanti:
            raise SystemExit(f'Produzione: parametri obbligatori mancanti: {mancanti}')
        endpoint = ENDPOINT_PROD
        user, password, pincode, cf_prop = args.user, args.password, args.pincode, args.cf_proprietario
        zip_path = Path(args.zip_path)
        verify = not args.insecure  # in produzione si verifica il TLS, salvo --insecure
        print('*** AMBIENTE DI PRODUZIONE — INVIO REALE ***')
    else:
        endpoint = ENDPOINT_TEST
        user, password, pincode, cf_prop = TEST_USER, TEST_PASSWORD, TEST_PINCODE, TEST_CF_PROPRIETARIO
        zip_path = Path(args.zip_path) if args.zip_path else (BASE_DIR / 'ts730' / 'provaMedico.zip')
        verify = False  # la CA del server di test non è pubblica
        print('*** AMBIENTE DI TEST (utente fittizio) ***')

    if not zip_path.is_file():
        raise SystemExit(f'File zip non trovato: {zip_path}')

    invia(endpoint, user, password, pincode, cf_prop, zip_path, verify=verify)


if __name__ == '__main__':
    main()

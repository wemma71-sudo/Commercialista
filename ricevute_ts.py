#!/usr/bin/env python3
"""
Recupero esito / ricevuta di un invio al Sistema Tessera Sanitaria (730 spese
sanitarie). Dato il numero di protocollo restituito da invia_ts.py, interroga:

- EsitoInvii      -> stato dell'invio (n. documenti inviati/accolti/warning/errori)
- RicevutaPdf     -> ricevuta in PDF (salvata su file)
- DettaglioErrori -> dettaglio errori (CSV) se presenti

Servizi SOAP semplici (non MTOM), autenticazione HTTP Basic, pinCode cifrato.

Uso (ambiente di test, utente fittizio):
    python ricevute_ts.py --protocollo 26060915384938763
    python ricevute_ts.py --protocollo 26060915384938763 --servizio esito
    python ricevute_ts.py --protocollo ... --servizio ricevuta --out ricevuta.pdf

Produzione:
    python ricevute_ts.py --prod --user <CF> --password <PWD> --pincode <PIN> \\
                          --protocollo <PROT>
"""

import argparse
import re
from base64 import b64decode, b64encode
from pathlib import Path

import requests
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.x509 import load_der_x509_certificate, load_pem_x509_certificate

try:
    import truststore
    truststore.inject_into_ssl()
except ImportError:
    pass

BASE_DIR = Path(__file__).parent
CERT_PATH = BASE_DIR / 'ts730' / 'SanitelCF.cer'

TEST_USER = 'PROVAX00X00X000Y'
TEST_PASSWORD = 'Salve123'
TEST_PINCODE = '1234567890'

# (namespace_prefix, namespace, operazione, endpoint_test, endpoint_prod)
SERVIZI = {
    'esito': (
        'esit', 'http://esitoinvio.p730.sanita.sogei.it/', 'EsitoInvii',
        'https://invioSS730pTest.sanita.finanze.it/EsitoStatoInviiWEB/EsitoInvioDatiSpesa730Service',
        'https://invioSS730p.sanita.finanze.it/EsitoStatoInviiWEB/EsitoInvioDatiSpesa730Service',
    ),
    'ricevuta': (
        'ric', 'http://ricevutapdf.p730.sanita.sogei.it/', 'RicevutaPdf',
        'https://invioSS730pTest.sanita.finanze.it/Ricevute730ServiceWeb/ricevutePdf',
        'https://invioSS730p.sanita.finanze.it/Ricevute730ServiceWeb/ricevutePdf',
    ),
    'errori': (
        'det', 'http://dettaglioerrori.p730.sanita.sogei.it/', 'DettaglioErrori',
        'https://invioSS730pTest.sanita.finanze.it/EsitoStatoInviiWEB/DettaglioErrori730Service',
        'https://invioSS730p.sanita.finanze.it/EsitoStatoInviiWEB/DettaglioErrori730Service',
    ),
}


def cifra(text: str) -> str:
    raw = CERT_PATH.read_bytes()
    try:
        cert = load_der_x509_certificate(raw)
    except Exception:
        cert = load_pem_x509_certificate(raw)
    return b64encode(cert.public_key().encrypt(text.encode('utf-8'), padding.PKCS1v15())).decode('ascii')


def chiama(servizio, protocollo, user, password, pincode, prod=False, verify=True):
    prefix, ns, op, ep_test, ep_prod = SERVIZI[servizio]
    endpoint = ep_prod if prod else ep_test
    pincode_cifrato = cifra(pincode)

    envelope = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f'<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:{prefix}="{ns}">'
        '<soapenv:Header/><soapenv:Body>'
        f'<{prefix}:{op}><DatiInputRichiesta>'
        f'<pinCode>{pincode_cifrato}</pinCode>'
        f'<protocollo>{protocollo}</protocollo>'
        f'</DatiInputRichiesta></{prefix}:{op}>'
        '</soapenv:Body></soapenv:Envelope>'
    )
    headers = {'Content-Type': 'text/xml; charset=utf-8', 'SOAPAction': '""'}
    if not verify:
        import urllib3
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    resp = requests.post(endpoint, data=envelope.encode('utf-8'), headers=headers,
                         auth=(user, password), timeout=120, verify=verify)
    print(f'[{servizio}] HTTP {resp.status_code}  ({endpoint})')
    return resp.text


def estrai(testo, tag):
    return [m.strip() for m in re.findall(rf'<{tag}>(.*?)</{tag}>', testo, re.S)]


def primo(testo, tag):
    v = estrai(testo, tag)
    return v[0] if v else None


def mostra_esito(testo):
    print('  esitoChiamata:', primo(testo, 'esitoChiamata'))
    desc = primo(testo, 'descrizioneEsito')
    if desc:
        print('  descrizioneEsito:', desc)
    # dettagli positivi
    for blocco in re.findall(r'<dettagliEsito>(.*?)</dettagliEsito>', testo, re.S):
        print('  --- invio ---')
        for campo in ('protocollo', 'dataInvio', 'stato', 'descrizione',
                      'nInviati', 'nAccolti', 'nWarnings', 'nErrori'):
            val = primo(blocco, campo)
            if val is not None:
                print(f'    {campo}: {val}')
    # dettagli negativi
    for blocco in re.findall(r'<dettaglioEsitoNegativo>(.*?)</dettaglioEsitoNegativo>', testo, re.S):
        print('  --- esito negativo ---')
        print('    codice:', primo(blocco, 'codice'))
        print('    descrizione:', primo(blocco, 'descrizione'))


def mostra_ricevuta(testo, out_path: Path):
    pdf_b64 = primo(testo, 'pdf')
    if pdf_b64:
        out_path.write_bytes(b64decode(pdf_b64))
        print(f'  Ricevuta PDF salvata: {out_path} ({out_path.stat().st_size} byte)')
    else:
        neg = re.findall(r'<dettaglioEsitoNegativo>(.*?)</dettaglioEsitoNegativo>', testo, re.S)
        if neg:
            for b in neg:
                print('  esito negativo:', primo(b, 'codice'), primo(b, 'descrizione'))
        else:
            print('  Nessun PDF nella risposta. esitoChiamata:', primo(testo, 'esitoChiamata'))


def mostra_errori(testo, out_path: Path):
    # Il dettaglio errori è tipicamente un CSV codificato base64
    for tag in ('csv', 'file', 'dettaglio', 'return'):
        b64 = primo(testo, tag)
        if b64 and len(b64) > 20:
            try:
                contenuto = b64decode(b64)
                out_path.write_bytes(contenuto)
                print(f'  Dettaglio errori salvato: {out_path} ({len(contenuto)} byte)')
                return
            except Exception:
                pass
    neg = re.findall(r'<dettaglioEsitoNegativo>(.*?)</dettaglioEsitoNegativo>', testo, re.S)
    if neg:
        for b in neg:
            print('  ', primo(b, 'codice'), primo(b, 'descrizione'))
    else:
        print('  Nessun dettaglio errori (probabile assenza di errori).')


def recupera_esito(protocollo, user, password, pincode, prod=False, verify=True) -> dict:
    """Interroga EsitoInvii e ritorna un dict strutturato con lo stato dell'invio."""
    testo = chiama('esito', protocollo, user, password, pincode, prod=prod, verify=verify)
    dett = re.search(r'<dettagliEsito>(.*?)</dettagliEsito>', testo, re.S)
    out = {'esitoChiamata': primo(testo, 'esitoChiamata')}
    if dett:
        b = dett.group(1)
        for c in ('protocollo', 'dataInvio', 'stato', 'descrizione',
                  'nInviati', 'nAccolti', 'nWarnings', 'nErrori'):
            out[c] = primo(b, c)
    neg = re.search(r'<dettaglioEsitoNegativo>(.*?)</dettaglioEsitoNegativo>', testo, re.S)
    if neg:
        out['negativo'] = {'codice': primo(neg.group(1), 'codice'),
                           'descrizione': primo(neg.group(1), 'descrizione')}
    return out


def scarica_ricevuta_pdf(protocollo, out_path, user, password, pincode, prod=False, verify=True):
    """Scarica la ricevuta PDF e la salva. Ritorna il Path se ottenuta, altrimenti None."""
    testo = chiama('ricevuta', protocollo, user, password, pincode, prod=prod, verify=verify)
    pdf_b64 = primo(testo, 'pdf')
    if not pdf_b64:
        return None
    out_path = Path(out_path)
    out_path.write_bytes(b64decode(pdf_b64))
    return out_path


def main():
    p = argparse.ArgumentParser(description='Recupero esito/ricevuta invio 730 dal Sistema TS.')
    p.add_argument('--protocollo', required=True, help='Numero di protocollo restituito dall\'invio.')
    p.add_argument('--servizio', choices=['esito', 'ricevuta', 'errori', 'all'], default='all')
    p.add_argument('--out', help='Percorso file di output per ricevuta PDF / dettaglio errori.')
    p.add_argument('--prod', action='store_true', help='Ambiente di produzione.')
    p.add_argument('--user', help='User (CF) — solo produzione.')
    p.add_argument('--password', help='Password — solo produzione.')
    p.add_argument('--pincode', help='Pincode in chiaro — solo produzione.')
    p.add_argument('--insecure', action='store_true', help='Disabilita verifica TLS.')
    args = p.parse_args()

    if not CERT_PATH.is_file():
        raise SystemExit(f'Certificato non trovato: {CERT_PATH}')

    if args.prod:
        mancanti = [n for n in ('user', 'password', 'pincode') if not getattr(args, n)]
        if mancanti:
            raise SystemExit(f'Produzione: parametri mancanti: {mancanti}')
        user, password, pincode = args.user, args.password, args.pincode
        verify = not args.insecure
        print('*** AMBIENTE DI PRODUZIONE ***')
    else:
        user, password, pincode = TEST_USER, TEST_PASSWORD, TEST_PINCODE
        verify = False  # CA di test non pubblica
        print('*** AMBIENTE DI TEST (utente fittizio) ***')

    servizi = ['esito', 'ricevuta', 'errori'] if args.servizio == 'all' else [args.servizio]
    for s in servizi:
        try:
            testo = chiama(s, args.protocollo, user, password, pincode, prod=args.prod, verify=verify)
        except requests.RequestException as e:
            print(f'[{s}] errore di rete: {e}')
            continue
        if s == 'esito':
            mostra_esito(testo)
        elif s == 'ricevuta':
            out = Path(args.out) if args.out else BASE_DIR / f'ricevuta_{args.protocollo}.pdf'
            mostra_ricevuta(testo, out)
        elif s == 'errori':
            out = Path(args.out) if args.out else BASE_DIR / f'errori_{args.protocollo}.csv'
            mostra_errori(testo, out)
        print()


if __name__ == '__main__':
    main()

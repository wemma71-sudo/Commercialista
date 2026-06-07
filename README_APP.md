# Applicazione Web Pipeline Fatture

Applicazione web per eseguire `run_pipeline.py` con selezione cartella da browser.

## Requisiti

- Python 3.7+
- Flask
- tkinter (di solito incluso con Python)

## Installazione

1. Apri un terminale nella cartella del progetto
2. Installa le dipendenze:

```bash
pip install flask
```

## Avvio

Dalla cartella del progetto, esegui:

```bash
python app.py
```

L'applicazione si avvierà e sarà accessibile da:
```
http://localhost:5000
```

Apri il browser e accedi a questo URL.

## Utilizzo

1. **Seleziona cartella**: Clicca su "Sfoglia..." per aprire il file picker di Windows oppure incolla il percorso direttamente nella casella di testo
2. **Esegui pipeline**: Clicca su "Esegui Pipeline"
3. L'applicazione eseguirà `run_pipeline.py` con la cartella selezionata
4. Vedrai lo stato in tempo reale e l'output al completamento

## Caratteristiche

- ✅ Interfaccia web moderna e responsiva
- ✅ Selezione cartella con file picker di Windows
- ✅ Input manuale della cartella
- ✅ Esecuzione asincrona della pipeline
- ✅ Salvataggio automatico dell'ultima cartella utilizzata
- ✅ Visualizzazione output in tempo reale
- ✅ Validazione input

## Arresto

Premi `Ctrl+C` nel terminale per fermare l'applicazione.

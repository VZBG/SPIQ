# SPIQ

Prima versione della suite documentale SPIQ.

## Modulo disponibile

- Carta intestata: applicazione di header e footer da un template `.docx` a un documento `.docx` caricato dall'utente.
- L'elaborazione avviene interamente nel browser.
- PDF/A: funzione predisposta nell'interfaccia ma non ancora attiva.

## Struttura

- `index.html`: interfaccia principale
- `assets/styles.css`: stile
- `js/app.js`: logica dell'interfaccia
- `js/docx-letterhead.js`: motore di modifica dei file Word
- `templates/letterheads/catalog.json`: elenco dei modelli disponibili
- `templates/letterheads/*.docx`: modelli Word (già presenti nel repository GitHub)

## Pubblicazione con GitHub Pages

Dopo aver caricato i file nel branch `main`:

1. Aprire `Settings` del repository.
2. Selezionare `Pages` nella colonna sinistra.
3. In `Build and deployment`, scegliere `Deploy from a branch`.
4. Selezionare branch `main` e cartella `/ (root)`.
5. Salvare.

GitHub pubblicherà il sito all'indirizzo indicato nella stessa pagina.

## Nota tecnica

La prima versione trasferisce i riferimenti a header/footer del modello nel documento caricato e copia le risorse collegate (per esempio immagini). Non modifica il corpo del documento né sostituisce i suoi stili. I documenti Word con strutture particolarmente complesse dovranno essere verificati durante il collaudo.

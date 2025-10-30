# Calendario Docenti — deploy rapido su Surge

## Requisiti
- Node.js
- Surge CLI (`npm i -g surge`)

## Deploy
```bash
cd docenti-surge
surge .
# oppure specifica un dominio:
surge . docenti-test.surge.sh
```
Per pubblicare aggiornamenti: riesegui `surge .` dalla stessa cartella.

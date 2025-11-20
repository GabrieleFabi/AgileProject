# 📘 ITS Calendario 2025-2027

## 🧭 Panoramica generale

**Nome progetto:** ITS Calendario 2025-2027  
**Obiettivo:** Fornire due portali dedicati per la consultazione dei calendari delle lezioni: uno per gli **Studenti** e uno per i **Docenti**. I dati vengono caricati da file Excel e visualizzati in modo interattivo.
**Utenti finali:** Studenti, docenti e personale amministrativo dell’istituto.

**Portale Studenti:** [https://cal-stud-itsaa.surge.sh/](https://cal-stud-itsaa.surge.sh/)  
**Portale Docenti:** [https://cal-doc-itsaa.surge.sh/](https://cal-doc-itsaa.surge.sh/)

---

## 🏗️ Architettura del sistema

Il progetto è diviso in due applicazioni web distinte, entrambe ospitate su Surge:

### 1️⃣ Portale Studenti (`/Studenti`)
Dedicato agli studenti, permette di:
- Selezionare il proprio **Anno** (1 o 2).
- Scegliere il **Corso** specifico (es. *FUST*, *CYSE*, *WEB*...).
- Visualizzare il calendario delle lezioni filtrato per il corso selezionato.
- Vedere a colpo d'occhio giornate "corte" o esami.

### 2️⃣ Portale Docenti (`/Docenti`)
Dedicato ai docenti, permette di:
- Visualizzare un elenco completo di tutti i docenti trovati nel file Excel.
- Cercare il proprio nome tramite una barra di ricerca.
- Accedere a una vista calendario personalizzata che aggrega tutte le lezioni del docente attraverso i vari corsi.
- Filtrare le lezioni per corso specifico tramite una legenda interattiva.

---

## ⚙️ Tecnologie utilizzate

| Tecnologia | Ruolo | Descrizione |
|-------------|--------|-------------|
| **HTML5** | Struttura | Base semantica, layout responsive. |
| **CSS3** | Stile | Styling moderno, variabili CSS, Flexbox/Grid, design responsive. |
| **JavaScript (ES6)** | Logica | Parsing Excel (SheetJS), logica di filtraggio, gestione DOM. |
| **SheetJS (xlsx)** | Libreria | Lettura e parsing dei file `.xlsx` direttamente nel browser. |
| **Surge** | Hosting | Hosting statico per il deploy delle due applicazioni. |

---

## 🧩 Funzionalità principali

### 🎓 Portale Studenti
1.  **Selezione Guidata:**
    *   Scelta Anno (1° o 2°).
    *   Scelta Corso (bottoni generati dinamicamente).
2.  **Visualizzazione Calendario:**
    *   Tabella chiara con Data, Orario, Modulo, Docente, Aula.
    *   **Evidenziazione Esami:** Le penultime lezioni di ogni modulo vengono evidenziate in verde.
    *   **Warning Orari:** Giornate con poche ore di lezione sono segnalate in arancione (<8h) o rosso (≤4h).
3.  **Filtri:**
    *   Barra di ricerca testuale (cerca per materia, docente, ecc.).
    *   Bottone "Mostra tutto" per vedere anche le lezioni passate (default: da oggi in poi).

### �‍🏫 Portale Docenti
1.  **Elenco Docenti:**
    *   Generazione automatica della lista docenti dai dati Excel.
    *   Ricerca rapida per nome.
2.  **Calendario Personale:**
    *   Vista aggregata di tutte le lezioni del docente selezionato.
    *   Le righe sono colorate in base al Corso di appartenenza.
3.  **Legenda Interattiva:**
    *   Permette di accendere/spegnere la visualizzazione dei singoli corsi per un'analisi più pulita.

---

## � Design Responsive

Entrambi i portali sono ottimizzati per l'uso da mobile:
- **Mobile:** Layout a colonna singola, orari compattati, tabelle scorrevoli orizzontalmente se necessario.
- **Desktop:** Layout esteso, controlli affiancati, tabelle ampie.

## � Deploy

Il progetto viene distribuito tramite **Surge.sh**.
Ogni cartella (`Studenti` e `Docenti`) viene deployata come sito indipendente.

Comandi di deploy (eseguiti dalla root del progetto):
```bash
# Deploy Studenti
npx surge "./Progetto Calendar/Studenti" cal-stud-itsaa.surge.sh

# Deploy Docenti
npx surge "./Progetto Calendar/Docenti" cal-doc-itsaa.surge.sh
```

# 📘 ITS Calendario Studs 2025

## 🧭 Panoramica generale

**Nome progetto:** ITS Calendario Studs 2025  
**Obiettivo:** Fornire un sistema interattivo per la visualizzazione dei calendari delle lezioni degli studenti ITS, generato dinamicamente da un file Excel aggiornato automaticamente ogni notte.  
**Utenti finali:** Studenti, docenti e personale amministrativo dell’istituto.  

**Hosting:**
- 🌐 **Sito pubblico:** Google Sites  
- 🚀 **Applicazione interattiva:** Netlify (deploy statico HTML/JS/CSS)  
- 🔗 **Esempio URL:** [https://its-calendar-2025-2027.netlify.app](https://its-calendar-2025-2027.netlify.app)

---

## 🏗️ Architettura del sistema

Il sistema è composto da due parti principali:

### 1️⃣ Frontend — Applicazione Web Interattiva
- Ospitata su **Netlify**, poi incorporata nel Google Site ufficiale tramite un *iframe*.  
- Contiene tutta la logica di visualizzazione, filtraggio e formattazione dei dati Excel.

### 2️⃣ Backend “leggero” — Automazione file
- Non esiste un server tradizionale.  
- L’aggiornamento del file Excel è gestito da uno **script Google Apps Script** che ogni notte rigenera il file `calendario.xlsx` e lo pubblica su Netlify.  
- Il sito legge sempre l’ultima versione per generare il calendario aggiornato.

---

## ⚙️ Tecnologie utilizzate

| Tecnologia | Ruolo | Descrizione |
|-------------|--------|-------------|
| **HTML5** | Struttura | Base semantica del sito, con layout responsive e sezioni dinamiche. |
| **CSS3** | Stile | Styling responsivo, gestione colori, legende e bordi condizionali. |
| **JavaScript (ES6)** | Logica | Lettura file Excel, filtraggio, ordinamento, ricerca e gestione viste. |
| **SheetJS (xlsx.js)** | Parsing Excel | Libreria per leggere e convertire file Excel in JSON. |
| **Netlify** | Hosting statico | Deploy rapido e gratuito di HTML/JS/CSS. |
| **Google Sites** | Integrazione | Piattaforma istituzionale per incorporare l’app tramite iframe. |
| **Google Drive / Sheets** | Sorgente dati | File Excel aggiornato automaticamente ogni notte. |
| **CSS Flexbox / Grid** | Layout | Struttura adattiva per desktop, tablet e mobile. |

---

## 🧩 Funzionalità principali

### 🔹 1. Selezione Anno e Corso
- Schermata iniziale per la scelta del corso (es. *FUST*, *CYSE*...).  
- Ogni corso carica dinamicamente le lezioni corrispondenti.  
- Al ricaricamento, la pagina iniziale è sempre la selezione corso per coerenza.

### 🔹 2. Caricamento e parsing automatico Excel
- Lettura automatica del file `calendario.xlsx` con **SheetJS (XLSX)**.  
- Le colonne vengono interpretate automaticamente (Data, Ora, Corso, Materia, Docente, Aula…).

### 🔹 3. Visualizzazione tabellare dinamica
- Tabella HTML responsive con intestazioni dinamiche.
- **Colonne principali:**  
  - Data (`dd/mm/aa`)  
  - Orario (3 righe: inizio / “–” / fine)  
  - Materia / Docente / Aula  
- Adattamento completo per mobile tramite CSS Grid.

### 🔹 4. Filtri e ricerca
- Barra di ricerca per testo libero (materia, docente, aula...).  
- Filtraggio in tempo reale della tabella.

### 🔹 5. Bottone “Mostra tutto”
- Mostra solo da oggi in poi per default.  
- Il bottone **“Mostra tutto”** ricarica l’intero calendario.

### 🔹 6. Evidenziazione orari insufficienti
| Condizione | Colore | Significato |
|-------------|---------|-------------|
| ≥ 8 ore | Nessun colore | Giorno pieno |
| < 8 ore | 🟧 Arancione chiaro | Giornata parziale |
| < 4 ore | 🔴 Rosso chiaro | Giornata ridotta |

### 🔹 7. Legenda grafica
- Box laterale con i colori `<8h`, `<4h`, `esame`.  
- Layout 50/50 con barra di ricerca.

### 🔹 8. Logo dinamico
- Rettangolo blu in alto a sinistra con il **nome del corso corrente**.  
- In home/selezione corso mostra “ITS”.

### 🔹 9. Gestione utenti multipli
- Sistema in sola lettura (read-only).  
- Tutti gli utenti accedono alla stessa istanza pubblicata.  
- Il file Excel è condiviso tramite account abilitato dell’istituto.

---

## 💻 Funzioni del codice JavaScript

### 🧾 Acquisizione & parsing Excel
- **`fetchAndLoadXlsx(url)`** → Scarica e carica il file Excel da Netlify.  
- **`handleFile(file)`** → Carica un file Excel locale.  
- **`excelToJson(ws)`** → Converte il foglio attivo in JSON normalizzato.

### 🧮 Pre-calcolo e formattazioni
- **`computeDayCoverage(headers, rows)`** → Calcola minuti giornalieri per evidenziazione `<8h` / `≤4h`.  
- **`computePenultimateKeys(headers, rows)`** → Identifica penultime lezioni (esame).  
- **Helpers** → `fmtDateIT`, `fmtTimeFromDate`, `fmtTimeFromFraction`, `prettyValue`.

### 🧱 Rendering & filtri
- **`renderTable(baseHeaders, allRows)`**
  1. Costruisce la tabella HTML con header originali.  
  2. Combina “Dalle–Alle” in “Orario” su mobile.  
  3. Applica il filtro “da oggi”.  
  4. Filtra in tempo reale da `#searchInput`.  
  5. Evidenzia giorni corti e “Esame”.

### 🔍 Ricerca & viste
- **Ricerca live:** integrata nel `renderTable(...)`.  
- **Filtro “Mostra tutto / da oggi”:** gestito da `showAll` + `updateToggleButton()`.  
- **Reset:** `clearAll()` riporta alla selezione corso e resetta logo ITS.

### 🌙 Aggiornamento notturno
- **`scheduleMidnightRefresh()`** → Ricarica automaticamente il calendario alle **00:05**, sincronizzato con Netlify.

### 🧩 Integrazione UI
- Gestione eventi: caricamento, selezione fogli, toggle, ritorno, scelta anno/corso.  
- `init()` lancia il caricamento automatico da `/data/calendario.xlsx`.

---

## 📱 Design responsive

**Mobile**
- Colonna orari su 3 righe.
- Layout a colonna singola.
- Padding ridotto.

**Desktop / Tablet**
- Barra comandi + legenda affiancate (50% ciascuna).
- Tabelle con padding maggiore.
- Righe e settimane ben separate.

---

## 🔄 Sincronizzazione notturna (Google Apps Script → Netlify)

Ogni notte lo script:
1. Esporta lo Sheet di Google in `.xlsx`.  
2. Legge il manifest dell’ultimo deploy Netlify.  
3. Aggiorna solo il file `/data/calendario.xlsx`.  
4. Crea un nuovo deploy con il manifest clonato.  
5. Carica solo il file aggiornato.

> 💡 In questo modo il sito resta identico e viene aggiornato solo il calendario, evitando rebuild completi e tempi d’attesa.

### ⚙️ Flusso di esecuzione (`nightlyUploadXlsxToNetlify()`)

1. Export XLSX via API Drive v3 (`files/{id}/export`).  
2. Legge il `published_deploy.id` di Netlify.  
3. Elenca file del deploy e costruisce `{ "/path": "sha1" }`.  
4. Sostituisce `TARGET_PATH` con nuovo digest SHA1.  
5. Crea un nuovo deploy con manifest clonato.  
6. Upload PUT solo del file aggiornato.  
7. Log finale con `deploy.id`.

> Netlify **non ricompila** i file statici: viene aggiornato solo `/data/calendario.xlsx`.

---

## 🚀 Distribuzione e Deploy

### 🟢 Netlify
- Deploy statico HTML/CSS/JS.  
- Upload via drag & drop su [app.netlify.com/drop](https://app.netlify.com/drop).  
- Genera URL pubblico, es. `https://its-calendar-2025-2027.netlify.app`.

### 🟣 Google Sites
- Nella pagina ufficiale ITS si inserisce un *iframe*:  

```html
<iframe src="https://its-calendar.netlify.app/" width="100%" height="900"></iframe>

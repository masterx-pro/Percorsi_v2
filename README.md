# 📱 Percorsi Pro - Ottimizzatore di Percorsi Android

App Android completa per l'ottimizzazione di percorsi con gestione file Excel, mappa integrata e navigazione.

## ✨ Funzionalità

### 📁 Caricamento File Excel
- Supporto file .xlsx, .xls, .csv
- Selezione colonne: Latitudine, Longitudine
- Colonne opzionali: Data, Operatore, Etichetta
- Auto-rilevamento colonne intelligente

### 🗺️ Mappa Integrata
- Visualizzazione percorso su mappa OpenStreetMap
- Markers colorati (verde=partenza, rosso=arrivo, blu=intermedi)
- Zoom e pan touch
- Traccia del percorso

### 📍 Gestione Percorsi
- Salvataggio automatico percorsi elaborati
- Lista percorsi salvati
- Visualizzazione dettaglio con ordine destinazioni
- Eliminazione percorsi

### 🧭 Navigazione
- Apertura in Google Maps per navigazione
- Divisione automatica percorsi lunghi in segmenti

### 💾 Export
- Export Excel (.xlsx) con scelta cartella
- Export GPX per navigatori
- Export KML per Google Earth
- Scelta cartella di destinazione

## 🚀 Installazione

### Scarica APK
1. Vai alla sezione [Releases](../../releases)
2. Scarica l'APK più recente
3. Installa sul dispositivo Android

### Compila da GitHub Actions
1. Fork questo repository
2. Vai su **Actions** → **Build Android APK**
3. Clicca **Run workflow**
4. Scarica l'APK da **Artifacts**

## 📖 Uso

### Caricamento Excel
1. Clicca "CARICA FILE EXCEL"
2. Seleziona il file
3. Scegli le colonne corrette
4. Clicca "CONFERMA E OTTIMIZZA"

### Inserimento Manuale
1. Clicca "INSERIMENTO MANUALE"
2. Inserisci coordinate (lat,lon per riga)
3. Opzionale: aggiungi etichette
4. Clicca "OTTIMIZZA PERCORSO"

### Visualizza Risultati
1. Vai su "PERCORSI SALVATI"
2. Seleziona un percorso
3. Usa i pulsanti:
   - 🗺️ MAPPA - Visualizza su mappa
   - 🧭 NAVIGA - Apri in Google Maps
   - 💾 ESPORTA - Salva file

## 🔧 Requisiti

- Android 5.0+ (API 21)
- Connessione internet (per calcolo distanze stradali)
- ~50MB spazio disco

## 📄 Licenza

MIT License

## 👤 Autore

**Mattia Prosperi**

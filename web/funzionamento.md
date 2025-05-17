# Funzionamento Applicativo

## 1. Gestione Schede di Lavoro

### 1.1 Creazione Scheda
- **Numero**: identificativo univoco della scheda
- **Data ricevimento**: data di ricezione della richiesta
- **Cliente**: selezione del cliente da anagrafica
- **Modello**: selezione del modello di macchina
- **Stato**: stato attuale della scheda (es. in lavorazione, completata)

### 1.2 Gestione Accessori
- Lista degli accessori disponibili
- Possibilità di aggiungere nuovi accessori
- Associazione degli accessori alla scheda

### 1.3 Gestione Problemi
- Lista dei problemi standard
- Descrizione del problema specifico
- Possibilità di inserire problemi personalizzati

### 1.4 Esito dell'Operazione
- Descrizione dell'esito
- Campo testo libero per note aggiuntive
- Gestione della garanzia (previa valutazione del centro assistenza)

## 2. Gestione Documenti di Trasporto

### 2.1 DDT (Documento di Trasporto)
- **Numero**: identificativo del DDT
- **Data**: data di emissione
- **Cliente**: cliente destinatario
- **Trasportatore**: selezione del trasportatore
- **Causale**: motivo del trasporto
- **Porto**: tipo di porto (es. franco, assegnato)
- **Colli**: numero di colli
- **Peso**: peso totale
- **Annotazioni**: note aggiuntive

### 2.2 Lettera di Accompagnamento
- Numero DDT
- Destinatario
- Numero ordine
- Indirizzo completo
- Trasporto a cura di
- Data ritiro

## 3. Gestione Clienti

### 3.1 Anagrafica Clienti
- **Nome/Ragione sociale**
- **Indirizzo**
- **Città**
- **Provincia**
- **CAP**
- **Telefono**
- **Fax**
- **Email**
- **Partita IVA**
- **Codice Fiscale**

### 3.2 Tipologie di Clienti
- Business
- Privati
- Distinzione tra i due tipi per la generazione di documenti

## 4. Gestione Articoli e Listini

### 4.1 Articoli
- Codice articolo
- Descrizione
- Prezzo base
- Marche disponibili
- Rivenditori associati

### 4.2 Listini
- Gestione multipli listini
- Associazione listini ai rivenditori
- Prezzi di listino
- Possibilità di aggiornamento

## 5. Gestione Email

### 5.1 Invio Email
- Selezione destinatario
- Allegato da inviare
- Testo personalizzato
- Campi precompilati con:
  - Numero scheda
  - Data ricevimento
  - Nome cliente
  - Email cliente

## 6. Gestione Trasportatori

### 6.1 Anagrafica Trasportatori
- Nome/ragione sociale
- Indirizzo
- Contatti
- Storico utilizzazioni

### 6.2 Ordinamento Trasportatori
- Ordinamento per frequenza di utilizzo
- Conteggio delle utilizzazioni
- Preferenza per i trasportatori più utilizzati

## 7. Gestione Costi

### 7.1 Costi di Trasporto
- Costi presa
- Costi consegna
- Costi totali
- Gestione costi per clienti privati
- Evidenziazione variazioni nei costi

### 7.2 Costi Manodopera
- Ore manodopera (obbligatorie)
- Costo orario
- Totale manodopera

## 8. Gestione Garanzia

### 8.1 Stato Garanzia
- In garanzia
- Richiesta in attesa di conferma
- Non in garanzia
- Previa valutazione del centro assistenza

### 8.2 Processo
- Richiesta di garanzia
- Valutazione del centro assistenza
- Conferma/denuncia della garanzia

## 9. Gestione PDF e Documenti

### 9.1 Generazione PDF
- Documenti di trasporto
- Lettere di accompagnamento
- Schede di lavoro
- Report vari

### 9.2 Archiviazione Documenti
- Allegati
- Documenti generati
- Storico documenti

## 10. Gestione Utenti e Accessi

### 10.1 Gestione Utenti
- Anagrafica utenti
- Livelli di accesso
- Permessi
- Storico accessi

### 10.2 Sicurezza
- Autenticazione
- Autorizzazione
- Log delle attività

## 11. Gestione Report e Export

### 11.1 Report
- Report schede
- Report trasporti
- Report costi
- Report garanzie

### 11.2 Export Excel
- Esportazione dati
- Formattazione specifica
- Gestione costi per clienti privati

## Menu Principale

Il menu principale dell'applicativo è definito nel file [intestazione.asp](cci:7://file:///Users/maurolarese/Dropbox/arma/infoschede/web/amministrazione/Infoschede/intestazione.asp:0:0-0:0) e è composto da queste voci principali:

1. **RICHIESTE ASSISTENZA**
   - Gestione delle richieste di assistenza in arrivo
   - Assegnazione delle richieste ai centri assistenza

2. **SCHEDE ASSISTENZA**
   - Gestione delle schede assegnate
   - Monitoraggio dello stato delle schede
   - Gestione delle attività di assistenza

3. **ANAGRAFICHE CLIENTI**
   - Gestione dell'anagrafica clienti
   - Visualizzazione e modifica dei dati
   - Ricerca clienti

4. **RITIRI**
   - Gestione delle richieste di ritiro
   - Organizzazione dei ritiri
   - Assegnazione dei trasportatori

5. **SPEDIZIONI**
   - Gestione delle spedizioni
   - Creazione di documenti di trasporto
   - Gestione dei trasportatori

6. **TABELLE**
   - Gestione delle tabelle di configurazione
   - Impostazioni sistema
   - Parametrizzazione

7. **Esci dall'applicazione**
   - Logout dall'applicativo
   - Chiusura sessione

## Note Tecniche

### 1. Database
- Tabelle principali:
  - `sgtb_schede`: schede di lavoro
  - `sgtb_ddt`: documenti di trasporto
  - `sgtb_accessori`: accessori
  - `sgtb_problemi`: problemi
  - `sgtb_esiti`: esiti
  - `tb_utenti`: utenti
  - `tb_indirizzario`: anagrafica
  - `gtb_articoli`: articoli
  - `gtb_listini`: listini prezzi

### 2. Sicurezza
- Controllo accessi per ruolo
- Validazione input
- Protezione dati sensibili

### 3. Performance
- Ottimizzazione query
- Cache
- Indicizzazione appropriata

### 4. Integrazioni
- Email
- File system
- Report PDF/Excel

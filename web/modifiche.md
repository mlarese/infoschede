# Modifiche da Implementare

## 1. Invio/Ricezione Mail da Applicativo
**File coinvolti:**
- `/Plugin/Bolla.ascx.cs` - Gestione email di conferma
- `/amministrazione/nextCom/ComunicazioniView.asp` - Gestione delle email
- `/amministrazione/nextCom/delete.asp` - Contiene riferimenti alla tabella `tb_email`

**Interventi:**
- Verificare il funzionamento del sistema di invio/ricezione email
- Controllare le configurazioni SMTP nel Web.Config
- Verificare il template delle email

## 2. Gestione Listini e Pezzi di Ricambio Duplicati
**File coinvolti:**
- `/amministrazione/Infoschede/_import_ricambi_*.asp` - Script di importazione listini
- `/amministrazione/nextB2B/Tools_B2B.asp` - Contiene funzioni di gestione listini
- Tabelle DB: `gtb_listini`, `gtb_prezzi`

**Interventi:**
- Modificare gli script di import per rilevare e gestire i duplicati
- Implementare un sistema di cancellazione logica per i vecchi listini
- Aggiungere un campo data/timestamp per identificare il listino più recente
- Aggiornare le query per filtrare solo gli articoli attivi (non cancellati logicamente)

## 3. Scheda Nuovo Lavoro
**File coinvolti:**
- `/Plugin/Scheda.ascx.cs` - Gestione interfaccia scheda lavoro
- `/amministrazione/Infoschede/SchedeNew.asp` - Creazione nuove schede
- `/amministrazione/Infoschede/SchedeSalva.asp` - Salvataggio dati schede

**Interventi:**
- Aggiungere campo Cod. Trasportatore con anagrafica
- Aggiungere funzionalità "inserisci nuovo accessorio" alla sezione Accessori
- Aggiornare tabella degli accessori

## 4. Sostituzione Voce "Garanzia"
**File coinvolti:**
- `/Plugin/Scheda.ascx.cs` - Contiene riferimenti alla garanzia
- `/Plugin/SchedaStampa.ascx.cs` - Stampa delle informazioni di garanzia
- `/Plugin/Bolla.ascx.cs` - Visualizzazione stato garanzia

**Interventi:**
- Sostituire "Garanzia" con "VALUTAZIONE DEL CENTRO ASSISTENZA"
- Modificare il sistema per inviare una mail al cliente finale con la valutazione

## 5. Aggiunta Campo Testo Libero in "Esito dell'operazione"
**File coinvolti:**
- `/Plugin/Scheda.ascx.cs` - Interfaccia utente della scheda
- Tabelle DB: potrebbe richiedere l'aggiunta di un campo

**Interventi:**
- Aggiungere campo di testo libero nella sezione "Esito dell'operazione"
- Aggiornare il database per includere il nuovo campo

## 6. Rendere Obbligatorio il Campo ORE MANODOPERA
**File coinvolti:**
- `/Plugin/Scheda.ascx.cs` - Validazione campi scheda
- `/amministrazione/Infoschede/SchedeSalva.asp` - Salvataggio dati

**Interventi:**
- Aggiungere validazione lato client per il campo ORE MANODOPERA
- Aggiungere validazione lato server
- Assicurarsi che il costo orario sia sempre specificato

## 7. Miglioramenti Layout Documento al Cliente
**File coinvolti:**
- `/Plugin/SchedaStampa.ascx.cs` - Layout di stampa per il cliente

**Interventi:**
- Modificare il layout per visualizzare "Spese trasporto Presa" e "Spese trasporto Consegna" in colonna

## 8. Gestione DDT per Clienti Privati
**File coinvolti:**
- `/Plugin/Bolla.ascx.cs` - Generazione DDT e lettere di trasporto

**Interventi:**
- Modificare la logica per non generare DDT per clienti di tipo privato

## 9. Unificazione Anagrafica Clienti
**File coinvolti:**
- `/amministrazione/Infoschede/ClientiNew.asp` (da verificare l'esistenza)
- `/amministrazione/nextCom/ContattiNew.asp` - Creazione contatti

**Interventi:**
- Unificare le anagrafiche tra clienti business e privati
- Aggiornare le interfacce di gestione contatti

## 10. Miglioramenti Ricerca Anagrafica
**File coinvolti:**
- Interfacce di ricerca clienti (da identificare)

**Interventi:**
- Estendere la ricerca a tutti i campi dell'anagrafica

## 11. Riduzione Categorie Accessori
**File coinvolti:**
- File di gestione categorie accessori (da identificare)

**Interventi:**
- Ridurre il numero di categorie degli accessori
- Aggiornare le interfacce di selezione

## 12. Ordinamento Lista Trasportatori
**File coinvolti:**
- File di visualizzazione trasportatori nei DDT (da identificare)

**Interventi:**
- Implementare ordinamento per frequenza di utilizzo
- Memorizzare e aggiornare i contatori di utilizzo dei trasportatori

## 13. Miglioramenti Export Excel
**File coinvolti:**
- File di generazione report Excel (da identificare)

**Interventi:**
- Esplicitare i costi di trasporto per i clienti privati
- Evidenziare i casi in cui (Nr - costi trasporto) ≠ (Nr - costi trasporto finale)

## 14. Sostituzione Logo
**File coinvolti:**
- File template che contengono il logo attuale

**Interventi:**
- Sostituire il logo "20 anni" con "35 anni"
- Aggiornare tutti i template che utilizzano il logo

## Punti da Investigare Ulteriormente

### 1. Gestione Listini e Pezzi di Ricambio Duplicati
- Non ho trovato il file principale che gestisce l'importazione dei listini attuale
- Non ho trovato il file che gestisce la logica di cancellazione degli articoli

### 2. Scheda Nuovo Lavoro
- Non ho trovato il file che gestisce l'anagrafica dei trasportatori
- Non ho trovato il file che gestisce la tabella degli accessori

### 3. Gestione DDT per Clienti Privati
- Non ho trovato il file che gestisce la logica di generazione del DDT
- Non ho trovato il file che gestisce la distinzione tra clienti business e privati

### 4. Unificazione Anagrafica Clienti
- Non ho trovato il file principale che gestisce l'anagrafica clienti
- Non ho trovato il file che gestisce la distinzione tra clienti business e privati

### 5. Miglioramenti Ricerca Anagrafica
- Non ho trovato il file che gestisce la logica di ricerca degli clienti
- Non ho trovato il file che gestisce l'interfaccia di ricerca

### 6. Riduzione Categorie Accessori
- Non ho trovato il file che gestisce la lista delle categorie
- Non ho trovato il file che gestisce la gestione delle categorie

### 7. Ordinamento Lista Trasportatori
- Non ho trovato il file che gestisce l'ordinamento della lista
- Non ho trovato il file che gestisce il conteggio delle utilizzazioni

### 8. Miglioramenti Export Excel
- Non ho trovato il file che gestisce la generazione del report Excel
- Non ho trovato il file che gestisce il calcolo dei costi di trasporto

### 9. Sostituzione Logo
- Non ho trovato tutti i file template che contengono il logo
- Non ho trovato il file che gestisce la gestione dei template

## Note
Per implementare correttamente queste modifiche, sarà necessario:
1. Identificare i file mancanti
2. Capire la logica attuale di gestione dei dati
3. Verificare eventuali dipendenze tra i componenti
4. Testare attentamente le modifiche implementate

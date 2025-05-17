Arma infoschede Aprile 2025

# Infoschede - Appunti e Richieste Interventi

## Accessi

### Jump Desktop
- **IP:** 136.144.247.246
- **Utente:** Administrator
- **Password:** Hdro31!

### Sito Web
- **URL Utente Esterno:** http://www.infoschede.it/it/
  - **Utente:** admin
  - **Password:** admin
- **URL Amministrazione:** http://www.infoschede.it/amministrazione/default.asp
  - **Utente:** HIDROSERVICES
  - **Password:** ATHENA69

### Login Hidroservice
- Dati di accesso: [da completare]

## Percorsi File

- **Cartella Sorgenti:** `C:\inetpub\wwwroot\infoschede.it\`
- **Cartella Upload:** `C:\CombiRoot\infoschede.it\upload\`
- **Database:** SQL Server
- **Cartella Admin:** `c:\inetpub\wwwroot\web`

## Lista Richieste Interventi (24-01-2025)

1. [x] **Gestione Email** (Completato 16/05/2025)
   - **File da modificare:** 
     - `/web/amministrazione/Infoschede/SchedeMod.asp` - Gestione schede e invio email
     - `/web/amministrazione/library/Class_Messages_CommonParts.asp` - Classe per invio email
   - **Problemi da risolvere:**
     - Verificare funzionamento invio/ricezione mail da applicativo nella nuova scheda
     - Risolvere errore di invio email dalla logistica (sezione SchedaMod)
   - **Note tecniche:**
     - Risolto il problema di permessi per l'eseguibile wkhtmltopdf e le sue DLL
     - Confermato il corretto funzionamento dell'invio email e generazione PDF
     - Le funzioni di invio email sono definite nella classe `Class_Messages_CommonParts.asp`
     - Controllare le funzioni `SendToContact`, `SendToAdmin` e `Save` per eventuali errori

2. **Gestione Pezzi di Ricambio**
   - **File da modificare:**
     - `/web/amministrazione/Infoschede/ArticoliSeleziona.asp` - Gestione selezione articoli
     - `/web/amministrazione/Infoschede/SchedeMod.asp` - Integrazione con articoli
   - **Problemi da risolvere:**
     - Eliminare i doppioni nei listini mantenendo solo il più recente
     - Nella scheda "Aggiungi Ricambio" si apre un popup con duplicati
   - **Note tecniche:**
     - Verificare la query SQL in ArticoliSeleziona.asp che seleziona gli articoli
     - Aggiungere filtro per escludere duplicati o mantenere solo il più recente

3. **Gestione Listini**
   - **File da modificare:**
     - Database: tabelle relative ai listini
     - File di import dei listini
   - **Problemi da risolvere:**
     - Attualmente i listini vanno in AGGIUNTA - modificare per eliminare i vecchi e mantenere SOLO i nuovi
     - Per gli articoli usati nelle vecchie schede, implementare cancellazione logica (non fisica)
     - Rendere gli articoli vecchi non visibili per le nuove schede
   - **Note tecniche:**
     - Aggiungere campo per flag di articolo attivo/disattivato
     - Modificare query di selezione articoli per filtrare solo quelli attivi
     - Durante l'import dei nuovi listini, disattivare articoli vecchi invece di eliminarli

4. **Scheda Nuovo Lavoro**
   - **File da modificare:**
     - `/web/amministrazione/Infoschede/SchedeMod.asp` - Form principale
     - `/web/amministrazione/Infoschede/SchedeSalva.asp` - Salvataggio dati
   - **Problemi da risolvere:**
     - Riferimento al punto 1 (gestione email)
     - Inserire il Codice Trasportatore (creare anagrafica dei Trasportatori)
     - Alla voce "Accessori" aggiungere in coda "Inserisci nuovo accessorio"
     - Il testo va aggiunto alla tabella degli accessori
   - **Note tecniche:**
     - Creare nuova tabella per anagrafica trasportatori se non esiste
     - Aggiungere campo per codice trasportatore nel form e nel DB
     - Aggiungere funzionalità per gestire nuovi accessori

5. **Modifica Garanzia**
   - **File da modificare:**
     - `/web/amministrazione/Infoschede/SchedeMod.asp` - Interfaccia utente
     - `/web/amministrazione/Infoschede/SchedeSalva.asp` - Persistenza dati
   - **Problemi da risolvere:**
     - Sostituire la voce "Garanzia" con "VALUTAZIONE DEL CENTRO ASSISTENZA"
     - Vedi punto 1 per l'invio mail
   - **Note tecniche:**
     - Modificare label nella UI senza alterare il comportamento funzionale
     - Assicurarsi che la modifica non influisca sui report e documenti collegati

6. **Esito dell'Operazione**
   - **File da modificare:**
     - `/web/amministrazione/Infoschede/SchedeMod.asp` - Interfaccia utente
     - `/web/amministrazione/Infoschede/SchedeSalva.asp` - Persistenza dati
     - Database: aggiungere campo nella tabella apposita
   - **Problemi da risolvere:**
     - Aggiungere un campo TXT ad inserimento libero nella sezione "Esito dell'operazione"
   - **Note tecniche:**
     - Creare nuovo campo nella tabella del database
     - Aggiungere il campo nel form di modifica
     - Aggiornare la funzione di salvataggio per gestire il nuovo campo

7. **Inserimento Ore Manodopera**
   - **File da modificare:**
     - `/web/amministrazione/Infoschede/SchedeMod.asp` - Validazione input
     - `/web/amministrazione/Infoschede/SchedeSalva.asp` - Controllo dati
   - **Problemi da risolvere:**
     - Rendere obbligatorio l'inserimento ORE MANODOPERA (con costo orario) nella scheda
   - **Note tecniche:**
     - Aggiungere validazione client-side con JavaScript
     - Implementare controllo server-side nella fase di salvataggio
     - Modificare messaggi di errore per informare l'utente

8. **Scheda 28682**
   - **Riferimento:**
     - Scheda ID 28682 come esempio di implementazione corretta
   - **Problemi da risolvere:**
     - Verificare l'implementazione della manodopera, costo presa e costo riconsegna
   - **Note tecniche:**
     - Utilizzare questa scheda come modello per le modifiche da implementare
     - Rendere coerenti tutte le schede secondo questo standard

9. **Spese Trasporto**
   - **File da modificare:**
     - Template di stampa per il documento inviato al cliente
     - `/web/amministrazione/Infoschede/Reports/` - Cartella dei report
   - **Problemi da risolvere:**
     - Sul documento al cliente, riorganizzare le voci "Spese Trasporto Presa" e "Spese Trasporto Consegna"
     - Metterle in colonna per migliorare leggibilità e comprensione
     - Aggiungere campi per manodopera, costo presa e costo riconsegna
   - **Note tecniche:**
     - Modificare il layout del template di stampa
     - Assicurarsi che tutti i campi siano correttamente allineati
     - Testare con la scheda 28682 come riferimento

10. **Gestione Clienti di tipo Privato**
   - **File da modificare:**
     - Moduli di generazione DDT e lettera di trasporto
   - **Problemi da risolvere:**
     - Per Clienti di tipo privato NON va generato il DDT e lettera di trasporto
   - **Note tecniche:**
     - Aggiungere controllo per il tipo di cliente prima della generazione dei documenti

11. **Anagrafica Cliente**
   - **File da modificare:**
     - Moduli di gestione anagrafica clienti
   - **Problemi da risolvere:**
     - Richiesta non approvata: Anagrafica Cliente UNICA (attualmente divisa tra Business e Privati)
     - Nel caso di nuovo cliente professionale, il radio button per tipologia (ente, ecc.) deve diventare il valore predefinito
   - **Note tecniche:**
     - Modificare la gestione dei radio button per i clienti professionali

12. **Ricerca Articoli**
   - **File da modificare:**
     - `/web/amministrazione/Infoschede/ArticoliSeleziona.asp`
   - **Problemi da risolvere:** 
     - Nell'anagrafica articoli attivare la ricerca su tutti i campi
     - Aggiungere pulsante [vedi tutti]
   - **Note tecniche:**
     - Estendere la funzione di ricerca per includere tutti i campi del database
     - Aggiungere pulsante per visualizzare risultati completi

13. **Categorie Accessori**
   - **File da modificare:**
     - Modulo di gestione delle categorie accessori
   - **Problemi da risolvere:**
     - Ridurre la lista delle categorie in cui sono suddivisi gli accessori
   - **Note tecniche:**
     - Analizzare le categorie esistenti e consolidare quelle simili
     - Aggiornare riferimenti nel database dopo la riduzione

14. **Ordinamento Trasportatori**
   - **File da modificare:**
     - Modulo di gestione trasportatori
   - **Problemi da risolvere:**
     - La lista dei trasportatori va ordinata per “utilizzo” i più usati vanno in testa
     - L’uso è determinato dall’associazione nei ddt (elenco trasportatori) scheda Clienti.asp PROFILO=trasportatori
     - Fare visibile non visibile nella anagrafica
   - **Note tecniche:**
     - Implementare ordinamento dinamico basato sull'utilizzo dei trasportatori

15. **Export Conteggi Fine Mese**
   - **File da modificare:**
     - Moduli di export dei conteggi fine mese
   - **Problemi da risolvere:**
     - Nell’export .xls dei conteggi fine mese, per i privati espicitare i costi del trasporto e manodopera
     - Nell’export .xls dei conteggi fine mese, per i privati evidenziare i casi in cui (Nr – costi trasporto) <> (Nr – costi trasporto finale)
   - **Note tecniche:**
     - Modificare la logica di export per includere i costi aggiuntivi per i clienti privati
     - Implementare evidenziazione per le differenze nei costi di trasporto

16. **Sostituzione Logo**
   - **File da modificare:**
     - File di immagine del logo
   - **Problemi da risolvere:**
     - Sostituire logo 20 anni con 35 anni
   - **Note tecniche:**
     - Sostituire il file di immagine del logo con la nuova versione
poi export per schede


## Problemi Tecnici da Risolvere

### Errore Permessi wkhtmltopdf.exe

**Errore rilevato:** Quando si tenta di inviare un preventivo dalla scheda SchedaMod, appare il seguente errore:
```
Access to the path 'c:\inetpub\wwwroot\infoschede.it\web\App_Data\wkhtmltopdf\wkhtmltopdf.exe' is denied.
```

**File coinvolti:**
- `c:\inetpub\wwwroot\infoschede.it\web\Plugin\InviaEmail.ascx.cs` (linea 267)
- `c:\inetpub\wwwroot\infoschede.it\web\App_Data\wkhtmltopdf\wkhtmltopdf.exe`

**Analisi tecnica dettagliata:**

1. Dal codice sorgente in `Plugin/InviaEmail.ascx.cs`, riga 268, vediamo che la generazione del PDF avviene chiamando `NextPdf.GetPdfFromPageUrl()`
2. Questa funzione appartiene alla libreria `NextPdfTools.dll` trovata nella cartella `Bin`
3. Internamente, questa libreria utilizza il componente `NReco.PdfGenerator` che richiama l'eseguibile wkhtmltopdf.exe
4. L'errore si verifica perché l'utente che esegue il processo ASP.NET non ha permessi di accesso a questo file eseguibile

**Causa del problema:**
In base al web.config esaminato, l'applicazione non specifica un'identità personalizzata tramite il tag <identity>, quindi viene utilizzata l'identità del pool applicativo IIS configurato per il sito. In questo caso, i pool applicativi sono:
- `IIS AppPool\infoschede.it`
- `IIS AppPool\DefaultAppPool`

Entrambi questi account necessitano di permessi di accesso al file `wkhtmltopdf.exe` e alla sua directory.

**Soluzione:**
1. Accedere al server tramite Remote Desktop (Jump Desktop)
2. Navigare alla cartella `c:\inetpub\wwwroot\infoschede.it\web\App_Data\wkhtmltopdf\`
3. Fare click con il tasto destro sul file `wkhtmltopdf.exe` e selezionare "Properties"
4. Andare alla scheda "Security"
5. Fare click su "Edit" e poi su "Add"
6. Nella finestra "Select Users, Computers, Service Accounts, or Groups":
   - Fare click su "Locations" e selezionare il server locale
   - Per cercare gli account IIS, inserire i nomi esatti come segue:
     * Per i pool applicativi usati da questo sito, digitare **esattamente**:
       - `IIS AppPool\infoschede.it`
       - `IIS AppPool\DefaultAppPool`
     * Per il servizio di rete: digitare esattamente `Network Service`
     * Per l'utente anonimo di IIS: digitare esattamente `IUSR` 
     * Per il gruppo di utenti IIS: digitare esattamente `IIS_IUSRS`
   - Dopo aver digitato il nome, fare click su "Check Names" per verificare che sia riconosciuto
   - Se non viene trovato, provare a fare click su "Advanced" → "Find Now" e cercare l'account nell'elenco

7. Per verificare il nome esatto del pool applicativo utilizzato:
   - Aprire IIS Manager (Start → Administrative Tools → Internet Information Services (IIS) Manager)
   - Espandere il server e fare click su "Application Pools"
   - Identificare quale pool viene utilizzato dall'applicazione infoschede.it

8. Una volta aggiunti gli account, assegnare a ciascuno i permessi di "Read & execute" e "Read"
9. Cliccare su "Apply" e poi su "OK"
10. Per applicare i permessi quando non è possibile aggiungere direttamente i pool IIS, usa uno dei seguenti metodi alternativi:

    **Metodo alternativo 1 - Utilizzare ICACLS da Command Prompt con privilegi di amministratore:**
    - Aprire Command Prompt come amministratore (tasto destro su Command Prompt → "Run as administrator")
    - Eseguire i seguenti comandi uno alla volta:
      ```
      icacls "c:\inetpub\wwwroot\infoschede.it\web\App_Data\wkhtmltopdf\*" /grant "IIS AppPool\infoschede.it":(OI)(CI)(RX) /T
      icacls "c:\inetpub\wwwroot\infoschede.it\web\App_Data\wkhtmltopdf\*" /grant "IIS AppPool\DefaultAppPool":(OI)(CI)(RX) /T
      icacls "c:\inetpub\wwwroot\infoschede.it\web\App_Data\wkhtmltopdf\*" /grant "Network Service":(OI)(CI)(RX) /T
      icacls "c:\inetpub\wwwroot\infoschede.it\web\App_Data\wkhtmltopdf\*" /grant "IUSR":(OI)(CI)(RX) /T
      icacls "c:\inetpub\wwwroot\infoschede.it\web\App_Data\wkhtmltopdf\*" /grant "IIS_IUSRS":(OI)(CI)(RX) /T
      ```
      (RX = Read and Execute permissions, /T = recursivo, OI = Object Inherit, CI = Container Inherit)

    **Metodo alternativo 2 - Utilizzare l'identità del processo applicativo:**
    - Identificare l'utente che esegue effettivamente il processo IIS (generalmente "NETWORK SERVICE" o "ApplicationPoolIdentity")
    - Assegnare i permessi a questo utente invece che ai pool specifici:
      ```
      icacls "c:\inetpub\wwwroot\infoschede.it\web\App_Data\wkhtmltopdf\*" /grant "NETWORK SERVICE":(OI)(CI)(RX) /T
      ```

    **Metodo alternativo 3 - Configurare l'application pool per utilizzare un account specifico:**
    - Aprire IIS Manager (Start → Administrative Tools → Internet Information Services (IIS) Manager)
    - Fare click su "Application Pools" e selezionare il pool "infoschede.it"
    - Fare click con il tasto destro e selezionare "Advanced Settings"
    - Sotto "Process Model", cambiare "Identity" in un account che abbia già i permessi necessari (come "NETWORK SERVICE" o un account amministrativo)

**Note aggiuntive:**
Questo errore è tipico quando l'applicazione viene aggiornata o spostata senza configurare correttamente i permessi per i file di sistema. Assicurarsi che tutti gli eseguibili utilizzati dall'applicazione abbiano i permessi corretti.
16.         Nell’export .xls dei conteggi fine mese, per i privati evidenziare i casi in cui 
(Nr – costi trasporto)   <> (Nr – costi trasporto finale) 

export per ricambi manda solo i ricambi espicitare i costo costi del trasporto e manodopera , 

17.          sostituire logo 20 anni con 35 anni

mia considerazione)   da gestione spedizione alla causale voce libere che mostra campo testo libero - ddt
in ArticoloSeleziona aggiung  (spedizioni - apri ddt - aggiungi articolo pulsante, se articolo non c’è devo poterlo inserire) 
lui va in scheda assistenza, aggiungere il ricambio quindi scheda ArticoloSelezione ha il pulsante aggiungi ricambio deve comparire in ddt

 
URL ACCESSO APPLICATIVO  INFOSCHEDE 
 http://www.infoschede.it/it/  (utente esterno) ut = admin pw = admin    
 
url login Amministrazione backend :
http://www.infoschede.it/amministrazione/default.asp ut = HIDROSERVICES pw = ATHENA69





***************  hidroservices.com 

migrato su register CODICE MIGRAZIONE hidroservices.com is: 1_!wr[1b

***************** hidroservices.it CODICE MIGRAZIONE  		hidroservices.it     N4SgsEIZ-Ixazn74T

*****************  infoschede.it CODICE MIGRAZIONE  	infoschede.it  		 FMPL1nhr-iotEs3Wl

********************************** VPS windows  ****************** 
su Register Infoschede accesso con jump desktop

IP = 136.144.247.246 
Windows = 
	ut= Administrator
	pw= Hdro31!


------------------------------
BUTTONS:
MSSQL
Arma - Arm@25
sa - Arm@25
hidro - Arm@25


￼

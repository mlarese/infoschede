# Analisi Sistema di Gestione Listini

## Struttura del Database

### Tabella Principale: `gtb_listini`
La tabella principale che gestisce i listini ha la seguente struttura:

```sql
CREATE TABLE dbo.gtb_listini (
    listino_id int IDENTITY (1, 1) NOT NULL,  -- ID univoco del listino
    listino_codice nvarchar(50) NULL,         -- Codice identificativo del listino
    listino_datacreazione smalldatetime NULL, -- Data di creazione del listino
    listino_datascadenza smalldatetime NULL,  -- Data di scadenza del listino
    listino_B2C bit NULL,                    -- Flag per listini B2C
    listino_offerte bit NULL,                -- Flag per listini offerte
    listino_base bit NULL,                   -- Flag per listino base
    listino_base_attuale bit NULL,           -- Flag per listino base attuale
    listino_ancestor_id int NULL,            -- Riferimento al listino genitore (per ereditarietà)
    listino_with_child bit NOT NULL,        -- Indica se ha figli
    listino_note ntext NULL,                 -- Note aggiuntive
    listino_Dataimport datetime NULL          -- Data ultimo import
)
```

### Tabelle Collegate
1. `gtb_prezzi` - Contiene i prezzi degli articoli per ogni listino
2. `grel_art_valori` - Contiene le varianti degli articoli
3. `gtb_articoli` - Tabella principale degli articoli
4. `gtb_marche` - Marche degli articoli
5. `gtb_tipologie` - Categorie degli articoli

## Flusso di Importazione Listini

I listini vengono importati tramite script ASP nella cartella `/web/amministrazione/Infoschede/` con prefisso `_import_ricambi_`. Ogni fornitore ha il proprio script di importazione.

### File Principali di Importazione:
- `_import_ricambi_comet_2014.asp`
- `_import_ricambi_karcher_2014.asp`
- `_import_ricambi_lavor_2013.asp`
- `_import_ricambi_lavor_2014.asp`
- `_import_ricambi_valex_2014.asp`

## Gestione Listini (NextB2B)

### File Principali:
1. `/web/amministrazione/nextB2B/ListiniMod.asp` - Interfaccia di modifica listini
2. `/web/amministrazione/nextB2B/ListiniSalva.asp` - Logica di salvataggio listini
3. `/web/amministrazione/nextB2B/Tools_B2B.asp` - Funzioni di utilità per la gestione prezzi

### Funzioni Principali:
1. `AggiornaPrezzoListini` - Aggiorna i prezzi per una variante articolo
2. `AggiornaPrezzoListiniBase` - Aggiorna i prezzi dei listini base
3. `AggiornaPrezzoListiniDaListinoBase` - Aggiorna i prezzi dei listini derivati

## Problemi Attuali e Soluzioni Proposte

### 1. Duplicazione dei Listini
**Problema**: I nuovi listini vengono aggiunti senza rimuovere i vecchi, causando duplicati.

**Soluzione**:
1. Aggiungere un campo `attivo` (bit) alla tabella `gtb_listini`
2. Modificare la procedura di importazione per:
   - Disattivare i listini esistenti dello stesso fornitore
   - Importare il nuovo listino come attivo
   - Aggiornare i riferimenti al vecchio listino

### 2. Gestione Articoli Obsoleti
**Problema**: Gli articoli non più presenti nei nuovi listini rimangono nel sistema.

**Soluzione**:
1. Aggiungere un campo `disattivato` (bit) alla tabella `gtb_articoli`
2. Durante l'importazione:
   - Contrassegnare come disattivati gli articoli non più presenti
   - Non mostrare gli articoli disattivati nelle nuove schede
   - Mantenere gli articoli disattivati per le schede esistenti

### 3. Query di Selezione
Modificare le viste e le query per escludere gli articoli disattivati:

```sql
-- Esempio di modifica alla vista gv_articoli
SELECT * FROM gtb_articoli 
WHERE ISNULL(art_disabilitato, 0) = 0 
AND ISNULL(art_disattivato, 0) = 0  -- Nuova condizione
```

## Modifiche Dettagliate ai File

### 1. Aggiunta Campi al Database
**File:** `/web/amministrazione/library/database/Update__library__b2b.asp` (circa riga 350)
```sql
-- Aggiungere dopo la creazione della tabella gtb_listini
ALTER TABLE gtb_listini ADD 
    listino_attivo bit NOT NULL DEFAULT 1,
    data_ultima_modifica datetime NULL,
    utente_ultima_modifica int NULL;

-- Aggiungere dopo la creazione della tabella gtb_articoli
ALTER TABLE gtb_articoli ADD 
    art_disattivato bit NOT NULL DEFAULT 0,
    data_disattivazione datetime NULL,
    motivo_disattivazione nvarchar(255) NULL;
```

### 2. Modifica Procedure di Importazione

#### File: `/web/amministrazione/Infoschede/_import_ricambi_comet_2014.asp`
1. **Riga ~50** (dopo la connessione al DB):
   ```asp
   ' Disattiva listini esistenti dello stesso fornitore
   sql = "UPDATE gtb_listini SET listino_attivo = 0, data_ultima_modifica = GETDATE() " & _
         "WHERE listino_codice LIKE 'COMET%'"
   conn.Execute(sql)
   ```

2. **Riga ~120** (dopo l'importazione dei dati):
   ```asp
   ' Imposta il nuovo listino come attivo
   sql = "UPDATE gtb_listini SET listino_attivo = 1, data_ultima_modifica = GETDATE(), " & _
         "utente_ultima_modifica = " & Session("ID_Admin") & " " & _
         "WHERE listino_id = SCOPE_IDENTITY()"
   conn.Execute(sql)
   ```

#### File: `/web/amministrazione/nextB2B/ListiniSalva.asp`
1. **Riga ~80** (prima del salvataggio):
   ```asp
   ' Aggiorna data modifica e utente
   rs("data_ultima_modifica") = Now()
   rs("utente_ultima_modifica") = Session("ID_Admin")
   ```

### 3. Modifica Viste e Query

#### File: `/web/amministrazione/library/database/Update__library__b2b.asp`
1. **Riga ~1015** (modifica vista gv_articoli):
   ```sql
   CREATE VIEW dbo.gv_articoli AS
   SELECT * FROM dbo.gtb_articoli 
   WHERE ISNULL(art_disabilitato, 0) = 0 
   AND ISNULL(art_disattivato, 0) = 0
   ```

2. **Riga ~1260** (modifica vista gv_listini):
   ```sql
   CREATE VIEW dbo.gv_listini AS
   SELECT * FROM gtb_listini 
   WHERE listino_attivo = 1
   ```

### 4. Aggiunta Log delle Modifiche
**File:** `/web/amministrazione/nextB2B/ListiniSalva.asp` (in cima al file)
```asp
' Log delle modifiche
Sub LogModificaListino(listinoId, azione, dettagli)
    Dim sql
    sql = "INSERT INTO log_listini (listino_id, data_operazione, utente_id, azione, dettagli) " & _
          "VALUES (" & listinoId & ", GETDATE(), " & Session("ID_Admin") & ", '" & azione & "', '" & dettagli & "')"
    conn.Execute(sql)
End Sub
```

## Procedure di Manutenzione

### 1. Pulizia Articoli Disattivati
**File:** `/web/amministrazione/nextB2B/ManutenzioneListini.asp` (nuovo file)
```asp
' Disattiva articoli non presenti nell'ultimo import
Sub DisattivaArticoliObsoleti(fornitore)
    ' Disattiva articoli non più presenti nell'ultimo listino
    sql = "UPDATE a SET art_disattivato = 1, data_disattivazione = GETDATE(), " & _
          "motivo_disattivazione = 'Rimosso dall''ultimo listino' " & _
          "FROM gtb_articoli a " & _
          "INNER JOIN grel_art_valori v ON a.art_id = v.rel_art_id " & _
          "WHERE a.art_marca_id = (SELECT mar_id FROM gtb_marche WHERE mar_nome = '" & fornitore & "') " & _
          "AND v.rel_id NOT IN (SELECT DISTINCT prz_variante_id FROM gtb_prezzi p " & _
          "INNER JOIN gtb_listini l ON p.prz_listino_id = l.listino_id " & _
          "WHERE l.listino_attivo = 1)"
    conn.Execute(sql)
End Sub
```

### 2. Report Listini
**File:** `/web/amministrazione/nextB2B/ReportListini.asp` (nuovo file)
```asp
' Genera report dei listini attivi e delle modifiche recenti
Sub GeneraReportListini()
    ' Query per ottenere lo stato dei listini
    sql = "SELECT listino_id, listino_codice, listino_datacreazione, " & _
          "listino_datascadenza, data_ultima_modifica, " & _
          "(SELECT nome FROM tb_admin WHERE ID_admin = utente_ultima_modifica) as utente_modifica " & _
          "FROM gtb_listini " & _
          "WHERE listino_attivo = 1 " & _
          "ORDER BY data_ultima_modifica DESC"
    ' Esegui query e genera report
    ' ...
End Sub
```

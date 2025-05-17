# Gestione Listini

## Problema
- Attualmente i listini vanno in AGGIUNTA - modificare per eliminare i vecchi e mantenere SOLO i nuovi
- Per gli articoli usati nelle vecchie schede, implementare cancellazione logica (non fisica)
- Rendere gli articoli vecchi non visibili per le nuove schede

## Files da modificare
- Database: tabelle relative ai listini
- File di import dei listini

## Soluzione

### Prima
Attualmente, quando vengono importati nuovi listini, questi vengono aggiunti al database senza eliminare quelli vecchi, causando duplicati e confusione:

```sql
-- Esempio procedura di importazione attuale
INSERT INTO ArticoliRicambi (CodiceArticolo, Descrizione, Prezzo, ...)
SELECT CodiceArticolo, Descrizione, Prezzo, ...
FROM TempImportazioneListini
```

### Dopo
Modificare il sistema di gestione listini per:
1. Aggiungere un campo "Attivo" nella tabella ArticoliRicambi per implementare la cancellazione logica
2. Durante l'import, disattivare tutti gli articoli vecchi e attivare solo quelli nuovi
3. Modificare tutte le query di selezione per filtrare solo gli articoli attivi

```sql
-- Aggiungere campo Attivo alla tabella ArticoliRicambi
ALTER TABLE ArticoliRicambi ADD Attivo bit NOT NULL DEFAULT 1;

-- Procedura di importazione aggiornata
BEGIN TRANSACTION;

-- Disattiva tutti gli articoli esistenti
UPDATE ArticoliRicambi SET Attivo = 0;

-- Inserisci nuovi articoli o aggiorna quelli esistenti
MERGE ArticoliRicambi AS target
USING TempImportazioneListini AS source
ON target.CodiceArticolo = source.CodiceArticolo
WHEN MATCHED THEN
    UPDATE SET target.Descrizione = source.Descrizione,
               target.Prezzo = source.Prezzo,
               target.Attivo = 1,
               target.DataModifica = GETDATE()
WHEN NOT MATCHED THEN
    INSERT (CodiceArticolo, Descrizione, Prezzo, Attivo, DataInserimento)
    VALUES (source.CodiceArticolo, source.Descrizione, source.Prezzo, 1, GETDATE());

COMMIT TRANSACTION;
```

Modificare tutte le query di selezione degli articoli per includere il filtro sul campo Attivo:

```vb
' Esempio di modifica a ArticoliSeleziona.asp
strSQL = "SELECT * FROM ArticoliRicambi WHERE Attivo = 1 AND 1=1 "
' Eventuali filtri aggiuntivi...
strSQL = strSQL & " ORDER BY Descrizione"
```

## Spiegazione
Il problema attuale è che i nuovi listini vengono aggiunti senza eliminare quelli vecchi, causando duplicati. Gli articoli vecchi che sono stati utilizzati nelle schede esistenti non possono essere eliminati fisicamente per mantenere l'integrità dei dati storici.

La soluzione proposta implementa una "cancellazione logica" tramite un campo "Attivo" nella tabella ArticoliRicambi. Durante l'importazione di nuovi listini:
1. Tutti gli articoli esistenti vengono contrassegnati come inattivi (Attivo = 0)
2. I nuovi articoli vengono inseriti o aggiornati e contrassegnati come attivi (Attivo = 1)
3. Tutte le query di selezione vengono modificate per filtrare solo gli articoli attivi

Questa soluzione garantisce che:
- Gli articoli vecchi non siano visibili per le nuove schede
- I riferimenti agli articoli vecchi nelle schede esistenti rimangano intatti
- Non ci siano duplicati quando si visualizzano gli articoli disponibili

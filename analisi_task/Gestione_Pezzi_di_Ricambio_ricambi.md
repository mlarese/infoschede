# Gestione Pezzi di Ricambio

## Problema
- Eliminare i doppioni nei listini mantenendo solo il più recente
- Nella scheda "Aggiungi Ricambio" si apre un popup con duplicati

## Files da modificare
- `/web/amministrazione/Infoschede/ArticoliSeleziona.asp` - Gestione selezione articoli
- `/web/amministrazione/Infoschede/SchedeMod.asp` - Integrazione con articoli

## Soluzione

### Prima
Attualmente, la query SQL in ArticoliSeleziona.asp seleziona tutti gli articoli disponibili, inclusi eventuali duplicati:

```vb
' Esempio di codice problematico in ArticoliSeleziona.asp
strSQL = "SELECT * FROM ArticoliRicambi WHERE 1=1 "
' Eventuali filtri aggiuntivi...
strSQL = strSQL & " ORDER BY Descrizione"
```

Questo causa la visualizzazione di duplicati nel popup di selezione ricambi.

### Dopo
Modificare la query SQL per selezionare solo l'ultima versione di ogni articolo, eliminando i duplicati:

```vb
' Codice corretto in ArticoliSeleziona.asp
strSQL = "SELECT a.* FROM ArticoliRicambi a " & _
         "INNER JOIN (SELECT CodiceArticolo, MAX(DataInserimento) as MaxData " & _
         "FROM ArticoliRicambi GROUP BY CodiceArticolo) b " & _
         "ON a.CodiceArticolo = b.CodiceArticolo AND a.DataInserimento = b.MaxData " & _
         "WHERE 1=1 "
' Eventuali filtri aggiuntivi...
strSQL = strSQL & " ORDER BY a.Descrizione"
```

In SchedeMod.asp, assicurarsi che il popup utilizzi questa query filtrata:

```vb
' Codice in SchedeMod.asp per il popup di selezione articoli
Function PopupSelezionaArticoli()
    ' Richiama la pagina ArticoliSeleziona.asp con i parametri appropriati
    ' Questo garantisce che vengano mostrati solo articoli unici (senza duplicati)
    Response.Write "<script>window.open('ArticoliSeleziona.asp?...);</script>"
End Function
```

## Spiegazione
Il problema è causato da una query SQL in ArticoliSeleziona.asp che seleziona tutti gli articoli senza filtrare i duplicati. Quando si importano nuovi listini, gli articoli esistenti vengono duplicati anziché aggiornati.

La soluzione consiste nel modificare la query per selezionare solo la versione più recente di ogni articolo, utilizzando:
1. Una subquery che trova la data di inserimento più recente per ciascun codice articolo
2. Un JOIN con la tabella principale per ottenere i dettagli completi solo degli articoli più recenti

Questa modifica garantirà che nel popup "Aggiungi Ricambio" vengano visualizzati solo gli articoli unici, migliorando l'esperienza utente e riducendo gli errori di selezione articoli duplicati.

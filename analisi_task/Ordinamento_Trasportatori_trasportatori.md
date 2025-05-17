# Ordinamento Trasportatori

## Problema
- La lista dei trasportatori va ordinata per "utilizzo" i più usati vanno in testa
- L'uso è determinato dall'associazione nei ddt (elenco trasportatori) scheda Clienti.asp PROFILO=trasportatori
- Fare visibile non visibile nella anagrafica

## Files da modificare
- Modulo di gestione trasportatori

## Soluzione

### Prima
Attualmente, la lista dei trasportatori è ordinata in modo statico, senza considerare la frequenza di utilizzo:

```vb
' Esempio di codice attuale per elencare i trasportatori
strSQL = "SELECT * FROM Trasportatori ORDER BY Nome"
```

Inoltre, manca un campo per gestire la visibilità dei trasportatori.

### Dopo
Modificare la struttura della tabella Trasportatori aggiungendo campi per la visibilità e il conteggio dell'utilizzo:

```sql
-- Aggiungere campo Visibile e NumeroUtilizzi alla tabella Trasportatori
ALTER TABLE Trasportatori ADD Visibile bit NOT NULL DEFAULT 1;
ALTER TABLE Trasportatori ADD NumeroUtilizzi int NOT NULL DEFAULT 0;

-- Aggiornare il conteggio dell'utilizzo basato sui DDT
UPDATE t
SET t.NumeroUtilizzi = ISNULL(count_table.CountUse, 0)
FROM Trasportatori t
LEFT JOIN (
    SELECT IDTrasportatore, COUNT(*) AS CountUse
    FROM DDT
    GROUP BY IDTrasportatore
) count_table ON t.ID = count_table.IDTrasportatore;
```

Modificare le query per ordinare i trasportatori per utilizzo, mostrando solo quelli visibili:

```vb
' Query modificata per elencare i trasportatori
strSQL = "SELECT * FROM Trasportatori WHERE Visibile = 1 ORDER BY NumeroUtilizzi DESC, Nome"
```

Aggiungere la gestione della visibilità nell'interfaccia di amministrazione:

```vb
' Modifica all'interfaccia di gestione dei trasportatori
<tr>
  <td>Nome:</td>
  <td><input type="text" name="Nome" value="<%= rs("Nome") %>"></td>
</tr>
<tr>
  <td>Indirizzo:</td>
  <td><input type="text" name="Indirizzo" value="<%= rs("Indirizzo") %>"></td>
</tr>
<!-- Altri campi... -->
<tr>
  <td>Visibile:</td>
  <td>
    <input type="checkbox" name="Visibile" value="1" <% If rs("Visibile") = True Then Response.Write "checked" %>>
    Mostra questo trasportatore nelle liste di selezione
  </td>
</tr>
<tr>
  <td>Utilizzo:</td>
  <td>
    Questo trasportatore è stato utilizzato <strong><%= rs("NumeroUtilizzi") %></strong> volte.
  </td>
</tr>
```

Aggiornare la procedura di salvataggio per gestire il campo Visibile:

```vb
' In ClienteSalva.asp (quando PROFILO=trasportatori)
strVisibile = "0"
If Request.Form("Visibile") = "1" Then strVisibile = "1"

' Aggiungere il campo Visibile alla query di aggiornamento
strSQL = "UPDATE Trasportatori SET " & _
         "Nome = '" & Replace(Request.Form("Nome"), "'", "''") & "', " & _
         "Indirizzo = '" & Replace(Request.Form("Indirizzo"), "'", "''") & "', " & _
         "Visibile = " & strVisibile & " " & _
         "WHERE ID = " & intID
```

Implementare un aggiornamento periodico (o trigger) per mantenere aggiornato il campo NumeroUtilizzi:

```vb
' Esempio di procedura di aggiornamento da eseguire periodicamente
Sub AggiornaContatoriTrasportatori()
    Conn.Execute("UPDATE t " & _
                 "SET t.NumeroUtilizzi = ISNULL(count_table.CountUse, 0) " & _
                 "FROM Trasportatori t " & _
                 "LEFT JOIN ( " & _
                 "  SELECT IDTrasportatore, COUNT(*) AS CountUse " & _
                 "  FROM DDT " & _
                 "  GROUP BY IDTrasportatore " & _
                 ") count_table ON t.ID = count_table.IDTrasportatore")
End Sub
```

## Spiegazione
L'implementazione introduce tre miglioramenti chiave:

1. **Conteggio dell'utilizzo dei trasportatori**:
   - Aggiunta di un campo `NumeroUtilizzi` che tiene traccia di quante volte ciascun trasportatore è stato utilizzato nei DDT
   - Implementazione di un meccanismo di aggiornamento di questo contatore (periodico o attraverso trigger)
   - Ordinamento della lista dei trasportatori in base a questo conteggio, mostrando i più utilizzati in cima

2. **Gestione della visibilità**:
   - Aggiunta di un campo `Visibile` alla tabella dei trasportatori
   - Implementazione di checkbox nell'interfaccia di amministrazione per attivare/disattivare la visibilità
   - Filtro nelle query per mostrare solo i trasportatori marcati come visibili nelle liste di selezione

3. **Miglioramenti all'interfaccia utente**:
   - Visualizzazione del conteggio di utilizzo nell'interfaccia di amministrazione
   - Ordinamento intuitivo con i trasportatori più utilizzati in cima
   - Possibilità di nascondere trasportatori obsoleti senza eliminarli dal database

Questi miglioramenti rendono la gestione dei trasportatori più efficiente:
- Gli utenti trovano più rapidamente i trasportatori che usano frequentemente
- I trasportatori obsoleti o raramente utilizzati possono essere nascosti senza perdere i riferimenti storici
- Il sistema mantiene automaticamente statistiche sull'utilizzo dei trasportatori

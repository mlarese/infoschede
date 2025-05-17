# Esito dell'Operazione

## Problema
- Aggiungere un campo TXT ad inserimento libero nella sezione "Esito dell'operazione"

## Files da modificare
- `/web/amministrazione/Infoschede/SchedeMod.asp` - Interfaccia utente
- `/web/amministrazione/Infoschede/SchedeSalva.asp` - Persistenza dati
- Database: aggiungere campo nella tabella apposita

## Soluzione

### Prima
Attualmente, la sezione "Esito dell'operazione" non include un campo di testo libero per inserire dettagli aggiuntivi:

```vb
' Esempio di codice in SchedeMod.asp
<tr>
  <td>Esito dell'operazione:</td>
  <td>
    <select name="EsitoOperazione" id="EsitoOperazione">
      <option value="">Seleziona...</option>
      <% 
        ' Query per elencare gli esiti possibili
        Set rs = Conn.Execute("SELECT * FROM EsitiOperazione ORDER BY Descrizione")
        While Not rs.EOF
          Response.Write "<option value=""" & rs("ID") & """>" & rs("Descrizione") & "</option>"
          rs.MoveNext
        Wend
        rs.Close
      %>
    </select>
  </td>
</tr>
```

### Dopo
Modificare la struttura del database per aggiungere un nuovo campo per i dettagli dell'esito:

```sql
-- Aggiunta del campo DettagliEsito alla tabella Schede
ALTER TABLE Schede ADD DettagliEsito nvarchar(MAX) NULL;
```

Modificare l'interfaccia utente per includere il campo di testo libero:

```vb
' Codice modificato in SchedeMod.asp
<tr>
  <td>Esito dell'operazione:</td>
  <td>
    <select name="EsitoOperazione" id="EsitoOperazione">
      <option value="">Seleziona...</option>
      <% 
        ' Query per elencare gli esiti possibili
        Set rs = Conn.Execute("SELECT * FROM EsitiOperazione ORDER BY Descrizione")
        While Not rs.EOF
          Response.Write "<option value=""" & rs("ID") & """>" & rs("Descrizione") & "</option>"
          rs.MoveNext
        Wend
        rs.Close
      %>
    </select>
  </td>
</tr>
<tr>
  <td>Dettagli esito:</td>
  <td>
    <textarea name="DettagliEsito" id="DettagliEsito" rows="4" cols="50"><%= rsScheda("DettagliEsito") %></textarea>
  </td>
</tr>
```

Aggiornare SchedeSalva.asp per salvare il nuovo campo:

```vb
' In SchedeSalva.asp, aggiungere il salvataggio dei dettagli dell'esito
strDettagliEsito = Request.Form("DettagliEsito")
If strDettagliEsito <> "" Then
  strSQL = strSQL & ", DettagliEsito = '" & Replace(strDettagliEsito, "'", "''") & "'"
Else
  strSQL = strSQL & ", DettagliEsito = NULL"
End If
```

Aggiornare anche i template dei report e delle email per includere il nuovo campo:

```vb
' Esempio di aggiornamento di un template di report
strReport = strReport & "<p><strong>Esito dell'operazione:</strong> " & GetEsitoOperazione(rsScheda("IDEsitoOperazione")) & "</p>"
If Not IsNull(rsScheda("DettagliEsito")) And rsScheda("DettagliEsito") <> "" Then
  strReport = strReport & "<p><strong>Dettagli esito:</strong> " & rsScheda("DettagliEsito") & "</p>"
End If
```

## Spiegazione
L'implementazione richiede:

1. **Modifica del database**:
   - Aggiunta di un nuovo campo `DettagliEsito` alla tabella `Schede` per memorizzare i dettagli testuali dell'esito dell'operazione
   - Il campo è di tipo `nvarchar(MAX)` per supportare testi di lunghezza variabile e caratteri speciali

2. **Modifica dell'interfaccia utente**:
   - Aggiunta di un campo `textarea` sotto il dropdown dell'esito dell'operazione
   - Il campo consente l'inserimento di testo libero multiplo con dimensioni adeguate (4 righe, 50 colonne)

3. **Aggiornamento della logica di salvataggio**:
   - Modifica di `SchedeSalva.asp` per raccogliere e salvare il nuovo campo
   - Gestione corretta delle virgolette singole per evitare errori SQL
   - Impostazione del campo a NULL se non viene inserito nessun testo

4. **Aggiornamento dei template**:
   - Modifica dei template di report e delle email per includere il nuovo campo quando presente
   - Presentazione ordinata con formattazione appropriata

Questa implementazione fornisce agli utenti la possibilità di inserire dettagli testuali liberi sull'esito dell'operazione, migliorando la qualità delle informazioni registrate e la comunicazione con i clienti.

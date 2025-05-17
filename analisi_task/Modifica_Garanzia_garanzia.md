# Modifica Garanzia

## Problema
- Sostituire la voce "Garanzia" con "VALUTAZIONE DEL CENTRO ASSISTENZA"
- Vedi punto 1 per l'invio mail

## Files da modificare
- `/web/amministrazione/Infoschede/SchedeMod.asp` - Interfaccia utente
- `/web/amministrazione/Infoschede/SchedeSalva.asp` - Persistenza dati

## Soluzione

### Prima
Attualmente, nel form è presente la voce "Garanzia":

```vb
' Esempio di codice in SchedeMod.asp
<tr>
  <td>Garanzia:</td>
  <td>
    <select name="Garanzia" id="Garanzia">
      <option value="0" <% If CInt(rsScheda("Garanzia")) = 0 Then Response.Write "selected" %>>No</option>
      <option value="1" <% If CInt(rsScheda("Garanzia")) = 1 Then Response.Write "selected" %>>Sì</option>
    </select>
  </td>
</tr>
```

### Dopo
Modificare l'etichetta da "Garanzia" a "VALUTAZIONE DEL CENTRO ASSISTENZA" mantenendo inalterato il comportamento funzionale:

```vb
' Codice modificato in SchedeMod.asp
<tr>
  <td>VALUTAZIONE DEL CENTRO ASSISTENZA:</td>
  <td>
    <select name="Garanzia" id="Garanzia">
      <option value="0" <% If CInt(rsScheda("Garanzia")) = 0 Then Response.Write "selected" %>>No</option>
      <option value="1" <% If CInt(rsScheda("Garanzia")) = 1 Then Response.Write "selected" %>>Sì</option>
    </select>
  </td>
</tr>
```

Modificare anche eventuali riferimenti alla voce "Garanzia" in altri file dell'applicazione, come nei template di email e nei report:

```vb
' Esempio di template email in Class_Messages_CommonParts.asp
Function GetTemplateEmail(iScheda)
    ' ...
    ' Sostituire "Garanzia: " con "VALUTAZIONE DEL CENTRO ASSISTENZA: "
    strTemplate = Replace(strTemplate, "Garanzia: ", "VALUTAZIONE DEL CENTRO ASSISTENZA: ")
    ' ...
End Function
```

## Spiegazione
La modifica è principalmente cosmetica e richiede solo la sostituzione dell'etichetta "Garanzia" con "VALUTAZIONE DEL CENTRO ASSISTENZA" nell'interfaccia utente. 

È importante notare che:
1. Il nome del campo nel database e nelle variabili del codice rimane invariato ("Garanzia")
2. La modifica riguarda solo il testo visualizzato all'utente
3. Occorre verificare che la modifica sia applicata coerentemente in tutti i punti in cui appare la voce "Garanzia", compresi:
   - Form di inserimento/modifica
   - Template di email
   - Report e documenti generati

Poiché questa modifica riguarda anche l'invio delle email, è necessario assicurarsi che i template delle email siano aggiornati per riflettere la nuova terminologia. Questo garantirà che i messaggi inviati agli utenti utilizzino la dicitura "VALUTAZIONE DEL CENTRO ASSISTENZA" anziché "Garanzia".

# Scheda Nuovo Lavoro

## Problema
- Riferimento al punto 1 (gestione email)
- Inserire il Codice Trasportatore (creare anagrafica dei Trasportatori)
- Alla voce "Accessori" aggiungere in coda "Inserisci nuovo accessorio"
- Il testo va aggiunto alla tabella degli accessori

## Files da modificare
- `/web/amministrazione/Infoschede/SchedeMod.asp` - Form principale
- `/web/amministrazione/Infoschede/SchedeSalva.asp` - Salvataggio dati

## Soluzione

### Prima
Attualmente, la scheda non prevede l'inserimento del codice trasportatore e non offre la possibilità di aggiungere nuovi accessori direttamente dal form:

```vb
' Esempio di sezione Accessori in SchedeMod.asp
<tr>
  <td>Accessori:</td>
  <td>
    <select name="Accessori" id="Accessori">
      <option value="">Seleziona...</option>
      <% 
        ' Query per elencare gli accessori esistenti
        Set rs = Conn.Execute("SELECT * FROM Accessori ORDER BY Descrizione")
        While Not rs.EOF
          Response.Write "<option value=""" & rs("ID") & """>" & rs("Descrizione") & "</option>"
          rs.MoveNext
        Wend
        rs.Close
      %>
    </select>
  </td>
</tr>

' Manca la gestione del codice trasportatore
```

### Dopo
Modificare il form per includere il codice trasportatore e la possibilità di aggiungere nuovi accessori:

```vb
' Aggiunta della sezione Trasportatore in SchedeMod.asp
<tr>
  <td>Trasportatore:</td>
  <td>
    <select name="Trasportatore" id="Trasportatore">
      <option value="">Seleziona...</option>
      <% 
        ' Query per elencare i trasportatori
        Set rs = Conn.Execute("SELECT * FROM Trasportatori WHERE Visibile = 1 ORDER BY NumeroUtilizzi DESC")
        While Not rs.EOF
          Response.Write "<option value=""" & rs("ID") & """>" & rs("Nome") & "</option>"
          rs.MoveNext
        Wend
        rs.Close
      %>
    </select>
  </td>
</tr>

' Modifica della sezione Accessori
<tr>
  <td>Accessori:</td>
  <td>
    <select name="Accessori" id="Accessori" onchange="gestioneAccessori()">
      <option value="">Seleziona...</option>
      <% 
        ' Query per elencare gli accessori esistenti
        Set rs = Conn.Execute("SELECT * FROM Accessori ORDER BY Descrizione")
        While Not rs.EOF
          Response.Write "<option value=""" & rs("ID") & """>" & rs("Descrizione") & "</option>"
          rs.MoveNext
        Wend
        rs.Close
      %>
      <option value="nuovo">Inserisci nuovo accessorio</option>
    </select>
    <div id="nuovoAccessorio" style="display:none">
      <input type="text" name="NuovoAccessorio" id="NuovoAccessorio" placeholder="Descrizione nuovo accessorio">
      <button type="button" onclick="salvaAccessorio()">Salva</button>
    </div>
  </td>
</tr>

<script>
  function gestioneAccessori() {
    var sel = document.getElementById('Accessori');
    var div = document.getElementById('nuovoAccessorio');
    if (sel.value === 'nuovo') {
      div.style.display = 'block';
    } else {
      div.style.display = 'none';
    }
  }
  
  function salvaAccessorio() {
    var desc = document.getElementById('NuovoAccessorio').value;
    if (desc === '') {
      alert('Inserire una descrizione per il nuovo accessorio');
      return;
    }
    
    // Invia richiesta AJAX per salvare il nuovo accessorio
    // ...
    
    // Aggiorna la dropdown degli accessori
    // ...
  }
</script>
```

Modificare SchedeSalva.asp per gestire il salvataggio del codice trasportatore e dei nuovi accessori:

```vb
' In SchedeSalva.asp, aggiungere il salvataggio del trasportatore
strTrasportatore = Request.Form("Trasportatore")
If strTrasportatore <> "" Then
  strSQL = strSQL & ", IDTrasportatore = " & strTrasportatore
End If

' Gestione del nuovo accessorio
strNuovoAccessorio = Request.Form("NuovoAccessorio")
If strNuovoAccessorio <> "" Then
  ' Inserisci il nuovo accessorio nella tabella
  Conn.Execute("INSERT INTO Accessori (Descrizione) VALUES ('" & Replace(strNuovoAccessorio, "'", "''") & "')")
  ' Recupera l'ID del nuovo accessorio
  Set rs = Conn.Execute("SELECT @@IDENTITY AS ID")
  strAccessorio = rs("ID")
  rs.Close
  ' Aggiorna la scheda con il nuovo accessorio
  strSQL = strSQL & ", IDAccessorio = " & strAccessorio
End If
```

## Spiegazione
Le modifiche implementano:

1. **Gestione Trasportatori**:
   - Creazione di una nuova tabella "Trasportatori" se non esiste già
   - Aggiunta di un campo dropdown per selezionare il trasportatore nella scheda
   - Salvataggio dell'ID del trasportatore selezionato nel database

2. **Gestione Nuovi Accessori**:
   - Aggiunta di un'opzione "Inserisci nuovo accessorio" alla fine della dropdown degli accessori
   - Implementazione di un campo di testo e un pulsante per salvare il nuovo accessorio
   - Salvataggio del nuovo accessorio nella tabella "Accessori"
   - Aggiornamento della scheda con il nuovo accessorio

Queste modifiche consentono di migliorare la funzionalità della scheda nuovo lavoro, permettendo agli utenti di selezionare un trasportatore dall'anagrafica e di aggiungere nuovi accessori direttamente dal form, senza dover passare per una gestione separata.

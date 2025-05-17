# Inserimento Ore Manodopera

## Problema
- Rendere obbligatorio l'inserimento ORE MANODOPERA (con costo orario) nella scheda

## Files da modificare
- `/web/amministrazione/Infoschede/SchedeMod.asp` - Validazione input
- `/web/amministrazione/Infoschede/SchedeSalva.asp` - Controllo dati

## Soluzione

### Prima
Attualmente, l'inserimento delle ore di manodopera non è obbligatorio e manca la validazione appropriata:

```vb
' Esempio di codice in SchedeMod.asp
<tr>
  <td>Ore Manodopera:</td>
  <td>
    <input type="text" name="OreManodopera" id="OreManodopera" value="<%= rsScheda("OreManodopera") %>">
  </td>
</tr>
<tr>
  <td>Costo Orario:</td>
  <td>
    <input type="text" name="CostoOrario" id="CostoOrario" value="<%= rsScheda("CostoOrario") %>">
  </td>
</tr>

' Nel JavaScript di validazione, non c'è controllo per questi campi
function ValidaForm() {
  // Validazione di altri campi...
  return true;
}
```

In SchedeSalva.asp, non c'è controllo sull'obbligatorietà dei campi:

```vb
' Esempio di codice in SchedeSalva.asp
strOreManodopera = Request.Form("OreManodopera")
strCostoOrario = Request.Form("CostoOrario")

' Salvataggio senza controlli di obbligatorietà
If strOreManodopera <> "" Then
  strSQL = strSQL & ", OreManodopera = " & Replace(strOreManodopera, ",", ".")
End If

If strCostoOrario <> "" Then
  strSQL = strSQL & ", CostoOrario = " & Replace(strCostoOrario, ",", ".")
End If
```

### Dopo
Modificare l'interfaccia utente per indicare che i campi sono obbligatori e aggiungere la validazione:

```vb
' Codice modificato in SchedeMod.asp
<tr>
  <td>Ore Manodopera: <span style="color: red;">*</span></td>
  <td>
    <input type="text" name="OreManodopera" id="OreManodopera" value="<%= rsScheda("OreManodopera") %>" required>
  </td>
</tr>
<tr>
  <td>Costo Orario: <span style="color: red;">*</span></td>
  <td>
    <input type="text" name="CostoOrario" id="CostoOrario" value="<%= rsScheda("CostoOrario") %>" required>
  </td>
</tr>

' Aggiungere la validazione JavaScript
<script>
function ValidaForm() {
  // Validazione di altri campi...
  
  // Validazione Ore Manodopera
  var oreManodopera = document.getElementById("OreManodopera").value;
  if (oreManodopera == "") {
    alert("Il campo 'Ore Manodopera' è obbligatorio.");
    document.getElementById("OreManodopera").focus();
    return false;
  }
  
  // Validazione Costo Orario
  var costoOrario = document.getElementById("CostoOrario").value;
  if (costoOrario == "") {
    alert("Il campo 'Costo Orario' è obbligatorio.");
    document.getElementById("CostoOrario").focus();
    return false;
  }
  
  // Verifica che i valori siano numerici
  if (isNaN(parseFloat(oreManodopera.replace(",", ".")))) {
    alert("Il valore inserito per 'Ore Manodopera' non è valido. Inserire un numero.");
    document.getElementById("OreManodopera").focus();
    return false;
  }
  
  if (isNaN(parseFloat(costoOrario.replace(",", ".")))) {
    alert("Il valore inserito per 'Costo Orario' non è valido. Inserire un numero.");
    document.getElementById("CostoOrario").focus();
    return false;
  }
  
  return true;
}
</script>
```

Modificare SchedeSalva.asp per aggiungere il controllo dell'obbligatorietà anche lato server:

```vb
' Codice modificato in SchedeSalva.asp
strOreManodopera = Request.Form("OreManodopera")
strCostoOrario = Request.Form("CostoOrario")

' Controllo obbligatorietà
If strOreManodopera = "" Then
  Response.Write "<script>alert('Il campo \"Ore Manodopera\" è obbligatorio.'); history.back();</script>"
  Response.End
End If

If strCostoOrario = "" Then
  Response.Write "<script>alert('Il campo \"Costo Orario\" è obbligatorio.'); history.back();</script>"
  Response.End
End If

' Controllo validità numerica
If Not IsNumeric(Replace(strOreManodopera, ",", ".")) Then
  Response.Write "<script>alert('Il valore inserito per \"Ore Manodopera\" non è valido. Inserire un numero.'); history.back();</script>"
  Response.End
End If

If Not IsNumeric(Replace(strCostoOrario, ",", ".")) Then
  Response.Write "<script>alert('Il valore inserito per \"Costo Orario\" non è valido. Inserire un numero.'); history.back();</script>"
  Response.End
End If

' Salvataggio con valori obbligatori
strSQL = strSQL & ", OreManodopera = " & Replace(strOreManodopera, ",", ".")
strSQL = strSQL & ", CostoOrario = " & Replace(strCostoOrario, ",", ".")
```

## Spiegazione
L'implementazione rende obbligatorio l'inserimento delle ore di manodopera e del costo orario attraverso:

1. **Modifiche all'interfaccia utente**:
   - Aggiunta di indicatori visivi (asterischi rossi) per segnalare i campi obbligatori
   - Aggiunta dell'attributo HTML5 "required" per la validazione nativa del browser
   - Mantenimento dei valori inseriti in caso di errore durante la validazione

2. **Validazione client-side**:
   - Controllo dell'obbligatorietà dei campi tramite JavaScript
   - Verifica che i valori inseriti siano numerici validi
   - Messaggi di errore chiari con focus automatico sul campo problematico

3. **Validazione server-side**:
   - Controllo dell'obbligatorietà dei campi lato server per garantire la sicurezza
   - Verifica della validità numerica dei valori
   - Reindirizzamento alla pagina del form con un messaggio di errore in caso di problemi

4. **Gestione dei dati**:
   - Sostituzione della virgola con il punto nei valori numerici per la corretta elaborazione SQL
   - Salvataggio garantito di entrambi i campi nel database

Questa implementazione assicura che ogni scheda includa sempre le informazioni relative alle ore di manodopera e al relativo costo orario, migliorando la completezza e l'accuratezza dei dati.

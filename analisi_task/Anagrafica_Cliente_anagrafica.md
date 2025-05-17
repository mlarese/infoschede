# Anagrafica Cliente

## Problema
- Richiesta non approvata: Anagrafica Cliente UNICA (attualmente divisa tra Business e Privati)
- Nel caso di nuovo cliente professionale, il radio button per tipologia (ente, ecc.) deve diventare il valore predefinito

## Files da modificare
- Moduli di gestione anagrafica clienti

## Soluzione

### Prima
Attualmente, nel form di inserimento di un nuovo cliente professionale, il radio button per la tipologia non viene preselezionato:

```vb
' Esempio di codice attuale per i radio button delle tipologie
<tr>
  <td>Tipologia:</td>
  <td>
    <input type="radio" name="Tipologia" id="TipologiaEnte" value="Ente"> Ente<br>
    <input type="radio" name="Tipologia" id="TipologiaAzienda" value="Azienda"> Azienda<br>
    <input type="radio" name="Tipologia" id="TipologiaAltro" value="Altro"> Altro
  </td>
</tr>
```

### Dopo
Modificare il form per preselezionare automaticamente il radio button più appropriato per i clienti professionali:

```vb
' Codice modificato per i radio button delle tipologie
<tr>
  <td>Tipologia:</td>
  <td>
    <% 
    ' Determina la tipologia predefinita
    Dim defaultTipologia
    If Request.QueryString("tipo") = "professionale" Then
      defaultTipologia = "Azienda"  ' Imposta "Azienda" come valore predefinito per i clienti professionali
    Else
      defaultTipologia = ""
    End If
    %>
    <input type="radio" name="Tipologia" id="TipologiaEnte" value="Ente" <% If defaultTipologia = "Ente" Then Response.Write "checked" %>> Ente<br>
    <input type="radio" name="Tipologia" id="TipologiaAzienda" value="Azienda" <% If defaultTipologia = "Azienda" Then Response.Write "checked" %>> Azienda<br>
    <input type="radio" name="Tipologia" id="TipologiaAltro" value="Altro" <% If defaultTipologia = "Altro" Then Response.Write "checked" %>> Altro
  </td>
</tr>
```

È importante notare che la richiesta di unificare l'anagrafica cliente (attualmente divisa tra Business e Privati) non è stata approvata, quindi manterremo la divisione esistente.

## Spiegazione
L'implementazione si concentra esclusivamente sulla preselazione del radio button per la tipologia dei clienti professionali:

1. **Determinazione del valore predefinito**:
   - Il codice verifica se il cliente è di tipo professionale controllando il parametro "tipo" nella query string
   - Se il cliente è professionale, imposta "Azienda" come tipologia predefinita
   - Questo valore può essere modificato in base alle esigenze specifiche dell'applicazione

2. **Preselazione del radio button**:
   - Il codice aggiunge l'attributo "checked" al radio button appropriato in base al valore predefinito
   - L'utente può comunque selezionare un'altra tipologia se necessario

3. **Mantenimento della separazione delle anagrafiche**:
   - Come specificato, la richiesta di unificare le anagrafiche non è stata approvata
   - Pertanto, il sistema continuerà a mantenere separate le anagrafiche per clienti business e privati

Questa modifica migliora l'usabilità del form di inserimento clienti professionali preselezionando automaticamente l'opzione più comune, riducendo i click necessari per l'utente e migliorando l'efficienza del processo di inserimento dati.

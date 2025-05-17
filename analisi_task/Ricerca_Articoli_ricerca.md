# Ricerca Articoli

## Problema
- Nell'anagrafica articoli attivare la ricerca su tutti i campi
- Aggiungere pulsante [vedi tutti]

## Files da modificare
- `/web/amministrazione/Infoschede/ArticoliSeleziona.asp`

## Soluzione

### Prima
Attualmente, la ricerca degli articoli è limitata solo ad alcuni campi specifici:

```vb
' Esempio di codice attuale in ArticoliSeleziona.asp
strSearch = Request.Form("txtSearch")
If strSearch <> "" Then
    strSQL = strSQL & " AND (CodiceArticolo LIKE '%" & strSearch & "%' OR Descrizione LIKE '%" & strSearch & "%')"
End If
```

Inoltre, non esiste un pulsante per visualizzare tutti i risultati.

### Dopo
Modificare la ricerca per includere tutti i campi rilevanti e aggiungere un pulsante "Vedi tutti":

```vb
' Codice modificato in ArticoliSeleziona.asp

' Aggiunta del pulsante "Vedi tutti" accanto al campo di ricerca
<form method="post" action="ArticoliSeleziona.asp" id="frmSearch">
    <input type="hidden" name="showAll" id="showAll" value="0">
    <div class="row">
        <div class="col-md-8">
            <div class="input-group">
                <input type="text" name="txtSearch" id="txtSearch" class="form-control" value="<%= strSearch %>" placeholder="Cerca...">
                <div class="input-group-append">
                    <button type="submit" class="btn btn-primary">Cerca</button>
                    <button type="button" class="btn btn-secondary" onclick="vediTutti()">Vedi tutti</button>
                </div>
            </div>
        </div>
    </div>
</form>

<script>
function vediTutti() {
    document.getElementById('showAll').value = "1";
    document.getElementById('txtSearch').value = "";
    document.getElementById('frmSearch').submit();
}
</script>

' Modifica della query di ricerca per incluere tutti i campi rilevanti
<%
strSearch = Request.Form("txtSearch")
strShowAll = Request.Form("showAll")

If strShowAll = "1" Then
    ' Non applicare filtri di ricerca
Else
    If strSearch <> "" Then
        strSQL = strSQL & " AND (CodiceArticolo LIKE '%" & strSearch & "%' " & _
                          "OR Descrizione LIKE '%" & strSearch & "%' " & _
                          "OR CodiceFabbrica LIKE '%" & strSearch & "%' " & _
                          "OR CodiceFornitore LIKE '%" & strSearch & "%' " & _
                          "OR CodiceOrigine LIKE '%" & strSearch & "%' " & _
                          "OR Categoria LIKE '%" & strSearch & "%' " & _
                          "OR Note LIKE '%" & strSearch & "%')"
    End If
End If
%>
```

## Spiegazione
L'implementazione migliora la funzionalità di ricerca degli articoli in due modi principali:

1. **Estensione della ricerca a tutti i campi rilevanti**:
   - Oltre a CodiceArticolo e Descrizione, la ricerca ora include:
     - CodiceFabbrica
     - CodiceFornitore
     - CodiceOrigine
     - Categoria
     - Note
   - Questo permette di trovare articoli anche quando l'utente ricorda solo informazioni parziali o alternative

2. **Aggiunta del pulsante "Vedi tutti"**:
   - Il nuovo pulsante "Vedi tutti" consente di visualizzare l'elenco completo di articoli senza filtri
   - Implementato tramite JavaScript che imposta un flag "showAll" e invia il form
   - Quando il flag "showAll" è attivo, la query SQL non applica filtri di ricerca

Il layout dell'interfaccia è stato migliorato con:
- Un campo di ricerca più ampio e visibile
- Pulsanti chiaramente etichettati e distinti per le diverse azioni
- Un design responsive che funziona bene su dispositivi di diverse dimensioni

Queste modifiche migliorano significativamente l'usabilità della funzione di ricerca articoli, rendendo più facile per gli utenti trovare gli articoli di cui hanno bisogno, anche quando non ricordano il codice o la descrizione esatta.

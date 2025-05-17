# Spese Trasporto

## Problema
- Sul documento al cliente, riorganizzare le voci "Spese Trasporto Presa" e "Spese Trasporto Consegna"
- Metterle in colonna per migliorare leggibilità e comprensione
- Aggiungere campi per manodopera, costo presa e costo riconsegna

## Files da modificare
- Template di stampa per il documento inviato al cliente
- `/web/amministrazione/Infoschede/Reports/` - Cartella dei report

## Soluzione

### Prima
Attualmente, il layout del documento inviato al cliente presenta le spese di trasporto in modo poco leggibile:

```html
<!-- Esempio di codice nel template attuale -->
<table>
  <tr>
    <td>Spese Trasporto Presa: <%= FormatCurrency(rsScheda("CostoPresa")) %></td>
    <td>Spese Trasporto Consegna: <%= FormatCurrency(rsScheda("CostoRiconsegna")) %></td>
  </tr>
  <!-- Altre informazioni... -->
</table>
```

### Dopo
Riorganizzare il layout per migliorare la leggibilità, mettendo le voci in colonna:

```html
<!-- Codice migliorato nel template -->
<table border="0" cellspacing="0" cellpadding="3" width="100%">
  <tr>
    <th align="left" width="70%">Descrizione</th>
    <th align="right" width="30%">Importo</th>
  </tr>
  
  <!-- Manodopera -->
  <tr>
    <td>Manodopera (<%= FormatNumber(rsScheda("OreManodopera"), 2) %> ore x <%= FormatCurrency(rsScheda("CostoOrario")) %>)</td>
    <td align="right"><%= FormatCurrency(rsScheda("OreManodopera") * rsScheda("CostoOrario")) %></td>
  </tr>
  
  <!-- Spese Trasporto Presa -->
  <tr>
    <td>Spese Trasporto Presa</td>
    <td align="right"><%= FormatCurrency(rsScheda("CostoPresa")) %></td>
  </tr>
  
  <!-- Spese Trasporto Consegna -->
  <tr>
    <td>Spese Trasporto Consegna</td>
    <td align="right"><%= FormatCurrency(rsScheda("CostoRiconsegna")) %></td>
  </tr>
  
  <!-- Parti di ricambio, se presenti -->
  <% If HasSpareParts Then %>
  <tr>
    <td>Parti di ricambio (dettagli in allegato)</td>
    <td align="right"><%= FormatCurrency(rsScheda("TotaleRicambi")) %></td>
  </tr>
  <% End If %>
  
  <!-- Totale -->
  <tr style="font-weight: bold; border-top: 1px solid #000;">
    <td>Totale</td>
    <td align="right"><%= FormatCurrency(rsScheda("OreManodopera") * rsScheda("CostoOrario") + rsScheda("CostoPresa") + rsScheda("CostoRiconsegna") + rsScheda("TotaleRicambi")) %></td>
  </tr>
</table>
```

## Spiegazione
L'implementazione migliora la leggibilità e la comprensione del documento inviato al cliente attraverso:

1. **Riorganizzazione del layout**:
   - Creazione di una tabella con due colonne chiare: "Descrizione" e "Importo"
   - Posizionamento delle voci in righe separate per una migliore organizzazione visiva
   - Allineamento dei valori monetari a destra per facilitare la lettura

2. **Aggiunta di dettagli sulla manodopera**:
   - Visualizzazione del numero di ore e del costo orario
   - Calcolo automatico del costo totale della manodopera

3. **Presentazione chiara delle spese di trasporto**:
   - "Spese Trasporto Presa" in una riga dedicata
   - "Spese Trasporto Consegna" in una riga separata
   - Visualizzazione chiara degli importi

4. **Miglioramento dell'aspetto generale**:
   - Uso di spaziature appropriate tra le celle
   - Separazione visiva tra le diverse sezioni del documento
   - Evidenziazione del totale con formattazione in grassetto e linea superiore

5. **Adattamento alla scheda 28682**:
   - Utilizzo della scheda 28682 come riferimento per la struttura e il formato
   - Mantenimento della coerenza con le altre modifiche implementate

Questa riorganizzazione rende il documento più professionale e più facile da comprendere per il cliente, migliorando la trasparenza riguardo ai costi addebitati.

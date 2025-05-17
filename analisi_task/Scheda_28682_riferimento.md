# Scheda 28682

## Problema
- Verificare l'implementazione della manodopera, costo presa e costo riconsegna
- Utilizzare questa scheda come modello per le modifiche da implementare

## Files da modificare
- Tutti i file relativi all'implementazione delle schede

## Soluzione

### Prima
Non esiste un problema "prima" in quanto la scheda 28682 è il modello di riferimento per l'implementazione corretta. La verifica consiste nell'analizzare questa scheda esistente per replicarne la struttura e il comportamento nelle nuove implementazioni.

### Dopo
Ogni nuova implementazione dovrà seguire il modello della scheda 28682, in particolare per quanto riguarda:
- La gestione della manodopera
- Il calcolo del costo presa
- Il calcolo del costo riconsegna

Il codice dovrebbe seguire lo schema utilizzato nella scheda di riferimento:

```vb
' Esempio di gestione manodopera come nella scheda 28682
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

' Esempio di gestione costi trasporto come nella scheda 28682
<tr>
  <td>Costo Presa:</td>
  <td>
    <input type="text" name="CostoPresa" id="CostoPresa" value="<%= rsScheda("CostoPresa") %>">
  </td>
</tr>
<tr>
  <td>Costo Riconsegna:</td>
  <td>
    <input type="text" name="CostoRiconsegna" id="CostoRiconsegna" value="<%= rsScheda("CostoRiconsegna") %>">
  </td>
</tr>

' Calcolo dei totali
<script>
function calcolaTotali() {
  var oreManodopera = parseFloat(document.getElementById("OreManodopera").value.replace(",", ".")) || 0;
  var costoOrario = parseFloat(document.getElementById("CostoOrario").value.replace(",", ".")) || 0;
  var costoPresa = parseFloat(document.getElementById("CostoPresa").value.replace(",", ".")) || 0;
  var costoRiconsegna = parseFloat(document.getElementById("CostoRiconsegna").value.replace(",", ".")) || 0;
  
  var costoManodopera = oreManodopera * costoOrario;
  var totale = costoManodopera + costoPresa + costoRiconsegna;
  
  document.getElementById("CostoManodopera").value = costoManodopera.toFixed(2).replace(".", ",");
  document.getElementById("TotaleCosti").value = totale.toFixed(2).replace(".", ",");
}
</script>
```

## Spiegazione
La scheda 28682 rappresenta l'implementazione corretta e deve essere utilizzata come riferimento per standardizzare tutte le altre schede. Essa contiene una corretta implementazione di:

1. **Gestione della manodopera**:
   - Campi per l'inserimento delle ore di manodopera
   - Calcolo automatico del costo totale della manodopera

2. **Gestione dei costi di trasporto**:
   - Campo per il costo di presa
   - Campo per il costo di riconsegna

3. **Calcolo dei totali**:
   - Somma automatica di tutti i costi
   - Visualizzazione chiara del totale per il cliente

4. **Layout e presentazione**:
   - Organizzazione chiara delle informazioni
   - Formattazione coerente dei campi numerici

Utilizzare questo modello garantirà coerenza e standardizzazione in tutte le schede, migliorando l'esperienza utente e semplificando la manutenzione del codice. La scheda 28682 dovrebbe essere analizzata nel dettaglio per comprendere tutte le sue caratteristiche e replicarle nelle nuove implementazioni.

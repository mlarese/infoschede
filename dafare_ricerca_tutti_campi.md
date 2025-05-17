# Ricerca su Tutti i Campi degli Articoli

## 1. File da modificare
- `/web/amministrazione/nextB2B/ArticoliSeleziona.asp`

## 2. Modifiche da fare

### 2.1 Aggiunta pulsante "Vedi tutti"
1. Posizionare il pulsante nella sezione di ricerca
2. Il pulsante deve essere visibile solo quando non ci sono criteri di ricerca attivi
3. Il pulsante deve avere classe "button" per la coerenza con lo stile

### 2.2 Implementazione della funzionalità
1. Quando il pulsante "Vedi tutti" viene cliccato:
   - Pulire tutti i campi di ricerca
   - Reset della sessione di ricerca
   - Reindirizzare alla pagina stessa

### 2.3 Codice da aggiungere
```asp
<!-- Aggiungere nella sezione dei pulsanti di ricerca -->
<% if Session("SelArt_codice") = "" AND Session("SelArt_nome") = "" AND Session("SelArt_descrizione") = "" AND Session("SelArt_categoria") = "" AND Session("SelArt_fornitore") = "" AND Session("SelArt_prezzo") = "" AND Session("SelArt_giacenza") = "" then %>
    <input type="button" class="button" value="Vedi tutti" onclick="location.href='ArticoliSeleziona.asp'" />
<% end if %>
```

### 2.4 Note tecniche
1. Il pulsante deve essere posizionato vicino ai campi di ricerca
2. La funzionalità deve essere implementata in modo non invasivo
3. Il pulsante deve essere visibile solo quando non ci sono criteri di ricerca attivi
4. La pulizia dei campi di ricerca deve essere gestita correttamente

## 3. Test da Eseguire
1. Verificare che il pulsante "Vedi tutti" sia visibile solo quando non ci sono criteri di ricerca
2. Verificare che cliccando su "Vedi tutti" vengano mostrati tutti gli articoli
3. Verificare che la funzionalità non interferisca con la ricerca normale
4. Testare la pulizia dei campi di ricerca

## 4. Considerazioni
1. La funzionalità deve essere coerente con lo stile esistente
2. La pulizia dei campi di ricerca deve essere gestita correttamente
3. La funzionalità deve essere testata in ambiente di sviluppo prima di essere portata in produzione

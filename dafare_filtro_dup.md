# Gestione Duplicati e Accessori

## 1. Eliminazione Duplicati nei Listini

### File da modificare:
- `/web/amministrazione/Infoschede/ArticoliSeleziona.asp`
- `/web/amministrazione/Infoschede/SchedeMod.asp`

### Modifiche da fare:
- Verificare e modificare la query SQL in ArticoliSeleziona.asp
- Aggiungere filtro per escludere duplicati o mantenere solo il più recente
- Implementare la logica di selezione dell'articolo più recente quando ci sono duplicati

## 2. Gestione Accessori

### File da modificare:
- `/web/amministrazione/Infoschede/SchedeMod.asp`
- `/web/amministrazione/Infoschede/AccessoriNew.asp`
- `/web/amministrazione/Infoschede/AccessoriSalva.asp`

### Modifiche da fare:
1. In SchedeMod.asp:
   - Aggiungere pulsante "Inserisci nuovo accessorio" nella sezione degli accessori
   - Posizionare il pulsante dopo il dropdown degli accessori (riga 387)
   - Il pulsante deve puntare a AccessoriNew.asp

2. In AccessoriNew.asp:
   - Verificare che il form per la creazione di nuovi accessori sia corretto
   - Assicurarsi che i campi obbligatori siano validati

3. In AccessoriSalva.asp:
   - Verificare che i dati vengano salvati correttamente
   - Gestire eventuali errori di validazione

## 3. Tabella Accessori

### Azioni da fare:
- Verificare che la tabella `sgtb_accessori` esista nel database
- Controllare che tutti i campi necessari siano presenti
- Verificare i vincoli di integrità referenziale

## 4. Test da Eseguire

### Test da fare:
1. Testare la selezione degli articoli:
   - Verificare che non ci siano duplicati
   - Verificare che venga selezionato sempre l'articolo più recente

2. Testare la gestione degli accessori:
   - Creare un nuovo accessorio
   - Selezionare un accessorio esistente
   - Verificare che il pulsante "Inserisci nuovo accessorio" funzioni correttamente

### Note:
- Tutte le modifiche devono essere testate in ambiente di sviluppo prima di essere portate in produzione
- Documentare tutte le modifiche apportate
- Verificare l'impact sulle altre funzionalità del sistema

# Categorie Accessori

## Problema
- Ridurre la lista delle categorie in cui sono suddivisi gli accessori

## Files da modificare
- Modulo di gestione delle categorie accessori

## Soluzione

### Prima
Attualmente, la lista delle categorie degli accessori è troppo ampia, creando confusione e difficoltà nella selezione:

```sql
-- Esempio di categorie attuali (troppe e potenzialmente ridondanti)
SELECT * FROM CategorieAccessori ORDER BY Descrizione
/*
ID  Descrizione
1   Accessori Elettrici
2   Accessori Meccanici
3   Accessori Idraulici
4   Accessori per modello X
5   Accessori per modello Y
6   Accessori per modello Z
7   Accessori vari tipo A
8   Accessori vari tipo B
9   Accessori specifici
10  Accessori opzionali
...
e altri simili
*/
```

### Dopo
Consolidare le categorie simili per ridurre e semplificare la lista:

```sql
-- Aggiornamento della tabella delle categorie per consolidare quelle simili
-- Prima creare una tabella di mappatura temporanea
CREATE TABLE #MappingCategorie (
    IDVecchio int,
    IDNuovo int
)

-- Inserire le mappature (esempio)
INSERT INTO #MappingCategorie VALUES
(1, 1),    -- Accessori Elettrici rimane
(2, 2),    -- Accessori Meccanici rimane
(3, 3),    -- Accessori Idraulici rimane
(4, 7),    -- Accessori per modello X -> Accessori per modelli
(5, 7),    -- Accessori per modello Y -> Accessori per modelli
(6, 7),    -- Accessori per modello Z -> Accessori per modelli
(7, 8),    -- Accessori vari tipo A -> Accessori vari
(8, 8),    -- Accessori vari tipo B -> Accessori vari
(9, 9),    -- Accessori specifici rimane
(10, 9)    -- Accessori opzionali -> Accessori specifici

-- Aggiornare i riferimenti nelle tabelle collegate
UPDATE Accessori
SET IDCategoria = m.IDNuovo
FROM Accessori a
JOIN #MappingCategorie m ON a.IDCategoria = m.IDVecchio

-- Eliminare le categorie non più necessarie
DELETE FROM CategorieAccessori
WHERE ID IN (
    SELECT IDVecchio 
    FROM #MappingCategorie 
    WHERE IDVecchio <> IDNuovo
)

-- Rinominare alcune categorie per maggiore chiarezza
UPDATE CategorieAccessori
SET Descrizione = 'Accessori per modelli' 
WHERE ID = 7

UPDATE CategorieAccessori
SET Descrizione = 'Accessori vari' 
WHERE ID = 8

-- Pulizia
DROP TABLE #MappingCategorie
```

Aggiornare anche l'interfaccia utente per riflettere le categorie consolidate:

```vb
' Esempio di codice per visualizzare le categorie consolidate
<select name="Categoria" id="Categoria">
    <option value="">Seleziona una categoria...</option>
    <% 
    Set rs = Conn.Execute("SELECT * FROM CategorieAccessori ORDER BY Descrizione")
    While Not rs.EOF
        Response.Write "<option value=""" & rs("ID") & """>" & rs("Descrizione") & "</option>"
        rs.MoveNext
    Wend
    rs.Close
    %>
</select>
```

## Spiegazione
L'implementazione mira a semplificare e razionalizzare le categorie degli accessori attraverso:

1. **Analisi delle categorie esistenti**:
   - Identificazione di categorie ridondanti o troppo specifiche
   - Determinazione di un insieme più piccolo e coerente di categorie

2. **Consolidamento delle categorie**:
   - Creazione di una tabella di mappatura per tracciare quali categorie vecchie confluiscono in quali nuove
   - Aggiornamento dei riferimenti nelle tabelle collegate (come la tabella Accessori)
   - Eliminazione delle categorie non più necessarie
   - Rinominazione di alcune categorie per maggiore chiarezza

3. **Vantaggi dell'approccio**:
   - Il processo preserva l'integrità dei dati esistenti
   - Tutte le associazioni tra accessori e categorie vengono mantenute
   - L'interfaccia utente risulta semplificata con meno opzioni da gestire

Il risultato finale è una lista di categorie più concisa e gestibile, che migliora l'esperienza utente durante la selezione degli accessori senza perdere la capacità di classificare correttamente gli elementi.

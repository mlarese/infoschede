# Gestione Clienti di tipo Privato

## Problema
- Per Clienti di tipo privato NON va generato il DDT e lettera di trasporto

## Files da modificare
- Moduli di generazione DDT e lettera di trasporto

## Soluzione

### Prima
Attualmente, il sistema genera DDT e lettera di trasporto per tutti i clienti, indipendentemente dalla loro tipologia:

```vb
' Esempio di codice per la generazione DDT
Function GeneraDDT(idScheda)
    ' Recupera i dati della scheda
    Set rsScheda = Conn.Execute("SELECT * FROM Schede WHERE ID = " & idScheda)
    
    ' Genera DDT senza controllare il tipo di cliente
    ' ...
    
    ' Genera lettera di trasporto
    GeneraLetteraTrasporto(idScheda)
    
    rsScheda.Close
End Function
```

### Dopo
Modificare la funzione di generazione DDT e lettera di trasporto per controllare il tipo di cliente:

```vb
' Codice modificato per la generazione DDT
Function GeneraDDT(idScheda)
    ' Recupera i dati della scheda e del cliente
    Dim strSQL
    strSQL = "SELECT s.*, c.TipoCliente FROM Schede s " & _
             "LEFT JOIN Clienti c ON s.IDCliente = c.ID " & _
             "WHERE s.ID = " & idScheda
    
    Set rsScheda = Conn.Execute(strSQL)
    
    ' Controlla se il cliente è privato
    If Not rsScheda.EOF Then
        If rsScheda("TipoCliente") = "Privato" Then
            ' Per i clienti privati, non generare DDT né lettera di trasporto
            GeneraDDT = "Cliente privato, DDT non richiesto"
            rsScheda.Close
            Exit Function
        Else
            ' Per i clienti business, genera DDT normalmente
            ' ...
            
            ' Genera lettera di trasporto
            GeneraLetteraTrasporto(idScheda)
        End If
    End If
    
    rsScheda.Close
End Function

' Funzione per verificare se un cliente è privato
Function IsClientePrivato(idCliente)
    Dim rsCliente
    Set rsCliente = Conn.Execute("SELECT TipoCliente FROM Clienti WHERE ID = " & idCliente)
    
    If Not rsCliente.EOF Then
        IsClientePrivato = (rsCliente("TipoCliente") = "Privato")
    Else
        IsClientePrivato = False
    End If
    
    rsCliente.Close
End Function
```

Modificare anche l'interfaccia utente per informare l'operatore che per i clienti privati non verranno generati DDT e lettera di trasporto:

```vb
' In SchedeMod.asp, aggiungere un messaggio informativo
<% If IsClientePrivato(rsScheda("IDCliente")) Then %>
<div class="alert alert-info">
    <strong>Nota:</strong> Per i clienti privati non vengono generati DDT e lettera di trasporto.
</div>
<% End If %>
```

## Spiegazione
L'implementazione introduce una distinzione tra clienti privati e business per quanto riguarda la generazione di documenti di trasporto:

1. **Verifica del tipo di cliente**:
   - Prima di generare DDT e lettera di trasporto, il sistema controlla se il cliente è di tipo "Privato"
   - Viene creata una funzione `IsClientePrivato()` per semplificare questa verifica in tutto il codice

2. **Logica condizionale**:
   - Se il cliente è privato, la generazione di DDT e lettera di trasporto viene saltata
   - Se il cliente è business, i documenti vengono generati normalmente

3. **Feedback all'utente**:
   - L'interfaccia avvisa l'operatore che per i clienti privati non verranno generati documenti di trasporto
   - Questo evita confusione e chiarisce il comportamento del sistema

4. **Query ottimizzata**:
   - La query recupera in un'unica operazione sia i dati della scheda che il tipo di cliente
   - Questo riduce il numero di chiamate al database e migliora le prestazioni

Questa modifica garantisce che i documenti di trasporto vengano generati solo quando necessario, evitando confusione e semplificando il processo per i clienti privati.

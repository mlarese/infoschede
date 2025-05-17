# Export Conteggi Fine Mese

## Problema
- Nell'export .xls dei conteggi fine mese, per i privati espicitare i costi del trasporto e manodopera
- Nell'export .xls dei conteggi fine mese, per i privati evidenziare i casi in cui (Nr – costi trasporto) <> (Nr – costi trasporto finale)
- Export per ricambi manda solo i ricambi, esplicitare i costi del trasporto e manodopera

## Files da modificare
- Moduli di export dei conteggi fine mese

## Soluzione

### Prima
Attualmente, l'export dei conteggi fine mese non esplicita i costi di trasporto e manodopera per i clienti privati e non evidenzia le discrepanze nei costi di trasporto:

```vb
' Esempio di codice attuale per l'export
Sub ExportConteggiFineMese(dataDa, dataA)
    ' Preparazione intestazioni
    strOutput = "Numero Scheda" & vbTab & "Data" & vbTab & "Cliente" & vbTab & "Importo Totale" & vbTab & "Ricambi" & vbCrLf
    
    ' Query per recuperare i dati
    strSQL = "SELECT s.ID, s.Data, c.RagioneSociale, s.ImportoTotale, s.TotaleRicambi " & _
             "FROM Schede s " & _
             "INNER JOIN Clienti c ON s.IDCliente = c.ID " & _
             "WHERE s.Data BETWEEN '" & dataDa & "' AND '" & dataA & "' " & _
             "ORDER BY s.ID"
    
    Set rs = Conn.Execute(strSQL)
    
    ' Generazione delle righe dell'export
    While Not rs.EOF
        strOutput = strOutput & rs("ID") & vbTab & _
                    FormatDateTime(rs("Data"), 2) & vbTab & _
                    rs("RagioneSociale") & vbTab & _
                    FormatNumber(rs("ImportoTotale"), 2) & vbTab & _
                    FormatNumber(rs("TotaleRicambi"), 2) & vbCrLf
        rs.MoveNext
    Wend
    
    rs.Close
    
    ' Output del file
    Response.ContentType = "application/vnd.ms-excel"
    Response.AddHeader "Content-Disposition", "attachment; filename=ConteggiFineMese.xls"
    Response.Write strOutput
End Sub
```

### Dopo
Modificare l'export per includere i costi di trasporto e manodopera per i clienti privati e per evidenziare eventuali discrepanze:

```vb
' Codice migliorato per l'export
Sub ExportConteggiFineMese(dataDa, dataA)
    ' Preparazione intestazioni con nuove colonne
    strOutput = "Numero Scheda" & vbTab & "Data" & vbTab & "Cliente" & vbTab & "Tipo Cliente" & vbTab & _
                "Importo Totale" & vbTab & "Ricambi" & vbTab & "Manodopera" & vbTab & _
                "Costo Presa" & vbTab & "Costo Riconsegna" & vbTab & "Trasporto Finale" & vbTab & _
                "Discrepanza" & vbCrLf
    
    ' Query ampliata per recuperare tutti i dati necessari
    strSQL = "SELECT s.ID, s.Data, c.RagioneSociale, c.TipoCliente, " & _
             "s.ImportoTotale, s.TotaleRicambi, " & _
             "(s.OreManodopera * s.CostoOrario) AS CostoManodopera, " & _
             "s.CostoPresa, s.CostoRiconsegna, s.CostoTrasportoFinale " & _
             "FROM Schede s " & _
             "INNER JOIN Clienti c ON s.IDCliente = c.ID " & _
             "WHERE s.Data BETWEEN '" & dataDa & "' AND '" & dataA & "' " & _
             "ORDER BY s.ID"
    
    Set rs = Conn.Execute(strSQL)
    
    ' Generazione delle righe dell'export con logica migliorata
    While Not rs.EOF
        ' Calcolo della discrepanza nei costi di trasporto
        Dim trasportoOrdinario, trasportoFinale, discrepanza
        trasportoOrdinario = CDbl(rs("CostoPresa")) + CDbl(rs("CostoRiconsegna"))
        trasportoFinale = CDbl(rs("CostoTrasportoFinale"))
        discrepanza = trasportoFinale - trasportoOrdinario
        
        ' Formattazione base della riga
        strRiga = rs("ID") & vbTab & _
                 FormatDateTime(rs("Data"), 2) & vbTab & _
                 rs("RagioneSociale") & vbTab & _
                 rs("TipoCliente") & vbTab & _
                 FormatNumber(rs("ImportoTotale"), 2) & vbTab & _
                 FormatNumber(rs("TotaleRicambi"), 2) & vbTab
        
        ' Per i clienti privati, aggiungiamo dettagli su manodopera e trasporto
        If rs("TipoCliente") = "Privato" Then
            strRiga = strRiga & FormatNumber(rs("CostoManodopera"), 2) & vbTab & _
                      FormatNumber(rs("CostoPresa"), 2) & vbTab & _
                      FormatNumber(rs("CostoRiconsegna"), 2) & vbTab & _
                      FormatNumber(rs("CostoTrasportoFinale"), 2) & vbTab
            
            ' Evidenziare discrepanze nei costi di trasporto
            If Abs(discrepanza) > 0.01 Then  ' Tolleranza per errori di arrotondamento
                strRiga = strRiga & "DISCREPANZA: " & FormatNumber(discrepanza, 2)
            Else
                strRiga = strRiga & "OK"
            End If
        Else
            ' Per i clienti business, inserire solo valori basici
            strRiga = strRiga & FormatNumber(rs("CostoManodopera"), 2) & vbTab & _
                      "-" & vbTab & "-" & vbTab & "-" & vbTab & "-"
        End If
        
        strOutput = strOutput & strRiga & vbCrLf
        rs.MoveNext
    Wend
    
    rs.Close
    
    ' Output del file
    Response.ContentType = "application/vnd.ms-excel"
    Response.AddHeader "Content-Disposition", "attachment; filename=ConteggiFineMese.xls"
    Response.Write strOutput
End Sub

' Modifica anche per l'export dei ricambi
Sub ExportRicambiFineMese(dataDa, dataA)
    ' Preparazione intestazioni con nuove colonne
    strOutput = "Numero Scheda" & vbTab & "Data" & vbTab & "Cliente" & vbTab & _
                "Ricambi" & vbTab & "Manodopera" & vbTab & _
                "Costo Presa" & vbTab & "Costo Riconsegna" & vbTab & _
                "Totale" & vbCrLf
    
    ' Query per recuperare i dati
    strSQL = "SELECT s.ID, s.Data, c.RagioneSociale, " & _
             "s.TotaleRicambi, (s.OreManodopera * s.CostoOrario) AS CostoManodopera, " & _
             "s.CostoPresa, s.CostoRiconsegna, s.ImportoTotale " & _
             "FROM Schede s " & _
             "INNER JOIN Clienti c ON s.IDCliente = c.ID " & _
             "WHERE s.Data BETWEEN '" & dataDa & "' AND '" & dataA & "' " & _
             "ORDER BY s.ID"
    
    Set rs = Conn.Execute(strSQL)
    
    ' Generazione delle righe dell'export
    While Not rs.EOF
        strOutput = strOutput & rs("ID") & vbTab & _
                    FormatDateTime(rs("Data"), 2) & vbTab & _
                    rs("RagioneSociale") & vbTab & _
                    FormatNumber(rs("TotaleRicambi"), 2) & vbTab & _
                    FormatNumber(rs("CostoManodopera"), 2) & vbTab & _
                    FormatNumber(rs("CostoPresa"), 2) & vbTab & _
                    FormatNumber(rs("CostoRiconsegna"), 2) & vbTab & _
                    FormatNumber(rs("ImportoTotale"), 2) & vbCrLf
        rs.MoveNext
    Wend
    
    rs.Close
    
    ' Output del file
    Response.ContentType = "application/vnd.ms-excel"
    Response.AddHeader "Content-Disposition", "attachment; filename=RicambiFineMese.xls"
    Response.Write strOutput
End Sub
```

## Spiegazione
L'implementazione migliora gli export dei conteggi di fine mese in diversi modi:

1. **Export Conteggi Fine Mese per Clienti Privati**:
   - Aggiunta di colonne specifiche per visualizzare i costi di manodopera e trasporto
   - Inclusione dei campi "Costo Presa", "Costo Riconsegna" e "Trasporto Finale"
   - Calcolo e visualizzazione di eventuali discrepanze tra la somma dei costi di trasporto (presa + riconsegna) e il costo di trasporto finale

2. **Evidenziazione delle Discrepanze**:
   - Calcolo della differenza tra (CostoPresa + CostoRiconsegna) e CostoTrasportoFinale
   - Aggiunta di un'indicazione esplicita "DISCREPANZA" quando viene rilevata una differenza significativa
   - Inclusione del valore della discrepanza per facilitare l'analisi

3. **Export Ricambi con Dettagli Aggiuntivi**:
   - Modifica dell'export dei ricambi per includere anche i costi di manodopera e trasporto
   - Visualizzazione chiara di tutti i componenti che contribuiscono al totale
   - Mantenimento del totale generale per facilità di verifica

4. **Miglioramenti Generali**:
   - Aggiunta di una colonna "Tipo Cliente" per distinguere facilmente tra clienti privati e business
   - Ottimizzazione della query SQL per recuperare tutti i dati necessari in un'unica operazione
   - Gestione differenziata delle informazioni in base al tipo di cliente

Questi miglioramenti rendono gli export più completi e utili, facilitando la riconciliazione dei costi e l'identificazione di eventuali incongruenze.

# Gestione Email

## Problema
- Verificare funzionamento invio/ricezione mail da applicativo nella nuova scheda
- Risolvere errore di invio email dalla logistica (sezione SchedaMod)

## Files da modificare
- `/web/amministrazione/Infoschede/SchedeMod.asp` - Gestione schede e invio email
- `/web/amministrazione/library/Class_Messages_CommonParts.asp` - Classe per invio email

## Soluzione

### Prima
Il codice attuale presenta problemi di permessi per l'eseguibile wkhtmltopdf e le sue DLL, causando errori nell'invio delle email e nella generazione dei PDF.

```vb
' Esempio di codice problematico in SchedeMod.asp
If bSendMail Then
    ' Tentativo di invio email con possibili errori di permessi
    SendEmail(...)
End If

' Possibile codice in Class_Messages_CommonParts.asp con problemi
Public Function SendToContact(...)
    ' Codice che utilizza wkhtmltopdf con problemi di permessi
    Dim oShell
    Set oShell = Server.CreateObject("WScript.Shell")
    oShell.Run("C:\path\to\wkhtmltopdf.exe ...")
End Function
```

### Dopo
Risoluzione del problema di permessi per l'eseguibile wkhtmltopdf e le sue DLL, assicurando il corretto funzionamento dell'invio email e generazione PDF.

```vb
' Codice aggiornato in SchedeMod.asp
If bSendMail Then
    ' Invio email con permessi corretti
    SendEmail(...)
End If

' Codice corretto in Class_Messages_CommonParts.asp
Public Function SendToContact(...)
    ' Codice con percorsi e permessi corretti
    Dim oShell
    Set oShell = Server.CreateObject("WScript.Shell")
    ' Percorso corretto e con permessi appropriati
    oShell.Run("C:\CombiRoot\infoschede.it\wkhtmltopdf\wkhtmltopdf.exe ...")
End Function
```

## Spiegazione
Il problema era dovuto a permessi insufficienti o percorsi errati per l'eseguibile wkhtmltopdf e le sue DLL. Questo impediva la corretta generazione dei PDF e l'invio delle email.

Le funzioni di invio email sono definite nella classe `Class_Messages_CommonParts.asp` e includono:
- `SendToContact` - Per inviare email ai contatti
- `SendToAdmin` - Per inviare email agli amministratori
- `Save` - Per salvare le email

La soluzione consiste nel verificare e correggere:
1. I percorsi all'eseguibile wkhtmltopdf
2. I permessi di accesso al file e alle sue DLL
3. Eventuali errori nelle funzioni di invio email

Questo intervento garantisce il corretto funzionamento dell'invio delle email dalla logistica (sezione SchedaMod) e dalla nuova scheda.

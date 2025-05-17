# Struttura delle cartelle amministrazione e Amministrazione2

## Cartella "amministrazione"

La cartella amministrazione contiene numerosi file e sottocartelle per la gestione di un sistema web basato su ASP.

### File Principali

- **default.asp**: Pagina principale del sistema di amministrazione. Contiene query per caricare i siti disponibili per l'amministratore corrente attraverso la tabella `rel_admin_sito` e `tb_siti`.
- **[Plugin/Login.ascx](cci:7://file:///Users/maurolarese/Dropbox/arma/infoschede/web/Plugin/Login.ascx:0:0-0:0)**: Componente ASP.NET per la gestione dell'autenticazione utenti. Contiene il form di login con campi per username e password, e il pulsante di accesso.

### Sottocartelle Principali

#### /amministrazione/library
Contiene classi e funzioni di utilità utilizzate in tutto il sistema:

- **ClassSalva.asp**: Gestisce le operazioni di salvataggio dati con funzioni come `AddForcedValue`, `UpdateForcedFields` e `SetUpdateParams`
- **ClassDelete.asp**: Gestisce le operazioni di eliminazione dati
- **FileManager.asp**: Interfaccia per la gestione dei file con funzioni come `Open_Delete`, `Open_Upload`, `Open_Multi` e `Open_File`
- **PickerDate.asp**: Componente per la selezione delle date
- **Class_Messages_CommonParts.asp**: Gestisce l'invio di messaggi agli utenti e agli amministratori
- **Class_Messages_EmailsParts.asp**: Funzioni per la gestione delle email, inclusa la manipolazione del contenuto HTML

#### /amministrazione/Infoschede
Gestione delle schede informative:

- **RitiriMod.asp**: Gestione dei ritiri con funzioni di modifica e rimozione
- **AgentiIntGestione.asp**: Gestione degli agenti intermediari

#### /amministrazione/nextWeb5
Sistema di gestione siti web:

- **SitoDominioMod.asp**: Modifica domini dei siti
- **SitoTemplatePagine.asp**: Gestione template delle pagine
- **SitoStili.asp**: Gestione degli stili CSS
- **SitoPagineTemplate.asp**: Gestione dei template delle pagine
- **SitoPagineSalva.asp**: Salvataggio delle pagine
- **Delete.asp**: Sistema di eliminazione avanzato che gestisce le relazioni tra tabelle

#### /amministrazione/nextCom
Sistema di comunicazioni:

- **ComunicazioniNew_Wizard_1.asp**: Creazione di nuove comunicazioni con procedura guidata

#### /amministrazione/nextB2B
Sistema B2B per la gestione delle relazioni commerciali

#### /amministrazione/nextMemo2
Sistema di gestione memo e appunti

#### /amministrazione/nextPassport
Sistema di autenticazione e accessi

### Query SQL Principali

Nel sistema vengono utilizzate numerose query SQL, principalmente per:

1. **Selezione di contenuti**:
   ```sql
   SELECT * FROM tb_siti WHERE id_sito IN (...)
   ```

2. **Join tra tabelle**:
   ```sql
   SELECT * FROM rel_admin_sito INNER JOIN tb_siti ON rel_admin_sito.sito_id = tb_siti.id_sito
   ```

3. **Subquery per relazioni**:
   ```sql
   SELECT * FROM tb_pagineSito WHERE id_pagSTAGE_[lingua]=[id]
   ```

4. **Query con filtri complessi**:
   ```sql
   SELECT idx_id FROM tb_contents_index WHERE idx_id IN (SELECT rip_idx_id FROM rel_index_pubblicazioni WHERE rip_pub_id= [ID]) AND idx_id NOT IN (SELECT rip_idx_id FROM rel_index_pubblicazioni WHERE rip_pub_id<>[ID])
   ```

## Cartella "Amministrazione2"

La cartella Amministrazione2 sembra essere una versione più recente o parallela del sistema di amministrazione.

### File Principali

- **Default.aspx**: Pagina principale del sistema (in formato ASPX)
- **Calendario.aspx**: Gestione del calendario
- **Web.Config**: Configurazione del sito

### Pagina Principale

La pagina principale del sito è `Default.aspx`, che funge da punto di ingresso principale dell'applicazione. Questa pagina è responsabile dell'inizializzazione dell'interfaccia utente e della configurazione del layout base.

Caratteristiche principali:
- Utilizza il framework NextFramework
- È una pagina ASP.NET con codice server in C#
- Contiene il NextForm principale dell'applicazione
- Gestisce l'inizializzazione del layout e dei componenti base

La pagina principale è responsabile di:
1. L'inizializzazione dei dati della pagina
2. La configurazione del tag head
3. La configurazione del tag body
4. L'inizializzazione del form principale e dei suoi componenti

### Sottocartelle Principali

#### /Amministrazione2/NextMemo2
Una versione aggiornata del sistema di memo

#### /Amministrazione2/NextComment
Sistema di gestione dei commenti 

#### /Amministrazione2/FileManager
Gestione file in formato ASPX invece che ASP classico

#### /Amministrazione2/App_Themes
Temi e stili per l'interfaccia utente

#### /Amministrazione2/bin
Contiene le librerie compilate utilizzate dall'applicazione

## Caratteristiche principali del sistema

Il sistema sembra essere un CMS (Content Management System) completo che gestisce:

1. Siti web con supporto multilingua (EN, FR, DE, ES, RU, CN, PT)
2. Templates e pagine
3. Stili CSS
4. Comunicazioni
5. Gestione intermediari/agenti
6. Sistema di autenticazione
7. Gestione file
8. Sistema B2B

Il database è relazionale con numerose tabelle interconnesse come:
- tb_siti
- tb_webs
- tb_admin
- tb_pagineSito
- tb_pages
- tb_layers
- tb_css_groups
- tb_css_styles
- tb_contents
- tb_contents_index

Le query SQL presenti nel codice mostrano una struttura complessa con molte relazioni tra tabelle per mantenere l'integrità dei dati e supportare funzionalità come il multilingua e la versione dei contenuti.

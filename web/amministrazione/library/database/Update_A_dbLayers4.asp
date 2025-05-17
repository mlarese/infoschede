<!--#INCLUDE FILE="Update__FileHeader.asp" -->
<% '........................................................................................... %>
<!--#INCLUDE FILE="../Tools.asp" -->
<!--#INCLUDE FILE="../Tools4Admin.asp" -->
<!--#INCLUDE FILE="../ToolsDescrittori.asp" -->
<%

'*******************************************************************************************
'AGGIORNAMENTO 36 
'...........................................................................................
'trasforma versioni nuove numerazioni per nextweb 4
'...........................................................................................
sql = "UPDATE AA_versione SET Versione=0"
CALL DB.Execute(sql, 36)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 2
'...........................................................................................
'aggiunta campi per sfondo della pagina
'...........................................................................................
sql = " ALTER TABLE tb_pages ADD COLUMN " + _
	  "		SfondoColore TEXT(7) WITH COMPRESSION NULL, " + _
	  "		SfondoImmagine TEXT(255) WITH COMPRESSION NULL ;" + _
	  " UPDATE tb_pages SET SfondoColore = Sfondo WHERE Sfondo LIKE '#%' ;" + _
	  " UPDATE tb_pages SET SfondoImmagine = Sfondo WHERE NOT Sfondo LIKE '#%' ;" + _
	  " UPDATE tb_pages SET SfondoColore = '#FFFFFF' WHERE ISNULL(SfondoColore) ; " + _
	  " UPDATE tb_pages SET SfondoImmagine = '' WHERE ISNULL(SfondoImmagine) "
CALL DB.Execute(sql, 2)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 3
'...........................................................................................
'aggiunge campo su tabella oggetti e campo oggetti
'...........................................................................................
sql = " ALTER TABLE tb_objects ADD COLUMN " + _
	  " 	name_objects TEXT(255) WITH COMPRESSION NULL; " + _ 
	  " UPDATE tb_objects SET name_objects = left(img_objects, (instr(1, img_objects, '.') - 1)) "
CALL DB.Execute(sql, 3)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 4
'...........................................................................................
'aggiorna nomi oggetti su layers
'...........................................................................................
sql = " UPDATE tb_layers INNER JOIN tb_objects ON tb_layers.id_objects = tb_objects.id_objects " + _
	  " SET tb_layers.nome = tb_objects.name_objects "
CALL DB.Execute(sql, 4)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 5
'...........................................................................................
'aggiunge campi su tb_pagineSito per gestione keywords e description
'...........................................................................................
sql = " ALTER TABLE tb_pagineSito ADD COLUMN " + _
	  " 	PAGE_keywords_IT TEXT(150) WITH COMPRESSION, " & _
	  " 	PAGE_keywords_EN TEXT(150) WITH COMPRESSION, " & _
	  " 	PAGE_keywords_FR TEXT(150) WITH COMPRESSION, " & _
	  " 	PAGE_keywords_DE TEXT(150) WITH COMPRESSION, " & _
	  " 	PAGE_keywords_ES TEXT(150) WITH COMPRESSION, " & _
	  " 	PAGE_description_IT TEXT(150) WITH COMPRESSION, " & _
	  " 	PAGE_description_EN TEXT(150) WITH COMPRESSION, " & _
	  " 	PAGE_description_FR TEXT(150) WITH COMPRESSION, " & _
	  " 	PAGE_description_DE TEXT(150) WITH COMPRESSION, " & _
	  " 	PAGE_description_ES TEXT(150) WITH COMPRESSION "
sql="select * from AA_versione"
CALL DB.Execute(sql, 5)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 6
'...........................................................................................
'aggiunge campo "lingua iniziale" su tb_webs
'...........................................................................................
sql = " ALTER TABLE tb_webs ADD COLUMN " + _
	  " 	lingua_iniziale TEXT(10) WITH COMPRESSION NULL "
CALL DB.Execute(sql, 6)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 7
'...........................................................................................
'aggiunge campi per nome menu in tutte le lingue
'...........................................................................................
sql = " ALTER TABLE tb_links ADD COLUMN " + _
	  " 	nomelink_it TEXT(255) WITH COMPRESSION NULL, " + _
	  " 	nomelink_en TEXT(255) WITH COMPRESSION NULL, " + _
	  " 	nomelink_fr TEXT(255) WITH COMPRESSION NULL, " + _
	  " 	nomelink_es TEXT(255) WITH COMPRESSION NULL, " + _
	  " 	nomelink_de TEXT(255) WITH COMPRESSION NULL; " + _
	  " UPDATE tb_links SET nomelink_it = nomelink; " + _
	  " ALTER TABLE tb_links DROP COLUMN " + _
	  " 	nomelink "
CALL DB.Execute(sql, 7)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 8
'...........................................................................................
'aggiunge contatori e tabelle per gestione nuove statistiche e archivio
'quando aggiorna resetta i contatori e archivia in storico
'...........................................................................................
sql = " ALTER TABLE tb_webs ADD COLUMN "& vbCrLf & _
	  " 	contUtenti INTEGER, "& vbCrLf & _
	  " 	contCrawler INTEGER , "& vbCrLf & _
	  " 	contAltro INTEGER; "& vbCrLf & _
	  " ALTER TABLE tb_pages ADD COLUMN "& vbCrLf & _
	  " 	contUtenti INTEGER, "& vbCrLf & _
	  " 	contCrawler INTEGER , "& vbCrLf & _
	  " 	contAltro INTEGER; "& vbCrLf & _
	  " CREATE TABLE tb_storico_webs ("& vbCrLf & _
	  "		sw_ID COUNTER CONSTRAINT PK_tb_storico_webs PRIMARY KEY, "& vbCrLf & _
	  "		sw_webs_id INTEGER NULL, "& vbCrLf & _
	  "		sw_data DATETIME NULL, "& vbCrLf & _
	  "		sw_contatore INTEGER NULL, "& vbCrLf & _
	  "		sw_contUtenti INTEGER NULL, "& vbCrLf & _
	  "		sw_contCrawler INTEGER NULL, "& vbCrLf & _
	  "		sw_contAltro INTEGER NULL "& vbCrLf & _
	  ");"& vbCrLf & _
	  " ALTER TABLE tb_storico_webs ADD CONSTRAINT FK_tb_storico_webs__tb_webs "& vbCrLf & _
   	  "		FOREIGN KEY (sw_webs_id) REFERENCES tb_webs (id_webs) "& vbCrLf & _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"& vbCrLf & _
	  " CREATE TABLE tb_storico_pages ("& vbCrLf & _
	  "		sp_ID COUNTER CONSTRAINT PK_tb_storico_pages PRIMARY KEY, "& vbCrLf & _
	  "		sp_page_id INTEGER NULL, "& vbCrLf & _
	  "		sp_pagineSito_id INTEGER NULL, "& vbCrLf & _
	  "		sp_nomepage VARCHAR(250) NULL, "& vbCrLf & _
	  "		sp_lingua VARCHAR(2) NULL, "& vbCrLf & _
	  "		sp_contatore INTEGER NULL, "& vbCrLf & _
	  "		sp_contUtenti INTEGER NULL, "& vbCrLf & _
	  "		sp_contCrawler INTEGER NULL, "& vbCrLf & _
	  "		sp_contAltro INTEGER NULL, "& vbCrLf & _
	  "		sp_sw_id INTEGER NULL "& vbCrLf & _
	  ");"& vbCrLf & _
	  " ALTER TABLE tb_storico_pages ADD CONSTRAINT FK_tb_storico_pages__tb_storico_webs "& vbCrLf & _
   	  "		FOREIGN KEY (sp_sw_id) REFERENCES tb_storico_webs (sw_id) "& vbCrLf & _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"& vbCrLf & _
	  " INSERT INTO tb_storico_webs(sw_webs_id, sw_data, sw_contatore, sw_contUtenti, sw_contCrawler, sw_contAltro) "& vbCrLf & _
	  " SELECT id_webs, NOW, contatore, 0, 0, contatore FROM tb_webs; "& vbCrLf & _
	  " INSERT INTO tb_storico_pages(sp_page_id, sp_pagineSito_id, sp_nomepage, sp_lingua, sp_contatore, sp_contUtenti, sp_contCrawler, sp_contAltro, sp_sw_id) "& vbCrLf & _
	  " SELECT id_page, (SELECT TOP 1 id_pagineSito FROM tb_pagineSito WHERE id_pagDyn_it=id_page OR id_pagDyn_en=id_page OR id_pagDyn_fr=id_page OR id_pagDyn_es=id_page OR id_pagDyn_de=id_page), nomepage, lingua, contatore, 0, 0, contatore, (SELECT sw_id FROM tb_storico_webs WHERE sw_webs_id=id_webs) FROM tb_pages WHERE (SELECT TOP 1 id_pagineSito FROM tb_pagineSito WHERE id_pagDyn_it=id_page OR id_pagDyn_en=id_page OR id_pagDyn_fr=id_page OR id_pagDyn_es=id_page OR id_pagDyn_de=id_page) > 0; "& vbCrLf & _
	  " UPDATE tb_webs SET contatore = 0, contRes = NOW; "& vbCrLf & _
	  " UPDATE tb_pages SET contatore = 0, contRes = NOW; "
CALL DB.Execute(sql, 8)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 9
'...........................................................................................
'aumenta dimensione campi per registrazione metatag
'...........................................................................................
sql = " ALTER TABLE tb_webs ALTER COLUMN META_Author TEXT(255) WITH COMPRESSION NULL; " + _
	  " ALTER TABLE tb_webs ALTER COLUMN META_keywords_IT TEXT(255) WITH COMPRESSION NULL; " + _
	  " ALTER TABLE tb_webs ALTER COLUMN META_keywords_EN TEXT(255) WITH COMPRESSION NULL; " + _
	  " ALTER TABLE tb_webs ALTER COLUMN META_keywords_FR TEXT(255) WITH COMPRESSION NULL; " + _
	  " ALTER TABLE tb_webs ALTER COLUMN META_keywords_DE TEXT(255) WITH COMPRESSION NULL; " + _
	  " ALTER TABLE tb_webs ALTER COLUMN META_keywords_ES TEXT(255) WITH COMPRESSION NULL; " + _
	  " ALTER TABLE tb_webs ALTER COLUMN META_description_IT TEXT(255) WITH COMPRESSION NULL; " + _
	  " ALTER TABLE tb_webs ALTER COLUMN META_description_EN TEXT(255) WITH COMPRESSION NULL; " + _
	  " ALTER TABLE tb_webs ALTER COLUMN META_description_FR TEXT(255) WITH COMPRESSION NULL; " + _
	  " ALTER TABLE tb_webs ALTER COLUMN META_description_DE TEXT(255) WITH COMPRESSION NULL; " + _
	  " ALTER TABLE tb_webs ALTER COLUMN META_description_ES TEXT(255) WITH COMPRESSION NULL; " + _
	  " ALTER TABLE tb_pagineSito ALTER COLUMN PAGE_keywords_IT TEXT(255) WITH COMPRESSION NULL; " + _
	  " ALTER TABLE tb_pagineSito ALTER COLUMN PAGE_keywords_EN TEXT(255) WITH COMPRESSION NULL; " + _
	  " ALTER TABLE tb_pagineSito ALTER COLUMN PAGE_keywords_FR TEXT(255) WITH COMPRESSION NULL; " + _
	  " ALTER TABLE tb_pagineSito ALTER COLUMN PAGE_keywords_DE TEXT(255) WITH COMPRESSION NULL; " + _
	  " ALTER TABLE tb_pagineSito ALTER COLUMN PAGE_keywords_ES TEXT(255) WITH COMPRESSION NULL; " + _
	  " ALTER TABLE tb_pagineSito ALTER COLUMN PAGE_description_IT TEXT(255) WITH COMPRESSION NULL; " + _
	  " ALTER TABLE tb_pagineSito ALTER COLUMN PAGE_description_EN TEXT(255) WITH COMPRESSION NULL; " + _
	  " ALTER TABLE tb_pagineSito ALTER COLUMN PAGE_description_FR TEXT(255) WITH COMPRESSION NULL; " + _
	  " ALTER TABLE tb_pagineSito ALTER COLUMN PAGE_description_DE TEXT(255) WITH COMPRESSION NULL; " + _
	  " ALTER TABLE tb_pagineSito ALTER COLUMN PAGE_description_ES TEXT(255) WITH COMPRESSION NULL; "
CALL DB.Execute(sql, 9)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 10
'...........................................................................................
'modifica percorso obj_credits.asp
'...........................................................................................
sql = " UPDATE tb_objects SET identif_objects='obj_credits.asp' WHERE identif_objects LIKE '%obj_credits%';" + _
	  " UPDATE tb_layers SET aspcode='obj_credits.asp' WHERE aspcode LIKE '%obj_credits%'"
CALL DB.Execute(sql, 10)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 11
'...........................................................................................
'modifica percorso obj_credits.asp
'...........................................................................................
sql = " UPDATE tb_webs SET contatore=0 WHERE IsNull(contatore); " + _
	  " UPDATE tb_webs SET contUtenti=0 WHERE IsNull(contUtenti); " + _
	  " UPDATE tb_webs SET contCrawler=0 WHERE IsNull(contCrawler); " + _
	  " UPDATE tb_webs SET contAltro=0 WHERE IsNull(contAltro); " + _
	  " UPDATE tb_pages SET contatore=0 WHERE IsNull(contatore); " + _
	  " UPDATE tb_pages SET contUtenti=0 WHERE IsNull(contUtenti); " + _
	  " UPDATE tb_pages SET contCrawler=0 WHERE IsNull(contCrawler); " + _
	  " UPDATE tb_pages SET contAltro=0 WHERE IsNull(contAltro); "
CALL DB.Execute(sql, 11)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 12
'...........................................................................................
'modifica tabella gestione versioni editor
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 12)
if DB.last_update_executed then
	if DB.FieldExistsInTable("AA_versione", "editor") then
		CALL DB.ReSyncTransactionAlways()
		CALL conn.execute("ALTER TABLE AA_versione DROP COLUMN editor")
	end if
	if not DB.FieldExistsInTable("AA_versione", "versione_editor") then
		CALL DB.ReSyncTransactionAlways()
		CALL conn.execute("ALTER TABLE AA_versione ADD versione_editor int NULL")
	end if
	CALL conn.execute("UPDATE AA_versione SET versione_editor=4")
	CALL DB.ReSyncTransactionAlways()
	CALL conn.execute("ALTER TABLE AA_versione ALTER COLUMN versione_editor int NOT NULL")
	CALL Aggiornamento_12_PuliziaDirectoryObjects()
end if 
'...........................................................................................
'dichiaro funzione per non avere interferenze con altre variabili d'ambiente
sub Aggiornamento_12_PuliziaDirectoryObjects()
	dim fso, FolderUpload, FolderSite, Path
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	
	set FolderUpload = fso.GetFolder(Application("IMAGE_PATH"))
	'scorre tutte le directory dei siti (solo con nome numerico)
	for each FolderSite in FolderUpload.SubFolders
		if isNumeric(FolderSite.name) then
			'rimuove cartella oggetti
			CALL FolderRemove(fso, FolderSite.path & "\objects", false)
			
			'rimuove file vuoto.jpg e/o obj_vuoto.jpg dalla cartella dei flash
			CALL FileRemove(fso, FolderSite.path & "\flash", "vuoto.jpg", true)
			CALL FileRemove(fso, FolderSite.path & "\flash", "obj_vuoto.jpg", true)
		end if
	next
	
	set fso = nothing
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 13
'...........................................................................................
'aggiunge tabella per descrizione ricerca per dati di tipo "contenuto"
'...........................................................................................
sql = " CREATE TABLE tb_ricerca_contenuti ( " + _
	  "		rc_id INT IDENTITY (1, 1) NOT NULL , " + _
	  "		rc_sql TEXT(250) WITH COMPRESSION NOT NULL , " + _
	  "		rc_campoID TEXT(50) WITH COMPRESSION NOT NULL , " + _
	  "		rc_campoTitolo TEXT(50) WITH COMPRESSION NOT NULL , " + _
	  "		rc_campoDescr TEXT(50) WITH COMPRESSION NOT NULL , " + _
	  "		rc_campiFiltro TEXT(250) WITH COMPRESSION NOT NULL , " + _
  	  "		rc_paginaSito_id INT NOT NULL , " + _
	  "		rc_urlID TEXT(100) WITH COMPRESSION NOT NULL " + _
	  " ); " + _
	  " ALTER TABLE tb_ricerca_contenuti ADD CONSTRAINT PK_tb_ricerca_contenuti PRIMARY KEY (rc_id);"
CALL DB.Execute(sql, 13)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 14
'...........................................................................................
'aggiunge campo web in tabella ricerca
'...........................................................................................
sql = " ALTER TABLE tb_ricerca_contenuti ADD " + _
	  "		rc_web_id INT NULL "
CALL DB.Execute(sql, 14)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 15
'...........................................................................................
'modifica campo sql in tabella descrizione contenuti
'...........................................................................................
sql = " ALTER TABLE tb_ricerca_contenuti ALTER COLUMN " + _
	  "		rc_sql TEXT WITH COMPRESSION NOT NULL "
CALL DB.Execute(sql, 15)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 16
'...........................................................................................
'aumenta dimensione campi per registrazione metatag
'...........................................................................................
sql = " ALTER TABLE tb_webs ALTER COLUMN META_Author TEXT(255) WITH COMPRESSION NULL; " + _
	  " ALTER TABLE tb_webs ALTER COLUMN META_keywords_IT MEMO WITH COMPRESSION NULL; " + _
	  " ALTER TABLE tb_webs ALTER COLUMN META_keywords_EN MEMO WITH COMPRESSION NULL; " + _
	  " ALTER TABLE tb_webs ALTER COLUMN META_keywords_FR MEMO WITH COMPRESSION NULL; " + _
	  " ALTER TABLE tb_webs ALTER COLUMN META_keywords_DE MEMO WITH COMPRESSION NULL; " + _
	  " ALTER TABLE tb_webs ALTER COLUMN META_keywords_ES MEMO WITH COMPRESSION NULL; " + _
	  " ALTER TABLE tb_webs ALTER COLUMN META_description_IT MEMO WITH COMPRESSION NULL; " + _
	  " ALTER TABLE tb_webs ALTER COLUMN META_description_EN MEMO WITH COMPRESSION NULL; " + _
	  " ALTER TABLE tb_webs ALTER COLUMN META_description_FR MEMO WITH COMPRESSION NULL; " + _
	  " ALTER TABLE tb_webs ALTER COLUMN META_description_DE MEMO WITH COMPRESSION NULL; " + _
	  " ALTER TABLE tb_webs ALTER COLUMN META_description_ES MEMO WITH COMPRESSION NULL; "
CALL DB.Execute(sql, 16)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 17
'...........................................................................................
'corregge errore di generazione pagine con contatori a NULL
'...........................................................................................
sql = " UPDATE tb_pages SET Contatore=0 WHERE IsNull(contatore) ; " + _
	  " UPDATE tb_pages SET contUtenti=0 WHERE IsNull(contUtenti) ; " + _
	  " UPDATE tb_pages SET contCrawler=0 WHERE IsNull(contCrawler) ; " + _
	  " UPDATE tb_pages SET contAltro=0 WHERE IsNull(contAltro) ; " + _
	  " UPDATE tb_webs SET Contatore=0 WHERE IsNull(contatore) ; " + _
	  " UPDATE tb_webs SET contUtenti=0 WHERE IsNull(contUtenti) ; " + _
	  " UPDATE tb_webs SET contCrawler=0 WHERE IsNull(contCrawler) ; " + _
	  " UPDATE tb_webs SET contAltro=0 WHERE IsNull(contAltro) ; "
CALL DB.Execute(sql, 17)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 18
'...........................................................................................
'aggiorna stato plugin credits
'...........................................................................................
sql = " UPDATE tb_objects SET identif_objects = 'amministrazione/library/plugins/obj_credits.asp' WHERE TRIM(identif_objects) LIKE '%credits.asp'; " + _
      " UPDATE tb_layers SET aspcode = 'amministrazione/library/plugins/obj_credits.asp' WHERE TRIM(aspcode) LIKE '%credits.asp'"
CALL DB.Execute(sql, 18)
'*******************************************************************************************

'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 19
'...........................................................................................
'aggiunge chiavi di verifica a google webmaster tools
'...........................................................................................
sql = " ALTER TABLE tb_webs ADD " + _
	  "		google_analytics_code " + SQL_CharField(Conn, 255) + " NULL, " + vbCrLf + _
	  "		google_webmaster_tools_verify_code " + SQL_CharField(Conn, 255) + " NULL" + vbCrLf
sql="select * from AA_versione"
CALL DB.Execute(sql, 19)
'*******************************************************************************************


'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 20
'...........................................................................................
'aggiorna stato plugin credits
'...........................................................................................
sql = "	ALTER TABLE tb_pages ADD" + vbCrLf + _
	  "		semplificata BIT NULL;"
CALL DB.Execute(sql, 20)
'*******************************************************************************************

%>
<% '........................................................................................... %>
<!--#INCLUDE FILE="Update__FileFooter.asp" -->
<% '........................................................................................... %>
<!--#INCLUDE FILE="Update__FileHeader.asp" -->
<% '........................................................................................... %>
<!--#INCLUDE FILE="../Tools.asp" -->
<!--#INCLUDE FILE="../Tools4Admin.asp" -->
<!--#INCLUDE FILE="../ToolsDescrittori.asp" -->
<%

'*******************************************************************************************
'AGGIORNAMENTO 1
'...........................................................................................
'cancellazione campi da tabella tb_webs
'...........................................................................................
sql = "ALTER TABLE tb_webs DROP COLUMN menu_visibile; " &_
	  "ALTER TABLE tb_webs DROP COLUMN titolo_webs"
CALL DB.Execute(sql, 1)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 2
'...........................................................................................
'aggiunta campi per gestione lingue su tb_webs
'...........................................................................................
sql = "ALTER TABLE tb_webs ADD COLUMN titolo_IT nvarchar(255), " & _
	  "lingua_EN bit, titolo_EN nvarchar(255), " & _
	  "lingua_FR bit, titolo_FR nvarchar(255), " & _
	  "lingua_DE bit, titolo_DE nvarchar(255), " & _
	  "lingua_ES bit, titolo_ES nvarchar(255)"
CALL DB.Execute(sql, 2)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 3
'...........................................................................................
'aggiunta campi per gestione delle lingue su tb_menuitem
'...........................................................................................
sql = " ALTER TABLE tb_menuItem ADD COLUMN " & _
	  " titolo_menuItem_IT nvarchar(255), link_menuItem_IT nvarchar(255), " & _
	  " titolo_menuItem_EN nvarchar(255), link_menuItem_EN nvarchar(255), " & _
	  " titolo_menuItem_FR nvarchar(255), link_menuItem_FR nvarchar(255), " & _
	  " titolo_menuItem_DE nvarchar(255), link_menuItem_DE nvarchar(255), " & _
	  " titolo_menuItem_ES nvarchar(255), link_menuItem_ES nvarchar(255);" &_
	  " UPDATE tb_menuItem SET titolo_menuItem_IT= titolo_menuItem, link_menuItem_IT=link_menuItem, " &_
	  " titolo_menuItem_EN=titolo_en_menuItem, link_menuitem_EN=link_en_menuitem; " &_
	  " ALTER TABLE tb_menuItem DROP COLUMN titolo_menuItem; " & _
	  " ALTER TABLE tb_menuItem DROP COLUMN titolo_en_menuItem; " & _
	  " ALTER TABLE tb_menuItem DROP COLUMN link_menuitem; " & _
	  " ALTER TABLE tb_menuItem DROP COLUMN link_en_menuitem "
CALL DB.Execute(sql, 3)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 4
'...........................................................................................
'toglie campi non piu' usati dai menu
'...........................................................................................
sql = " ALTER TABLE tb_links DROP COLUMN asplink; " & _
	  " ALTER TABLE tb_links DROP COLUMN nomelink_EN; " & _
	  " ALTER TABLE tb_links DROP COLUMN visilink; " & _
	  " ALTER TABLE tb_links DROP COLUMN class; " & _
	  " ALTER TABLE tb_links DROP COLUMN lingua; " & _
	  " ALTER TABLE tb_links DROP COLUMN data_creaz; " & _
	  " ALTER TABLE tb_links DROP COLUMN data_Scad"
CALL DB.Execute(sql, 4)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 5
'...........................................................................................
'aggiunge campo alla tabella tb_objects e duplica struttura oggetti per ogni sito
'...........................................................................................
sql = " ALTER TABLE tb_objects ADD COLUMN id_webs integer; " &_
	  " ALTER TABLE tb_objects DROP COLUMN method_objects; " & _
	  " INSERT INTO tb_objects (identif_objects, img_objects, param_list, id_webs) " &_
	  " SELECT identif_objects, img_objects, param_list, tb_webs.id_webs " &_
	  " FROM tb_webs, tb_objects ; " &_
	  " DELETE FROM tb_objects WHERE ISNULL(id_webs); " &_
	  " ALTER TABLE tb_objects ADD CONSTRAINT FK_tb_objects FOREIGN KEY (id_webs) " &_
	  " REFERENCES tb_webs(id_webs) ON UPDATE CASCADE ON DELETE CASCADE "
CALL DB.Execute(sql, 5)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 6 e 7
'...........................................................................................
'aggiunge campo alla tabella tb_pages per stabilirne la corrispondenza con il sito
'...........................................................................................
sql = "ALTER TABLE tb_pages ADD COLUMN id_webs integer"
CALL DB.Execute(sql, 6)
'...........................................................................................
if DB.last_update_executed then
	sql = "SELECT * FROM tb_pages"
	rs.open sql, conn, aDopenStatic, adLockOptimistic, adAsyncFetch
	while not rs.eof
		if rs("template") then
			sql = "SELECT (id_webs) AS ID FROM tb_templates WHERE id_pagina=" & rs("id_page")
		else
			sql = "SELECT (id_web) AS ID FROM tb_pagineSito WHERE id_pagDyn=" & rs("id_page") & _
				  " OR id_pagStage=" & rs("id_page") & " OR id_pagDynEN=" & rs("id_page") & _
				  " OR id_pagStageEN=" & rs("id_page")
		end if
		rst.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdtext
		if not rst.eof then
			rs("id_webs")=rst("ID")
			rs.update
		end if
		rst.close
		rs.Movenext
	wend
	rs.close
end if
'...........................................................................................
sql = " DROP TABLE tb_templates; " &_
	  " ALTER TABLE tb_pages ADD CONSTRAINT FK_tb_pages FOREIGN KEY (id_webs) " &_
	  " REFERENCES tb_webs(id_webs) ON UPDATE CASCADE ON DELETE CASCADE "
CALL DB.Execute(sql, 7)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 8
'...........................................................................................
'aggiunge id dell'oggetto utilizzato nel layer
'...........................................................................................
sql = " ALTER TABLE tb_layers ADD COLUMN id_objects integer"
CALL DB.Execute(sql, 8)
if DB.last_update_executed then
	sql = "SELECT * FROM tb_layers WHERE id_tipo=4"
	rs.open sql, conn, aDopenStatic, adLockOptimistic, adAsyncFetch
	while not rs.eof
		sql = "SELECT id_objects FROM tb_objects WHERE img_objects='" & rs("nome") & "'"
		rst.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdtext
		if not rst.eof then
			rs("id_objects")=rst("id_objects")
			rs.update
		end if
		rst.close
		rs.Movenext
	wend
	rs.close
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 9
'...........................................................................................
'toglie campi non piu' usati nella tabella tb_pages
'...........................................................................................
sql = " ALTER TABLE tb_pages DROP COLUMN asppage; " & _
	  " ALTER TABLE tb_pages DROP COLUMN stylpage; " & _
	  " ALTER TABLE tb_pages DROP COLUMN no_pagina; " & _
	  " ALTER TABLE tb_pages DROP COLUMN data_creaz; " & _
	  " ALTER TABLE tb_pages DROP COLUMN data_scad  "
CALL DB.Execute(sql, 9)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 10
'...........................................................................................
'toglie indice e relazione link dei menu e tb_pages
'...........................................................................................
sql = " DROP INDEX id_link ON tb_pages; " &_
	  " ALTER TABLE tb_pages DROP COLUMN id_link; "
CALL DB.Execute(sql, 10)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 11
'...........................................................................................
'aggiunge campi al tb_paginesito per gestione 5 lingue
'...........................................................................................
sql = " ALTER TABLE tb_pagineSito ADD COLUMN nome_ps_IT nvarchar(255)," &_
	  " nome_ps_FR nvarchar(255), " &_
	  " nome_ps_DE nvarchar(255), " &_
	  " nome_ps_ES nvarchar(255), " &_
	  " id_pagDyn_IT integer, " &_
	  " id_pagDyn_EN integer, " &_
	  " id_pagDyn_FR integer, " &_
	  " id_pagDyn_DE integer, " &_
	  " id_pagDyn_ES integer, " &_
	  " id_pagStage_IT integer, " &_
	  " id_pagStage_EN integer, " &_
	  " id_pagStage_FR integer, " &_
	  " id_pagStage_DE integer, " &_
	  " id_pagStage_ES integer ;" &_
	  " UPDATE tb_PagineSito SET nome_ps_IT=nome_ps, " &_
	  " id_pagDyn_IT=id_pagDyn, "  &_
	  " id_pagDyn_EN=id_pagDynEN, " &_
	  " id_pagStage_IT=id_pagStage, " &_
	  " id_pagStage_EN=id_pagStageEN; " &_
	  " DROP INDEX id_pagDyn ON tb_PagineSito; " &_
	  " DROP INDEX id_pagDyn1 ON tb_PagineSito; " &_
	  " DROP INDEX id_pagStage ON tb_PagineSito; " &_
	  " DROP INDEX id_pagStage1 ON tb_PagineSito; " &_
	  " ALTER TABLE tb_paginesito DROP COLUMN lingua_ps, nome_ps, id_pagDyn, id_pagStage, id_pagDynEN, id_pagStageEN "
CALL DB.Execute(sql, 11)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 12
'...........................................................................................
'aggiunge campi per utilizzo contatore
'...........................................................................................
sql = " ALTER TABLE tb_webs ADD COLUMN contatore integer, contRes smalldatetime"
CALL DB.Execute(sql, 12)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 13
'...........................................................................................
'aggiunge campo per la correlazione tra una pagina ed un record esterno.
'...........................................................................................
sql = " ALTER TABLE tb_pages ADD COLUMN external_ID nvarchar(50)"
CALL DB.Execute(sql, 13)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 14
'...........................................................................................
'aggiunge compressione unicode ai campi del menu
'...........................................................................................
sql = " ALTER TABLE tb_menuItem ADD COLUMN " & _
	  " b_titolo_menuItem_IT TEXT(255) WITH COMPRESSION, " &_
	  " b_link_menuItem_IT TEXT(255) WITH COMPRESSION, " &_
	  " b_titolo_menuItem_EN TEXT(255) WITH COMPRESSION, " &_
	  " b_link_menuItem_EN TEXT(255) WITH COMPRESSION, " &_
	  " b_titolo_menuItem_FR TEXT(255) WITH COMPRESSION, " &_
	  " b_link_menuItem_FR  TEXT(255) WITH COMPRESSION, " &_
	  " b_titolo_menuItem_DE TEXT(255) WITH COMPRESSION, " &_
	  " b_link_menuItem_DE TEXT(255) WITH COMPRESSION, " &_
	  " b_titolo_menuItem_ES TEXT(255) WITH COMPRESSION, " &_
	  " b_link_menuItem_ES TEXT(255) WITH COMPRESSION; " &_
	  " UPDATE tb_menuItem SET b_titolo_menuItem_IT = titolo_menuItem_IT, " &_
	  " b_link_menuItem_IT = link_menuItem_IT, b_titolo_menuItem_EN = titolo_menuItem_EN, " &_
	  " b_link_menuItem_EN = link_menuItem_EN, b_titolo_menuItem_FR = titolo_menuItem_FR, " &_
	  " b_link_menuItem_FR = link_menuItem_FR, b_titolo_menuItem_DE = titolo_menuItem_DE, " &_
	  " b_link_menuItem_DE = link_menuItem_DE, b_titolo_menuItem_ES = titolo_menuItem_ES, " &_
	  " b_link_menuItem_ES = link_menuItem_ES; " &_
	  " ALTER TABLE tb_menuItem DROP COLUMN titolo_menuItem_IT, link_menuItem_IT, " & _
	  " titolo_menuItem_EN, link_menuItem_EN, titolo_menuItem_FR, link_menuItem_FR, " & _
	  " titolo_menuItem_DE, link_menuItem_DE, titolo_menuItem_ES, link_menuItem_ES; " &_
	  " ALTER TABLE tb_menuItem ADD COLUMN " & _
	  " titolo_menuItem_IT TEXT(255) WITH COMPRESSION, " &_
	  " link_menuItem_IT TEXT(255) WITH COMPRESSION, " &_
	  " titolo_menuItem_EN TEXT(255) WITH COMPRESSION, " &_
	  " link_menuItem_EN TEXT(255) WITH COMPRESSION, " &_
	  " titolo_menuItem_FR TEXT(255) WITH COMPRESSION, " &_
	  " link_menuItem_FR  TEXT(255) WITH COMPRESSION, " &_
	  " titolo_menuItem_DE TEXT(255) WITH COMPRESSION, " &_
	  " link_menuItem_DE TEXT(255) WITH COMPRESSION, " &_
	  " titolo_menuItem_ES TEXT(255) WITH COMPRESSION, " &_
	  " link_menuItem_ES TEXT(255) WITH COMPRESSION; " &_
	  " UPDATE tb_menuItem SET titolo_menuItem_IT = b_titolo_menuItem_IT, " &_
	  " link_menuItem_IT = b_link_menuItem_IT, titolo_menuItem_EN = b_titolo_menuItem_EN, " &_
	  " link_menuItem_EN = b_link_menuItem_EN, titolo_menuItem_FR = b_titolo_menuItem_FR, " &_
	  " link_menuItem_FR = b_link_menuItem_FR, titolo_menuItem_DE = b_titolo_menuItem_DE, " &_
	  " link_menuItem_DE = b_link_menuItem_DE, titolo_menuItem_ES = b_titolo_menuItem_ES, " &_
	  " link_menuItem_ES = b_link_menuItem_ES; " &_
	  " ALTER TABLE tb_menuItem DROP COLUMN b_titolo_menuItem_IT, b_link_menuItem_IT, " & _
	  " b_titolo_menuItem_EN, b_link_menuItem_EN, b_titolo_menuItem_FR, b_link_menuItem_FR, " & _
	  " b_titolo_menuItem_DE, b_link_menuItem_DE, b_titolo_menuItem_ES, b_link_menuItem_ES; "
CALL DB.Execute(sql, 14)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 15
'...........................................................................................
'aggiunge compressione unicode ai campi del sito
'...........................................................................................
sql = " ALTER TABLE tb_webs ADD COLUMN " & _
	  " b_titolo_IT TEXT(255) WITH COMPRESSION, " &_
	  " b_titolo_EN TEXT(255) WITH COMPRESSION, " &_
	  " b_titolo_FR TEXT(255) WITH COMPRESSION, " &_
	  " b_titolo_DE TEXT(255) WITH COMPRESSION, " &_
	  " b_titolo_ES TEXT(255) WITH COMPRESSION; " &_
	  " UPDATE tb_webs SET " &_
	  " b_titolo_IT = titolo_IT, " &_
	  " b_titolo_EN = titolo_EN, " &_
	  " b_titolo_FR = titolo_FR, " &_
	  " b_titolo_DE = titolo_DE, " &_
	  " b_titolo_ES = titolo_ES; " &_
	  " ALTER TABLE tb_webs DROP COLUMN " &_
	  " titolo_IT, " &_
	  " titolo_EN, " &_
	  " titolo_FR, " &_
	  " titolo_DE, " &_
	  " titolo_ES; " &_
	  " ALTER TABLE tb_webs ADD COLUMN " & _
	  " titolo_IT TEXT(255) WITH COMPRESSION, " &_
	  " titolo_EN TEXT(255) WITH COMPRESSION, " &_
	  " titolo_FR TEXT(255) WITH COMPRESSION, " &_
	  " titolo_DE TEXT(255) WITH COMPRESSION, " &_
	  " titolo_ES TEXT(255) WITH COMPRESSION; " &_
	  " UPDATE tb_webs SET " &_
	  " titolo_IT = b_titolo_IT, " &_
	  " titolo_EN = b_titolo_EN, " &_
	  " titolo_FR = b_titolo_FR, " &_
	  " titolo_DE = b_titolo_DE, " &_
	  " titolo_ES = b_titolo_ES; " &_
	  " ALTER TABLE tb_webs DROP COLUMN " &_
	  " b_titolo_IT, " &_
	  " b_titolo_EN, " &_
	  " b_titolo_FR, " &_
	  " b_titolo_DE, " &_
	  " b_titolo_ES; "
CALL DB.Execute(sql, 15)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 16
'...........................................................................................
'aggiunge compressione unicode ai campi della tabella paginesito
'...........................................................................................
sql = " ALTER TABLE tb_pagineSito ADD COLUMN " & _
	  " b_nome_ps_IT TEXT(255) WITH COMPRESSION, " &_
	  " b_nome_ps_EN TEXT(255) WITH COMPRESSION, " &_
	  " b_nome_ps_FR TEXT(255) WITH COMPRESSION, " &_
	  " b_nome_ps_DE TEXT(255) WITH COMPRESSION, " &_
	  " b_nome_ps_ES TEXT(255) WITH COMPRESSION; " &_
	  " UPDATE tb_pagineSito SET " &_
	  " b_nome_ps_IT = nome_ps_IT, " &_
	  " b_nome_ps_EN = nome_ps_EN, " &_
	  " b_nome_ps_FR = nome_ps_FR, " &_
	  " b_nome_ps_DE = nome_ps_DE, " &_
	  " b_nome_ps_ES = nome_ps_ES; " &_
	  " ALTER TABLE tb_pagineSito DROP COLUMN " &_
	  " nome_ps_IT, " &_
	  " nome_ps_EN, " &_
	  " nome_ps_FR, " &_
	  " nome_ps_DE, " &_
	  " nome_ps_ES; " &_
	  " ALTER TABLE tb_pagineSito ADD COLUMN " & _
	  " nome_ps_IT TEXT(255) WITH COMPRESSION, " &_
	  " nome_ps_EN TEXT(255) WITH COMPRESSION, " &_
	  " nome_ps_FR TEXT(255) WITH COMPRESSION, " &_
	  " nome_ps_DE TEXT(255) WITH COMPRESSION, " &_
	  " nome_ps_ES TEXT(255) WITH COMPRESSION; " &_
	  " UPDATE tb_pagineSito SET " &_
	  " nome_ps_IT = b_nome_ps_IT, " &_
	  " nome_ps_EN = b_nome_ps_EN, " &_
	  " nome_ps_FR = b_nome_ps_FR, " &_
	  " nome_ps_DE = b_nome_ps_DE, " &_
	  " nome_ps_ES = b_nome_ps_ES; " &_
	  " ALTER TABLE tb_pagineSito DROP COLUMN " &_
	  " b_nome_ps_IT, " &_
	  " b_nome_ps_EN, " &_
	  " b_nome_ps_FR, " &_
	  " b_nome_ps_DE, " &_
	  " b_nome_ps_ES; "
CALL DB.Execute(sql, 16)
'*******************************************************************************************


'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 17
'...........................................................................................
'aggiunge compressione unicode ai campi della tabella pages
'...........................................................................................
sql = " ALTER TABLE tb_pages ADD COLUMN b_external_ID TEXT(50) WITH COMPRESSION; " &_
	  " UPDATE tb_pages SET b_external_ID = external_ID; " &_
	  " ALTER TABLE tb_pages DROP COLUMN external_ID; " &_
	  " ALTER TABLE tb_pages ADD COLUMN external_ID TEXT(255) WITH COMPRESSION; " &_
	  " UPDATE tb_pages SET external_ID = b_external_ID; " &_
	  " ALTER TABLE tb_pages DROP COLUMN b_external_ID; "
CALL DB.Execute(sql, 17)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 18
'...........................................................................................
'aggiunge compressione unicode ai campi della tabella pages
'...........................................................................................
sql = " ALTER TABLE tb_links ADD COLUMN b_nomelink TEXT(50) WITH COMPRESSION; " &_
	  " UPDATE tb_links SET b_nomelink = nomelink; " &_
	  " ALTER TABLE tb_links DROP COLUMN nomelink; " &_
	  " ALTER TABLE tb_links ADD COLUMN nomelink TEXT(255) WITH COMPRESSION; " &_
	  " UPDATE tb_links SET nomelink = b_nomelink; " &_
	  " ALTER TABLE tb_links DROP COLUMN b_nomelink; "
CALL DB.Execute(sql, 18)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 19
'*******************************************************************************************
'aggiunge campi per metatag su tb_webs
'*******************************************************************************************
sql = " ALTER TABLE tb_webs ADD COLUMN " & _
	  " META_Author TEXT(150) WITH COMPRESSION, " & _
	  " META_keywords_IT TEXT(150) WITH COMPRESSION, " & _
	  " META_keywords_EN TEXT(150) WITH COMPRESSION, " & _
	  " META_keywords_FR TEXT(150) WITH COMPRESSION, " & _
	  " META_keywords_DE TEXT(150) WITH COMPRESSION, " & _
	  " META_keywords_ES TEXT(150) WITH COMPRESSION, " & _
	  " META_description_IT TEXT(150) WITH COMPRESSION, " & _
	  " META_description_EN TEXT(150) WITH COMPRESSION, " & _
	  " META_description_FR TEXT(150) WITH COMPRESSION, " & _
	  " META_description_DE TEXT(150) WITH COMPRESSION, " & _
	  " META_description_ES TEXT(150) WITH COMPRESSION "
CALL DB.Execute(sql, 19)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 20
'...........................................................................................
'aggiunta campi per utilizzo di immagini come etichette per voci di menu
'...........................................................................................
sql = " ALTER TABLE tb_menuItem ADD COLUMN " & _
	  " 	image_menuItem_IT TEXT(150) WITH COMPRESSION," & _
	  " 	image_menuItem_EN TEXT(150) WITH COMPRESSION," & _
	  " 	image_menuItem_FR TEXT(150) WITH COMPRESSION," & _
	  " 	image_menuItem_DE TEXT(150) WITH COMPRESSION," & _
	  " 	image_menuItem_ES TEXT(150) WITH COMPRESSION"
CALL DB.Execute(sql, 20)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 21
'...........................................................................................
'corregge errore su collegamento id dell'oggetto utilizzato nel layer
'...........................................................................................
sql = " SELECT id_objects FROM tb_objects"
CALL DB.Execute(sql, 21)
if DB.last_update_executed then
	sql = "SELECT * FROM tb_layers INNER JOIN tb_pages ON tb_layers.id_pag=tb_pages.id_page WHERE id_tipo=4"
	rs.open sql, conn, aDopenStatic, adLockOptimistic, adAsyncFetch
	while not rs.eof
		sql = "SELECT id_objects FROM tb_objects WHERE img_objects='" & rs("nome") & "' AND id_webs=" & rs("id_webs")
		rst.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdtext
		if not rst.eof then
			rs("id_objects")=rst("id_objects")
			rs.update
		end if
		rst.close
		rs.MoveNext
	wend
	rs.close
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 22
'...........................................................................................
'inserimento oggetto di default x gestione credits
'...........................................................................................
sql = "INSERT INTO tb_objects(identif_objects, param_list, img_objects, id_webs) "& _
	  "SELECT '../../../next-framework/library/obj_credits.asp', 'coloreLink=black;;"& vbCrLf &"coloreLinkHover=blue;;"& vbCrLf &"coloreTesto=black;;"& vbCrLf &"', 'obj_credits.jpg', id_webs FROM tb_webs"
CALL DB.Execute(sql, 22)
'*******************************************************************************************


'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 23
'...........................................................................................
'modifica campo tab objects x gestione credits
'...........................................................................................
sql = "ALTER TABLE tb_objects ALTER COLUMN "& _
	  "identif_objects TEXT(70) WITH COMPRESSION NULL"
CALL DB.Execute(sql, 23)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 24
'...........................................................................................
'aggiunta campo lingua su ogni pagina
'...........................................................................................
sql = " ALTER TABLE tb_pages ADD COLUMN " & _
	  "		lingua TEXT(2) WITH COMPRESSION NULL; " & _
	  " UPDATE tb_pages SET lingua='it' WHERE id_page IN (SELECT id_pagDyn_IT FROM tb_pagineSito) " & _
	  " OR id_page IN (SELECT id_pagStage_IT FROM tb_pagineSito); " & _
	  " UPDATE tb_pages SET lingua='en' WHERE id_page IN (SELECT id_pagDyn_EN FROM tb_pagineSito) " & _
	  " OR id_page IN (SELECT id_pagStage_EN FROM tb_pagineSito); " & _
	  " UPDATE tb_pages SET lingua='fr' WHERE id_page IN (SELECT id_pagDyn_FR FROM tb_pagineSito) " & _
	  " OR id_page IN (SELECT id_pagStage_FR FROM tb_pagineSito); " & _
	  " UPDATE tb_pages SET lingua='de' WHERE id_page IN (SELECT id_pagDyn_DE FROM tb_pagineSito) " & _
	  " OR id_page IN (SELECT id_pagStage_DE FROM tb_pagineSito); " & _
	  " UPDATE tb_pages SET lingua='es' WHERE id_page IN (SELECT id_pagDyn_ES FROM tb_pagineSito) " & _
	  " OR id_page IN (SELECT id_pagStage_ES FROM tb_pagineSito)"
CALL DB.Execute(sql, 24)
'*******************************************************************************************


'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 25
'...........................................................................................
'aggiunta campo data modifica pagina
'...........................................................................................
sql = " ALTER TABLE tb_paginesito ADD COLUMN DataModifica DATETIME NULL "
CALL DB.Execute(sql, 25)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 26
'...........................................................................................
'aggiunta campo data modifica pagina
'...........................................................................................
sql = "  UPDATE tb_pagineSito SET DataModifica=NOW()"
CALL DB.Execute(sql, 26)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 27
'...........................................................................................
'modifica lunghezza campo per nome layer
'...........................................................................................
sql = " ALTER TABLE tb_layers ALTER COLUMN nome TEXT(250) WITH COMPRESSION NULL"
CALL DB.Execute(sql, 27)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 28
'...........................................................................................
'aggiunge ampi per la gestione del tag title
'...........................................................................................
sql = " ALTER TABLE tb_menuItem ADD COLUMN " & _
	  " tag_title_IT TEXT(255) WITH COMPRESSION NULL, " &_
	  " tag_title_EN TEXT(255) WITH COMPRESSION NULL, " &_
	  " tag_title_FR TEXT(255) WITH COMPRESSION NULL, " &_
	  " tag_title_DE TEXT(255) WITH COMPRESSION NULL, " &_
	  " tag_title_ES TEXT(255) WITH COMPRESSION NULL; " &_
	  ""
CALL DB.Execute(sql, 28)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 29
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
	  " ALTER TABLE tb_webs ALTER COLUMN META_description_ES TEXT(255) WITH COMPRESSION NULL; "
CALL DB.Execute(sql, 29)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 30
'...........................................................................................
'modifica percorso obj_credits.asp
'...........................................................................................
sql = " UPDATE tb_objects SET identif_objects='obj_credits.asp' WHERE identif_objects LIKE '%obj_credits%';" + _
	  " UPDATE tb_layers SET aspcode='obj_credits.asp' WHERE aspcode LIKE '%obj_credits%'"
CALL DB.Execute(sql, 30)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 31
'...........................................................................................
'modifica tabella gestione versioni editor
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 31)
if DB.last_update_executed then
	sql = "SELECT * FROM AA_versione"
	rs.open sql, conn, adOpenStatic, adLockOptimistic
	if rs.fields.count>1 then
		sql = "ALTER TABLE AA_versione DROP COLUMN " & rs(1).name
	else
		sql = ""
	end if
	rs.close
	CALL DB.ReSyncTransactionAlways()
	if sql<>"" then
		CALL conn.execute(sql)
	end if
	CALL conn.execute("ALTER TABLE AA_versione ADD versione_editor int NULL")
	CALL conn.execute("UPDATE AA_versione SET versione_editor=3")
	CALL conn.execute("ALTER TABLE AA_versione ALTER COLUMN versione_editor int NOT NULL")
end if 
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 32
'...........................................................................................
'aggiorna stato plugin credits
'...........................................................................................
sql = " UPDATE tb_objects SET identif_objects = 'amministrazione/library/plugins/obj_credits.asp' WHERE TRIM(identif_objects) LIKE '%credits.asp'; " + _
      " UPDATE tb_layers SET aspcode = 'amministrazione/library/plugins/obj_credits.asp' WHERE TRIM(aspcode) LIKE '%credits.asp'"
CALL DB.Execute(sql, 32)
'*******************************************************************************************
'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 33
'...........................................................................................
'aggiunge chiavi di verifica a google webmaster tools
'...........................................................................................
sql = " ALTER TABLE tb_webs ADD " + _
	  "		google_analytics_code " + SQL_CharField(Conn, 255) + " NULL, " + vbCrLf + _
	  "		google_webmaster_tools_verify_code " + SQL_CharField(Conn, 255) + " NULL" + vbCrLf
CALL DB.Execute(sql, 33)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 34
'...........................................................................................
'aggiunge campi su tb_pagineSito per gestione keywords e description
'...........................................................................................
sql = " ALTER TABLE tb_pagineSito ADD COLUMN " + _
	  " 	PAGE_keywords_IT TEXT(255) WITH COMPRESSION, " & _
	  " 	PAGE_keywords_EN TEXT(255) WITH COMPRESSION, " & _
	  " 	PAGE_keywords_FR TEXT(255) WITH COMPRESSION, " & _
	  " 	PAGE_keywords_DE TEXT(255) WITH COMPRESSION, " & _
	  " 	PAGE_keywords_ES TEXT(255) WITH COMPRESSION, " & _
	  " 	PAGE_description_IT TEXT(255) WITH COMPRESSION, " & _
	  " 	PAGE_description_EN TEXT(255) WITH COMPRESSION, " & _
	  " 	PAGE_description_FR TEXT(255) WITH COMPRESSION, " & _
	  " 	PAGE_description_DE TEXT(255) WITH COMPRESSION, " & _
	  " 	PAGE_description_ES TEXT(255) WITH COMPRESSION "
CALL DB.Execute(sql, 34)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 35
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
	  " ALTER TABLE tb_webs ALTER COLUMN META_description_ES TEXT(255) WITH COMPRESSION NULL; "
CALL DB.Execute(sql, 35)
'*******************************************************************************************

%>
<% '........................................................................................... %>
<!--#INCLUDE FILE="Update__FileFooter.asp" -->
<% '........................................................................................... %>
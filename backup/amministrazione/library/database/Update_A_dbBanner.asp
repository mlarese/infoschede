<!--#INCLUDE FILE="Update__FileHeader.asp" -->
<% '........................................................................................... %>
<!--#INCLUDE FILE="../Tools.asp" -->
<!--#INCLUDE FILE="../Tools4Admin.asp" -->
<!--#INCLUDE FILE="../ToolsDescrittori.asp" -->
<%

'*******************************************************************************************
'AGGIORNAMENTO 2
'...........................................................................................
'importa tutti i dati del vecchio dbBanner nella struttura del dbContent
'...........................................................................................
sql = "SELECT TOP 1 * FROM tb_banner"
CALL DB.Execute(sql, 2)		
'...........................................................................................
if DB.last_update_executed then
	CALL Aggiornamento_3()		'inserito in una funzione per togliere interferenza delle variabili con altri script
end  if
'...........................................................................................
sub Aggiornamento_3()
	
	dim connContent, rsb_s, rsb_d, rsp_s, rsp_d
	set connContent = server.createobject("adodb.connection")
	connContent.open Application("DATA_ConnectionString")
	
	'cancella tutti i dati preesistenti nelle tabelle di destinazione
	sql = "DELETE FROM rel_rub_ind WHERE id_rubrica=" & Application("NextBanner_RubricaAziende") & "; " + _
		  "DELETE FROM tb_indirizzario WHERE idElencoIndirizzi NOT IN (SELECT id_indirizzo FROM rel_rub_ind ); " + _
		  "DELETE FROM tb_banner; " + _
		  "DELETE FROM tb_tipiBanner; " + _
		  "DELETE FROM rel_banner_pagine; " + _
		  "DELETE FROM tb_pagine; " + _
		  "DELETE FROM tb_applicativi"
	CALL ExecuteMultipleSql(connContent, sql, true)
	'compatta il database per azzerare tutti gli i contatori
	CALL CompactDatabase(connContent)
	'inserisce dummy record su tabelle applicativi e pagine per spostare l'id di base
	sql = "INSERT INTO tb_applicativi (sito_nome, sito_url) VALUES('dummy', 'dummy'); " + _
		  "INSERT INTO tb_pagine (pag_sito, pag_url) VALUES(1, 'dummy');" + _
		  "INSERT INTO tb_pagine (pag_sito, pag_url) VALUES(1, 'dummy');" + _
		  "INSERT INTO tb_pagine (pag_sito, pag_url) VALUES(1, 'dummy');" + _
		  "INSERT INTO tb_pagine (pag_sito, pag_url) VALUES(1, 'dummy');" + _
		  "INSERT INTO tb_pagine (pag_sito, pag_url) VALUES(1, 'dummy');" + _
		  "INSERT INTO tb_pagine (pag_sito, pag_url) VALUES(1, 'dummy');" + _
		  "INSERT INTO tb_pagine (pag_sito, pag_url) VALUES(1, 'dummy');" + _
		  "INSERT INTO tb_pagine (pag_sito, pag_url) VALUES(1, 'dummy');" + _
		  "INSERT INTO tb_pagine (pag_sito, pag_url) VALUES(1, 'dummy');" + _
		  "DELETE FROM tb_applicativi; " + _
		  "DELETE FROM tb_pagine; "
	CALL ExecuteMultipleSql(connContent, sql, true)
	
	connContent.beginTrans

	'trasferisce dati da dbBanner a dbContent
	rs.open "tb_tipiBanner", connContent, adOpenForwardOnly, adLockOptimistic, adCmdTable
	rsT.open "tb_tipiBanner", conn, adOpenForwardOnly, adLockOptimistic, adCmdTable
	while not rsT.eof
		rs.AddNew
		rs("tipoB_nome") = rsT("tipoB_nome")
		rs.Update
		rsT.MoveNext
	wend
	rs.close
	rsT.close
	
	rs.open "tb_Applicativi", connContent, adOpenForwardOnly, adLockOptimistic, adCmdTable
	rsT.open "tb_Siti", conn, adOpenForwardOnly, adLockOptimistic, adCmdTable
	while not rsT.eof
		rs.AddNew
		rs("sito_nome") = rsT("sito_nome")
		rs("sito_URL") = rsT("sito_URL")
		rs.Update
		rsT.MoveNext
	wend
	rs.close
	rsT.close
		
	rs.open "tb_pagine", connContent, adOpenForwardOnly, adLockOptimistic, adCmdTable
	rsT.open "tb_pagine", conn, adOpenForwardOnly, adLockOptimistic, adCmdTable
	while not rsT.eof
		rs.AddNew
		rs("pag_url") = rsT("pag_url")
		rs("pag_cat") = rsT("pag_cat")
		rs("pag_sito") = rsT("pag_sito")
		rs.Update
		rsT.MoveNext
	wend
	rs.close
	rsT.close
	
	set rsb_s = Server.CreateObject("ADODB.RecordSet")
	set rsb_d = Server.CreateObject("ADODB.RecordSet")
	set rsp_s = Server.CreateObject("ADODB.RecordSet")
	set rsp_d = Server.CreateObject("ADODB.RecordSet")
	'inserisco le aziende in [dbContent] come contatti e modifico l'ID in tb_banner per integrita'
	rs.open "tb_indirizzario", connContent, adOpenKeySet, adLockOptimistic, adCmdTable
	rsb_d.open "tb_banner", connContent, adOpenKeySet, adLockOptimistic, adCmdTable
	rsp_d.open "rel_banner_pagine", connContent, adOpenKeySet, adLockOptimistic, adCmdTable
	
	rsT.open "tb_aziende", conn, adOpenForwardOnly, adLockOptimistic, adCmdTable
	while not rsT.eof
		rs.AddNew
		rs("NomeOrganizzazioneElencoIndirizzi") = rsT("az_nome")
		rs("IndirizzoElencoIndirizzi") = rsT("az_ind")
		rs("NoteElencoIndirizzi") = "da applicativo banner campo az_cnt: "& rsT("az_cnt")
		rs("IsSocieta") = 1
		rs("ModoRegistra") = rsT("az_nome")
		rs("DataIscrizione") = Date
		rs("LockedByApplication") = 1
		rs("ApplicationsLocker") = NEXTBANNER &", "
		rs.update
		
		sql = " INSERT INTO rel_rub_ind (id_indirizzo, id_rubrica) " &_
			  " VALUES(" & rs("IDElencoIndirizzi") & ", " & Application("NextBanner_RubricaAziende") & ")"
		CALL connContent.execute(sql)
		
		'inserisce banner relativi all'azienda
		sql = "SELECT * FROM tb_banner WHERE ban_az=" & rsT("az_id")
		rsb_s.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
		while not rsb_s.eof
			rsb_d.addNew
			rsb_d("ban_nome") = rsb_s("ban_nome")
			rsb_d("ban_image") = rsb_s("ban_image")
			rsb_d("ban_link") = rsb_s("ban_link")
			rsb_d("ban_alt") = rsb_s("ban_alt")
			rsb_d("ban_tipo") = rsb_s("ban_tipo")
			rsb_d("ban_az") = rs("IDElencoIndirizzi")
			rsb_d.update
			
			'inserisce i collegamenti con le posizioni di pubblicazione di ogni banner
			sql = "SELECT * FROM rel_banner_pagine WHERE rbp_banner=" & rsb_s("ban_id")
			rsp_s.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
			while not rsp_s.eof
				rsp_d.addNew
				rsp_d("rbp_impress_iniz") = rsp_s("rbp_impress_iniz")
				rsp_d("rbp_impress") = rsp_s("rbp_impress")
				rsp_d("rbp_data_iniz") = rsp_s("rbp_data_iniz")
				rsp_d("rbp_data_fine") = rsp_s("rbp_data_fine")
				rsp_d("rbp_click_iniz") = rsp_s("rbp_click_iniz")
				rsp_d("rbp_click") = rsp_s("rbp_click")
				rsp_d("rbp_pag") = rsp_s("rbp_pag")
				rsp_d("rbp_banner") = rsb_d("ban_id")
				rsp_d.update
				rsp_s.movenext
			wend
			rsp_s.close
			rsb_s.movenext
		wend
		rsb_s.close
		rsT.movenext
	wend
	rs.close
	rsb_d.close
	rsp_d.close
	rsT.close
	
	
	connContent.CommitTrans
	connContent.Close
	set connContent = nothing
end sub
'...........................................................................................
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 3
'...........................................................................................
'aggiunge campo su tabella dei log per registrazione header richiesta
'...........................................................................................
sql = "ALTER TABLE rel_log_click ADD COLUMN rlc_request TEXT WITH COMPRESSION NULL"
CALL DB.Execute(sql, 3)
'*******************************************************************************************


'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 4
'...........................................................................................
'cancellazione relazioni vecchia struttra tabellare non piu' utlizzata 
'...........................................................................................
sql = "ALTER TABLE rel_log_impress DROP CONSTRAINT rel_banner_paginerel_log_impress; " & _
	  "ALTER TABLE tb_aziende  DROP CONSTRAINT tb_admintb_aziende; " & _
	  "ALTER TABLE tb_siti DROP CONSTRAINT tb_admintb_siti; " & _
	  "ALTER TABLE tb_banner DROP CONSTRAINT tb_aziendetb_bunner; " & _
	  "ALTER TABLE rel_banner_pagine DROP CONSTRAINT tb_bunnerrel_bunner_pagine; " & _
	  "ALTER TABLE rel_banner_pagine DROP CONSTRAINT tb_paginerel_banner_pagine; " & _
	  "ALTER TABLE tb_pagine DROP CONSTRAINT tb_sititb_pagine; " & _
	  "ALTER TABLE tb_banner DROP CONSTRAINT tb_tipiBunnertb_bunner; " & _
	  "ALTER TABLE rel_log_click DROP CONSTRAINT rel_banner_paginerel_log_click; "
CALL DB.Execute(sql, 4)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 5
'...........................................................................................
'cancellazione indici vecchia struttra tabellare non piu' utlizzata 
'...........................................................................................
sql = "DROP INDEX adm_id ON tb_admin; " &_
	  "DROP INDEX bun_id ON tb_banner; " &_
	  "DROP INDEX cli_id ON tb_siti; " &_
	  "DROP INDEX idAzienda ON tb_aziende; " &_
	  "DROP INDEX pag_id ON tb_pagine; " &_
	  "DROP INDEX rbp_id ON rel_banner_pagine; " &_
	  "DROP INDEX rl_id ON rel_log_click; " &_
	  "DROP INDEX rl_id ON rel_log_impress; " &_
	  "DROP INDEX tipoB_id ON tb_tipiBanner; "
CALL DB.Execute(sql, 5)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 6
'...........................................................................................
'cancellazione tabelle vecchia struttra tabellare non piu' utlizzata 
'...........................................................................................
sql = "DROP TABLE rel_log_impress; " &_
	  "DROP TABLE tb_pagine; " &_
	  "DROP TABLE tb_siti;" &_
	  "DROP TABLE tb_tipiBanner;" &_
	  "DROP TABLE tb_admin;" &_
	  "DROP TABLE tb_banner;" &_
	  "DROP TABLE tb_aziende;" &_
	  "DROP TABLE rel_banner_pagine;"
CALL DB.Execute(sql, 6)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 7
'...........................................................................................
'aggiunge tabella per i log degli impressions
'...........................................................................................
sql = "CREATE TABLE rel_log_impress (" &_
	  "		rli_ID COUNTER CONSTRAINT PK_rel_log_impress PRIMARY KEY, " &_
	  "		rli_data DATETIME NULL, " &_
	  "		rli_banner INTEGER NULL, " &_
	  "		rli_pagina INTEGER NULL " &_
	  ")"
CALL DB.Execute(sql, 7)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 8
'...........................................................................................
'aggiunge tabella per lo storico per ora dei log degli impressions
'...........................................................................................
sql = "CREATE TABLE tb_storico_impress (" & _
	  "		sti_ID COUNTER CONSTRAINT PK_tb_storico_impress PRIMARY KEY, " & _
	  "		sti_data DATETIME NULL, " & _
	  "		sti_ora INTEGER NULL, " & _
	  "		sti_count INTEGER NULL, " & _
	  "		sti_banner INTEGER NULL, " & _
	  "		sti_pagina INTEGER NULL " & _
	  ")"
CALL DB.Execute(sql, 8)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 9
'...........................................................................................
'aggiunge tabella per lo storico per ora dei log degli impressions
'...........................................................................................
sql = " ALTER TABLE rel_log_click ADD "+ _
	  "		rlc_ip VARCHAR(15) NULL "
CALL DB.Execute(sql, 9)
'*******************************************************************************************

%>
<% '........................................................................................... %>
<!--#INCLUDE FILE="Update__FileFooter.asp" -->
<% '........................................................................................... %>
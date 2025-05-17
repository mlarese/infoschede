<%
'...........................................................................................
'........................................................................................... 
'libreria di funzioni che contiene tutti gli aggiornamenti per il NEXT-memo
'...........................................................................................
'...........................................................................................

'*******************************************************************************************
'INSTALLAZIONE QUESTIONARIO
'...........................................................................................
function Install__Questionario(conn)
	Install__Questionario = _
		"CREATE TABLE dbo.qtb_questionari ( " + _
		"quest_id " & SQL_PrimaryKey(conn, "qtb_questionari") + ", " + _
		"quest_nome_it " + SQL_CharField(Conn, 255) + ", " + _
		"quest_cod_inizio " + SQL_CharField(Conn, 10) + ", " + _
		"quest_cod_fine " + SQL_CharField(Conn, 10) + ", " + _
		"quest_archiviato bit NULL " + _
		"); " + _
		"CREATE TABLE dbo.qtb_domande ( " + _
		"dom_id " & SQL_PrimaryKey(conn, "qtb_domande") + ", " + _
		"dom_testo_it " + SQL_CharField(Conn, 0) +" NULL, " + _
		"dom_codice " + SQL_CharField(Conn, 10) + ", " + _
		"dom_ordine INT NULL " + _
		"); " + _
		"CREATE TABLE dbo.qtb_rel_quest_dom ( " + _
		"rqd_id " & SQL_PrimaryKey(conn, "qtb_rel_quest_dom") + ", " + _
		"rqd_questionario_id INT NULL, " + _
		"rqd_domanda_id INT NULL, " + _
		"rqd_ordine INT NULL " + _
		"); " + _
		SQL_AddForeignKey(conn, "qtb_rel_quest_dom", "rqd_questionario_id", "qtb_questionari", "quest_id", true, "") + _
		SQL_AddForeignKey(conn, "qtb_rel_quest_dom", "rqd_domanda_id", "qtb_domande", "dom_id", true, "") + _
		"CREATE TABLE dbo.qtb_risposte ( " + _
		"risp_id " & SQL_PrimaryKey(conn, "qtb_risposte") + ", " + _
		"risp_domanda_id INT NULL, " + _
		"risp_testo_it " + SQL_CharField(Conn, 0) + ", " + _
		"risp_codice " + SQL_CharField(Conn, 10) + ", " + _
		"risp_ordine INT NULL " + _	
		"); " + _
		SQL_AddForeignKey(conn, "qtb_risposte", "risp_domanda_id", "qtb_domande", "dom_id", true, "") + _
		"CREATE TABLE dbo.qtb_compilazioni ( " + _
		"comp_id " & SQL_PrimaryKey(conn, "qtb_compilazioni") + ", " + _
		"comp_questionario_id INT NULL, " + _
		"comp_data smalldatetime NULL, " +_
		"comp_operatore_id INT NULL " + _
		"); " + _
		SQL_AddForeignKey(conn, "qtb_compilazioni", "comp_questionario_id", "qtb_questionari", "quest_id", false, "") + _
		"CREATE TABLE dbo.qtb_risposte_date ( " + _
		"risp_data_id " & SQL_PrimaryKey(conn, "qtb_risposte_date") + ", " + _
		"risp_data_compilazione_id INT NULL, " + _
		"risp_data_risposta_id INT NULL " + _
		"); " + _	
		SQL_AddForeignKey(conn, "qtb_risposte_date", "risp_data_compilazione_id", "qtb_compilazioni", "comp_id", true, "") + _
		
		SQL_AddForeignKey(conn, "qtb_risposte_date", "risp_data_risposta_id", "qtb_risposte", "risp_id", true, "") 
		
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO QUESTIONARIO 1
'...........................................................................................
'aggiunta campi testo per lo storico delle compilazioni
'...........................................................................................
function Aggiornamento__Questionario__1(conn)
	Aggiornamento__Questionario__1 = _
		" ALTER TABLE qtb_risposte_date ADD " + _
		" 	risp_data_testo_risposta_it" + SQL_CharField(Conn, 0) + " NULL, " + _
		"	risp_data_testo_domanda_it " + SQL_CharField(Conn, 0) +" NULL; " + _
		" ALTER TABLE qtb_compilazioni ADD " + _
		"	comp_nome_questionario_it " + SQL_CharField(Conn, 255) + " ;"
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO QUESTIONARIO 2
'...........................................................................................
'correzioni su in vincolo d'integrità e sulla lunghezza dei campi "codice"
'...........................................................................................
function Aggiornamento__Questionario__2(conn)
	Aggiornamento__Questionario__2 = _
		SQL_RemoveForeignKey(conn, "qtb_risposte_date", "", "", true, "FK_qtb_risposte_date__qtb_risposte") + _		
		SQL_AddForeignKey(conn, "qtb_risposte_date", "risp_data_risposta_id", "qtb_risposte", "risp_id", false, "") + _
		" ALTER TABLE qtb_questionari ALTER COLUMN quest_cod_inizio " + SQL_CharField(Conn, 20) + "; " + _
		" ALTER TABLE qtb_questionari ALTER COLUMN quest_cod_fine " + SQL_CharField(Conn, 20) + "; " + _
		" ALTER TABLE qtb_domande ALTER COLUMN dom_codice " + SQL_CharField(Conn, 20) + "; " + _
		" ALTER TABLE qtb_risposte ALTER COLUMN risp_codice " + SQL_CharField(Conn, 20) + "; "		
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO QUESTIONARIO 3
'...........................................................................................
'Aggiunge gestione domande a scelta multipla
'...........................................................................................
function Aggiornamento__Questionario__3(conn)
	Aggiornamento__Questionario__3 = _
		" ALTER TABLE dbo.qtb_domande  ADD " + _
		"	dom_scelta_multipla bit NOT NULL DEFAULT 0" + _
		" ;"
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO QUESTIONARIO 4
'...........................................................................................
'Aggiunge gestione ordinamento delle domande e risposte
'...........................................................................................
function Aggiornamento__Questionario__4(conn)
	Aggiornamento__Questionario__4 = _
		" ALTER TABLE dbo.qtb_risposte_date  ADD " + _
		"	risp_data_ordine_domanda INT NOT NULL DEFAULT 0," + _
		"	risp_data_ordine_risposta INT NOT NULL DEFAULT 0" + _
		" ;"
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO QUESTIONARIO 5
'...........................................................................................
' Creazione tabella rilevazione accessi
'...........................................................................................
function Aggiornamento__Questionario__5(conn)
	Aggiornamento__Questionario__5 = _
		" CREATE TABLE dbo.qtb_log_accessi_iat (" +_
		"qtb_log_id " & SQL_PrimaryKey(conn, "qtb_log_accessi_iat") + ", " + _
		"qtb_log_iat_id INT NULL, " + _
		"qtb_log_data smalldatetime NULL, " +_
		"qtb_log_num_per_in INT NOT NULL DEFAULT 0, " + _
		"qtb_log_num_per_out INT NOT NULL DEFAULT 0, " + _
		"qtb_log_num_group_in INT NOT NULL DEFAULT 0 " + _
		"); " + _
		SQL_AddForeignKey(conn, "qtb_log_accessi_iat", "qtb_log_iat_id", "tb_iat", "iat_id", false, "") + _
		" ;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO QUESTIONARIO 6
'...........................................................................................
'Aggiunge risposta di default tra le domande
'...........................................................................................
function Aggiornamento__Questionario__6(conn)
	Aggiornamento__Questionario__6 = _
		" ALTER TABLE dbo.qtb_risposte  ADD " + _
		"	risp_default bit NOT NULL DEFAULT 0" + _
		" ;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO QUESTIONARIO 7
'...........................................................................................
'Aggiunge colonna con i dati row per accelerare la ricerca rilevazione accessi
'...........................................................................................
function Aggiornamento__Questionario__7(conn)
	Aggiornamento__Questionario__7 = _
		" ALTER TABLE dbo.qtb_log_accessi_iat  ADD " + _
		"	row_data " + SQL_CharField(Conn, 27) + " NOT NULL DEFAULT '' " + _
		" ;" +_
		" CREATE UNIQUE INDEX Idx_log_accessi_rowdata ON qtb_log_accessi_iat(row_data);"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO QUESTIONARIO 8
'...........................................................................................
'	Nicola
'	23/07/2010
'...........................................................................................
'Aggiunge indice per velocizzare il recupero dei dati delle compilazioni
'...........................................................................................
function Aggiornamento__Questionario__8(conn)
	Aggiornamento__Questionario__8 = _
		" CREATE NONCLUSTERED INDEX IX_qtb_risposte_date ON qtb_risposte_date "  + _
		"	( risp_data_compilazione_id DESC ) "
end function
'*******************************************************************************************


%>
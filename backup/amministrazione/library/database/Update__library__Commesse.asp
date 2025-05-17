<%
'...........................................................................................
'........................................................................................... 
'libreria di funzioni che contiene tutti gli aggiornamenti per l'applicativo Gestione commesse
'...........................................................................................
'...........................................................................................



'*******************************************************************************************
'AGGIORNAMENTO COMMESSE 1
'...........................................................................................
' Nicola 20/11/2013
'...........................................................................................
' creazione nuove tabelle per l'applicativo gestione commesse
'...........................................................................................
function Aggiornamento__COMMESSE__1(conn)
    Aggiornamento__COMMESSE__1 = _
			"CREATE TABLE " + SQL_Dbo(conn) + "gctb_commesse(" + _
			"	gco_id " + SQL_PrimaryKey(conn, "gctb_commesse") + ", " + _
			"	gco_nome " + SQL_CharField(Conn, 500) + " NULL, " + _
			" 	gco_cliente_id int NULL, " + _
			"	gco_note " + SQL_CharField(Conn, 0) + " NULL, " + _
			" 	gco_data_stipula DATETIME NULL, " + _
			" 	gco_data_consegna DATETIME NULL, " + _
			" 	gco_data_chiusura DATETIME NULL, " + _
			" 	gco_insAdmin_id int NULL, " + _
			" 	gco_insData DATETIME NULL, " + _
			" 	gco_modAdmin_id int NULL, " + _
			" 	gco_modData DATETIME NULL " + _
			"); " + _
			SQL_AddForeignKey(conn, "gctb_commesse", "gco_cliente_id", "tb_indirizzario", "IdElencoIndirizzi", true, "") + _
			"CREATE TABLE " + SQL_Dbo(conn) + "gctb_tipo_attivita(" + _
			"	ta_id " + SQL_PrimaryKey(conn, "gctb_tipo_attivita") + ", " + _
			"	ta_nome " + SQL_CharField(Conn, 250) + " NULL" + _
			"); " + _
			"CREATE TABLE " + SQL_Dbo(conn) + "gctb_attivita(" + _
			"	gat_id " + SQL_PrimaryKey(conn, "gctb_attivita") + ", " + _
			" 	gat_commessa_id int NULL, " + _
			" 	gat_tipo_id int NULL, " + _
			" 	gat_data DATETIME NULL, " + _
			" 	gat_data_registrazione DATETIME NULL, " + _
			" 	gat_ora_inizio DATETIME NULL, " + _
			" 	gat_ora_fine DATETIME NULL, " + _
			" 	gat_tempo_minuti int NULL, " + _
			"	gat_lavoro_eseguito " + SQL_CharField(Conn, 500) + " NULL, " + _
			"	gat_operatore_id INT NULL, " + _
			"	gat_note " + SQL_CharField(Conn, 0) + " NULL, " + _
			"	gat_spese money NULL, " + _
			" 	gat_insAdmin_id int NULL, " + _
			" 	gat_insData DATETIME NULL, " + _
			" 	gat_modAdmin_id int NULL, " + _
			" 	gat_modData DATETIME NULL " + _
			"); " + _
			SQL_AddForeignKey(conn, "gctb_attivita", "gat_commessa_id", "gctb_commesse", "gco_id", true, "") + _
			SQL_AddForeignKey(conn, "gctb_attivita", "gat_tipo_id", "gctb_tipo_attivita", "ta_id", true, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO SPECIALE COMMESSE 2
'...........................................................................................
' Giacomo 16/01/2014
'...........................................................................................
' aggiunge parametro per indicare l'id della rubrica dei clienti attivi
'...........................................................................................
function Aggiornamento__COMMESSE__2(conn)
	Aggiornamento__COMMESSE__2 = " SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__COMMESSE__2(conn)
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & COMMESSE)) <> "" then
		CALL AddParametroSito(conn, "COMMESSE_ID_RUBRICA_CLIENTI_ATTIVI", _
									0, _
									"id della rubrica dei clienti attivi.", _
									"", _
									adIDispatch, _
									0, _
									"", _
									1, _
									1, _
									COMMESSE, _
									null, null, null, null, null)
	end if
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO COMMESSE 3
'...........................................................................................
' Giacomo 19/12/2014
'...........................................................................................
' creazione nuove tabelle per l'applicativo gestione commesse
'...........................................................................................
function Aggiornamento__COMMESSE__3(conn)
    Aggiornamento__COMMESSE__3 = _
		"CREATE TABLE " + SQL_Dbo(conn) + "gctb_attivita_spese(" + _
		"	gas_id " + SQL_PrimaryKey(conn, "gctb_attivita_spese") + ", " + _
		"	gas_descrizione " + SQL_CharField(Conn, 500) + " NULL, " + _
		"	gas_importo money NULL, " + _
		"	gas_id_attivita int NOT NULL " + _
		"); " + _
		SQL_AddForeignKey(conn, "gctb_attivita_spese", "gas_id_attivita", "gctb_attivita", "gat_id", true, "") + _
		"CREATE TABLE " + SQL_Dbo(conn) + "tb_admin_orario(" + _
		"	ao_id " + SQL_PrimaryKey(conn, "tb_admin_orario") + ", " + _
		"	ao_id_admin int NOT NULL, " + _
		" 	ao_data_dal DATETIME NULL, " + _
		" 	ao_data_al DATETIME NULL, " + _
		" 	ao_min_lav_lun int NULL, " + _
		" 	ao_min_lav_mar int NULL, " + _
		" 	ao_min_lav_mer int NULL, " + _
		" 	ao_min_lav_gio int NULL, " + _
		" 	ao_min_lav_ven int NULL, " + _
		" 	ao_min_lav_sab int NULL, " + _
		" 	ao_min_lav_dom int NULL " + _
		"); " + _
		SQL_AddForeignKey(conn, "tb_admin_orario", "ao_id_admin", "tb_admin", "ID_admin", true, "") + _
		" INSERT INTO gctb_attivita_spese (gas_id_attivita, gas_importo, gas_descrizione) " + _
		" SELECT gat_id, gat_spese, '' FROM gctb_attivita WHERE gat_spese > 0; " + _
		" ALTER TABLE gctb_attivita DROP COLUMN gat_spese; "
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO COMMESSE 4
'...........................................................................................
' Giacomo 22/12/2014
'...........................................................................................
' creazione nuova tabella per log invio attività giornaliere
'...........................................................................................
function Aggiornamento__COMMESSE__4(conn)
    Aggiornamento__COMMESSE__4 = _
		"CREATE TABLE " + SQL_Dbo(conn) + "gctb_log_giorni_completati(" + _
		"	lgc_id " + SQL_PrimaryKey(conn, "gctb_log_giorni_completati") + ", " + _
		"	lgc_id_operatore int NOT NULL, " + _
		"	lgc_giorno_completato smalldatetime NOT NULL, " + _
		"	lgc_min_previsti int NULL, " + _
		"	lgc_min_effettuati int NULL, " + _
		"	lgc_data_invio smalldatetime NOT NULL " + _
		"); "
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO COMMESSE 
'...........................................................................................
' Nicola 07/01/2015
'...........................................................................................
' creazione nuovo campo tabella commesse
'...........................................................................................
function Aggiornamento__COMMESSE__5(conn)
    Aggiornamento__COMMESSE__5 = _
		"ALTER TABLE " + SQL_Dbo(conn) + "gctb_commesse ADD " + _
			"	gco_codice " + SQL_CharField(Conn, 20) + " NULL " + _
		"; "
end function
'*******************************************************************************************



%>
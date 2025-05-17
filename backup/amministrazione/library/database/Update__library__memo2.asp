<%
'...........................................................................................
'........................................................................................... 
'libreria di funzioni che contiene tutti gli aggiornamenti per il NEXT-memo 2.0
'...........................................................................................
'...........................................................................................

'*******************************************************************************************
'INSTALLAZIONE MEMO 2.0
'...........................................................................................
function Install__MEMO2(conn)
	Install__MEMO2 = _
		"CREATE TABLE " + SQL_Dbo(conn) + "mtb_documenti_categorie (" + _
		"	catC_id  " & SQL_PrimaryKey(conn, "mtb_documenti_categorie") + ", " + _
			SQL_MultiLanguageFieldComplete(conn, "catC_nome_<lingua> " + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
		"	catC_codice " + SQL_CharField(Conn, 100) + " NULL, " + _
		"	catC_foglia BIT NULL ," + _
		"	catC_livello INTEGER NULL ," + _
		"	catC_padre_id INTEGER NULL ," + _
		"	catC_ordine INTEGER NULL ," + _
		"	catC_ordine_assoluto " + SQL_CharField(Conn, 255) + " NULL ," + _				
			SQL_MultiLanguageFieldComplete(conn, "catC_descr_<lingua> " + SQL_CharField(Conn, 0) + " NULL ") + ", " + _
		"	catC_tipologia_padre_base INTEGER NULL ," + _
		"	catC_visibile BIT NULL , " + _
		"	catC_albero_visibile BIT NULL , " + _
		"	catC_tipologie_padre_lista " + SQL_CharField(Conn, 255) + " NULL" + _
		"); " + _
		"CREATE TABLE " + SQL_Dbo(conn) + "mtb_documenti (" + _
		"	doc_id " & SQL_PrimaryKey(conn, "mtb_documenti") + ", " + _
		"	doc_numero " + SQL_CharField(Conn, 100) + " NULL ," + _
			SQL_MultiLanguageFieldComplete(conn, "doc_titolo_<lingua> " + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
			SQL_MultiLanguageFieldComplete(conn, "doc_estratto_<lingua> " + SQL_CharField(Conn, 0) + " NULL ") + ", " + _
		"	doc_pubblicazione DATETIME NULL ," + _
		"	doc_scadenza DATETIME NULL ," + _
			SQL_MultiLanguageFieldComplete(conn, "doc_file_<lingua> " + SQL_CharField(Conn, 0) + " NULL ") + ", " + _
		"	doc_visibile BIT NULL ," + _
		"	doc_protetto BIT NULL ," + _
		"	doc_categoria_id INTEGER NOT NULL " + "); " + _	
		"CREATE TABLE " + SQL_Dbo(conn) + "log_documenti (" + _
		"	log_id " & SQL_PrimaryKey(conn, "log_documenti") + ", " + _
		"	log_ut_id INTEGER NULL, " + _
		"	log_dip_id INTEGER NULL, " + _
		"	log_doc_id INTEGER NULL, " + _
		"	log_data DATETIME NULL" + _
		"); " + _	
		" CREATE TABLE " + SQL_Dbo(conn) + "mtb_profili (" + _
		"	pro_id " & SQL_PrimaryKey(conn, "mtb_profili") + ", " + _
			SQL_MultiLanguageFieldComplete(conn, "pro_nome_<lingua> " + SQL_CharField(Conn, 255) + " NULL ") + _
		"); " + _
		"CREATE TABLE " + SQL_Dbo(conn) + "mrel_doc_profili (" + _
		"	rdp_id " & SQL_PrimaryKey(conn, "mrel_doc_profili") + ", " + _
		"	rdp_doc_id INTEGER NULL, " + _
		"	rdp_profilo_id INTEGER NULL " + _
		"); " + _
		"CREATE TABLE " + SQL_Dbo(conn) + "mrel_doc_admin (" + _
		"	rda_id " & SQL_PrimaryKey(conn, "mrel_doc_admin") + ", " + _
		"	rda_doc_id INTEGER NULL, " + _
		"	rda_admin_id INTEGER NULL " + _
		"); " + _
		"CREATE TABLE " + SQL_Dbo(conn) + "mrel_doc_utenti (" + _
		"	rdu_id " & SQL_PrimaryKey(conn, "mrel_doc_utenti") + ", " + _
		"	rdu_doc_id INTEGER NULL, " + _
		"	rdu_utenti_id INTEGER NULL " + _
		"); " + _
		"CREATE TABLE " + SQL_Dbo(conn) + "mrel_profili_admin (" + _
		"	rpa_id " & SQL_PrimaryKey(conn, "mrel_profili_admin") + ", " + _
		"	rpa_profilo_id INTEGER NULL, " + _
		"	rpa_admin_id INTEGER NULL " + _
		"); " + _
		"CREATE TABLE " + SQL_Dbo(conn) + "mrel_profili_utenti (" + _
		"	rpu_id " & SQL_PrimaryKey(conn, "mrel_profili_utenti") + ", " + _
		"	rpu_profilo_id INTEGER NULL, " + _
		"	rpu_utenti_id INTEGER NULL " + _
		"); " + _
		SQL_AddForeignKey(conn, "mrel_profili_utenti", "rpu_profilo_id", "mtb_profili", "pro_id", true, "") + _
		SQL_AddForeignKey(conn, "mrel_profili_utenti", "rpu_utenti_id", "tb_utenti", "ut_ID", true, "") + _
		SQL_AddForeignKey(conn, "mtb_documenti", "doc_categoria_id", "mtb_documenti_categorie", "catC_id", true, "") + _
		SQL_AddForeignKey(conn, "mrel_doc_profili", "rdp_doc_id", "mtb_documenti", "doc_id", true, "") + _
		SQL_AddForeignKey(conn, "mrel_doc_profili", "rdp_profilo_id", "mtb_profili", "pro_id", true, "") + _
		SQL_AddForeignKey(conn, "log_documenti", "log_doc_id", "mtb_documenti", "doc_id", true, "") + _
		SQL_AddForeignKey(conn, "mrel_doc_admin", "rda_doc_id", "mtb_documenti", "doc_id", true, "") + _
		SQL_AddForeignKey(conn, "mrel_doc_admin", "rda_admin_id", "tb_admin", "ID_admin", true, "") + _
		SQL_AddForeignKey(conn, "mrel_doc_utenti", "rdu_doc_id", "mtb_documenti", "doc_id", true, "") + _
		SQL_AddForeignKey(conn, "mrel_doc_utenti", "rdu_utenti_id", "tb_utenti", "ut_ID", true, "") + _
		SQL_AddForeignKey(conn, "mrel_profili_admin", "rpa_profilo_id", "mtb_profili", "pro_id", true, "") + _
		SQL_AddForeignKey(conn, "mrel_profili_admin", "rpa_admin_id", "tb_admin", "ID_admin", true, "") + _
		";"
end function
'*******************************************************************************************	

				
'*******************************************************************************************
'AGGIORNAMENTO SPECIALE NEXT-MEMO 2.0  1
'...........................................................................................
'ClassCategorie: aggiunge il campo per la gestione della lista degli IDs dei padri
'...........................................................................................
function AggiornamentoSpeciale__MEMO2__1(DB, rs, version)
    CALL AggiornamentoSpeciale__FRAMEWORK_CORE__ListaPadriCategorie(DB, rs, version, "mtb_documenti_categorie", "catC")
    AggiornamentoSpeciale__MEMO2__1 = "SELECT * FROM AA_versione"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-MEMO 2.0  1
'...........................................................................................
' Giacomo 10/02/2011
'...........................................................................................
' Creazioni tabelle per gestire gli impegni (appuntamenti) - SEZIONE AGENDA
'...........................................................................................
function Aggiornamento__MEMO2__1(conn)
    Aggiornamento__MEMO2__1 = _
			"CREATE TABLE " + SQL_Dbo(conn) + "mtb_configurazione_impegni(" + _
			"	coi_id " & SQL_PrimaryKey(conn, "mtb_configurazione_impegni") + ", " + _
			"	coi_giorno int NULL, " + _
			"	coi_dal DATETIME NULL ," + _
			"	coi_al DATETIME NULL" + _
			"); " + _
			"CREATE TABLE " + SQL_Dbo(conn) + "mtb_tipi_impegni(" + _
			"	tim_id " & SQL_PrimaryKey(conn, "mtb_tipi_impegni") + ", " + _
			SQL_MultiLanguageFieldComplete(conn, "tim_nome_<lingua> " + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
			"	tim_colore " + SQL_CharField(Conn, 7) + " NULL, " + _
			SQL_MultiLanguageFieldComplete(conn, "tim_descrizione_<lingua> " + SQL_CharField(Conn, 0) + " NULL ") + _
			"); " + _
			"CREATE TABLE " + SQL_Dbo(conn) + "mtb_impegni(" + _
			"	imp_id " & SQL_PrimaryKey(conn, "mtb_impegni") + ", " + _
			SQL_MultiLanguageFieldComplete(conn, "imp_titolo_<lingua> " + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
			SQL_MultiLanguageFieldComplete(conn, "imp_descrizione_<lingua> " + SQL_CharField(Conn, 0) + " NULL ") + ", " + _
			"	imp_tipo_id INTEGER NULL, " + _
			"	imp_data_ora_inizio DATETIME NULL ," + _
			"	imp_data_ora_fine DATETIME NULL ," + _
			"	imp_protetto BIT NULL" + _			
			"); " + _
			"CREATE TABLE " + SQL_Dbo(conn) + "mrel_impegni_utenti(" + _
			"	riu_id " & SQL_PrimaryKey(conn, "mrel_impegni_utenti") + ", " + _
			"	riu_impegno_id INTEGER NULL, " + _
			"	riu_utente_id INTEGER NULL " + _
			"); " + _
			"CREATE TABLE " + SQL_Dbo(conn) + "mrel_impegni_profili(" + _
			"	rip_id " & SQL_PrimaryKey(conn, "mrel_impegni_profili") + ", " + _
			"	rip_impegno_id INTEGER NULL, " + _
			"	rip_profilo_id INTEGER NULL " + _
			"); " + _
			SQL_AddForeignKey(conn, "mtb_impegni", "imp_tipo_id", "mtb_tipi_impegni", "tim_id", false, "") + _
			SQL_AddForeignKey(conn, "mrel_impegni_utenti", "riu_impegno_id", "mtb_impegni", "imp_id", true, "") + _
			SQL_AddForeignKey(conn, "mrel_impegni_utenti", "riu_utente_id", "tb_utenti", "ut_ID", true, "") + _
			SQL_AddForeignKey(conn, "mrel_impegni_profili", "rip_impegno_id", "mtb_impegni", "imp_id", true, "") + _
			SQL_AddForeignKey(conn, "mrel_impegni_profili", "rip_profilo_id", "mtb_profili", "pro_id", true, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO SPECIALE NEXT-MEMO 2.0  2
'...........................................................................................
' Giacomo 10/02/2011
'...........................................................................................
' aggiunge parametro per gestire gli impegni (appuntamenti) - SEZIONE AGENDA
'...........................................................................................
function Aggiornamento__MEMO2__2(conn)
	Aggiornamento__MEMO2__2 = " INSERT INTO tb_siti_descrittori_raggruppamenti(sdr_titolo_it, sdr_ordine,sdr_personalizzato) " & _
							  " SELECT 'Agenda Next-Memo 2',MAX(sdr_ordine)+1,1 FROM tb_siti_descrittori_raggruppamenti "
end function

function AggiornamentoSpeciale__MEMO2__2(conn)
	dim sql, id_ragg
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & NEXTMEMO2)) <> "" then
		sql = "SELECT sdr_id FROM tb_siti_descrittori_raggruppamenti WHERE sdr_titolo_it LIKE 'Agenda Next-Memo 2'" 
		id_ragg = cIntero(GetValueList(conn, NULL, sql))
		CALL AddParametroSito(conn, "AGENDA_ATTIVA", _
									id_ragg, _
									"attiva la gestione degli impegni/appuntamenti", _
									"", _
									adBoolean, _
									0, _
									"", _
									1, _
									1, _
									NEXTMEMO2, _
									"false", null, null, null, null)
		CALL AddParametroSito(conn, "AGENDA_INTERVALLO_CALENDARIO", _
									id_ragg, _
									"imposta l'intervallo di tempo, in minuti, tra due periodi", _
									"", _
									adNumeric, _
									0, _
									"", _
									1, _
									1, _
									NEXTMEMO2, _
									30, null, null, null, null)
		AggiornamentoSpeciale__MEMO2__2 = " SELECT * FROM AA_Versione "
	end if
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO NEXT-MEMO 2.0  3
'...........................................................................................
' Giacomo 23/02/2011
'...........................................................................................
' Creazione tabella per log degli avvisi di impegno spediti
' e aggiunta alcuni campi a mtb_impegni
'...........................................................................................
function Aggiornamento__MEMO2__3(conn)
    Aggiornamento__MEMO2__3 = _
			"CREATE TABLE " + SQL_Dbo(conn) + "mtb_log_avvisi_spediti(" + _
			"	las_id " & SQL_PrimaryKey(conn, "mtb_log_avvisi_spediti") + ", " + _
			"	las_impegno_id int NULL, " + _
			"	las_data_spedizione DATETIME NULL ," + _
			"	las_id_admin_mittente int NULL, " + _
			"	las_id_utente_destinatario int NULL" + _
			"); " + _
			SQL_AddForeignKey(conn, "mtb_log_avvisi_spediti", "las_impegno_id", "mtb_impegni", "imp_id", false, "") + _
			SQL_AddForeignKey(conn, "mtb_log_avvisi_spediti", "las_id_admin_mittente", "tb_admin", "ID_admin", false, "") + _
			SQL_AddForeignKey(conn, "mtb_log_avvisi_spediti", "las_id_utente_destinatario", "tb_utenti", "ut_ID", false, "") + _
			
			" ALTER TABLE mtb_impegni ADD " + _
			" 	imp_invia_avviso bit NULL, " + _
			"	imp_anticipo_avviso int NULL, " + _
			" 	imp_insAdmin_id int NULL, " + _
			" 	imp_insData DATETIME NULL, " + _
			" 	imp_modAdmin_id int NULL, " + _
			" 	imp_modData DATETIME NULL " + _
			"; " + _
			SQL_AddForeignKey(conn, "mtb_impegni", "imp_insAdmin_id", "tb_admin", "ID_admin", false, "FK_mtb_impegni__tb_admin_2") + _
			SQL_AddForeignKey(conn, "mtb_impegni", "imp_modAdmin_id", "tb_admin", "ID_admin", false, "FK_mtb_impegni__tb_admin_3")
end function
'*******************************************************************************************

			
			
'*******************************************************************************************
'AGGIORNAMENTO SPECIALE NEXT-MEMO 2.0  3
'...........................................................................................
' Giacomo 23/02/2011
'...........................................................................................
' aggiunge parametro per gestire gli avvisi degli impegni
'...........................................................................................
function Aggiornamento__MEMO2__4(conn)
	Aggiornamento__MEMO2__4 = " SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__MEMO2__4(conn)
	dim sql, id_ragg
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & NEXTMEMO2)) <> "" then
		sql = "SELECT sdr_id FROM tb_siti_descrittori_raggruppamenti WHERE sdr_titolo_it LIKE 'Agenda Next-Memo 2'" 
		id_ragg = cIntero(GetValueList(conn, NULL, sql))
		CALL AddParametroSito(conn, "ID_PAGINA_AVVISO", _
									id_ragg, _
									"id della pagina sito da utilizzare nella e-mail di avviso di un impegno", _
									"", _
									adNumeric, _
									0, _
									"", _
									1, _
									1, _
									NEXTMEMO2, _
									null, null, null, null, null)
		CALL AddParametroSito(conn, "OGGETTO_PAGINA_AVVISO", _
									id_ragg, _
									"oggetto della e-mail di avviso di un impegno", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									NEXTMEMO2, _
									null, null, null, null, null)
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO SPECIALE NEXT-MEMO 2.0  5
'...........................................................................................
' Giacomo 24/03/2011
'...........................................................................................
' cambio nome dei un parametro perchè il codice coincideva con un paramentro del next-NEWS
'...........................................................................................
function Aggiornamento__MEMO2__5(conn)
	dim sql, id_ragg
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & NEXTMEMO2)) <> "" then
		sql = "(SELECT sdr_id FROM tb_siti_descrittori_raggruppamenti WHERE sdr_titolo_it LIKE 'Attivazione agenda')"
		id_ragg = cIntero(GetValueList(conn, NULL, sql))
		Aggiornamento__MEMO2__5 = " DELETE FROM rel_siti_descrittori WHERE rsd_sito_id="&NEXTMEMO2&" AND " & _
								  " 	rsd_descrittore_id IN (SELECT sid_id FROM tb_siti_descrittori WHERE sid_codice LIKE 'AGENDA_ATTIVA');" & _
								  " UPDATE tb_siti_descrittori SET sid_raggruppamento_id = "&id_ragg&", sid_nome_it = 'attiva le gestione dell''agenda' " & _
								  " WHERE sid_codice LIKE 'agenda_attiva'; "
	else
		Aggiornamento__MEMO2__5 = " SELECT * FROM AA_versione "
	end if
end function

function AggiornamentoSpeciale__MEMO2__5(conn)
	dim sql, id_ragg
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & NEXTMEMO2)) <> "" then
		sql = "SELECT sdr_id FROM tb_siti_descrittori_raggruppamenti WHERE sdr_titolo_it LIKE 'Agenda Next-Memo 2'" 
		id_ragg = cIntero(GetValueList(conn, NULL, sql))
		CALL AddParametroSito(conn, "AGENDA_MEMO2_ATTIVA", _
									id_ragg, _
									"attiva la gestione degli impegni/appuntamenti", _
									"", _
									adBoolean, _
									0, _
									"", _
									1, _
									1, _
									NEXTMEMO2, _
									"false", null, null, null, null)
	end if
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO NEXT-MEMO 2.0 6
'...........................................................................................
' Giacomo 30/05/2011
'...........................................................................................
' aggiungo campi per gestire i cataloghi sul memo 2
'...........................................................................................
function Aggiornamento__MEMO2__6(conn)
    Aggiornamento__MEMO2__6 = _
		" ALTER TABLE mtb_documenti ADD " & _
		SQL_MultiLanguageFieldComplete(conn, "doc_url_catalogo_<lingua> " + SQL_CharField(Conn, 500) + " NULL ") + ", " + _
		SQL_MultiLanguageFieldComplete(conn, "doc_data_modifica_catalogo_<lingua> smalldatetime NULL") + ", " + _
		" doc_catalogo_sfogliabile bit NULL;"
end function
'*******************************************************************************************
	
	

'*******************************************************************************************
'AGGIORNAMENTO SPECIALE NEXT-MEMO 2.0  7
'...........................................................................................
' Giacomo 30/05/2011
'...........................................................................................
' aggiunge parametro per attivare la gestione dei cataloghi 
'...........................................................................................
function Aggiornamento__MEMO2__7(conn)
	Aggiornamento__MEMO2__7 = " INSERT INTO tb_siti_descrittori_raggruppamenti(sdr_titolo_it, sdr_ordine,sdr_personalizzato) " & _
							  " SELECT 'Cataloghi Next-Memo 2',MAX(sdr_ordine)+1,1 FROM tb_siti_descrittori_raggruppamenti "
end function

function AggiornamentoSpeciale__MEMO2__7(conn)
	dim sql, id_ragg
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & NEXTMEMO2)) <> "" then
		sql = "SELECT sdr_id FROM tb_siti_descrittori_raggruppamenti WHERE sdr_titolo_it LIKE 'Cataloghi Next-Memo 2'" 
		id_ragg = cIntero(GetValueList(conn, NULL, sql))
		CALL AddParametroSito(conn, "MEMO2_CATALOGHI_ATTIVI", _
									id_ragg, _
									"attivazione gestione cataloghi speciali", _
									"", _
									adBoolean, _
									0, _
									"", _
									1, _
									1, _
									NEXTMEMO2, _
									null, null, null, null, null)

	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO SPECIALE NEXT-MEMO 2.0  8
'...........................................................................................
' Giacomo 30/05/2011
'...........................................................................................
' aggiunge parametri per gestire i cataloghi (parte pubblica)
'...........................................................................................
function Aggiornamento__MEMO2__8(conn)
	Aggiornamento__MEMO2__8 = "SELECT * FROM AA_versione "
end function

function AggiornamentoSpeciale__MEMO2__8(conn)
	dim sql, id_ragg
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & NEXTMEMO2)) <> "" then
		sql = "SELECT sdr_id FROM tb_siti_descrittori_raggruppamenti WHERE sdr_titolo_it LIKE 'Cataloghi Next-Memo 2'" 
		id_ragg = cIntero(GetValueList(conn, NULL, sql))
		CALL AddParametroSito(conn, "MEMO2_CATALOGO_PATH", _
									id_ragg, _
									"percorso base per la creazione dei cataloghi", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									NEXTMEMO2, _
									"catalogo", null, null, null, null)
		CALL AddParametroSito(conn, "MEMO2_CATALOGO_IMAGE_WIDTH", _
									id_ragg, _
									"larghezza immagini create per il catalogo", _
									"", _
									adNumeric, _
									0, _
									"", _
									1, _
									1, _
									NEXTMEMO2, _
									600, null, null, null, null)
		CALL AddParametroSito(conn, "MEMO2_CATALOGO_IMAGE_HEIGHT", _
									id_ragg, _
									"altezza immagini create per il catalogo", _
									"", _
									adNumeric, _
									0, _
									"", _
									1, _
									1, _
									NEXTMEMO2, _
									800, null, null, null, null)									
		CALL AddParametroSito(conn, "MEMO2_CATALOGO_BGCOLOR", _
									id_ragg, _
									"colore di sfondo", _
									"", _
									adPropVariant, _
									0, _
									"", _
									1, _
									1, _
									NEXTMEMO2, _
									"#CCCCCC", null, null, null, null)
		CALL AddParametroSito(conn, "MEMO2_CATALOGO_LOADERCOLOR", _
									id_ragg, _
									"colore", _
									"", _
									adPropVariant, _
									0, _
									"", _
									1, _
									1, _
									NEXTMEMO2, _
									"#FFFFFF", null, null, null, null)
		CALL AddParametroSito(conn, "MEMO2_CATALOGO_PANELCOLOR", _
									id_ragg, _
									"colore", _
									"", _
									adPropVariant, _
									0, _
									"", _
									1, _
									1, _
									NEXTMEMO2, _
									"#5D5D61", null, null, null, null)
		CALL AddParametroSito(conn, "MEMO2_CATALOGO_BUTTONCOLOR", _
									id_ragg, _
									"colore dei pulsanti", _
									"", _
									adPropVariant, _
									0, _
									"", _
									1, _
									1, _
									NEXTMEMO2, _
									"#5D5D61", null, null, null, null)
		CALL AddParametroSito(conn, "MEMO2_CATALOGO_TEXTCOLOR", _
									id_ragg, _
									"colore del testo", _
									"", _
									adPropVariant, _
									0, _
									"", _
									1, _
									1, _
									NEXTMEMO2, _
									"#FFFFFF", null, null, null, null)										
	end if
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO NEXT-MEMO 2.0  9
'...........................................................................................
' Giacomo 27/06/2011
'...........................................................................................
' Creazione tabelle per aggiungere i descrittori ai documenti
'...........................................................................................
function Aggiornamento__MEMO2__9(conn)
    Aggiornamento__MEMO2__9 = _
			"CREATE TABLE " + SQL_Dbo(conn) + " mtb_carattech (" + _
			" ct_id " & SQL_PrimaryKey(conn, "mtb_carattech") + ", " + _
			SQL_MultiLanguageFieldComplete(conn, "ct_nome_<lingua> " + SQL_CharField(Conn, 510) + " NULL ") + ", " + _
			SQL_MultiLanguageFieldComplete(conn, "ct_unita_<lingua> " + SQL_CharField(Conn, 100) + " NULL ") + ", " + _
			" ct_tipo int NULL, " + _
			" ct_codice " + SQL_CharField(Conn, 255) + " NULL, " + _
			" ct_per_ricerca bit NULL, " + _
			" ct_per_confronto bit NULL, " + _
			" ct_img " + SQL_CharField(Conn, 255) + " NULL, " + _
			" ct_raggruppamento_id int NULL " + _
			"); " + _
			"CREATE TABLE " + SQL_Dbo(conn) + " mtb_carattech_raggruppamenti (" + _
			" ctr_id " & SQL_PrimaryKey(conn, "mtb_carattech_raggruppamenti") + ", " + _
			SQL_MultiLanguageFieldComplete(conn, "ctr_titolo_<lingua> " + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
			" ctr_ordine int NULL, " + _
			" ctr_codice " + SQL_CharField(Conn, 255) + " NULL, " + _
			" ctr_di_sistema bit NULL " + _
			"); " + _
			SQL_AddForeignKey(conn, "mtb_carattech", "ct_raggruppamento_id", "mtb_carattech_raggruppamenti", "ctr_id", false, "") + _
			"CREATE TABLE " + SQL_Dbo(conn) + " mrel_categ_ctech (" + _
			" rcc_id " & SQL_PrimaryKey(conn, "mrel_categ_ctech") + ", " + _
			" rcc_ctech_id int NULL, " + _
			" rcc_ordine int NULL, " + _
			" rcc_categoria_id int NULL " + _
			"); " + _
			SQL_AddForeignKey(conn, "mrel_categ_ctech", "rcc_ctech_id", "mtb_carattech", "ct_id", false, "") + _
			SQL_AddForeignKey(conn, "mrel_categ_ctech", "rcc_categoria_id", "mtb_documenti_categorie", "catC_id", false, "") + _
			"CREATE TABLE " + SQL_Dbo(conn) + " mrel_doc_ctech (" + _
			" rdc_id " & SQL_PrimaryKey(conn, "mrel_doc_ctech") + ", " + _
			" rdc_doc_id int NULL, " + _
			" rdc_ctech_id int NULL, " + _
			SQL_MultiLanguageFieldComplete(conn, "rdc_valore_<lingua> " + SQL_CharField(Conn, 0) + " NULL ") + _
			"); " + _
			SQL_AddForeignKey(conn, "mrel_doc_ctech", "rdc_doc_id", "mtb_documenti", "doc_id", false, "") + _
			SQL_AddForeignKey(conn, "mrel_doc_ctech", "rdc_ctech_id", "mtb_carattech", "ct_id", false, "")
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO NEXT-MEMO 2.0 10
'...........................................................................................
' Giacomo 29/06/2011
'...........................................................................................
' cambio il tipo dei campi file
'...........................................................................................
function Aggiornamento__MEMO2__10(conn)
	Dim sql, rs	
	set rs = Server.CreateObject("ADODB.Recordset")
	rs.open "SELECT * FROM tb_cnt_lingue", conn, adOpenKeySet, adLockOptimistic, adCmdText
	while not rs.eof 
		sql = sql & " ALTER TABLE mtb_documenti ADD temp_"&rs("lingua_codice")&" "&SQL_CharField(Conn, 0)&" NULL; " & vbCrLF
		rs.moveNext
	wend
	
	rs.moveFirst
	sql = sql & " UPDATE mtb_documenti SET "
	while not rs.eof 
		sql = sql & " temp_"&rs("lingua_codice")&"=doc_file_"&rs("lingua_codice")&","
		rs.moveNext
	wend
	sql = sql & "---"
	sql = Replace(sql, ",---", ";")
	sql = sql & vbCrLF
	
	rs.moveFirst
	while not rs.eof 
		sql = sql & " ALTER TABLE mtb_documenti DROP COLUMN doc_file_"&rs("lingua_codice")&"; " & vbCrLF
		sql = sql & " ALTER TABLE mtb_documenti ADD doc_file_"&rs("lingua_codice")&" "&SQL_CharField(Conn, 500)&" NULL; " & vbCrLF
		rs.moveNext
	wend
	
	rs.moveFirst
	sql = sql & " UPDATE mtb_documenti SET "
	while not rs.eof 
		sql = sql & " doc_file_"&rs("lingua_codice")&"=temp_"&rs("lingua_codice")&","
		rs.moveNext
	wend
	sql = sql & "---"
	sql = Replace(sql, ",---", ";")
	sql = sql & vbCrLF
	
	rs.moveFirst
	while not rs.eof 
		sql = sql & " ALTER TABLE mtb_documenti DROP COLUMN temp_"&rs("lingua_codice")&"; " & vbCrLF
		rs.moveNext
	wend
	
	rs.close
	
    Aggiornamento__MEMO2__10 = sql
	
	set rs = nothing
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-MEMO 2.0 11
'...........................................................................................
' Giacomo 13/04/2012
'...........................................................................................
' aggiungo campo foto sulle categorie
'...........................................................................................
function Aggiornamento__MEMO2__11(conn)
    Aggiornamento__MEMO2__11 = _
		" ALTER TABLE mtb_documenti_categorie ADD " & _
		" catC_foto " + SQL_CharField(Conn, 250) + " NULL;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-MEMO 2.0 12
'...........................................................................................
' Giacomo 21/05/2013
'...........................................................................................
' aggiungo campi inserimento e modifica
'...........................................................................................
function Aggiornamento__MEMO2__12(conn)
    Aggiornamento__MEMO2__12 = _
			" ALTER TABLE mtb_documenti ADD " + _
			" 	doc_insAdmin_id int NULL, " + _
			" 	doc_insData DATETIME NULL, " + _
			" 	doc_modAdmin_id int NULL, " + _
			" 	doc_modData DATETIME NULL " + _
			"; " + _
			SQL_AddForeignKey(conn, "mtb_documenti", "doc_insAdmin_id", "tb_admin", "ID_admin", false, "FK_mtb_documenti__tb_admin_2") + _
			SQL_AddForeignKey(conn, "mtb_documenti", "doc_modAdmin_id", "tb_admin", "ID_admin", false, "FK_mtb_documenti__tb_admin_3")
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO NEXT-MEMO 2.0 13
'...........................................................................................
' Giacomo 30/01/2014
'...........................................................................................
' aggiungo campo foto sulle categorie
'...........................................................................................
function Aggiornamento__MEMO2__13(conn)
    Aggiornamento__MEMO2__13 = _
		" ALTER TABLE mtb_carattech ADD " & _
		" ct_principale bit NULL;"
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO NEXT-MEMO 2.0 14
'...........................................................................................
' Giacomo 12/09/2014
'...........................................................................................
' estendo lunghezza campo titolo documento
'...........................................................................................
function Aggiornamento__MEMO2__14(conn)
	Aggiornamento__MEMO2__14 = ""
	
	set rs = Server.CreateObject("ADODB.Recordset")
	rs.open "SELECT * FROM tb_cnt_lingue", conn, adOpenKeySet, adLockOptimistic, adCmdText
	while not rs.eof
		Aggiornamento__MEMO2__14 = Aggiornamento__MEMO2__14 & _
			" ALTER TABLE mtb_documenti " & _
			" ALTER COLUMN doc_titolo_" & rs("lingua_codice") & " " & SQL_CharField(Conn, 500) & " NULL; "
		rs.moveNext
	wend
	rs.close
    
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO SPECIALE NEXT-MEMO 2.0  13
'...........................................................................................
' Giacomo 21/11/2013
'...........................................................................................
' aggiunge parametro per avere solo gli utenti dell'area riservata del memo2 nella gestione utenti in area amministrativa
'...........................................................................................
'function Aggiornamento__MEMO2__13(conn)
'	Aggiornamento__MEMO2__13 = "SELECT * FROM AA_versione "
'end function
'
'function AggiornamentoSpeciale__MEMO2__13(conn)
'	dim sql, id_ragg
'	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & NEXTMEMO2)) <> "" then
'		CALL AddParametroSito(conn, "MEMO2_PERMESSI_AREA_RISERVATA", _
'									0, _
'									"lista permessi area riservata memo2 (se più di uno separati da ,)", _
'									"", _
'									adVarChar, _
'									0, _
'									"", _
'									1, _
'									1, _
'									NEXTMEMO2, _
'									null, null, null, null, null)
'	end if
'end function
'*******************************************************************************************

				
%>
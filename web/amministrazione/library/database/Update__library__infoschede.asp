<%
'...........................................................................................
'........................................................................................... 
'libreria di funzioni che contiene tutti gli aggiornamenti per l'applicativo INFOSCHEDE
'...........................................................................................
'...........................................................................................



'*******************************************************************************************
'AGGIORNAMENTO INFOSCHEDE 1
'...........................................................................................
' Giacomo 01/03/2011
'...........................................................................................
' creazione nuove tabelle per l'applicativo infoschede
'...........................................................................................
function Aggiornamento__INFOSCHEDE__1(conn)
    Aggiornamento__INFOSCHEDE__1 = _
			"CREATE TABLE " + SQL_Dbo(conn) + "sgtb_problemi(" + _
			"	prb_id " + SQL_PrimaryKey(conn, "sgtb_problemi") + ", " + _
			SQL_MultiLanguageFieldComplete(conn, "prb_nome_<lingua> " + SQL_CharField(Conn, 255) + " NULL ") + "," + _
			SQL_MultiLanguageFieldComplete(conn, "prb_descrizione_<lingua> " + SQL_CharField(Conn, 0) + " NULL ") + "," + _
			" 	prb_modalita_easy bit NULL, " + _
			SQL_MultiLanguageFieldComplete(conn, "prb_avviso_per_conferma_<lingua> " + SQL_CharField(Conn, 0) + " NULL ") + "," + _
			" 	prb_ordine int NULL, " + _
			" 	prb_visibile bit NULL, " + _
			" 	prb_riscontrato bit NULL, " + _
			"	prb_img_th " + SQL_CharField(Conn, 500) + " NULL, " + _
			"	prb_img_zo " + SQL_CharField(Conn, 500) + " NULL, " + _
			" 	prb_insAdmin_id int NULL, " + _
			" 	prb_insData DATETIME NULL, " + _
			" 	prb_modAdmin_id int NULL, " + _
			" 	prb_modData DATETIME NULL " + _
			"); " + _
			SQL_AddForeignKey(conn, "sgtb_problemi", "prb_insAdmin_id", "tb_admin", "ID_admin", false, "") + _
			SQL_AddForeignKey(conn, "sgtb_problemi", "prb_modAdmin_id", "tb_admin", "ID_admin", false, "FK_sgtb_problemi__tb_admin_2") + _
			" " + _
			"CREATE TABLE " + SQL_Dbo(conn) + "srel_problemi_mar_tip(" + _
			"	rpm_id " + SQL_PrimaryKey(conn, "srel_problemi_mar_tip") + ", " + _
			" 	rpm_problema_id int NULL, " + _
			" 	rpm_marchio_id int NULL, " + _
			" 	rpm_tipologia_id int NULL " + _
			"); " + _
			SQL_AddForeignKey(conn, "srel_problemi_mar_tip", "rpm_problema_id", "sgtb_problemi", "prb_id", true, "") + _
			SQL_AddForeignKey(conn, "srel_problemi_mar_tip", "rpm_marchio_id", "gtb_marche", "mar_id", false, "") + _
			SQL_AddForeignKey(conn, "srel_problemi_mar_tip", "rpm_tipologia_id", "gtb_tipologie", "tip_id", false, "") + _
			" " + _
			"CREATE TABLE " + SQL_Dbo(conn) + "srel_problemi_profili(" + _
			"	rpp_id " + SQL_PrimaryKey(conn, "srel_problemi_profili") + ", " + _
			" 	rpp_problema_id int NULL, " + _
			" 	rpp_profilo_id int NULL, " + _
			"); " + _
			SQL_AddForeignKey(conn, "srel_problemi_profili", "rpp_problema_id", "sgtb_problemi", "prb_id", true, "") + _
			SQL_AddForeignKey(conn, "srel_problemi_profili", "rpp_profilo_id", "gtb_profili", "pro_id", true, "") + _
			" " + _
			"CREATE TABLE " + SQL_Dbo(conn) + "srel_problemi_articoli(" + _
			"	rpa_id " + SQL_PrimaryKey(conn, "srel_problemi_articoli") + ", " + _
			" 	rpa_problema_id int NULL, " + _
			" 	rpa_articolo_id int NULL, " + _
			"); " + _
			SQL_AddForeignKey(conn, "srel_problemi_articoli", "rpa_problema_id", "sgtb_problemi", "prb_id", true, "") + _
			SQL_AddForeignKey(conn, "srel_problemi_articoli", "rpa_articolo_id", "gtb_articoli", "art_id", true, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO INFOSCHEDE 2
'...........................................................................................
' Giacomo 16/03/2011
'...........................................................................................
' creazione nuove tabelle per l'applicativo infoschede (schede)
'...........................................................................................
function Aggiornamento__INFOSCHEDE__2(conn)
    Aggiornamento__INFOSCHEDE__2 = _
			"CREATE TABLE " + SQL_Dbo(conn) + "sgtb_stati_schede(" + _
			"	sts_id " + SQL_PrimaryKey(conn, "sgtb_stati_schede") + ", " + _
			SQL_MultiLanguageFieldComplete(conn, "sts_nome_<lingua> " + SQL_CharField(Conn, 255) + " NULL ") + "," + _
			" 	sts_ordine int NULL, " + _
			" 	sts_visibile_admin bit NULL, " + _
			" 	sts_modifica_admin bit NULL, " + _
			" 	sts_visibile_officina bit NULL, " + _
			" 	sts_modifica_officina bit NULL, " + _
			" 	sts_visibile_centr_assist bit NULL, " + _
			" 	sts_modifica_centr_assist bit NULL " + _			
			"); " + _
			"CREATE TABLE " + SQL_Dbo(conn) + "sgtb_accessori(" + _
			"	acc_id " + SQL_PrimaryKey(conn, "sgtb_accessori") + ", " + _
			SQL_MultiLanguageFieldComplete(conn, "acc_nome_<lingua> " + SQL_CharField(Conn, 500) + " NULL ") + _
			"); " + _
			"CREATE TABLE " + SQL_Dbo(conn) + "sgtb_esiti(" + _
			"	esi_id " + SQL_PrimaryKey(conn, "sgtb_esiti") + ", " + _
			SQL_MultiLanguageFieldComplete(conn, "esi_nome_<lingua> " + SQL_CharField(Conn, 255) + " NULL ") + _
			"); " + _
			"CREATE TABLE " + SQL_Dbo(conn) + "sgtb_schede(" + _
			"	sc_id " + SQL_PrimaryKey(conn, "sgtb_schede") + ", " + _
			" 	sc_stato_id int NULL, " + _
			" 	sc_numero int NULL, " + _
			" 	sc_data_ricevimento SMALLDATETIME NULL, " + _
			"	sc_cliente_id int NULL, " + _
			"	sc_centro_assistenza_id int NULL, " + _
			"	sc_modello_id int NULL, " + _
			"	sc_matricola " + SQL_CharField(Conn, 255) + " NULL, " + _
			" 	sc_data_acquisto SMALLDATETIME NULL, " + _
			"	sc_numero_scontrino " + SQL_CharField(Conn, 100) + " NULL, " + _
			" 	sc_in_garanzia bit NULL, " + _
			" 	sc_guasto_segnalato_id int NULL, " + _
			" 	sc_guasto_segnalato_altro " + SQL_CharField(Conn, 500) + " NULL, " + _
			" 	sc_guasto_riscontrato_id int NULL, " + _
			" 	sc_guasto_riscontrato_altro " + SQL_CharField(Conn, 500) + " NULL, " + _
			" 	sc_esito_intervento_id int NULL, " + _
			" 	sc_accessori_presenti_id int NULL, " + _
			"	sc_accessori_presenti_altro " + SQL_CharField(Conn, 500) + " NULL, " + _
			"	sc_note_cliente " + SQL_CharField(Conn, 0) + " NULL, " + _
			"	sc_note_chiusura " + SQL_CharField(Conn, 0) + " NULL, " + _
			" 	sc_data_fine_lavoro SMALLDATETIME NULL, " + _
			"	sc_trasportatore_id int NULL, " + _
			"	sc_costo_presa money NULL, " + _
			"	sc_costo_riconsegna money NULL, " + _
			"	sc_rif_DDT_di_carico " + SQL_CharField(Conn, 0) + " NULL, " + _
			"	sc_rif_DDT_di_resa_id int NULL, " + _
			"	sc_rif_cliente " + SQL_CharField(Conn, 0) + " NULL, " + _
			"	sc_ora_manodopera_intervento int NULL, " + _
			"	sc_prezzo_manodopera money NULL, " + _
			" 	sc_insAdmin_id int NULL, " + _
			" 	sc_insData DATETIME NULL, " + _
			" 	sc_modAdmin_id int NULL, " + _
			" 	sc_modData DATETIME NULL " + _
			"); " + _
			SQL_AddForeignKey(conn, "sgtb_schede", "sc_stato_id", "sgtb_stati_schede", "sts_id", true, "") + _
			SQL_AddForeignKey(conn, "sgtb_schede", "sc_cliente_id", "tb_Indirizzario", "IDElencoIndirizzi", true, "") + _
			SQL_AddForeignKey(conn, "sgtb_schede", "sc_centro_assistenza_id", "gtb_agenti", "ag_id", false, "") + _
			SQL_AddForeignKey(conn, "sgtb_schede", "sc_modello_id", "grel_art_valori", "rel_id", true, "") + _
			SQL_AddForeignKey(conn, "sgtb_schede", "sc_guasto_segnalato_id", "sgtb_problemi", "prb_id", false, "") + _
			SQL_AddForeignKey(conn, "sgtb_schede", "sc_guasto_riscontrato_id", "sgtb_problemi", "prb_id", false, "FK_sgtb_schede__sgtb_problemi_2") + _
			SQL_AddForeignKey(conn, "sgtb_schede", "sc_esito_intervento_id", "sgtb_esiti", "esi_id", false, "") + _
			SQL_AddForeignKey(conn, "sgtb_schede", "sc_accessori_presenti_id", "sgtb_accessori", "acc_id", false, "") + _
			SQL_AddForeignKey(conn, "sgtb_schede", "sc_trasportatore_id", "gtb_rivenditori", "riv_id", false, "") + _
			SQL_AddForeignKey(conn, "sgtb_schede", "sc_rif_DDT_di_resa_id", "gtb_ordini", "ord_id", false, "") + _
			SQL_AddForeignKey(conn, "sgtb_schede", "sc_insAdmin_id", "tb_admin", "ID_admin", false, "") + _
			SQL_AddForeignKey(conn, "sgtb_schede", "sc_modAdmin_id", "tb_admin", "ID_admin", false, "FK_sgtb_schede__tb_admin_2") + _
			" " + _
			"CREATE TABLE " + SQL_Dbo(conn) + "sgtb_dettagli_schede(" + _
			"	dts_id " + SQL_PrimaryKey(conn, "sgtb_dettagli_schede") + ", " + _
			"	dts_ricambio_id int NULL, " + _
			" 	dts_ricambio_codice " + SQL_CharField(Conn, 100) + " NULL, " + _
			" 	dts_ricambio_nome " + SQL_CharField(Conn, 255) + " NULL, " + _
			" 	dts_ricambio_prezzo money NULL, " + _
			" 	dts_ricambio_qta int NULL, " + _
			" 	dts_ricambio_sconto real NULL " + _
			"); " + _
			SQL_AddForeignKey(conn, "sgtb_dettagli_schede", "dts_ricambio_id", "grel_art_valori", "rel_id", true, "")			
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO INFOSCHEDE 3
'...........................................................................................
' Giacomo 18/03/2011
'...........................................................................................
function Aggiornamento__INFOSCHEDE__3(conn)
    Aggiornamento__INFOSCHEDE__3 = _
			"ALTER TABLE sgtb_dettagli_schede ADD dts_scheda_id int NULL; " + _
			SQL_AddForeignKeyExtended(conn, "sgtb_dettagli_schede", "dts_scheda_id", "sgtb_schede", "sc_id", true, false, "")	
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO INFOSCHEDE 4
'...........................................................................................
' Giacomo 18/03/2011
'...........................................................................................
'	Correzione relazione problemi - articoli
'...........................................................................................
function Aggiornamento__INFOSCHEDE__4(conn)
    Aggiornamento__INFOSCHEDE__4 = _
			SQL_RemoveForeignKey(conn, "srel_problemi_articoli", "rpa_articolo_id", "gtb_articoli", true, "FK_srel_problemi_articoli__gtb_articoli") + _
			"ALTER TABLE srel_problemi_articoli DROP COLUMN rpa_articolo_id; " + _
			"ALTER TABLE srel_problemi_articoli ADD rpa_articolo_rel_id int NULL; " + _
			SQL_AddForeignKey(conn, "srel_problemi_articoli", "rpa_articolo_rel_id", "grel_art_valori", "rel_id", true, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO INFOSCHEDE 5
'...........................................................................................
' Giacomo 25/03/2011
'...........................................................................................
' creazione nuove tabelle per aggiungere i descrittori alle schede
'...........................................................................................
function Aggiornamento__INFOSCHEDE__5(conn)
    Aggiornamento__INFOSCHEDE__5 = _
			"CREATE TABLE " & SQL_Dbo(Conn) & "sgtb_descrittori_raggruppamenti(" + _
			"	rag_id " & SQL_PrimaryKey(conn, "sgtb_descrittori_raggruppamenti") + ", " +_ 
			SQL_MultiLanguageFieldComplete(conn, "rag_titolo_<lingua>" + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
			"	rag_ordine INT NULL, " + _
			"	rag_codice " + SQL_CharField(Conn, 255) + ", " + _
			"	rag_note " + SQL_CharField(Conn, 0) + _
			" ); " + _
			"CREATE TABLE " + SQL_Dbo(conn) + "sgtb_descrittori(" + _
			"	des_id " + SQL_PrimaryKey(conn, "sgtb_descrittori") + ", " + _
			SQL_MultiLanguageFieldComplete(conn, "des_nome_<lingua> " + SQL_CharField(Conn, 255) + " NULL ")  + ", " + _
			"	des_raggruppamento_id int NULL, " + _
			"	des_tipo int NULL, " + _
			"	des_principale BIT NULL, " + _
			SQL_MultiLanguageFieldComplete(conn, "des_unita_<lingua>" + SQL_CharField(Conn, 255) + " NULL ")  + ", " + _
			"	des_img " + SQL_CharField(Conn, 255) + _
			");" + _
			"CREATE TABLE " + SQL_Dbo(conn) + "srel_descrittori_schede (" + _
			"	rds_id " + SQL_PrimaryKey(conn, "srel_descrittori_schede") + ", " + _
			"	rds_descrittore_id int NULL, " + _
			"	rds_scheda_id int NULL, " + _
			"	rds_ordine int NULL, " + _
			SQL_MultiLanguageFieldComplete(conn, "rds_valore_<lingua>" + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
			SQL_MultiLanguageFieldComplete(conn, "rds_memo_<lingua>" + SQL_CharField(Conn, 0) + " NULL ") + ", " + _
			"	rds_data SMALLDATETIME NULL " + _
			");" + _
			SQL_AddForeignKey(conn, "sgtb_descrittori", "des_raggruppamento_id", "sgtb_descrittori_raggruppamenti", "rag_id", false, "") + _
			SQL_AddForeignKey(conn, "srel_descrittori_schede", "rds_descrittore_id", "sgtb_descrittori", "des_id", true, "") + _
			SQL_AddForeignKey(conn, "srel_descrittori_schede", "rds_scheda_id", "sgtb_schede", "sc_id", true, "")
end function
'*******************************************************************************************

			
'*******************************************************************************************
'AGGIORNAMENTO INFOSCHEDE 6
'...........................................................................................
' Giacomo 25/03/2011
'...........................................................................................
' alcune correzioni
'...........................................................................................
function Aggiornamento__INFOSCHEDE__6(conn)
    Aggiornamento__INFOSCHEDE__6 = _
		" ALTER TABLE gtb_rivenditori ADD riv_sconto real NULL;" + _
		" ALTER TABLE sgtb_schede DROP COLUMN sc_rif_DDT_di_carico;" + _
		" ALTER TABLE sgtb_schede ADD sc_numero_DDT_di_carico int NULL;" + _
		" ALTER TABLE sgtb_schede ADD sc_data_DDT_di_carico smalldatetime NULL;" + _
		" ALTER TABLE sgtb_schede DROP COLUMN sc_rif_cliente;" + _
		" ALTER TABLE sgtb_schede ADD sc_rif_cliente " + SQL_CharField(Conn, 255) + " NULL;"
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO INFOSCHEDE 7
'...........................................................................................
' Giacomo 28/03/2011
'...........................................................................................
' creo la tabella dei DDT, da usare al posto di gtb_ordini
'...........................................................................................
function Aggiornamento__INFOSCHEDE__7(conn)
    Aggiornamento__INFOSCHEDE__7 = _
		"CREATE TABLE " & SQL_Dbo(Conn) & "sgtb_ddt_categorie(" + _
		"	cat_id " + SQL_PrimaryKey(conn, "sgtb_ddt_categorie") + ", " + _
		SQL_MultiLanguageFieldComplete(conn, "cat_nome_<lingua>" + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
		"	cat_pagina_id int NULL " + _ 
		");" + _
		SQL_AddForeignKey(conn, "sgtb_ddt_categorie", "cat_pagina_id", "tb_paginesito", "id_pagineSito", false, "") + _
		"CREATE TABLE " & SQL_Dbo(Conn) & "sgtb_ddt(" + _
		"	ddt_id " + SQL_PrimaryKey(conn, "sgtb_ddt") + ", " + _
		"	ddt_categoria_id int NULL, " + _
		"	ddt_numero int NULL, " + _
		"	ddt_causale_id int NULL, " + _
		"	ddt_trasportatore_id int NULL, " + _
		"	ddt_cliente_id int NULL, " + _
		"	ddt_destinazione_id int NULL, " + _
		"	ddt_data smalldatetime NULL, " + _
		"	ddt_note " + SQL_CharField(Conn, 0) + _
		");" + _
		SQL_AddForeignKey(conn, "sgtb_ddt", "ddt_categoria_id", "sgtb_ddt_categorie", "cat_id", true, "") + _
		SQL_AddForeignKey(conn, "sgtb_ddt", "ddt_causale_id", "gtb_stati_ordine", "so_id", false, "") + _
		SQL_AddForeignKey(conn, "sgtb_ddt", "ddt_trasportatore_id", "gtb_rivenditori", "riv_id", false, "") + _
		SQL_AddForeignKey(conn, "sgtb_ddt", "ddt_cliente_id", "gtb_rivenditori", "riv_id", false, "FK_sgtb_ddt__gtb_rivenditori_2") + _
		SQL_AddForeignKey(conn, "sgtb_ddt", "ddt_destinazione_id", "tb_indirizzario", "IDElencoIndirizzi", false, "") + _
		SQL_RemoveForeignKey(conn, "sgtb_schede", "sc_trasportatore_id", "gtb_rivenditori", false, "") + _
		" ALTER TABLE sgtb_schede DROP COLUMN sc_trasportatore_id;" + _
		SQL_RemoveForeignKey(conn, "sgtb_schede", "sc_rif_DDT_di_resa_id", "gtb_ordini", false, "") + _
		SQL_AddForeignKey(conn, "sgtb_schede", "sc_rif_DDT_di_resa_id", "sgtb_ddt", "ddt_id", false, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO INFOSCHEDE 8
'...........................................................................................
' Giacomo 29/03/2011
'...........................................................................................
' correggo il riferimento della chiave esterna sc_cliente_id
'...........................................................................................
function Aggiornamento__INFOSCHEDE__8(conn)
    Aggiornamento__INFOSCHEDE__8 = _
		SQL_RemoveForeignKey(conn, "sgtb_schede", "sc_cliente_id", "tb_Indirizzario", true, "") + _
		SQL_AddForeignKey(conn, "sgtb_schede", "sc_cliente_id", "gtb_rivenditori", "riv_id", true, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO INFOSCHEDE 9
'...........................................................................................
' Giacomo 31/03/2011
'...........................................................................................
' aggiungo colonna prezzo totale a sgtb_dettagli_schede
'...........................................................................................
function Aggiornamento__INFOSCHEDE__9(conn)
    Aggiornamento__INFOSCHEDE__9 = _
		" ALTER TABLE sgtb_dettagli_schede ADD dts_prezzo_totale money NULL;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO INFOSCHEDE 10
'...........................................................................................
' Giacomo 14/04/2011
'...........................................................................................
' alcune correzioni
'...........................................................................................
function Aggiornamento__INFOSCHEDE__10(conn)
    Aggiornamento__INFOSCHEDE__10 = _
		" ALTER TABLE sgtb_schede ALTER COLUMN sc_numero_DDT_di_carico " + SQL_CharField(Conn, 255) + " NULL;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO INFOSCHEDE 11
'...........................................................................................
' Giacomo 14/04/2011
'...........................................................................................
' aggiungo campo
'...........................................................................................
function Aggiornamento__INFOSCHEDE__11(conn)
    Aggiornamento__INFOSCHEDE__11 = _
		" ALTER TABLE sgtb_schede ADD sc_negozio_acquisto " + SQL_CharField(Conn, 255) + " NULL;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO INFOSCHEDE 12
'...........................................................................................
' Giacomo 14/04/2011
'...........................................................................................
' aggiungo campi
'...........................................................................................
function Aggiornamento__INFOSCHEDE__12(conn)
    Aggiornamento__INFOSCHEDE__12 = _
		" ALTER TABLE sgtb_schede ADD sc_documento_ritiro_id int NULL;" & _
		SQL_AddForeignKey(conn, "sgtb_schede", "sc_documento_ritiro_id", "sgtb_ddt", "ddt_id", false, "_2") & _
		" ALTER TABLE sgtb_ddt ADD ddt_peso " + SQL_CharField(Conn, 255) + " NULL;" & _
		" ALTER TABLE sgtb_ddt ADD ddt_volume " + SQL_CharField(Conn, 255) + " NULL;" & _
		" ALTER TABLE sgtb_ddt ADD ddt_numero_colli " + SQL_CharField(Conn, 255) + " NULL;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO SPECIALE INFOSCHEDE 13
'...........................................................................................
' Giacomo 14/04/2011
'...........................................................................................
' aggiunge parametri per l'invio di preventivo e consuntivo
'...........................................................................................
function Aggiornamento__INFOSCHEDE__13(conn)
	Aggiornamento__INFOSCHEDE__13 = "SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__INFOSCHEDE__13(conn)
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & INFOSCHEDE)) <> "" then
		CALL AddParametroSito(conn, "INFOSCHEDE_ID_PAG_PREVENTIVO", _
									0, _
									"pagina di generazione del preventivo", _
									"", _
									adNumeric, _
									0, _
									"", _
									1, _
									1, _
									INFOSCHEDE, _
									null, null, null, null, null)
		CALL AddParametroSito(conn, "INFOSCHEDE_TESTO_PREVENTIVO", _
									0, _
									"testo e-mail per l'invio del preventivo", _
									"", _
									adLongVarChar, _
									0, _
									"", _
									1, _
									1, _
									INFOSCHEDE, _
									null, null, null, null, null)
		CALL AddParametroSito(conn, "INFOSCHEDE_ID_PAG_CONSUNTIVO", _
									0, _
									"pagina di generazione del consuntivo", _
									"", _
									adNumeric, _
									0, _
									"", _
									1, _
									1, _
									INFOSCHEDE, _
									null, null, null, null, null)
		CALL AddParametroSito(conn, "INFOSCHEDE_TESTO_CONSUNTIVO", _
									0, _
									"testo e-mail per l'invio del consuntivo", _
									"", _
									adLongVarChar, _
									0, _
									"", _
									1, _
									1, _
									INFOSCHEDE, _
									null, null, null, null, null)
		AggiornamentoSpeciale__INFOSCHEDE__13 = " SELECT * FROM AA_Versione "
	end if
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO INFOSCHEDE 14
'...........................................................................................
' Giacomo 27/04/2011
'...........................................................................................
' correggo il riferimento della chiave esterna dei ricambi
'...........................................................................................
function Aggiornamento__INFOSCHEDE__14(conn)
    Aggiornamento__INFOSCHEDE__14 = _
		SQL_RemoveForeignKey(conn, "sgtb_dettagli_schede", "dts_ricambio_id", "grel_art_valori", true, "") + _
		SQL_AddForeignKey(conn, "sgtb_dettagli_schede", "dts_ricambio_id", "grel_art_valori", "rel_id", false, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO SPECIALE INFOSCHEDE 15
'...........................................................................................
' Giacomo 27/04/2011
'...........................................................................................
' aggiunge parametri per le sezioni di invio e di spedizione, e visualizzazione della scheda
'...........................................................................................
function Aggiornamento__INFOSCHEDE__15(conn)
	Aggiornamento__INFOSCHEDE__15 = "SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__INFOSCHEDE__15(conn)
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & INFOSCHEDE)) <> "" then
		CALL AddParametroSito(conn, "INFOSCHEDE_ID_PAG_INVIO_RICH_RIT", _
									0, _
									"pagina per l'invio della richiesta di ritiro al trasportatore", _
									"", _
									adNumeric, _
									0, _
									"", _
									1, _
									1, _
									INFOSCHEDE, _
									null, null, null, null, null)
		CALL AddParametroSito(conn, "INFOSCHEDE_TESTO_INVIO_RICH_RIT", _
									0, _
									"testo e-mail per l'invio della richiesta di ritiro al trasportatore", _
									"", _
									adLongVarChar, _
									0, _
									"", _
									1, _
									1, _
									INFOSCHEDE, _
									null, null, null, null, null)
		CALL AddParametroSito(conn, "INFOSCHEDE_ID_PAG_INVIO_LETT_VETT", _
									0, _
									"pagina per l'invio della lettera di vettura al trasportatore", _
									"", _
									adNumeric, _
									0, _
									"", _
									1, _
									1, _
									INFOSCHEDE, _
									null, null, null, null, null)
		CALL AddParametroSito(conn, "INFOSCHEDE_TESTO_INVIO_LETT_VETT", _
									0, _
									"testo e-mail per l'invio della lettera di vettura al trasportatore", _
									"", _
									adLongVarChar, _
									0, _
									"", _
									1, _
									1, _
									INFOSCHEDE, _
									null, null, null, null, null)
		CALL AddParametroSito(conn, "INFOSCHEDE_ID_PAG_RICH_RITIRO", _
									0, _
									"pagina per visualizzare la richiesta di ritiro", _
									"", _
									adNumeric, _
									0, _
									"", _
									1, _
									1, _
									INFOSCHEDE, _
									null, null, null, null, null)
		CALL AddParametroSito(conn, "INFOSCHEDE_ID_PAG_DDT", _
									0, _
									"pagina per visualizzare il ddt (bolla)", _
									"", _
									adNumeric, _
									0, _
									"", _
									1, _
									1, _
									INFOSCHEDE, _
									null, null, null, null, null)
		CALL AddParametroSito(conn, "INFOSCHEDE_ID_PAG_LETT_VETT", _
									0, _
									"pagina per visualizzare la lettera di vettura", _
									"", _
									adNumeric, _
									0, _
									"", _
									1, _
									1, _
									INFOSCHEDE, _
									null, null, null, null, null)
		CALL AddParametroSito(conn, "INFOSCHEDE_ID_PAG_STAMPA_SCHEDA", _
									0, _
									"pagina per visualizzare la scheda", _
									"", _
									adNumeric, _
									0, _
									"", _
									1, _
									1, _
									INFOSCHEDE, _
									null, null, null, null, null)						
		AggiornamentoSpeciale__INFOSCHEDE__15 = " SELECT * FROM AA_Versione "
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO INFOSCHEDE 16
'...........................................................................................
' Giacomo 27/04/2011
'...........................................................................................
' aggiungo campo per salvare id  da import schede
'...........................................................................................
function Aggiornamento__INFOSCHEDE__16(conn)
    Aggiornamento__INFOSCHEDE__16 = _
		" ALTER TABLE sgtb_schede ADD sc_external_id " + SQL_CharField(Conn, 255) + " NULL;"
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO SPECIALE INFOSCHEDE 17
'...........................................................................................
' Giacomo 02/05/2011
'...........................................................................................
' aggiunge parametro
'...........................................................................................
function Aggiornamento__INFOSCHEDE__17(conn)
	Aggiornamento__INFOSCHEDE__17 = "SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__INFOSCHEDE__17(conn)
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & INFOSCHEDE)) <> "" then
		CALL AddParametroSito(conn, "INFOSCHEDE_ID_PAG_ASSEGNA_CENTRO", _
									0, _
									"pagina che viene spedita ad un centro assistenza dopo che gli è stata assegnata una scheda", _
									"", _
									adNumeric, _
									0, _
									"", _
									1, _
									1, _
									INFOSCHEDE, _
									null, null, null, null, null)						
		AggiornamentoSpeciale__INFOSCHEDE__17 = " SELECT * FROM AA_Versione "
	end if
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO INFOSCHEDE 18
'...........................................................................................
' Giacomo 27/05/2011
'...........................................................................................
' aggiungo campi
'...........................................................................................
function Aggiornamento__INFOSCHEDE__18(conn)
    Aggiornamento__INFOSCHEDE__18 = _
		" ALTER TABLE sgtb_schede ADD sc_richiesta_garanzia bit NULL;" & _
		" ALTER TABLE sgtb_schede ADD sc_data_garanzia_controllata smalldatetime NULL;" & _
		" ALTER TABLE sgtb_schede ADD sc_modello_altro " + SQL_CharField(Conn, 500) + " NULL;"
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO SPECIALE INFOSCHEDE 19
'...........................................................................................
' Giacomo 27/05/2011
'...........................................................................................
' aggiunge parametro
'...........................................................................................
function Aggiornamento__INFOSCHEDE__19(conn)
	Aggiornamento__INFOSCHEDE__19 = "SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__INFOSCHEDE__19(conn)
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & INFOSCHEDE)) <> "" then
		CALL AddParametroSito(conn, "INFOSCHEDE_COD_ART_DEFAULT", _
									0, _
									"codice articolo di default", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									INFOSCHEDE, _
									null, null, null, null, null)						
		AggiornamentoSpeciale__INFOSCHEDE__19 = " SELECT * FROM AA_Versione "
	end if
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO SPECIALE INFOSCHEDE 20
'...........................................................................................
' Matteo 30/05/2011
'...........................................................................................
' aggiunge parametro
'...........................................................................................
function Aggiornamento__INFOSCHEDE__20(conn)
	Aggiornamento__INFOSCHEDE__20 = "SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__INFOSCHEDE__20(conn)
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & INFOSCHEDE)) <> "" then
		CALL AddParametroSito(conn, "INFOSCHEDE_ID_PAG_LETT_ACCOMP", _
									20, _
									"pagina per visualizzare la lettera d'accompagnamento", _
									"", _
									adNumeric, _
									0, _
									"", _
									1, _
									1, _
									INFOSCHEDE, _
									null, null, null, null, null)
		CALL AddParametroSito(conn, "INFOSCHEDE_ID_PAG_INVIO_DDT", _
									20, _
									"pagina per l'invio del DDT", _
									"", _
									adNumeric, _
									0, _
									"", _
									1, _
									1, _
									INFOSCHEDE, _
									null, null, null, null, null)
		CALL AddParametroSito(conn, "INFOSCHEDE_TESTO_INVIO_DDT", _
									20, _
									"testo e-mail per l'invio del DDT", _
									"", _
									adLongVarChar, _
									0, _
									"", _
									1, _
									1, _
									INFOSCHEDE, _
									null, null, null, null, null)
		CALL AddParametroSito(conn, "INFOSCHEDE_ID_PAG_INVIO_LETT_ACC", _
									20, _
									"pagina per l'invio della lettera d'accompagnamento", _
									"", _
									adNumeric, _
									0, _
									"", _
									1, _
									1, _
									INFOSCHEDE, _
									null, null, null, null, null)
		CALL AddParametroSito(conn, "INFOSCHEDE_TESTO_INVIO_LETT_ACC", _
									20, _
									"testo e-mail per l'invio della lettera d'accompagnamento", _
									"", _
									adLongVarChar, _
									0, _
									"", _
									1, _
									1, _
									INFOSCHEDE, _
									null, null, null, null, null)						
									
									
		AggiornamentoSpeciale__INFOSCHEDE__20 = " SELECT * FROM AA_Versione "		
	end if
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO INFOSCHEDE 21
'...........................................................................................
' Giacomo 01/08/2011
'...........................................................................................
' aggiungo tabelle per ddt e riferimenti nella tabella ddt
'...........................................................................................
function Aggiornamento__INFOSCHEDE__21(conn)
    Aggiornamento__INFOSCHEDE__21 = _
			"CREATE TABLE " & SQL_Dbo(Conn) & "sgtb_ddt_causali(" + _
			"	cau_id " & SQL_PrimaryKey(conn, "sgtb_ddt_causali") + ", " +_ 
			SQL_MultiLanguageFieldComplete(conn, "cau_titolo_<lingua>" + SQL_CharField(Conn, 255) + " NULL ")  + ", " + _
			"	cau_ordine int NULL " + _
			" ); " + _
			"CREATE TABLE " & SQL_Dbo(Conn) & "sgtb_ddt_porto(" + _
			"	por_id " & SQL_PrimaryKey(conn, "sgtb_ddt_porto") + ", " +_ 
			SQL_MultiLanguageFieldComplete(conn, "por_titolo_<lingua>" + SQL_CharField(Conn, 255) + " NULL ") + _
			" ); " + _
			"CREATE TABLE " & SQL_Dbo(Conn) & "sgtb_ddt_trasporto(" + _
			"	tra_id " & SQL_PrimaryKey(conn, "sgtb_ddt_trasporto") + ", " +_ 
			SQL_MultiLanguageFieldComplete(conn, "tra_titolo_<lingua>" + SQL_CharField(Conn, 255) + " NULL ") + _
			" ); " + _
			SQL_RemoveForeignKey(conn, "sgtb_ddt", "ddt_causale_id", "gtb_stati_ordine", false, "") + _
			SQL_AddForeignKey(conn, "sgtb_ddt", "ddt_causale_id", "sgtb_ddt_causali", "cau_id", false, "") + _
			" ALTER TABLE sgtb_ddt ADD ddt_porto_id int NULL;" + _
			" ALTER TABLE sgtb_ddt ADD ddt_trasporto_id int NULL;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO INFOSCHEDE 22
'...........................................................................................
' Giacomo 01/08/2011
'...........................................................................................
' aggiungo campo per salvare id da import DDT
'...........................................................................................
function Aggiornamento__INFOSCHEDE__22(conn)
    Aggiornamento__INFOSCHEDE__22 = _
		" ALTER TABLE sgtb_ddt ADD ddt_external_id " + SQL_CharField(Conn, 255) + " NULL;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO INFOSCHEDE 23
'...........................................................................................
' Giacomo 01/08/2011
'...........................................................................................
' aggiungo chiavi esterne su ddt
'...........................................................................................
function Aggiornamento__INFOSCHEDE__23(conn)
    Aggiornamento__INFOSCHEDE__23 = _
		SQL_AddForeignKey(conn, "sgtb_ddt", "ddt_porto_id", "sgtb_ddt_porto", "por_id", false, "") + _
		SQL_AddForeignKey(conn, "sgtb_ddt", "ddt_trasporto_id", "sgtb_ddt_trasporto", "tra_id", false, "")
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO INFOSCHEDE 24
'...........................................................................................
' Giacomo 05/08/2011
'...........................................................................................
' cambio formato campo ore_manodopera
'...........................................................................................
function Aggiornamento__INFOSCHEDE__24(conn)
    Aggiornamento__INFOSCHEDE__24 = _
		" ALTER TABLE sgtb_schede ALTER COLUMN sc_ora_manodopera_intervento real NULL;"
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO INFOSCHEDE 25
'...........................................................................................
' Nicola 08/08/2011
'...........................................................................................
' aggiunge suddivisione a stati delle schede per visualizzazione nei ddt e nei ritiri
'...........................................................................................
function Aggiornamento__INFOSCHEDE__25(conn)
    Aggiornamento__INFOSCHEDE__25 = _
		" ALTER TABLE sgtb_stati_schede ADD " + _
		"	sts_elenco_ddt_da_consegnare BIT null, " + _
		"	sts_elenco_ddt_da_ritirare BIT null; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO INFOSCHEDE 26
'...........................................................................................
' Giacomo 27/10/2014
'...........................................................................................
' aggiunge tabella dettagli ddt
'...........................................................................................
function Aggiornamento__INFOSCHEDE__26(conn)
    Aggiornamento__INFOSCHEDE__26 = _
			"CREATE TABLE " + SQL_Dbo(conn) + "sgtb_dettagli_ddt(" + _
			"	dtd_id " + SQL_PrimaryKey(conn, "sgtb_dettagli_ddt") + ", " + _
			"	dtd_ddt_id int NULL, " + _
			"	dtd_articolo_id int NULL, " + _
			" 	dtd_articolo_codice " + SQL_CharField(Conn, 100) + " NULL, " + _
			" 	dtd_articolo_nome " + SQL_CharField(Conn, 255) + " NULL, " + _
			" 	dtd_articolo_qta int NULL, " + _
			" 	dtd_rif_vs_ddt " + SQL_CharField(Conn, 100) + " NULL, " + _
			" 	dtd_in_garanzia bit NULL " + _
			"); " + _
			SQL_AddForeignKey(conn, "sgtb_dettagli_ddt", "dtd_articolo_id", "grel_art_valori", "rel_id", false, "") + _
			SQL_AddForeignKeyExtended(conn, "sgtb_dettagli_ddt", "dtd_ddt_id", "sgtb_ddt", "ddt_id", true, false, "")
	
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO INFOSCHEDE 27
'...........................................................................................
' Giacomo 31/10/2014
'...........................................................................................
' aggiunge campi su ddt e dettagli ddt
'...........................................................................................
function Aggiornamento__INFOSCHEDE__27(conn)
    Aggiornamento__INFOSCHEDE__27 = _
		" ALTER TABLE sgtb_dettagli_ddt ADD " + _
		" 	dtd_articolo_prezzo_unitario money NULL, " + _
		" 	dtd_articolo_sconto real NULL; " + _
		" ALTER TABLE sgtb_ddt ADD " + _
		"	ddt_contrassegno " + SQL_CharField(Conn, 100)
end function
'*******************************************************************************************



%>
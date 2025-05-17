<%
'...........................................................................................
'...........................................................................................
'libreria di funzioni che contiene tutti gli aggiornamenti per il NEXT-booking3
'...........................................................................................
'...........................................................................................


'*******************************************************************************************
'INSTALLAZIONE booking3
'...........................................................................................
function Install__BOOKING2(conn)
	Select case DB_Type(conn)
		case DB_Access
			Install__BOOKING2 = _
				"CREATE TABLE htb_strutture (" + _
				"str_id COUNTER CONSTRAINT PK_htb_strutture PRIMARY KEY ," + _
				"str_paginaSito_id INTEGER NULL, " + _
				"str_nome TEXT(250) WITH COMPRESSION NULL ," + _
				"str_descrizione_it TEXT WITH COMPRESSION NULL ," + _
				"str_descrizione_en TEXT WITH COMPRESSION NULL ," + _
				"str_descrizione_fr TEXT WITH COMPRESSION NULL ," + _
				"str_descrizione_es TEXT WITH COMPRESSION NULL ," + _
				"str_descrizione_de TEXT WITH COMPRESSION NULL, " + _
				"str_sito TEXT(250) WITH COMPRESSION NULL ," + _
				"str_ordine INTEGER NULL, " + _
				"str_logo TEXT(250) WITH COMPRESSION NULL, " + _
				"str_paginaAccetta_id INTEGER NULL, " + _
				"str_paginaAnnulla_id INTEGER NULL" + _
				"); " + _
				"CREATE TABLE hrel_strutture_admin (" + _
				"rsa_id COUNTER CONSTRAINT PK_hrel_strutture_admin PRIMARY KEY ," + _
				"rsa_struttura_id INTEGER NULL, " + _
				"rsa_admin_id INTEGER NULL" + _
				"); " + _
				"CREATE TABLE htb_tipiCamera (" + _
				"tipC_id COUNTER CONSTRAINT PK_htb_tipiCamera PRIMARY KEY ," + _
				"tipC_struttura_id INTEGER NULL, " + _
				"tipC_nome_it TEXT(50) WITH COMPRESSION NULL ," + _
				"tipC_nome_en TEXT(50) WITH COMPRESSION NULL ," + _
				"tipC_nome_fr TEXT(50) WITH COMPRESSION NULL ," + _
				"tipC_nome_es TEXT(50) WITH COMPRESSION NULL ," + _
				"tipC_nome_de TEXT(50) WITH COMPRESSION NULL, " + _
				"tipC_descrizione_it TEXT WITH COMPRESSION NULL ," + _
				"tipC_descrizione_en TEXT WITH COMPRESSION NULL ," + _
				"tipC_descrizione_fr TEXT WITH COMPRESSION NULL ," + _
				"tipC_descrizione_es TEXT WITH COMPRESSION NULL ," + _
				"tipC_descrizione_de TEXT WITH COMPRESSION NULL, " + _
				"tipC_ordine INTEGER NULL, " + _
				"tipC_numero INTEGER NULL, " + _
				"tipC_immagine TEXT(250) WITH COMPRESSION NULL" + _
				"); " + _
				"CREATE TABLE htb_disponibilita (" + _
				"dis_id COUNTER CONSTRAINT PK_htb_disponibilita PRIMARY KEY ," + _
				"dis_prezzo CURRENCY NULL ," + _
				"dis_data DATETIME NULL ," + _
				"dis_disponibilita INTEGER NULL ," + _
				"dis_tipo_id INTEGER NULL, " + _
				"dis_min_stay INTEGER NULL, " + _
				"dis_promozione BIT NULL, " + _
				"dis_bloccata BIT NULL" + _
				"); " + _
				"CREATE TABLE htb_listini (" + _
				"lis_id COUNTER CONSTRAINT PK_htb_listini PRIMARY KEY ," + _
				"lis_struttura_id INTEGER NULL ," + _
				"lis_nome_it TEXT(100) WITH COMPRESSION NULL ," + _
				"lis_nome_en TEXT(100) WITH COMPRESSION NULL ," + _
				"lis_nome_fr TEXT(100) WITH COMPRESSION NULL ," + _
				"lis_nome_es TEXT(100) WITH COMPRESSION NULL ," + _
				"lis_nome_de TEXT(100) WITH COMPRESSION NULL ," + _
				"lis_condizioni_it TEXT WITH COMPRESSION NULL ," + _
				"lis_condizioni_en TEXT WITH COMPRESSION NULL ," + _
				"lis_condizioni_fr TEXT WITH COMPRESSION NULL ," + _
				"lis_condizioni_es TEXT WITH COMPRESSION NULL ," + _
				"lis_condizioni_de TEXT WITH COMPRESSION NULL ," + _
				"lis_data DATETIME NULL" + _
				"); " + _
				"CREATE TABLE htb_listini_tipiCamera (" + _
				"rlt_id COUNTER CONSTRAINT PK_htb_listini_tipiCamera PRIMARY KEY ," + _
				"rlt_prezzo CURRENCY NULL ," + _
				"rlt_listino_id INTEGER NULL ," + _
				"rlt_tipo_id INTEGER NULL" + _
				"); " + _
				"CREATE TABLE htb_prenotazioni (" + _
				"pre_id COUNTER CONSTRAINT PK_htb_prenotazioni PRIMARY KEY ," + _
				"pre_stato_id INTEGER NULL, " + _
				"pre_data DATETIME NULL ," + _
				"pre_data_inizio DATETIME NULL ," + _
				"pre_data_fine DATETIME NULL ," + _
				"pre_note TEXT WITH COMPRESSION NULL ," + _
				"pre_cliente_id INTEGER NULL, " + _
				"pre_nomeCC TEXT(255) WITH COMPRESSION NULL, " + _
				"pre_numeroCC TEXT(255) WITH COMPRESSION NULL, " + _
				"pre_dataCC DATETIME NULL, " + _
				"pre_tipoCC TEXT(50) WITH COMPRESSION NULL, " + _
				"pre_totale INTEGER NULL, " + _
				"pre_meseCC INTEGER NULL, " + _
				"pre_annoCC INTEGER NULL, " + _
				"pre_cvcCC TEXT(5) NULL, " + _
				"pre_chiave TEXT(10) NULL" + _
				"); " + _
				"CREATE TABLE htb_prenotazioni_tipiCamera (" + _
				"rpt_id COUNTER CONSTRAINT PK_htb_prenotazioni_tipiCamera PRIMARY KEY ," + _
				"rpt_numero INTEGER NULL ," + _
				"rpt_prenotazione_id INTEGER NULL ," + _
				"rpt_tipo_id INTEGER NULL" + _
				"); " + _
				"CREATE TABLE htb_prenotazioniStati (" + _
				"pst_id COUNTER CONSTRAINT PK_htb_prenotazioniStati PRIMARY KEY ," + _
				"pst_nome_it TEXT(50) WITH COMPRESSION NULL ," + _
				"pst_nome_en TEXT(50) WITH COMPRESSION NULL ," + _
				"pst_nome_fr TEXT(50) WITH COMPRESSION NULL ," + _
				"pst_nome_es TEXT(50) WITH COMPRESSION NULL ," + _
				"pst_nome_de TEXT(50) WITH COMPRESSION NULL " + _
				"); " + _
				"INSERT INTO htb_prenotazioniStati(pst_nome_it, pst_nome_en, pst_nome_fr, pst_nome_es, pst_nome_de) "+ _
				"VALUES ('Richiesta', 'Requested', 'Demandé', 'Solicitado', 'Verlangt');"+ _
				"INSERT INTO htb_prenotazioniStati(pst_nome_it, pst_nome_en, pst_nome_fr, pst_nome_es, pst_nome_de) "+ _
				"VALUES ('Accettata', 'Accepted', 'Admis', 'Aceptado', 'Angenommen');"+ _
				"INSERT INTO htb_prenotazioniStati(pst_nome_it, pst_nome_en, pst_nome_fr, pst_nome_es, pst_nome_de) "+ _
				"VALUES ('Annullata', 'Cancelled', 'Décommandé', 'Cancelado', 'Annulliert');"+ _
				"INSERT INTO htb_prenotazioniStati(pst_nome_it, pst_nome_en, pst_nome_fr, pst_nome_es, pst_nome_de) "+ _
				"VALUES ('Confermata', 'Confirmed', 'Confirmé', 'Confirmado', 'Bestätigt');"+ _
				"ALTER TABLE hrel_strutture_admin ADD CONSTRAINT FK_hrel_strutture_admin__htb_strutture " + _
				"FOREIGN KEY (rsa_struttura_id) REFERENCES htb_strutture (str_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;" + _
				"ALTER TABLE hrel_strutture_admin ADD CONSTRAINT FK_hrel_strutture_admin__tb_admin " + _
				"FOREIGN KEY (rsa_admin_id) REFERENCES tb_admin (id_admin) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;" + _
				"ALTER TABLE htb_tipiCamera ADD CONSTRAINT FK_htb_tipiCamera__htb_strutture " + _
				"FOREIGN KEY (tipC_struttura_id) REFERENCES htb_strutture (str_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;" + _
				"ALTER TABLE htb_listini ADD CONSTRAINT FK_htb_listini__htb_strutture " + _
				"FOREIGN KEY (lis_struttura_id) REFERENCES htb_strutture (str_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;" + _
				"ALTER TABLE htb_disponibilita ADD CONSTRAINT FK_htb_disponibilita__htb_tipiCamera " + _
				"FOREIGN KEY (dis_tipo_id) REFERENCES htb_tipiCamera (tipC_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;" + _
				"ALTER TABLE htb_listini_tipiCamera ADD CONSTRAINT FK_htb_listini_tipiCamera__htb_tipiCamera " + _
				"FOREIGN KEY (rlt_tipo_id) REFERENCES htb_tipiCamera (tipC_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;" + _
				"ALTER TABLE htb_listini_tipiCamera ADD CONSTRAINT FK_htb_listini_tipiCamera__htb_listini " + _
				"FOREIGN KEY (rlt_listino_id) REFERENCES htb_listini (lis_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;" + _
				"ALTER TABLE htb_prenotazioni_tipiCamera ADD CONSTRAINT FK_htb_prenotazioni_tipiCamera__htb_tipiCamera " + _
				"FOREIGN KEY (rpt_tipo_id) REFERENCES htb_tipiCamera (tipC_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;" + _
				"ALTER TABLE htb_prenotazioni_tipiCamera ADD CONSTRAINT FK_htb_prenotazioni_tipiCamera__htb_prenotazioni " + _
				"FOREIGN KEY (rpt_prenotazione_id) REFERENCES htb_prenotazioni (pre_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;"+ _
				"ALTER TABLE htb_prenotazioni ADD CONSTRAINT FK_htb_prenotazioni_tb_indirizzario " + _
				"FOREIGN KEY (pre_cliente_id) REFERENCES tb_indirizzario (IDElencoIndirizzi) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;"+ _
				"ALTER TABLE htb_prenotazioni ADD CONSTRAINT FK_htb_prenotazioni_htb_prenotazioniStati " + _
				"FOREIGN KEY (pre_stato_id) REFERENCES htb_prenotazioniStati (pst_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;"

		case DB_SQL
			Install__BOOKING2 = _
				"CREATE TABLE dbo.htb_strutture (" + _
				"str_id COUNTER CONSTRAINT PK_htb_strutture PRIMARY KEY ," + _
				"str_paginaSito_id INTEGER NULL, " + _
				"str_nome NVARCHAR(250) NULL ," + _
				"str_descrizione_it NTEXT NULL ," + _
				"str_descrizione_en NTEXT NULL ," + _
				"str_descrizione_fr NTEXT NULL ," + _
				"str_descrizione_es NTEXT NULL ," + _
				"str_descrizione_de NTEXT NULL, " + _
				"str_sito NVARCHAR(250) NULL ," + _
				"str_ordine INTEGER NULL, " + _
				"str_logo NVARCHAR(250) NULL, " + _
				"str_paginaAccetta_id INTEGER NULL, " + _
				"str_paginaAnnulla_id INTEGER NULL" + _
				"); " + _
				"CREATE TABLE dbo.hrel_strutture_admin (" + _
				"rsa_id COUNTER CONSTRAINT PK_hrel_strutture_admin PRIMARY KEY ," + _
				"rsa_struttura_id INTEGER NULL, " + _
				"rsa_admin_id INTEGER NULL" + _
				"); " + _
				"CREATE TABLE dbo.htb_tipiCamera (" + _
				"tipC_id COUNTER CONSTRAINT PK_htb_tipiCamera PRIMARY KEY ," + _
				"tipC_struttura_id INTEGER NULL, " + _
				"tipC_nome_it NVARCHAR(50) NULL ," + _
				"tipC_nome_en NVARCHAR(50) NULL ," + _
				"tipC_nome_fr NVARCHAR(50) NULL ," + _
				"tipC_nome_es NVARCHAR(50) NULL ," + _
				"tipC_nome_de NVARCHAR(50) NULL, " + _
				"tipC_descrizione_it NTEXT NULL ," + _
				"tipC_descrizione_en NTEXT NULL ," + _
				"tipC_descrizione_fr NTEXT NULL ," + _
				"tipC_descrizione_es NTEXT NULL ," + _
				"tipC_descrizione_de NTEXT NULL, " + _
				"tipC_ordine INTEGER NULL, " + _
				"tipC_numero INTEGER NULL, " + _
				"tipC_immagine NVARCHAR(250) NULL" + _
				"); " + _
				"CREATE TABLE dbo.htb_disponibilita (" + _
				"dis_id COUNTER CONSTRAINT PK_htb_disponibilita PRIMARY KEY ," + _
				"dis_prezzo CURRENCY NULL ," + _
				"dis_data DATETIME NULL ," + _
				"dis_disponibilita INTEGER NULL ," + _
				"dis_tipo_id INTEGER NULL, " + _
				"dis_min_stay INTEGER NULL, " + _
				"dis_promozione BIT NULL, " + _
				"dis_bloccata BIT NULL" + _
				"); " + _
				"CREATE TABLE dbo.htb_listini (" + _
				"lis_id COUNTER CONSTRAINT PK_htb_listini PRIMARY KEY ," + _
				"lis_struttura_id INTEGER NULL ," + _
				"lis_nome_it NVARCHAR(100) NULL ," + _
				"lis_nome_en NVARCHAR(100) NULL ," + _
				"lis_nome_fr NVARCHAR(100) NULL ," + _
				"lis_nome_es NVARCHAR(100) NULL ," + _
				"lis_nome_de NVARCHAR(100) NULL ," + _
				"lis_condizioni_it NTEXT NULL ," + _
				"lis_condizioni_en NTEXT NULL ," + _
				"lis_condizioni_fr NTEXT NULL ," + _
				"lis_condizioni_es NTEXT NULL ," + _
				"lis_condizioni_de NTEXT NULL ," + _
				"lis_data DATETIME NULL" + _
				"); " + _
				"CREATE TABLE dbo.htb_listini_tipiCamera (" + _
				"rlt_id COUNTER CONSTRAINT PK_htb_listini_tipiCamera PRIMARY KEY ," + _
				"rlt_prezzo CURRENCY NULL ," + _
				"rlt_listino_id INTEGER NULL ," + _
				"rlt_tipo_id INTEGER NULL" + _
				"); " + _
				"CREATE TABLE dbo.htb_prenotazioni (" + _
				"pre_id COUNTER CONSTRAINT PK_htb_prenotazioni PRIMARY KEY ," + _
				"pre_stato_id INTEGER NULL, " + _
				"pre_data DATETIME NULL ," + _
				"pre_data_inizio DATETIME NULL ," + _
				"pre_data_fine DATETIME NULL ," + _
				"pre_note NTEXT NULL ," + _
				"pre_cliente_id INTEGER NULL, " + _
				"pre_nomeCC NVARCHAR(255) NULL, " + _
				"pre_numeroCC NVARCHAR(255) NULL, " + _
				"pre_dataCC DATETIME NULL, " + _
				"pre_tipoCC NVARCHAR(50) NULL, " + _
				"pre_totale INTEGER NULL, " + _
				"pre_meseCC INTEGER NULL, " + _
				"pre_annoCC INTEGER NULL, " + _
				"pre_cvcCC NVARCHAR(5) NULL, " + _
				"pre_chiave NVARCHAR(10) NULL" + _
				"); " + _
				"CREATE TABLE dbo.htb_prenotazioni_tipiCamera (" + _
				"rpt_id COUNTER CONSTRAINT PK_htb_prenotazioni_tipiCamera PRIMARY KEY ," + _
				"rpt_numero INTEGER NULL ," + _
				"rpt_prenotazione_id INTEGER NULL ," + _
				"rpt_tipo_id INTEGER NULL" + _
				"); " + _
				"CREATE TABLE dbo.htb_prenotazioniStati (" + _
				"pst_id COUNTER CONSTRAINT PK_htb_prenotazioniStati PRIMARY KEY ," + _
				"pst_nome_it NVARCHAR(50) NULL ," + _
				"pst_nome_en NVARCHAR(50) NULL ," + _
				"pst_nome_fr NVARCHAR(50) NULL ," + _
				"pst_nome_es NVARCHAR(50) NULL ," + _
				"pst_nome_de NVARCHAR(50) NULL " + _
				"); " + _
				"INSERT INTO htb_prenotazioniStati(pst_nome_it, pst_nome_en, pst_nome_fr, pst_nome_es, pst_nome_de) "+ _
				"VALUES ('Richiesta', 'Requested', 'Demandé', 'Solicitado', 'Verlangt');"+ _
				"INSERT INTO htb_prenotazioniStati(pst_nome_it, pst_nome_en, pst_nome_fr, pst_nome_es, pst_nome_de) "+ _
				"VALUES ('Accettata', 'Accepted', 'Admis', 'Aceptado', 'Angenommen');"+ _
				"INSERT INTO htb_prenotazioniStati(pst_nome_it, pst_nome_en, pst_nome_fr, pst_nome_es, pst_nome_de) "+ _
				"VALUES ('Annullata', 'Cancelled', 'Décommandé', 'Cancelado', 'Annulliert');"+ _
				"INSERT INTO htb_prenotazioniStati(pst_nome_it, pst_nome_en, pst_nome_fr, pst_nome_es, pst_nome_de) "+ _
				"VALUES ('Confermata', 'Confirmed', 'Confirmé', 'Confirmado', 'Bestätigt');"+ _
				"ALTER TABLE hrel_strutture_admin ADD CONSTRAINT FK_hrel_strutture_admin__htb_strutture " + _
				"FOREIGN KEY (rsa_struttura_id) REFERENCES htb_strutture (str_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;" + _
				"ALTER TABLE hrel_strutture_admin ADD CONSTRAINT FK_hrel_strutture_admin__tb_admin " + _
				"FOREIGN KEY (rsa_admin_id) REFERENCES tb_admin (id_admin) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;" + _
				"ALTER TABLE htb_tipiCamera ADD CONSTRAINT FK_htb_tipiCamera__htb_strutture " + _
				"FOREIGN KEY (tipC_struttura_id) REFERENCES htb_strutture (str_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;" + _
				"ALTER TABLE htb_listini ADD CONSTRAINT FK_htb_listini__htb_strutture " + _
				"FOREIGN KEY (lis_struttura_id) REFERENCES htb_strutture (str_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;" + _
				"ALTER TABLE htb_disponibilita ADD CONSTRAINT FK_htb_disponibilita__htb_tipiCamera " + _
				"FOREIGN KEY (dis_tipo_id) REFERENCES htb_tipiCamera (tipC_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;" + _
				"ALTER TABLE htb_listini_tipiCamera ADD CONSTRAINT FK_htb_listini_tipiCamera__htb_tipiCamera " + _
				"FOREIGN KEY (rlt_tipo_id) REFERENCES htb_tipiCamera (tipC_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;" + _
				"ALTER TABLE htb_listini_tipiCamera ADD CONSTRAINT FK_htb_listini_tipiCamera__htb_listini " + _
				"FOREIGN KEY (rlt_listino_id) REFERENCES htb_listini (lis_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;" + _
				"ALTER TABLE htb_prenotazioni_tipiCamera ADD CONSTRAINT FK_htb_prenotazioni_tipiCamera__htb_tipiCamera " + _
				"FOREIGN KEY (rpt_tipo_id) REFERENCES htb_tipiCamera (tipC_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;" + _
				"ALTER TABLE htb_prenotazioni_tipiCamera ADD CONSTRAINT FK_htb_prenotazioni_tipiCamera__htb_prenotazioni " + _
				"FOREIGN KEY (rpt_prenotazione_id) REFERENCES htb_prenotazioni (pre_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;"+ _
				"ALTER TABLE htb_prenotazioni ADD CONSTRAINT FK_htb_prenotazioni_tb_indirizzario " + _
				"FOREIGN KEY (pre_cliente_id) REFERENCES tb_indirizzario (IDElencoIndirizzi) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;"+ _
				"ALTER TABLE htb_prenotazioni ADD CONSTRAINT FK_htb_prenotazioni_htb_prenotazioniStati " + _
				"FOREIGN KEY (pre_stato_id) REFERENCES htb_prenotazioniStati (pst_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;"

	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO booking3 1
'...........................................................................................
'aggiungo campi per descrivere e-mail
'...........................................................................................
function Aggiornamento__BOOKING2__1(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__BOOKING2__1 = "ALTER TABLE htb_strutture ADD "+ _
											"str_emailAdmin_id INT NULL, "+ _
											"str_emailAccettaObj_it TEXT(100) WITH COMPRESSION NULL, "+ _
											"str_emailAccettaObj_en TEXT(100) WITH COMPRESSION NULL, "+ _
											"str_emailAccettaObj_fr TEXT(100) WITH COMPRESSION NULL, "+ _
											"str_emailAccettaObj_es TEXT(100) WITH COMPRESSION NULL, "+ _
											"str_emailAccettaObj_de TEXT(100) WITH COMPRESSION NULL, "+ _
											"str_emailAnnullaObj_it TEXT(100) WITH COMPRESSION NULL, "+ _
											"str_emailAnnullaObj_en TEXT(100) WITH COMPRESSION NULL, "+ _
											"str_emailAnnullaObj_fr TEXT(100) WITH COMPRESSION NULL, "+ _
											"str_emailAnnullaObj_es TEXT(100) WITH COMPRESSION NULL, "+ _
											"str_emailAnnullaObj_de TEXT(100) WITH COMPRESSION NULL, "+ _
											"str_emailConfermaObj_it TEXT(100) WITH COMPRESSION NULL, "+ _
											"str_emailConfermaObj_en TEXT(100) WITH COMPRESSION NULL, "+ _
											"str_emailConfermaObj_fr TEXT(100) WITH COMPRESSION NULL, "+ _
											"str_emailConfermaObj_es TEXT(100) WITH COMPRESSION NULL, "+ _
											"str_emailConfermaObj_de TEXT(100) WITH COMPRESSION NULL;"
		case DB_SQL
			Aggiornamento__BOOKING2__1 = "ALTER TABLE htb_strutture ADD "+ _
											"str_emailAdmin_id INT NULL, "+ _
											"str_emailAccettaObj_it NVARCHAR(100) NULL, "+ _
											"str_emailAccettaObj_en NVARCHAR(100) NULL, "+ _
											"str_emailAccettaObj_fr NVARCHAR(100) NULL, "+ _
											"str_emailAccettaObj_es NVARCHAR(100) NULL, "+ _
											"str_emailAccettaObj_de NVARCHAR(100) NULL, "+ _
											"str_emailAnnullaObj_it NVARCHAR(100) NULL, "+ _
											"str_emailAnnullaObj_en NVARCHAR(100) NULL, "+ _
											"str_emailAnnullaObj_fr NVARCHAR(100) NULL, "+ _
											"str_emailAnnullaObj_es NVARCHAR(100) NULL, "+ _
											"str_emailAnnullaObj_de NVARCHAR(100) NULL, "+ _
											"str_emailConfermaObj_it NVARCHAR(100) NULL, "+ _
											"str_emailConfermaObj_en NVARCHAR(100) NULL, "+ _
											"str_emailConfermaObj_fr NVARCHAR(100) NULL, "+ _
											"str_emailConfermaObj_es NVARCHAR(100) NULL, "+ _
											"str_emailConfermaObj_de NVARCHAR(100) NULL;"
	end select
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO booking3 2
'...........................................................................................
'aggiungo campi per descrivere e-mail
'...........................................................................................
function Aggiornamento__BOOKING2__2(conn)
	Select case DB_Type(conn)
		case DB_Access, DB_SQL
			Aggiornamento__BOOKING2__2 = "ALTER TABLE htb_strutture ADD "+ _
											"str_paginaConferma_id INT NULL;"
	end select
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO booking3 3
'...........................................................................................
'aggiungo campi per descrivere e-mail
'...........................................................................................
function Aggiornamento__BOOKING2__3(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__BOOKING2__3 = "ALTER TABLE htb_strutture ADD "+ _
											"str_paginaRichiesta_id INT NULL, "+ _
											"str_emailRichiestaObj_it TEXT(100) WITH COMPRESSION NULL, "+ _
											"str_emailRichiestaObj_en TEXT(100) WITH COMPRESSION NULL, "+ _
											"str_emailRichiestaObj_fr TEXT(100) WITH COMPRESSION NULL, "+ _
											"str_emailRichiestaObj_es TEXT(100) WITH COMPRESSION NULL, "+ _
											"str_emailRichiestaObj_de TEXT(100) WITH COMPRESSION NULL;"
		case DB_SQL
			Aggiornamento__BOOKING2__3 = "ALTER TABLE htb_strutture ADD "+ _
											"str_paginaRichiesta_id INT NULL, "+ _
											"str_emailRichiestaObj_it NVARCHAR(100) NULL, "+ _
											"str_emailRichiestaObj_en NVARCHAR(100) NULL, "+ _
											"str_emailRichiestaObj_fr NVARCHAR(100) NULL, "+ _
											"str_emailRichiestaObj_es NVARCHAR(100) NULL, "+ _
											"str_emailRichiestaObj_de NVARCHAR(100) NULL;"
	end select
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO booking3 4
'...........................................................................................
'aggiungo campo rubrica all'hotel
'...........................................................................................
function Aggiornamento__BOOKING2__4(conn)
	Select case DB_Type(conn)
		case DB_Access, DB_SQL
			Aggiornamento__BOOKING2__4 = "DROP TABLE hrel_strutture_admin;"+ _
											"ALTER TABLE htb_strutture ADD "+ _
											"str_gruppo_id INT NULL, "+ _
											"str_rubricaRichieste_id INT NULL, "+ _
											"str_rubricaPrenotazioni_id INT NULL;"
	end select
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO booking3 5
'...........................................................................................
'sposta l'impostazione dell'applicativo per hotel
'...........................................................................................
function Aggiornamento__BOOKING2__5(conn)
	Select case DB_Type(conn)
		case DB_Access, DB_SQL
			Aggiornamento__BOOKING2__5 = "ALTER TABLE htb_strutture ADD "+ _
											"str_param_enableNCamere BIT NULL, "+ _
											"str_param_enableHttps BIT NULL, "+ _
											"str_param_disableDispo BIT NULL, "+ _
											"str_param_enableDifferita BIT NULL, "+ _
											"str_param_disableDecrementa BIT NULL;"
	end select
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO booking3 6
'...........................................................................................
'aggiunge un campo di risposta alla prenotazione
'...........................................................................................
function Aggiornamento__BOOKING2__6(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__BOOKING2__6 = "ALTER TABLE htb_prenotazioni ADD "+ _
											"pre_risposta TEXT WITH COMPRESSION NULL;"
		case DB_SQL
			Aggiornamento__BOOKING2__6 = "ALTER TABLE htb_prenotazioni ADD "+ _
											"pre_risposta NTEXT NULL;"
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO booking3 7
'...........................................................................................
'aggiunge un campo numero di persone associate alla prenotazione
'...........................................................................................
function Aggiornamento__BOOKING2__7(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__BOOKING2__7 = "ALTER TABLE htb_prenotazioni ADD "+ _
											"pre_numero_persone INT NULL;"
		case DB_SQL
			Aggiornamento__BOOKING2__7 = "ALTER TABLE htb_prenotazioni ADD "+ _
											"pre_numero_persone INT NULL;"
	end select
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO booking3 8
'...........................................................................................
'aggiunge un campo numero di persone associate alla prenotazione
'...........................................................................................
function Aggiornamento__BOOKING2__8(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__BOOKING2__8 = "ALTER TABLE htb_tipicamera ADD "+ _
											"tipC_posti_letto INT NULL;"
		case DB_SQL
			Aggiornamento__BOOKING2__8 = "ALTER TABLE htb_tipicamera ADD "+ _
											"tipC_posti_letto INT NULL;"
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO booking3 9
'.............................................................................................
'aggiunge i campi per la gestione delle date alternative e la gestione del cliente ospite  
'.............................................................................................
function Aggiornamento__BOOKING2__9(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__BOOKING2__9 = "ALTER TABLE htb_prenotazioni ADD "+ _
											"pre_data_inizio_richiesta datetime NULL, pre_data_fine_richiesta datetime NULL, " + _
											"pre_nome_ospite TEXT(255) WITH COMPRESSION NULL, pre_cognome_ospite TEXT(255) WITH COMPRESSION NULL;"
		case DB_SQL
			Aggiornamento__BOOKING2__9 = "ALTER TABLE htb_prenotazioni ADD "+ _
											"pre_data_inizio_richiesta smalldatetime NULL, pre_data_fine_richiesta datetime NULL, " + _
											"pre_nome_ospite NVARCHAR(255) NULL, pre_cognome_ospite NVARCHAR(255) NULL;"
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO booking3 10
'...........................................................................................
'aggiungo campi per descrivere e-mail
'...........................................................................................
function Aggiornamento__BOOKING2__10(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__BOOKING2__10 = "ALTER TABLE htb_strutture ADD "+ _
											"str_paginaDataAlternativa_id INT NULL, "+ _
											"str_emailDataAlternativaObj_it TEXT(100) WITH COMPRESSION NULL, "+ _
											"str_emailDataAlternativaObj_en TEXT(100) WITH COMPRESSION NULL, "+ _
											"str_emailDataAlternativaObj_fr TEXT(100) WITH COMPRESSION NULL, "+ _
											"str_emailDataAlternativaObj_es TEXT(100) WITH COMPRESSION NULL, "+ _
											"str_emailDataAlternativaObj_de TEXT(100) WITH COMPRESSION NULL;"
		case DB_SQL
			Aggiornamento__BOOKING2__10 = "ALTER TABLE htb_strutture ADD "+ _
											"str_paginaDataAlternativa_id INT NULL, "+ _
											"str_emailDataAlternativaObj_it NVARCHAR(100) NULL, "+ _
											"str_emailDataAlternativaObj_en NVARCHAR(100) NULL, "+ _
											"str_emailDataAlternativaObj_fr NVARCHAR(100) NULL, "+ _
											"str_emailDataAlternativaObj_es NVARCHAR(100) NULL, "+ _
											"str_emailDataAlternativaObj_de NVARCHAR(100) NULL;"
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO booking3 11
'...........................................................................................
'rimuove disabilitazione campo inutile
'...........................................................................................
function Aggiornamento__BOOKING2__11(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__BOOKING2__11 = "ALTER TABLE htb_strutture DROP str_param_enableNCamere ; "
		case DB_SQL
			Aggiornamento__BOOKING2__11 = "ALTER TABLE htb_strutture DROP COLUMN str_param_enableNCamere ; "
	end select
end function
'*******************************************************************************************
%>
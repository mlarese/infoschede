<%
'...........................................................................................
'........................................................................................... 
'libreria di funzioni che contiene tutti gli aggiornamenti per il NEXT-BOOKING3
'...........................................................................................
'...........................................................................................


'*******************************************************************************************
'INSTALLAZIONE BOOKING3
'...........................................................................................
function Install__BOOKING3(conn)
	Select case DB_Type(conn)
		case DB_Access
			Install__BOOKING3 = _
				"CREATE TABLE vtb_strutture (" + _
				"str_id COUNTER CONSTRAINT PK_vtb_strutture PRIMARY KEY ," + _
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
				"CREATE TABLE vrel_strutture_admin (" + _
				"rsa_id COUNTER CONSTRAINT PK_vrel_strutture_admin PRIMARY KEY ," + _
				"rsa_struttura_id INTEGER NULL, " + _
				"rsa_admin_id INTEGER NULL" + _
				"); " + _
				"CREATE TABLE vtb_tipiCameraBase (" + _
				"tipCB_id COUNTER CONSTRAINT PK_vtb_tipiCameraBase PRIMARY KEY ," + _
				"tipCB_nome_it TEXT(50) WITH COMPRESSION NULL ," + _
				"tipCB_nome_en TEXT(50) WITH COMPRESSION NULL ," + _
				"tipCB_nome_fr TEXT(50) WITH COMPRESSION NULL ," + _
				"tipCB_nome_es TEXT(50) WITH COMPRESSION NULL ," + _
				"tipCB_nome_de TEXT(50) WITH COMPRESSION NULL, " + _
				"tipCB_descrizione_it TEXT WITH COMPRESSION NULL ," + _
				"tipCB_descrizione_en TEXT WITH COMPRESSION NULL ," + _
				"tipCB_descrizione_fr TEXT WITH COMPRESSION NULL ," + _
				"tipCB_descrizione_es TEXT WITH COMPRESSION NULL ," + _
				"tipCB_descrizione_de TEXT WITH COMPRESSION NULL, " + _
				"tipCB_ordine INTEGER NULL, " + _
				"tipCB_numero INTEGER NULL, " + _
				"tipCB_immagine TEXT(250) WITH COMPRESSION NULL" + _
				"); " + _
				"CREATE TABLE vtb_tipiCamera (" + _
				"tipC_id COUNTER CONSTRAINT PK_vtb_tipiCamera PRIMARY KEY ," + _
				"tipC_struttura_id INTEGER NULL, " + _
				"tipC_tipoCameraBase_id INTEGER NULL, " + _
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
				"CREATE TABLE vtb_disponibilita (" + _
				"dis_id COUNTER CONSTRAINT PK_vtb_disponibilita PRIMARY KEY ," + _
				"dis_prezzo CURRENCY NULL ," + _
				"dis_data DATETIME NULL ," + _
				"dis_disponibilita INTEGER NULL ," + _
				"dis_tipo_id INTEGER NULL, " + _
				"dis_min_stay INTEGER NULL, " + _
				"dis_promozione BIT NULL, " + _
				"dis_bloccata BIT NULL" + _
				"); " + _
				"CREATE TABLE vtb_listini (" + _
				"lis_id COUNTER CONSTRAINT PK_vtb_listini PRIMARY KEY ," + _
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
				"CREATE TABLE vtb_listini_tipiCamera (" + _
				"rlt_id COUNTER CONSTRAINT PK_vtb_listini_tipiCamera PRIMARY KEY ," + _
				"rlt_prezzo CURRENCY NULL ," + _
				"rlt_listino_id INTEGER NULL ," + _
				"rlt_tipo_id INTEGER NULL" + _
				"); " + _
				"CREATE TABLE vtb_prenotazioni (" + _
				"pre_id COUNTER CONSTRAINT PK_vtb_prenotazioni PRIMARY KEY ," + _
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
				"CREATE TABLE vtb_prenotazioni_tipiCamera (" + _
				"rpt_id COUNTER CONSTRAINT PK_vtb_prenotazioni_tipiCamera PRIMARY KEY ," + _
				"rpt_numero INTEGER NULL ," + _
				"rpt_prenotazione_id INTEGER NULL ," + _
				"rpt_tipo_id INTEGER NULL" + _
				"); " + _
				"CREATE TABLE vtb_prenotazioniStati (" + _
				"pst_id COUNTER CONSTRAINT PK_vtb_prenotazioniStati PRIMARY KEY ," + _
				"pst_nome_it TEXT(50) WITH COMPRESSION NULL ," + _
				"pst_nome_en TEXT(50) WITH COMPRESSION NULL ," + _
				"pst_nome_fr TEXT(50) WITH COMPRESSION NULL ," + _
				"pst_nome_es TEXT(50) WITH COMPRESSION NULL ," + _
				"pst_nome_de TEXT(50) WITH COMPRESSION NULL " + _
				"); " + _
				"INSERT INTO vtb_prenotazioniStati(pst_nome_it, pst_nome_en, pst_nome_fr, pst_nome_es, pst_nome_de) "+ _
				"VALUES ('Richiesta', 'Requested', 'DemandÃ©', 'Solicitado', 'Verlangt');"+ _
				"INSERT INTO vtb_prenotazioniStati(pst_nome_it, pst_nome_en, pst_nome_fr, pst_nome_es, pst_nome_de) "+ _
				"VALUES ('Accettata', 'Accepted', 'Admis', 'Aceptado', 'Angenommen');"+ _
				"INSERT INTO vtb_prenotazioniStati(pst_nome_it, pst_nome_en, pst_nome_fr, pst_nome_es, pst_nome_de) "+ _
				"VALUES ('Annullata', 'Cancelled', 'DÃ©commandÃ©', 'Cancelado', 'Annulliert');"+ _
				"INSERT INTO vtb_prenotazioniStati(pst_nome_it, pst_nome_en, pst_nome_fr, pst_nome_es, pst_nome_de) "+ _
				"VALUES ('Confermata', 'Confirmed', 'ConfirmÃ©', 'Confirmado', 'BestÃ¤tigt');"+ _
				"ALTER TABLE vrel_strutture_admin ADD CONSTRAINT FK_hrel_strutture_admin__vtb_strutture " + _
				"FOREIGN KEY (rsa_struttura_id) REFERENCES vtb_strutture (str_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;" + _
				"ALTER TABLE vrel_strutture_admin ADD CONSTRAINT FK_hrel_strutture_admin__tb_admin " + _
				"FOREIGN KEY (rsa_admin_id) REFERENCES tb_admin (id_admin) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;" + _
				"ALTER TABLE vtb_tipiCamera ADD CONSTRAINT FK_vtb_tipiCamera__vtb_strutture " + _
				"FOREIGN KEY (tipC_struttura_id) REFERENCES vtb_strutture (str_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;" + _
				"ALTER TABLE vtb_tipiCamera ADD CONSTRAINT FK_vtb_tipiCamera__vtb_tipiCameraBase " + _
				"FOREIGN KEY (tipC_tipoCameraBase_id) REFERENCES vtb_tipiCameraBase (tipCB_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;" + _
				"ALTER TABLE vtb_listini ADD CONSTRAINT FK_vtb_listini__vtb_strutture " + _
				"FOREIGN KEY (lis_struttura_id) REFERENCES vtb_strutture (str_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;" + _
				"ALTER TABLE vtb_disponibilita ADD CONSTRAINT FK_vtb_disponibilita__vtb_tipiCamera " + _
				"FOREIGN KEY (dis_tipo_id) REFERENCES vtb_tipiCamera (tipC_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;" + _
				"ALTER TABLE vtb_listini_tipiCamera ADD CONSTRAINT FK_vtb_listini_tipiCamera__vtb_tipiCamera " + _
				"FOREIGN KEY (rlt_tipo_id) REFERENCES vtb_tipiCamera (tipC_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;" + _
				"ALTER TABLE vtb_listini_tipiCamera ADD CONSTRAINT FK_vtb_listini_tipiCamera__vtb_listini " + _
				"FOREIGN KEY (rlt_listino_id) REFERENCES vtb_listini (lis_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;" + _
				"ALTER TABLE vtb_prenotazioni_tipiCamera ADD CONSTRAINT FK_vtb_prenotazioni_tipiCamera__vtb_tipiCamera " + _
				"FOREIGN KEY (rpt_tipo_id) REFERENCES vtb_tipiCamera (tipC_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;" + _
				"ALTER TABLE vtb_prenotazioni_tipiCamera ADD CONSTRAINT FK_vtb_prenotazioni_tipiCamera__vtb_prenotazioni " + _
				"FOREIGN KEY (rpt_prenotazione_id) REFERENCES vtb_prenotazioni (pre_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;"+ _
				"ALTER TABLE vtb_prenotazioni ADD CONSTRAINT FK_vtb_prenotazioni_tb_indirizzario " + _
				"FOREIGN KEY (pre_cliente_id) REFERENCES tb_indirizzario (IDElencoIndirizzi) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;"+ _
				"ALTER TABLE vtb_prenotazioni ADD CONSTRAINT FK_vtb_prenotazioni_vtb_prenotazioniStati " + _
				"FOREIGN KEY (pre_stato_id) REFERENCES vtb_prenotazioniStati (pst_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;"

		case DB_SQL
			Install__BOOKING3 = _
				"CREATE TABLE dbo.vtb_strutture (" + _
				"str_id int IDENTITY (1, 1) NOT NULL CONSTRAINT PK_vtb_strutture PRIMARY KEY  CLUSTERED ," + _
				"str_paginaSito_id int NULL, " + _
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
				"str_paginaAnnulla_id INTEGER NULL," + _
				"str_dataCreazione SMALLDATETIME NULL" + _
				"); " + _
				"CREATE TABLE dbo.vrel_strutture_admin (" + _
				"rsa_id int IDENTITY (1, 1) NOT NULL CONSTRAINT PK_vrel_strutture_admin PRIMARY KEY  CLUSTERED ," + _
				"rsa_struttura_id int NULL, " + _
				"rsa_admin_id int NULL" + _
				"); " + _
				"CREATE TABLE dbo.vtb_tipiCameraBase (" + _
				"tipCB_id int IDENTITY (1, 1) NOT NULL CONSTRAINT PK_vtb_tipiCameraBase PRIMARY KEY  CLUSTERED ," + _
				"tipCB_nome_it NVARCHAR(50) NULL ," + _
				"tipCB_nome_en NVARCHAR(50) NULL ," + _
				"tipCB_nome_fr NVARCHAR(50) NULL ," + _
				"tipCB_nome_es NVARCHAR(50) NULL ," + _
				"tipCB_nome_de NVARCHAR(50) NULL, " + _
				"tipCB_descrizione_it NTEXT NULL ," + _
				"tipCB_descrizione_en NTEXT NULL ," + _
				"tipCB_descrizione_fr NTEXT NULL ," + _
				"tipCB_descrizione_es NTEXT NULL ," + _
				"tipCB_descrizione_de NTEXT NULL, " + _
				"tipCB_ordine int NULL, " + _
				"tipCB_numero int NULL, " + _
				"tipCB_immagine NVARCHAR(250) NULL" + _
				"); " + _
				"CREATE TABLE dbo.vtb_tipiCamera (" + _
				"tipC_id int IDENTITY (1, 1) NOT NULL CONSTRAINT PK_vtb_tipiCamera PRIMARY KEY  CLUSTERED ," + _
				"tipC_struttura_id INTEGER NULL, " + _
				"tipC_tipoCameraBase_id INTEGER NULL, " + _
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
				"tipC_ordine int NULL, " + _
				"tipC_numero int NULL, " + _
				"tipC_immagine NVARCHAR(250) NULL" + _
				"); " + _
				"CREATE TABLE dbo.vtb_disponibilita (" + _
				"dis_id int IDENTITY (1, 1) NOT NULL CONSTRAINT PK_vtb_disponibilita PRIMARY KEY  CLUSTERED ," + _
				"dis_prezzo SMALLMONEY NULL ," + _
				"dis_data SMALLDATETIME NULL ," + _
				"dis_disponibilita INTEGER NULL ," + _
				"dis_tipo_id INTEGER NULL, " + _
				"dis_min_stay INTEGER NULL, " + _
				"dis_promozione BIT NULL, " + _
				"dis_bloccata BIT NULL" + _
				"); " + _
				"CREATE TABLE dbo.vtb_listini (" + _
				"lis_id int IDENTITY (1, 1) NOT NULL CONSTRAINT PK_vtb_listini PRIMARY KEY  CLUSTERED ," + _
				"lis_struttura_id int NULL ," + _
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
				"lis_data SMALLDATETIME NULL" + _
				"); " + _
				"CREATE TABLE dbo.vtb_listini_tipiCamera (" + _
				"rlt_id int IDENTITY (1, 1) NOT NULL CONSTRAINT PK_vtb_listini_tipiCamera PRIMARY KEY  CLUSTERED ," + _
				"rlt_prezzo SMALLMONEY NULL ," + _
				"rlt_listino_id int NULL ," + _
				"rlt_tipo_id int NULL" + _
				"); " + _
				"CREATE TABLE dbo.vtb_prenotazioni (" + _
				"pre_id int IDENTITY (1, 1) NOT NULL CONSTRAINT PK_vtb_prenotazioni PRIMARY KEY  CLUSTERED ," + _
				"pre_stato_id int NULL, " + _
				"pre_data SMALLDATETIME NULL ," + _
				"pre_data_inizio SMALLDATETIME NULL ," + _
				"pre_data_fine SMALLDATETIME NULL ," + _
				"pre_note NTEXT NULL ," + _
				"pre_cliente_id INTEGER NULL, " + _
				"pre_nomeCC NVARCHAR(255) NULL, " + _
				"pre_numeroCC NVARCHAR(255) NULL, " + _
				"pre_dataCC SMALLDATETIME NULL, " + _
				"pre_tipoCC NVARCHAR(50) NULL, " + _
				"pre_totale INTEGER NULL, " + _
				"pre_meseCC int NULL, " + _
				"pre_annoCC int NULL, " + _
				"pre_cvcCC NVARCHAR(5) NULL, " + _
				"pre_chiave NVARCHAR(10) NULL" + _
				"); " + _
				"CREATE TABLE dbo.vtb_prenotazioni_tipiCamera (" + _
				"rpt_id int IDENTITY (1, 1) NOT NULL CONSTRAINT PK_vtb_prenotazioni_tipiCamera PRIMARY KEY  CLUSTERED ," + _
				"rpt_numero int NULL ," + _
				"rpt_prenotazione_id int NULL ," + _
				"rpt_tipo_id int NULL" + _
				"); " + _
				"CREATE TABLE dbo.vtb_prenotazioniStati (" + _
				"pst_id int IDENTITY (1, 1) NOT NULL CONSTRAINT PK_vtb_prenotazioniStati PRIMARY KEY  CLUSTERED ," + _
				"pst_nome_it NVARCHAR(50) NULL ," + _
				"pst_nome_en NVARCHAR(50) NULL ," + _
				"pst_nome_fr NVARCHAR(50) NULL ," + _
				"pst_nome_es NVARCHAR(50) NULL ," + _
				"pst_nome_de NVARCHAR(50) NULL " + _
				"); " + _
				"INSERT INTO vtb_prenotazioniStati(pst_nome_it, pst_nome_en, pst_nome_fr, pst_nome_es, pst_nome_de) "+ _
				"VALUES ('Richiesta', 'Requested', 'DemandÃ©', 'Solicitado', 'Verlangt');"+ _
				"INSERT INTO vtb_prenotazioniStati(pst_nome_it, pst_nome_en, pst_nome_fr, pst_nome_es, pst_nome_de) "+ _
				"VALUES ('Accettata', 'Accepted', 'Admis', 'Aceptado', 'Angenommen');"+ _
				"INSERT INTO vtb_prenotazioniStati(pst_nome_it, pst_nome_en, pst_nome_fr, pst_nome_es, pst_nome_de) "+ _
				"VALUES ('Annullata', 'Cancelled', 'DÃ©commandÃ©', 'Cancelado', 'Annulliert');"+ _
				"INSERT INTO vtb_prenotazioniStati(pst_nome_it, pst_nome_en, pst_nome_fr, pst_nome_es, pst_nome_de) "+ _
				"VALUES ('Confermata', 'Confirmed', 'ConfirmÃ©', 'Confirmado', 'BestÃ¤tigt');"+ _
				"ALTER TABLE vrel_strutture_admin ADD CONSTRAINT FK_hrel_strutture_admin__vtb_strutture " + _
				"FOREIGN KEY (rsa_struttura_id) REFERENCES vtb_strutture (str_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;" + _
				"ALTER TABLE vrel_strutture_admin ADD CONSTRAINT FK_hrel_strutture_admin__tb_admin " + _
				"FOREIGN KEY (rsa_admin_id) REFERENCES tb_admin (id_admin) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;" + _
				"ALTER TABLE vtb_tipiCamera ADD CONSTRAINT FK_vtb_tipiCamera__vtb_strutture " + _
				"FOREIGN KEY (tipC_struttura_id) REFERENCES vtb_strutture (str_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;" + _
				"ALTER TABLE vtb_tipiCamera ADD CONSTRAINT FK_vtb_tipiCamera__vtb_tipiCameraBase " + _
				"FOREIGN KEY (tipC_tipoCameraBase_id) REFERENCES vtb_tipiCameraBase (tipCB_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;" + _
				"ALTER TABLE vtb_listini ADD CONSTRAINT FK_vtb_listini__vtb_strutture " + _
				"FOREIGN KEY (lis_struttura_id) REFERENCES vtb_strutture (str_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;" + _
				"ALTER TABLE vtb_disponibilita ADD CONSTRAINT FK_vtb_disponibilita__vtb_tipiCamera " + _
				"FOREIGN KEY (dis_tipo_id) REFERENCES vtb_tipiCamera (tipC_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;" + _
				"ALTER TABLE vtb_listini_tipiCamera ADD CONSTRAINT FK_vtb_listini_tipiCamera__vtb_tipiCamera " + _
				"FOREIGN KEY (rlt_tipo_id) REFERENCES vtb_tipiCamera (tipC_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;" + _
				"ALTER TABLE vtb_listini_tipiCamera ADD CONSTRAINT FK_vtb_listini_tipiCamera__vtb_listini " + _
				"FOREIGN KEY (rlt_listino_id) REFERENCES vtb_listini (lis_id) " + _
				"ON UPDATE NO ACTION ON DELETE NO ACTION;" + _
				"ALTER TABLE vtb_prenotazioni_tipiCamera ADD CONSTRAINT FK_vtb_prenotazioni_tipiCamera__vtb_tipiCamera " + _
				"FOREIGN KEY (rpt_tipo_id) REFERENCES vtb_tipiCamera (tipC_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;" + _
				"ALTER TABLE vtb_prenotazioni_tipiCamera ADD CONSTRAINT FK_vtb_prenotazioni_tipiCamera__vtb_prenotazioni " + _
				"FOREIGN KEY (rpt_prenotazione_id) REFERENCES vtb_prenotazioni (pre_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;"+ _
				"ALTER TABLE vtb_prenotazioni ADD CONSTRAINT FK_vtb_prenotazioni_tb_indirizzario " + _
				"FOREIGN KEY (pre_cliente_id) REFERENCES tb_indirizzario (IDElencoIndirizzi) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;"+ _
				"ALTER TABLE vtb_prenotazioni ADD CONSTRAINT FK_vtb_prenotazioni_vtb_prenotazioniStati " + _
				"FOREIGN KEY (pre_stato_id) REFERENCES vtb_prenotazioniStati (pst_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;"

	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 1
'...........................................................................................
'aggiungo campi per descrivere e-mail
'...........................................................................................
function Aggiornamento__BOOKING3__1(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__BOOKING3__1 = "ALTER TABLE vtb_strutture ADD "+ _
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
			Aggiornamento__BOOKING3__1 = "ALTER TABLE vtb_strutture ADD "+ _
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
'AGGIORNAMENTO BOOKING3 2
'...........................................................................................
'aggiungo campi per descrivere e-mail
'...........................................................................................
function Aggiornamento__BOOKING3__2(conn)
	Select case DB_Type(conn)
		case DB_Access, DB_SQL
			Aggiornamento__BOOKING3__2 = "ALTER TABLE vtb_strutture ADD "+ _
											"str_paginaConferma_id INT NULL;"
	end select
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 3
'...........................................................................................
'aggiungo campi per descrivere e-mail
'...........................................................................................
function Aggiornamento__BOOKING3__3(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__BOOKING3__3 = "ALTER TABLE vtb_strutture ADD "+ _
											"str_paginaRichiesta_id INT NULL, "+ _
											"str_emailRichiestaObj_it TEXT(100) WITH COMPRESSION NULL, "+ _
											"str_emailRichiestaObj_en TEXT(100) WITH COMPRESSION NULL, "+ _
											"str_emailRichiestaObj_fr TEXT(100) WITH COMPRESSION NULL, "+ _
											"str_emailRichiestaObj_es TEXT(100) WITH COMPRESSION NULL, "+ _
											"str_emailRichiestaObj_de TEXT(100) WITH COMPRESSION NULL;"
		case DB_SQL
			Aggiornamento__BOOKING3__3 = "ALTER TABLE vtb_strutture ADD "+ _
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
'AGGIORNAMENTO BOOKING3 4
'...........................................................................................
'aggiungo campo rubrica all'hotel
'...........................................................................................
function Aggiornamento__BOOKING3__4(conn)
	Select case DB_Type(conn)
		case DB_Access, DB_SQL
			Aggiornamento__BOOKING3__4 = "DROP TABLE vrel_strutture_admin;"+ _
											"ALTER TABLE vtb_strutture ADD "+ _
											"str_gruppo_id INT NULL, "+ _
											"str_rubricaRichieste_id INT NULL, "+ _
											"str_rubricaPrenotazioni_id INT NULL;"
	end select
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 5
'...........................................................................................
'sposta l'impostazione dell'applicativo per hotel
'...........................................................................................
function Aggiornamento__BOOKING3__5(conn)
	Select case DB_Type(conn)
		case DB_Access, DB_SQL
			Aggiornamento__BOOKING3__5 = "ALTER TABLE vtb_strutture ADD "+ _
											"str_param_enableHttps BIT NULL, "+ _
											"str_param_disableDispo BIT NULL, "+ _
											"str_param_enableDifferita BIT NULL, "+ _
											"str_param_disableDecrementa BIT NULL;"
	end select
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 6
'...........................................................................................
'aggiunge un campo di risposta alla prenotazione
'...........................................................................................
function Aggiornamento__BOOKING3__6(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__BOOKING3__6 = "ALTER TABLE vtb_prenotazioni ADD "+ _
											"pre_risposta TEXT WITH COMPRESSION NULL;"
		case DB_SQL
			Aggiornamento__BOOKING3__6 = "ALTER TABLE vtb_prenotazioni ADD "+ _
											"pre_risposta NTEXT NULL;"
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 7
'...........................................................................................
'aggiunge un campo numero di persone associate alla prenotazione
'...........................................................................................
function Aggiornamento__BOOKING3__7(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__BOOKING3__7 = "ALTER TABLE vtb_prenotazioni ADD "+ _
											"pre_numero_persone INT NULL;"
		case DB_SQL
			Aggiornamento__BOOKING3__7 = "ALTER TABLE vtb_prenotazioni ADD "+ _
											"pre_numero_persone INT NULL;"
	end select
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 8
'...........................................................................................
'aggiunge un campo numero di persone associate alla prenotazione
'...........................................................................................
function Aggiornamento__BOOKING3__8(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__BOOKING3__8 = "ALTER TABLE vtb_tipicamera ADD "+ _
											"tipC_posti_letto INT NULL;"
		case DB_SQL
			Aggiornamento__BOOKING3__8 = "ALTER TABLE vtb_tipicamera ADD "+ _
											"tipC_posti_letto INT NULL;"
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 9
'.............................................................................................
'aggiunge i campi per la gestione delle date alternative e la gestione del cliente ospite  
'.............................................................................................
function Aggiornamento__BOOKING3__9(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__BOOKING3__9 = "ALTER TABLE vtb_prenotazioni ADD "+ _
											"pre_data_inizio_richiesta datetime NULL, pre_data_fine_richiesta datetime NULL, " + _
											"pre_nome_ospite TEXT(255) WITH COMPRESSION NULL, pre_cognome_ospite TEXT(255) WITH COMPRESSION NULL; "
		case DB_SQL
			Aggiornamento__BOOKING3__9 = "ALTER TABLE vtb_prenotazioni ADD "+ _
											"pre_data_inizio_richiesta smalldatetime NULL, pre_data_fine_richiesta datetime NULL, " + _
											"pre_nome_ospite NVARCHAR(255) NULL, pre_cognome_ospite NVARCHAR(255) NULL; "
	end select
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 10
'...........................................................................................
'aggiungo campi per descrivere e-mail
'...........................................................................................
function Aggiornamento__BOOKING3__10(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__BOOKING3__10 = "ALTER TABLE vtb_strutture ADD "+ _
											"str_paginaDataAlternativa_id INT NULL, "+ _
											"str_emailDataAlternativaObj_it TEXT(100) WITH COMPRESSION NULL, "+ _
											"str_emailDataAlternativaObj_en TEXT(100) WITH COMPRESSION NULL, "+ _
											"str_emailDataAlternativaObj_fr TEXT(100) WITH COMPRESSION NULL, "+ _
											"str_emailDataAlternativaObj_es TEXT(100) WITH COMPRESSION NULL, "+ _
											"str_emailDataAlternativaObj_de TEXT(100) WITH COMPRESSION NULL;"
		case DB_SQL
			Aggiornamento__BOOKING3__10 = "ALTER TABLE vtb_strutture ADD "+ _
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
'AGGIORNAMENTO BOOKING3 11
'...........................................................................................
'rimuove colonne di configurazione vecchio booking plus
'...........................................................................................
function Aggiornamento__BOOKING3__11(conn)
	Select case DB_Type(conn)
		case DB_Access, DB_SQL
			Aggiornamento__BOOKING3__11 = "ALTER TABLE vtb_strutture DROP COLUMN str_param_enableHttps; "+ _
											  "ALTER TABLE vtb_strutture DROP COLUMN str_param_disableDispo; "+ _
											  "ALTER TABLE vtb_strutture DROP COLUMN str_param_enableDifferita; "+ _
											  "ALTER TABLE vtb_strutture DROP COLUMN str_param_disableDecrementa; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 12
'...........................................................................................
'aggiunge relazioni tra struttura e relativo gruppo di lavoro e rubriche del next-com
'...........................................................................................
function Aggiornamento__BOOKING3__12(conn)
	Select case DB_Type(conn)
		case DB_Access, DB_SQL
			Aggiornamento__BOOKING3__12 = " ALTER TABLE vtb_strutture ADD CONSTRAINT FK_vtb_strutture__tb_gruppi " + _
											  " 	FOREIGN KEY (str_gruppo_id) REFERENCES tb_gruppi (id_gruppo) " + _
											  " 	ON UPDATE NO ACTION ON DELETE NO ACTION;" + _
											  " ALTER TABLE vtb_strutture ADD CONSTRAINT FK_vtb_strutture__tb_rubriche__RICHIESTE " + _
											  " 	FOREIGN KEY (str_rubricaRichieste_id) REFERENCES tb_rubriche (id_rubrica) " + _
											  " 	ON UPDATE NO ACTION ON DELETE NO ACTION;" + _
											  " ALTER TABLE vtb_strutture ADD CONSTRAINT FK_vtb_strutture__tb_rubriche__PRENOTAZIONI " + _
											  " 	FOREIGN KEY (str_rubricaPrenotazioni_id) REFERENCES tb_rubriche (id_rubrica) " + _
											  " 	ON UPDATE NO ACTION ON DELETE NO ACTION;"
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 13
'...........................................................................................
'rimuove relazione e tabella con tipi camera base
'...........................................................................................
function Aggiornamento__BOOKING3__13(conn)
	Select case DB_Type(conn)
		case DB_Access, DB_SQL
			Aggiornamento__BOOKING3__13 = " ALTER TABLE vtb_tipiCamera DROP CONSTRAINT FK_vtb_tipiCamera__vtb_tipiCameraBase; " + _
											  " ALTER TABLE vtb_tipiCamera DROP COLUMN tipC_tipoCameraBase_id; " + _
											  " DROP TABLE vtb_tipiCameraBase ;" 
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 14
'...........................................................................................
'aggiunge campi per configurazione comportamento struttura ricettiva
'...........................................................................................
function Aggiornamento__BOOKING3__14(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__BOOKING3__14 = " ALTER TABLE vtb_strutture ADD " + _
											  " 	str_param_listini BIT NOT NULL, " + _
											  "		str_param_gestione_giornaliera_prezzo BIT NOT NULL, " + _
											  "		str_param_gestione_giornaliera_dispo BIT NOT NULL, " + _
											  "		str_param_dispo_minima INT NULL, " +_
											  "		str_param_pre_richiesta BIT NOT NULL, " + _
											  "		str_param_pre_richiesta_prezzo_visibile BIT NOT NULL, " + _
											  "		str_param_pre_immediata BIT NOT NULL, " + _
											  "		str_param_pre_differita BIT NOT NULL ; "
		case DB_SQL
			Aggiornamento__BOOKING3__14 = " ALTER TABLE vtb_strutture ADD " + _
											  " 	str_param_listini BIT NULL, " + _
											  "		str_param_gestione_giornaliera_prezzo BIT NULL, " + _
											  "		str_param_gestione_giornaliera_dispo BIT NULL, " + _
											  "		str_param_dispo_minima INT NULL, " +_
											  "		str_param_pre_richiesta BIT NULL, " + _
											  "		str_param_pre_richiesta_prezzo_visibile BIT NULL, " + _
											  "		str_param_pre_immediata BIT NULL, " + _
											  "		str_param_pre_differita BIT NULL; " + _
											  " ALTER TABLE vtb_strutture ALTER COLUMN str_param_listini BIT NOT NULL; " + _
											  " ALTER TABLE vtb_strutture ALTER COLUMN str_param_gestione_giornaliera_prezzo BIT NOT NULL; " + _
											  " ALTER TABLE vtb_strutture ALTER COLUMN str_param_gestione_giornaliera_dispo BIT NOT NULL; " + _
											  " ALTER TABLE vtb_strutture ALTER COLUMN str_param_pre_richiesta BIT NOT NULL; " + _
											  " ALTER TABLE vtb_strutture ALTER COLUMN str_param_pre_richiesta_prezzo_visibile BIT NOT NULL; " + _
											  " ALTER TABLE vtb_strutture ALTER COLUMN str_param_pre_immediata BIT NOT NULL; " + _
											  " ALTER TABLE vtb_strutture ALTER COLUMN str_param_pre_differita BIT NOT NULL; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 15
'...........................................................................................
'rimuove campi non piu' necessari
'...........................................................................................
function Aggiornamento__BOOKING3__15(conn)
	Select case DB_Type(conn)
		case DB_Access, DB_SQL
			Aggiornamento__BOOKING3__15 = " ALTER TABLE vtb_strutture DROP COLUMN str_paginaAccetta_id; " + _
											  " ALTER TABLE vtb_strutture DROP COLUMN str_paginaAnnulla_id; " + _
											  " ALTER TABLE vtb_strutture DROP COLUMN str_emailAccettaObj_it; " + _
											  " ALTER TABLE vtb_strutture DROP COLUMN str_emailAccettaObj_en; " + _
											  " ALTER TABLE vtb_strutture DROP COLUMN str_emailAccettaObj_fr; " + _
											  " ALTER TABLE vtb_strutture DROP COLUMN str_emailAccettaObj_de; " + _
											  " ALTER TABLE vtb_strutture DROP COLUMN str_emailAccettaObj_es; " + _
											  " ALTER TABLE vtb_strutture DROP COLUMN str_emailAnnullaObj_it; " + _
											  " ALTER TABLE vtb_strutture DROP COLUMN str_emailAnnullaObj_en; " + _
											  " ALTER TABLE vtb_strutture DROP COLUMN str_emailAnnullaObj_fr; " + _
											  " ALTER TABLE vtb_strutture DROP COLUMN str_emailAnnullaObj_de; " + _
											  " ALTER TABLE vtb_strutture DROP COLUMN str_emailAnnullaObj_es; " + _
											  " ALTER TABLE vtb_strutture DROP COLUMN str_emailConfermaObj_it; " + _
											  " ALTER TABLE vtb_strutture DROP COLUMN str_emailConfermaObj_en; " + _
											  " ALTER TABLE vtb_strutture DROP COLUMN str_emailConfermaObj_fr; " + _
											  " ALTER TABLE vtb_strutture DROP COLUMN str_emailConfermaObj_de; " + _
											  " ALTER TABLE vtb_strutture DROP COLUMN str_emailConfermaObj_es; " + _
											  " ALTER TABLE vtb_strutture DROP COLUMN str_paginaConferma_id; " + _
											  " ALTER TABLE vtb_strutture DROP COLUMN str_paginaRichiesta_id; " + _
											  " ALTER TABLE vtb_strutture DROP COLUMN str_emailRichiestaObj_it; " + _
											  " ALTER TABLE vtb_strutture DROP COLUMN str_emailRichiestaObj_en; " + _
											  " ALTER TABLE vtb_strutture DROP COLUMN str_emailRichiestaObj_fr; " + _
											  " ALTER TABLE vtb_strutture DROP COLUMN str_emailRichiestaObj_de; " + _
											  " ALTER TABLE vtb_strutture DROP COLUMN str_emailRichiestaObj_es; " + _
											  " ALTER TABLE vtb_strutture DROP COLUMN str_paginaDataAlternativa_id; " + _
											  " ALTER TABLE vtb_strutture DROP COLUMN str_emailDataAlternativaObj_it; " + _
											  " ALTER TABLE vtb_strutture DROP COLUMN str_emailDataAlternativaObj_en; " + _
											  " ALTER TABLE vtb_strutture DROP COLUMN str_emailDataAlternativaObj_fr; " + _
											  " ALTER TABLE vtb_strutture DROP COLUMN str_emailDataAlternativaObj_de; " + _
											  " ALTER TABLE vtb_strutture DROP COLUMN str_emailDataAlternativaObj_es; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 16
'...........................................................................................
'aggiunge struttura dati per gestione testi della prenotazione/richiesta
'...........................................................................................
function Aggiornamento__BOOKING3__16(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__BOOKING3__16 = " CREATE TABLE vtb_strutture_impostazioni ( " + _
											  "	i_str_id INT NOT NULL, " + _
											  "	i_str_richiesta_form_pagina INT NULL, " + _
											  " i_str_richiesta_form_testo_it TEXT WITH COMPRESSION NULL, " + _
											  " i_str_richiesta_form_testo_en TEXT WITH COMPRESSION NULL, " + _
											  " i_str_richiesta_form_testo_fr TEXT WITH COMPRESSION NULL, " + _
											  " i_str_richiesta_form_testo_de TEXT WITH COMPRESSION NULL, " + _
											  " i_str_richiesta_form_testo_es TEXT WITH COMPRESSION NULL, " + _
											  "	i_str_richiesta_inviata_pagina INT NULL, " + _
											  " i_str_richiesta_inviata_testo_it TEXT WITH COMPRESSION NULL, " + _
											  " i_str_richiesta_inviata_testo_en TEXT WITH COMPRESSION NULL, " + _
											  " i_str_richiesta_inviata_testo_fr TEXT WITH COMPRESSION NULL, " + _
											  " i_str_richiesta_inviata_testo_de TEXT WITH COMPRESSION NULL, " + _
											  " i_str_richiesta_inviata_testo_es TEXT WITH COMPRESSION NULL, " + _
											  "	i_str_richiesta_email_pagina INT NULL, " + _
											  " i_str_richiesta_email_oggetto_it TEXT WITH COMPRESSION NULL, " + _
											  " i_str_richiesta_email_oggetto_en TEXT WITH COMPRESSION NULL, " + _
											  " i_str_richiesta_email_oggetto_fr TEXT WITH COMPRESSION NULL, " + _
											  " i_str_richiesta_email_oggetto_de TEXT WITH COMPRESSION NULL, " + _
											  " i_str_richiesta_email_oggetto_es TEXT WITH COMPRESSION NULL, " + _
											  " i_str_richiesta_email_testo_it TEXT WITH COMPRESSION NULL, " + _
											  " i_str_richiesta_email_testo_en TEXT WITH COMPRESSION NULL, " + _
											  " i_str_richiesta_email_testo_fr TEXT WITH COMPRESSION NULL, " + _
											  " i_str_richiesta_email_testo_de TEXT WITH COMPRESSION NULL, " + _
											  " i_str_richiesta_email_testo_es TEXT WITH COMPRESSION NULL, " + _
											  _
											  "	i_str_pre_immediata_form_pagina INT NULL, " + _
											  " i_str_pre_immediata_form_testo_it TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_immediata_form_testo_en TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_immediata_form_testo_fr TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_immediata_form_testo_de TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_immediata_form_testo_es TEXT WITH COMPRESSION NULL, " + _
											  "	i_str_pre_immediata_inviata_pagina INT NULL, " + _
											  " i_str_pre_immediata_inviata_testo_it TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_immediata_inviata_testo_en TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_immediata_inviata_testo_fr TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_immediata_inviata_testo_de TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_immediata_inviata_testo_es TEXT WITH COMPRESSION NULL, " + _
											  "	i_str_pre_immediata_email_pagina INT NULL, " + _
											  " i_str_pre_immediata_email_oggetto_it TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_immediata_email_oggetto_en TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_immediata_email_oggetto_fr TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_immediata_email_oggetto_de TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_immediata_email_oggetto_es TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_immediata_email_testo_it TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_immediata_email_testo_en TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_immediata_email_testo_fr TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_immediata_email_testo_de TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_immediata_email_testo_es TEXT WITH COMPRESSION NULL, " + _
											  _
											  "	i_str_pre_differita_form_pagina INT NULL, " + _
											  " i_str_pre_differita_form_testo_it TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_form_testo_en TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_form_testo_fr TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_form_testo_de TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_form_testo_es TEXT WITH COMPRESSION NULL, " + _
											  "	i_str_pre_differita_inviata_pagina INT NULL, " + _
											  " i_str_pre_differita_inviata_testo_it TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_inviata_testo_en TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_inviata_testo_fr TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_inviata_testo_de TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_inviata_testo_es TEXT WITH COMPRESSION NULL, " + _
											  "	i_str_pre_differita_email_pagina INT NULL, " + _
											  " i_str_pre_differita_email_oggetto_it TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_email_oggetto_en TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_email_oggetto_fr TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_email_oggetto_de TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_email_oggetto_es TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_email_testo_it TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_email_testo_en TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_email_testo_fr TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_email_testo_de TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_email_testo_es TEXT WITH COMPRESSION NULL, " + _
											  _
											  "	i_str_pre_differita_accettazione_email_pagina INT NULL, " + _
											  " i_str_pre_differita_accettazione_email_oggetto_it TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_accettazione_email_oggetto_en TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_accettazione_email_oggetto_fr TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_accettazione_email_oggetto_de TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_accettazione_email_oggetto_es TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_accettazione_email_testo_it TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_accettazione_email_testo_en TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_accettazione_email_testo_fr TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_accettazione_email_testo_de TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_accettazione_email_testo_es TEXT WITH COMPRESSION NULL, " + _
											  _
											  "	i_str_pre_differita_rifiuto_email_pagina INT NULL, " + _
											  " i_str_pre_differita_rifiuto_email_oggetto_it TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_rifiuto_email_oggetto_en TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_rifiuto_email_oggetto_fr TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_rifiuto_email_oggetto_de TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_rifiuto_email_oggetto_es TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_rifiuto_email_testo_it TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_rifiuto_email_testo_en TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_rifiuto_email_testo_fr TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_rifiuto_email_testo_de TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_rifiuto_email_testo_es TEXT WITH COMPRESSION NULL, " + _
											  _
											  "	i_str_pre_differita_alternativa_email_pagina INT NULL, " + _
											  " i_str_pre_differita_alternativa_email_oggetto_it TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_alternativa_email_oggetto_en TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_alternativa_email_oggetto_fr TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_alternativa_email_oggetto_de TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_alternativa_email_oggetto_es TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_alternativa_email_testo_it TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_alternativa_email_testo_en TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_alternativa_email_testo_fr TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_alternativa_email_testo_de TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_alternativa_email_testo_es TEXT WITH COMPRESSION NULL, " + _
											  _
											  "	i_str_pre_differita_login_pagina INT NULL, " + _
											  " i_str_pre_differita_login_testo_it TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_login_testo_en TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_login_testo_fr TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_login_testo_de TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_login_testo_es TEXT WITH COMPRESSION NULL, " + _
											  _
											  "	i_str_pre_differita_completata_form_pagina INT NULL, " + _
											  " i_str_pre_differita_completata_form_testo_it TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_completata_form_testo_en TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_completata_form_testo_fr TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_completata_form_testo_de TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_completata_form_testo_es TEXT WITH COMPRESSION NULL, " + _
											  "	i_str_pre_differita_completata_inviata_pagina INT NULL, " + _
											  " i_str_pre_differita_completata_inviata_testo_it TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_completata_inviata_testo_en TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_completata_inviata_testo_fr TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_completata_inviata_testo_de TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_completata_inviata_testo_es TEXT WITH COMPRESSION NULL, " + _
											  "	i_str_pre_differita_completata_email_pagina INT NULL, " + _
											  " i_str_pre_differita_completata_email_oggetto_it TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_completata_email_oggetto_en TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_completata_email_oggetto_fr TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_completata_email_oggetto_de TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_completata_email_oggetto_es TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_completata_email_testo_it TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_completata_email_testo_en TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_completata_email_testo_fr TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_completata_email_testo_de TEXT WITH COMPRESSION NULL, " + _
											  " i_str_pre_differita_completata_email_testo_es TEXT WITH COMPRESSION NULL " + _
											  " ) ; " + _
											  " ALTER TABLE vtb_strutture_impostazioni ADD CONSTRAINT FK_vtb_strutture_impostazioni__vtb_strutture " + _
											  " 	FOREIGN KEY (i_str_id) REFERENCES vtb_Strutture (str_id) " + _
											  " 	ON UPDATE CASCADE ON DELETE CASCADE;"
		case DB_SQL
			Aggiornamento__BOOKING3__16 = " CREATE TABLE vtb_strutture_impostazioni ( " + _
											  "	i_str_id INT NOT NULL, " + _
											  "	i_str_richiesta_form_pagina INT NULL, " + _
											  " i_str_richiesta_form_testo_it ntext NULL, " + _
											  " i_str_richiesta_form_testo_en ntext NULL, " + _
											  " i_str_richiesta_form_testo_fr ntext NULL, " + _
											  " i_str_richiesta_form_testo_de ntext NULL, " + _
											  " i_str_richiesta_form_testo_es ntext NULL, " + _
											  "	i_str_richiesta_inviata_pagina INT NULL, " + _
											  " i_str_richiesta_inviata_testo_it ntext NULL, " + _
											  " i_str_richiesta_inviata_testo_en ntext NULL, " + _
											  " i_str_richiesta_inviata_testo_fr ntext NULL, " + _
											  " i_str_richiesta_inviata_testo_de ntext NULL, " + _
											  " i_str_richiesta_inviata_testo_es ntext NULL, " + _
											  "	i_str_richiesta_email_pagina INT NULL, " + _
											  " i_str_richiesta_email_oggetto_it ntext NULL, " + _
											  " i_str_richiesta_email_oggetto_en ntext NULL, " + _
											  " i_str_richiesta_email_oggetto_fr ntext NULL, " + _
											  " i_str_richiesta_email_oggetto_de ntext NULL, " + _
											  " i_str_richiesta_email_oggetto_es ntext NULL, " + _
											  " i_str_richiesta_email_testo_it ntext NULL, " + _
											  " i_str_richiesta_email_testo_en ntext NULL, " + _
											  " i_str_richiesta_email_testo_fr ntext NULL, " + _
											  " i_str_richiesta_email_testo_de ntext NULL, " + _
											  " i_str_richiesta_email_testo_es ntext NULL, " + _
											  _
											  "	i_str_pre_immediata_form_pagina INT NULL, " + _
											  " i_str_pre_immediata_form_testo_it ntext NULL, " + _
											  " i_str_pre_immediata_form_testo_en ntext NULL, " + _
											  " i_str_pre_immediata_form_testo_fr ntext NULL, " + _
											  " i_str_pre_immediata_form_testo_de ntext NULL, " + _
											  " i_str_pre_immediata_form_testo_es ntext NULL, " + _
											  "	i_str_pre_immediata_inviata_pagina INT NULL, " + _
											  " i_str_pre_immediata_inviata_testo_it ntext NULL, " + _
											  " i_str_pre_immediata_inviata_testo_en ntext NULL, " + _
											  " i_str_pre_immediata_inviata_testo_fr ntext NULL, " + _
											  " i_str_pre_immediata_inviata_testo_de ntext NULL, " + _
											  " i_str_pre_immediata_inviata_testo_es ntext NULL, " + _
											  "	i_str_pre_immediata_email_pagina INT NULL, " + _
											  " i_str_pre_immediata_email_oggetto_it ntext NULL, " + _
											  " i_str_pre_immediata_email_oggetto_en ntext NULL, " + _
											  " i_str_pre_immediata_email_oggetto_fr ntext NULL, " + _
											  " i_str_pre_immediata_email_oggetto_de ntext NULL, " + _
											  " i_str_pre_immediata_email_oggetto_es ntext NULL, " + _
											  " i_str_pre_immediata_email_testo_it ntext NULL, " + _
											  " i_str_pre_immediata_email_testo_en ntext NULL, " + _
											  " i_str_pre_immediata_email_testo_fr ntext NULL, " + _
											  " i_str_pre_immediata_email_testo_de ntext NULL, " + _
											  " i_str_pre_immediata_email_testo_es ntext NULL, " + _
											  _
											  "	i_str_pre_differita_form_pagina INT NULL, " + _
											  " i_str_pre_differita_form_testo_it ntext NULL, " + _
											  " i_str_pre_differita_form_testo_en ntext NULL, " + _
											  " i_str_pre_differita_form_testo_fr ntext NULL, " + _
											  " i_str_pre_differita_form_testo_de ntext NULL, " + _
											  " i_str_pre_differita_form_testo_es ntext NULL, " + _
											  "	i_str_pre_differita_inviata_pagina INT NULL, " + _
											  " i_str_pre_differita_inviata_testo_it ntext NULL, " + _
											  " i_str_pre_differita_inviata_testo_en ntext NULL, " + _
											  " i_str_pre_differita_inviata_testo_fr ntext NULL, " + _
											  " i_str_pre_differita_inviata_testo_de ntext NULL, " + _
											  " i_str_pre_differita_inviata_testo_es ntext NULL, " + _
											  "	i_str_pre_differita_email_pagina INT NULL, " + _
											  " i_str_pre_differita_email_oggetto_it ntext NULL, " + _
											  " i_str_pre_differita_email_oggetto_en ntext NULL, " + _
											  " i_str_pre_differita_email_oggetto_fr ntext NULL, " + _
											  " i_str_pre_differita_email_oggetto_de ntext NULL, " + _
											  " i_str_pre_differita_email_oggetto_es ntext NULL, " + _
											  " i_str_pre_differita_email_testo_it ntext NULL, " + _
											  " i_str_pre_differita_email_testo_en ntext NULL, " + _
											  " i_str_pre_differita_email_testo_fr ntext NULL, " + _
											  " i_str_pre_differita_email_testo_de ntext NULL, " + _
											  " i_str_pre_differita_email_testo_es ntext NULL, " + _
											  _
											  "	i_str_pre_differita_accettazione_email_pagina INT NULL, " + _
											  " i_str_pre_differita_accettazione_email_oggetto_it ntext NULL, " + _
											  " i_str_pre_differita_accettazione_email_oggetto_en ntext NULL, " + _
											  " i_str_pre_differita_accettazione_email_oggetto_fr ntext NULL, " + _
											  " i_str_pre_differita_accettazione_email_oggetto_de ntext NULL, " + _
											  " i_str_pre_differita_accettazione_email_oggetto_es ntext NULL, " + _
											  " i_str_pre_differita_accettazione_email_testo_it ntext NULL, " + _
											  " i_str_pre_differita_accettazione_email_testo_en ntext NULL, " + _
											  " i_str_pre_differita_accettazione_email_testo_fr ntext NULL, " + _
											  " i_str_pre_differita_accettazione_email_testo_de ntext NULL, " + _
											  " i_str_pre_differita_accettazione_email_testo_es ntext NULL, " + _
											  _
											  "	i_str_pre_differita_rifiuto_email_pagina INT NULL, " + _
											  " i_str_pre_differita_rifiuto_email_oggetto_it ntext NULL, " + _
											  " i_str_pre_differita_rifiuto_email_oggetto_en ntext NULL, " + _
											  " i_str_pre_differita_rifiuto_email_oggetto_fr ntext NULL, " + _
											  " i_str_pre_differita_rifiuto_email_oggetto_de ntext NULL, " + _
											  " i_str_pre_differita_rifiuto_email_oggetto_es ntext NULL, " + _
											  " i_str_pre_differita_rifiuto_email_testo_it ntext NULL, " + _
											  " i_str_pre_differita_rifiuto_email_testo_en ntext NULL, " + _
											  " i_str_pre_differita_rifiuto_email_testo_fr ntext NULL, " + _
											  " i_str_pre_differita_rifiuto_email_testo_de ntext NULL, " + _
											  " i_str_pre_differita_rifiuto_email_testo_es ntext NULL, " + _
											  _
											  "	i_str_pre_differita_alternativa_email_pagina INT NULL, " + _
											  " i_str_pre_differita_alternativa_email_oggetto_it ntext NULL, " + _
											  " i_str_pre_differita_alternativa_email_oggetto_en ntext NULL, " + _
											  " i_str_pre_differita_alternativa_email_oggetto_fr ntext NULL, " + _
											  " i_str_pre_differita_alternativa_email_oggetto_de ntext NULL, " + _
											  " i_str_pre_differita_alternativa_email_oggetto_es ntext NULL, " + _
											  " i_str_pre_differita_alternativa_email_testo_it ntext NULL, " + _
											  " i_str_pre_differita_alternativa_email_testo_en ntext NULL, " + _
											  " i_str_pre_differita_alternativa_email_testo_fr ntext NULL, " + _
											  " i_str_pre_differita_alternativa_email_testo_de ntext NULL, " + _
											  " i_str_pre_differita_alternativa_email_testo_es ntext NULL, " + _
											  _
											  "	i_str_pre_differita_login_pagina INT NULL, " + _
											  " i_str_pre_differita_login_testo_it ntext NULL, " + _
											  " i_str_pre_differita_login_testo_en ntext NULL, " + _
											  " i_str_pre_differita_login_testo_fr ntext NULL, " + _
											  " i_str_pre_differita_login_testo_de ntext NULL, " + _
											  " i_str_pre_differita_login_testo_es ntext NULL, " + _
											  _
											  "	i_str_pre_differita_completata_form_pagina INT NULL, " + _
											  " i_str_pre_differita_completata_form_testo_it ntext NULL, " + _
											  " i_str_pre_differita_completata_form_testo_en ntext NULL, " + _
											  " i_str_pre_differita_completata_form_testo_fr ntext NULL, " + _
											  " i_str_pre_differita_completata_form_testo_de ntext NULL, " + _
											  " i_str_pre_differita_completata_form_testo_es ntext NULL, " + _
											  "	i_str_pre_differita_completata_inviata_pagina INT NULL, " + _
											  " i_str_pre_differita_completata_inviata_testo_it ntext NULL, " + _
											  " i_str_pre_differita_completata_inviata_testo_en ntext NULL, " + _
											  " i_str_pre_differita_completata_inviata_testo_fr ntext NULL, " + _
											  " i_str_pre_differita_completata_inviata_testo_de ntext NULL, " + _
											  " i_str_pre_differita_completata_inviata_testo_es ntext NULL, " + _
											  "	i_str_pre_differita_completata_email_pagina INT NULL, " + _
											  " i_str_pre_differita_completata_email_oggetto_it ntext NULL, " + _
											  " i_str_pre_differita_completata_email_oggetto_en ntext NULL, " + _
											  " i_str_pre_differita_completata_email_oggetto_fr ntext NULL, " + _
											  " i_str_pre_differita_completata_email_oggetto_de ntext NULL, " + _
											  " i_str_pre_differita_completata_email_oggetto_es ntext NULL, " + _
											  " i_str_pre_differita_completata_email_testo_it ntext NULL, " + _
											  " i_str_pre_differita_completata_email_testo_en ntext NULL, " + _
											  " i_str_pre_differita_completata_email_testo_fr ntext NULL, " + _
											  " i_str_pre_differita_completata_email_testo_de ntext NULL, " + _
											  " i_str_pre_differita_completata_email_testo_es ntext NULL " + _
											  " ) ; " + _
											  " ALTER TABLE vtb_strutture_impostazioni ADD CONSTRAINT FK_vtb_strutture_impostazioni__vtb_strutture " + _
											  " 	FOREIGN KEY (i_str_id) REFERENCES vtb_Strutture (str_id) " + _
											  " 	ON UPDATE CASCADE ON DELETE CASCADE;"
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 17
'...........................................................................................
'aggiunge campo di gestione della prenotazione che indica il tipo di inserimento.
'...........................................................................................
function Aggiornamento__BOOKING3__17(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__BOOKING3__17 = "ALTER TABLE vtb_prenotazioni ADD "+ _
											  " pre_tipo TEXT(100) NOT NULL, "+ _
											  " pre_insData DATETIME NULL ," + _
											  " pre_insAdmin_id INT NULL ," + _
											  " pre_modData DATETIME NULL ," + _
											  " pre_modAdmin_id INT NULL " + _
											  " ;"
		case DB_SQL
			Aggiornamento__BOOKING3__17 = "ALTER TABLE vtb_prenotazioni ADD "+ _
											  " pre_tipo nvarchar(100) NULL, "+ _
											  " pre_insData DATETIME NULL ," + _
											  " pre_insAdmin_id INT NULL ," + _
											  " pre_modData DATETIME NULL ," + _
											  " pre_modAdmin_id INT NULL " + _
											  " ;" + _
											  " ALTER TABLE vtb_prenotazioni ALTER COLUMN pre_tipo nvarchar(100) NOT NULL; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 18
'...........................................................................................
'aggiunge campo di gestione della prenotazione che indica il tipo di inserimento.
'...........................................................................................
function Aggiornamento__BOOKING3__18(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__BOOKING3__18 = "ALTER TABLE vtb_prenotazioni ADD "+ _
											  " pre_str_id INT NOT NULL ;" + _
											  " ALTER TABLE vtb_prenotazioni ADD CONSTRAINT FK_vtb_prenotazioni__vtb_strutture " + _
											  " 	FOREIGN KEY (pre_str_id) REFERENCES vtb_strutture (str_id) " + _
											  " 	ON UPDATE NO ACTION ON DELETE NO ACTION ; "
		case DB_SQL
			Aggiornamento__BOOKING3__18 = "ALTER TABLE vtb_prenotazioni ADD "+ _
											  " pre_str_id INT NULL ;" + _
											  " ALTER TABLE vtb_prenotazioni ALTER COLUMN pre_str_id INT NOT NULL; " + _
											  " ALTER TABLE vtb_prenotazioni ADD CONSTRAINT FK_vtb_prenotazioni__vtb_strutture " + _
											  " 	FOREIGN KEY (pre_str_id) REFERENCES vtb_strutture (str_id) " + _
											  " 	ON UPDATE NO ACTION ON DELETE NO ACTION ; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 19
'...........................................................................................
'aggiunge campo per indicazione prenotazione alternativa
'...........................................................................................
function Aggiornamento__BOOKING3__19(conn)
	Aggiornamento__BOOKING3__19 = "ALTER TABLE vtb_prenotazioni ADD " + _
									  "		pre_alternativa_precedente_id INT NULL, " + _
									  "		pre_alternativa_successiva_id INT NULL ; "
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 20
'...........................................................................................
'aggiunge campo per gestire il momento in cui viene scalata la disponibilità
' 
'...........................................................................................
function Aggiornamento__BOOKING3__20(conn)
	Aggiornamento__BOOKING3__20 = "ALTER TABLE vtb_strutture ADD " + _
									   "		str_param_pre_scala_dispo_accett BIT NULL;  "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 21
'...........................................................................................
'aggiunge campo per gestire il momento in cui viene scalata la disponibilità
' 
'...........................................................................................
function Aggiornamento__BOOKING3__21(conn)
Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__BOOKING3__21 = "ALTER TABLE vtb_prenotazioni ALTER COLUMN "+ _
											  " pre_totale CURRENCY NULL;"
		case DB_SQL
			Aggiornamento__BOOKING3__21 = "ALTER TABLE vtb_prenotazioni ALTER COLUMN  "+ _
											  " pre_totale SMALLMONEY NULL;"
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 22
'...........................................................................................
'aggiunge campo per gestire il momento in cui viene scalata la disponibilità
' 
'...........................................................................................
function Aggiornamento__BOOKING3__22(conn)
	Aggiornamento__BOOKING3__22 = "ALTER TABLE vtb_strutture ADD " + _
									   "		str_info_external_key INT NULL;  "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 23
'...........................................................................................
'aggiunge tabella per codifica tipologie base di camere
' 
'...........................................................................................
function Aggiornamento__BOOKING3__23(conn)
    Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__BOOKING3__23 = _
                "CREATE TABLE vtb_tipiCameraBase (" + _
				    " tipCB_id COUNTER CONSTRAINT PK_vtb_tipiCameraBase PRIMARY KEY ," + _
				    " tipCB_nome_it TEXT(50) WITH COMPRESSION NULL ," + _
                    " tipCB_nome_en TEXT(50) WITH COMPRESSION NULL ," + _
                    " tipCB_nome_fr TEXT(50) WITH COMPRESSION NULL ," + _
                    " tipCB_nome_es TEXT(50) WITH COMPRESSION NULL ," + _
                    " tipCB_nome_de TEXT(50) WITH COMPRESSION NULL, " + _
                    " tipCB_ordine INTEGER NULL, " + _
                    " tipCB_posti_letto INTEGER NULL, " + _
                    " tipCB_immagine TEXT(250) WITH COMPRESSION NULL" + _
			    " ); "
		case DB_SQL
			Aggiornamento__BOOKING3__23 = _
                "CREATE TABLE dbo.vtb_tipiCameraBase (" + _
				    " tipCB_id int IDENTITY (1, 1) NOT NULL CONSTRAINT PK_vtb_tipiCameraBase PRIMARY KEY  CLUSTERED ," + _
				    " tipCB_nome_it nvarchar(250) NULL ," + _
                    " tipCB_nome_en nvarchar(250) NULL ," + _
                    " tipCB_nome_fr nvarchar(250) NULL ," + _
                    " tipCB_nome_es nvarchar(250) NULL ," + _
                    " tipCB_nome_de nvarchar(250) NULL ," + _
                    " tipCB_ordine int NULL, " + _
                    " tipCB_posti_letto int NULL, " + _
                    " tipCB_immagine nvarchar(250) NULL " + _
			    " ); "
    end select
    Aggiornamento__BOOKING3__23 = Aggiornamento__BOOKING3__23 + _
        " ALTER TABLE vtb_tipiCamera ADD tipC_tipoCameraBase_id INT NULL; " + _
        " INSERT INTO vtb_tipiCameraBase (tipCB_nome_it, tipCB_ordine, tipCB_posti_letto) VALUES ('Camera', 0, 0); " + _
        " UPDATE vtb_tipiCamera SET tipC_tipoCameraBase_id = 1; " + _
        " ALTER TABLE vtb_tipiCamera ALTER COLUMN tipC_tipoCameraBase_id INT NOT NULL; " + _
        " ALTER TABLE vtb_tipiCamera ADD CONSTRAINT FK_vtb_tipiCamera__vtb_tipiCameraBase " + _
	    "   FOREIGN KEY (tipC_tipoCameraBase_id) REFERENCES vtb_tipiCameraBase (tipCB_id) " + _
        "   ON UPDATE CASCADE ON DELETE CASCADE;"
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 24
'...........................................................................................
' aggiunge campo per gestire la classificazione degli hotel 
' valori validi da 0 a 6: 0=NON CLASSIFICATO 6=QUINTA CATEGORIA LUSSO
'...........................................................................................
function Aggiornamento__BOOKING3__24(conn)
	Aggiornamento__BOOKING3__24 = "ALTER TABLE vtb_strutture ADD " + _
									   "		str_stars INT NULL;  "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 25
'...........................................................................................
' aggiunge tabella fifo per la gestione del log dei comandi di sicronizzazione
'...........................................................................................
function Aggiornamento__BOOKING3__25(conn)
Select case DB_Type(conn)
		case DB_Access
		Aggiornamento__BOOKING3__25 = "CREATE TABLE vtb_command_queue_log (" + _
										" cql_id COUNTER CONSTRAINT PK_vtb_command_queue_log PRIMARY KEY ," + _
										" cql_codice_op TEXT(250) WITH COMPRESSION NULL ," + _
										" cql_timestamp DATETIME NULL ," +_
										" cql_struttura_id int NULL" + _
									   ");"
		case DB_SQL 
		Aggiornamento__BOOKING3__25 = "CREATE TABLE vtb_command_queue_log (" + _
										" cql_id int IDENTITY (1, 1) NOT NULL CONSTRAINT PK_vtb_command_queue_log PRIMARY KEY  CLUSTERED ," + _
										" cql_codice_op nvarchar(250) NULL ," + _
										" cql_timestamp DATETIME NULL ," + _
										" cql_struttura_id int NULL" + _
										");"
		end select
		Aggiornamento__BOOKING3__25 = Aggiornamento__BOOKING3__25 + _
				"ALTER TABLE vtb_command_queue_log ADD CONSTRAINT FK_vtb_command_queue_log__vtb_strutture " + _
				"FOREIGN KEY (cql_struttura_id) REFERENCES vtb_strutture (str_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 26
'...........................................................................................
' Aggiunge un campo per la gestione della scadenza delle prenotazioni singole
' 
'...........................................................................................
function Aggiornamento__BOOKING3__26(conn)
    Aggiornamento__BOOKING3__26 = "ALTER TABLE vtb_prenotazioni ADD "+ _
                                  " pre_scadenza DATETIME NULL;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 27
'...........................................................................................
' Aggiunge un campo per la gestione dell'orario di arrivo nella prenotazione e relativo flag 
'   di abilitazione nel form di conferma prenotazione
'...........................................................................................
function Aggiornamento__BOOKING3__27(conn)
    Aggiornamento__BOOKING3__27 = _
        " ALTER TABLE vtb_prenotazioni ADD " + _
        "   pre_orario_arrivo " + SQL_CharField(Conn, 255) + " NULL;" + _
        " ALTER TABLE vtb_strutture ADD " + _
        "   str_param_pre_orario_arrivo BIT NULL; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 28
'...........................................................................................
' rimuove struttura dati per sincronizzazione: non pi&ugrave; necessaria.
'...........................................................................................
function Aggiornamento__BOOKING3__28(conn)
    Aggiornamento__BOOKING3__28 = _
        "ALTER TABLE vtb_command_queue_log DROP CONSTRAINT FK_vtb_command_queue_log__vtb_strutture; " + _
        DropObject(conn, "vtb_command_queue_log", "TABLE")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 29
'...........................................................................................
' aggiunge nuovo stato della prenotazione
'...........................................................................................
function Aggiornamento__BOOKING3__29(conn)
    Aggiornamento__BOOKING3__29 = _
        " INSERT INTO vtb_prenotazionistati (pst_nome_it, pst_nome_en, pst_nome_fr, pst_nome_es, pst_nome_de) " + _
        "       VALUES ('Definitiva', 'Definitive', '', '', '') ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 30
'...........................................................................................
' aggiunge colonne alle impostazioni delle strutture.
'...........................................................................................
function Aggiornamento__BOOKING3__30(conn)
	Select case DB_Type(conn)
		case DB_Access
		Aggiornamento__BOOKING3__30 = _
		" ALTER TABLE vtb_strutture_impostazioni ADD " + _
		"   i_str_pre_differita_definitiva_email_pagina INT NULL, " + _
        "   i_str_pre_differita_definitiva_email_oggetto_it TEXT WITH COMPRESSION NULL, " + _
        "   i_str_pre_differita_definitiva_email_oggetto_en TEXT WITH COMPRESSION NULL, " + _
        "   i_str_pre_differita_definitiva_email_oggetto_fr TEXT WITH COMPRESSION NULL, " + _
        "   i_str_pre_differita_definitiva_email_oggetto_de TEXT WITH COMPRESSION NULL, " + _
        "   i_str_pre_differita_definitiva_email_oggetto_es TEXT WITH COMPRESSION NULL, " + _
        "   i_str_pre_differita_definitiva_email_testo_it TEXT WITH COMPRESSION NULL, " + _
        "   i_str_pre_differita_definitiva_email_testo_en TEXT WITH COMPRESSION NULL, " + _
        "   i_str_pre_differita_definitiva_email_testo_fr TEXT WITH COMPRESSION NULL, " + _
        "   i_str_pre_differita_definitiva_email_testo_de TEXT WITH COMPRESSION NULL, " + _
        "   i_str_pre_differita_definitiva_email_testo_es TEXT WITH COMPRESSION NULL " + _
		 " ; " + _
        " ALTER TABLE vtb_strutture ADD " + _
        "   str_param_pre_differita_definitiva_differita BIT NULL, " + _
        "   str_param_pre_immediata_definitiva_differita BIT NULL; "
		case DB_SQL 
		Aggiornamento__BOOKING3__30 = _
        " ALTER TABLE vtb_strutture_impostazioni ADD " + _
        "   i_str_pre_immediata_definitiva_email_pagina INT NULL, " + _
        "   i_str_pre_immediata_definitiva_email_oggetto_it ntext NULL, " + _
        "   i_str_pre_immediata_definitiva_email_oggetto_en ntext NULL, " + _
        "   i_str_pre_immediata_definitiva_email_oggetto_fr ntext NULL, " + _
        "   i_str_pre_immediata_definitiva_email_oggetto_de ntext NULL, " + _
        "   i_str_pre_immediata_definitiva_email_oggetto_es ntext NULL, " + _
        "   i_str_pre_immediata_definitiva_email_testo_it ntext NULL, " + _
        "   i_str_pre_immediata_definitiva_email_testo_en ntext NULL, " + _
        "   i_str_pre_immediata_definitiva_email_testo_fr ntext NULL, " + _
        "   i_str_pre_immediata_definitiva_email_testo_de ntext NULL, " + _
        "   i_str_pre_immediata_definitiva_email_testo_es ntext NULL " + _
		" ; " + _
        " ALTER TABLE vtb_strutture ADD " + _
        "   str_param_pre_differita_definitiva_differita BIT NULL, " + _
        "   str_param_pre_immediata_definitiva_differita BIT NULL; "
		end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 31
'...........................................................................................
' Aggiunge un campo per l'abilitazione della visualizzazione della data di scadenza
' per la conferma della prenotazione al momento dell'accettazione.
'...........................................................................................
function Aggiornamento__BOOKING3__31(conn)
    Aggiornamento__BOOKING3__31 = _
        " ALTER TABLE vtb_strutture ADD " + _
        "   str_param_pre_scadenza_prenotazione BIT NULL; " + _
        " UPDATE vtb_strutture SET str_param_pre_scadenza_prenotazione = 1 WHERE " + SQL_IsTrue(conn, "str_param_pre_differita")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 32
'...........................................................................................
' Aggiunge un campo per il calcolo della data di scadenza
'...........................................................................................
function Aggiornamento__BOOKING3__32(conn)
    Aggiornamento__BOOKING3__32 = _
        " ALTER TABLE vtb_strutture ADD " + _
        "   str_param_pre_scadenza_giorni INT NULL; " + _
        " UPDATE vtb_strutture SET str_param_pre_scadenza_giorni = 2 "
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 32
'...........................................................................................
' Aggiunge un campo per il calcolo della data di scadenza
'...........................................................................................
function Aggiornamento__BOOKING3__33(conn)
    Aggiornamento__BOOKING3__33 = _
		" ALTER TABLE " & SQL_Dbo(Conn) & "vtb_strutture ADD str_NextCom_id INT NULL;"& vbCrLf & _
        SQL_AddForeignKey(conn, "vtb_strutture", "str_NextCom_id", "tb_indirizzario", "idElencoIndirizzi", false, "") & vbCrLf
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 33
'...........................................................................................
' aggiunge colonne alle impostazioni delle strutture.
'...........................................................................................
function Aggiornamento_speciale__BOOKING3__33(conn)
	Select case DB_Type(conn)
		case DB_SQL 
		Aggiornamento_speciale__BOOKING3__33 = _
        " ALTER TABLE vtb_strutture_impostazioni ADD " + _
		"   i_str_pre_differita_definitiva_email_pagina INT NULL, " + _
        "   i_str_pre_differita_definitiva_email_oggetto_it ntext NULL, " + _
        "   i_str_pre_differita_definitiva_email_oggetto_en ntext NULL, " + _
        "   i_str_pre_differita_definitiva_email_oggetto_fr ntext NULL, " + _
        "   i_str_pre_differita_definitiva_email_oggetto_de ntext NULL, " + _
        "   i_str_pre_differita_definitiva_email_oggetto_es ntext NULL, " + _
        "   i_str_pre_differita_definitiva_email_testo_it ntext NULL, " + _
        "   i_str_pre_differita_definitiva_email_testo_en ntext NULL, " + _
        "   i_str_pre_differita_definitiva_email_testo_fr ntext NULL, " + _
        "   i_str_pre_differita_definitiva_email_testo_de ntext NULL, " + _
        "   i_str_pre_differita_definitiva_email_testo_es ntext NULL " + _
		" ; "
		end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 34
'...........................................................................................
' Aggiunge un campo per la gestione dell'orario di arrivo nella prenotazione e relativo flag 
'   di abilitazione nel form di conferma prenotazione
'...........................................................................................
function Aggiornamento__BOOKING3__34(conn)
    Aggiornamento__BOOKING3__34 = _
        " ALTER TABLE vtb_prenotazioni ADD " + _
        "   pre_numero_camere INT NULL;" 
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 35
'...........................................................................................
' aggiunge colonne alle impostazioni delle strutture.
'...........................................................................................
function Aggiornamento__BOOKING3__35(conn)
	Select case DB_Type(conn)
		case DB_Access
		Aggiornamento__BOOKING3__35 = _
		" ALTER TABLE vtb_strutture_impostazioni ADD " + _
		"   i_str_pre_term_testo_it TEXT WITH COMPRESSION NULL, " + _
        "   i_str_pre_term_testo_en TEXT WITH COMPRESSION NULL, " + _
        "   i_str_pre_term_testo_fr TEXT WITH COMPRESSION NULL, " + _
        "   i_str_pre_term_testo_de TEXT WITH COMPRESSION NULL, " + _
        "   i_str_pre_term_testo_es TEXT WITH COMPRESSION NULL, " + _
        "   i_str_pre_maggiori_dettagli_testo_it TEXT WITH COMPRESSION NULL, " + _
        "   i_str_pre_maggiori_dettagli_testo_en TEXT WITH COMPRESSION NULL, " + _
        "   i_str_pre_maggiori_dettagli_testo_fr TEXT WITH COMPRESSION NULL, " + _
        "   i_str_pre_maggiori_dettagli_testo_de TEXT WITH COMPRESSION NULL, " + _
        "   i_str_pre_maggiori_dettagli_testo_es TEXT WITH COMPRESSION NULL " + _
		 " ; " 
		case DB_SQL 
		Aggiornamento__BOOKING3__35 = _
        " ALTER TABLE vtb_strutture_impostazioni ADD " + _
        "   i_str_pre_term_testo_it ntext NULL, " + _
        "   i_str_pre_term_testo_en ntext NULL, " + _
        "   i_str_pre_term_testo_fr ntext NULL, " + _
        "   i_str_pre_term_testo_de ntext NULL, " + _
        "   i_str_pre_term_testo_es ntext NULL, " + _
        "   i_str_pre_maggiori_dettagli_testo_it ntext NULL, " + _
        "   i_str_pre_maggiori_dettagli_testo_en ntext NULL, " + _
        "   i_str_pre_maggiori_dettagli_testo_fr ntext NULL, " + _
        "   i_str_pre_maggiori_dettagli_testo_de ntext NULL, " + _
        "   i_str_pre_maggiori_dettagli_testo_es ntext NULL " + _
		" ; "
		end select
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 36
'...........................................................................................
' Aggiunge un campo per la gestione dell'orario di arrivo nella prenotazione e relativo flag 
'   di abilitazione nel form di conferma prenotazione
'...........................................................................................
function Aggiornamento__BOOKING3__36(conn)
    Aggiornamento__BOOKING3__36 = _
        " ALTER TABLE vtb_strutture ADD " + _
        "   str_param_pre_no_gest_camere BIT NULL;" 
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 37
'...........................................................................................
'	Giacomo, 17/05/2011
'...........................................................................................
'  	aggiornamento per criptare i numeri delle carte di credito
'...........................................................................................
function Aggiornamento__BOOKING3__37(conn)
	Aggiornamento__BOOKING3__37 = _
		" ALTER TABLE vtb_prenotazioni " & _
		" ALTER COLUMN pre_cvcCC " & SQL_CharField(Conn, 255) & " NULL;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 38
'...........................................................................................
'	Giacomo, 16/05/2011
'...........................................................................................
'  	aggiornamento per criptare i numeri delle carte di credito
'...........................................................................................
function Aggiornamento__BOOKING3__38(conn)
	Aggiornamento__BOOKING3__38 = "SELECT * FROM AA_Versione "
end function

sub AggiornamentoSpeciale__BOOKING3__38(conn)
	dim numero_carta, cvc_carta
	sql = "SELECT * FROM vtb_prenotazioni WHERE pre_numeroCC <> '' OR pre_cvcCC <> '' "
	if rs.State = adStateOpen then
		rs.close
	end if
	rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
	while not rs.eof
		numero_carta = EncryptCreditCard(rs("pre_data"), rs("pre_str_id"), rs("pre_numeroCC"))
		rs("pre_numeroCC") = numero_carta
		
		cvc_carta = EncryptCreditCard(rs("pre_data"), rs("pre_str_id"), rs("pre_cvcCC"))
		rs("pre_cvcCC") = cvc_carta
		
		rs.update
		rs.moveNext
	wend
	rs.close
end sub
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 39
'...........................................................................................
'	Giacomo, 21/09/2011
'...........................................................................................
'  	aggiornamento per aggiungere i campi lingua
'...........................................................................................
function Aggiornamento__BOOKING3__39(conn, lingua_abbr)
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__BOOKING3__39 = _
			  " ALTER TABLE vtb_strutture ADD " + vbCrLf + _
			  " 	str_descrizione_" + lingua_abbr + " NTEXT NULL; " + vbCrLf + _
			  " " + vbCrLf + _
			  " ALTER TABLE vtb_tipiCameraBase ADD " + vbCrLf + _
			  " 	tipCB_nome_" + lingua_abbr + " nvarchar(250) NULL; " + vbCrLf + _
			  " " + vbCrLf + _
			  " ALTER TABLE vtb_tipiCamera ADD " + vbCrLf + _
			  " 	tipC_nome_" + lingua_abbr + " NVARCHAR(50) NULL, " + vbCrLf + _
			  " 	tipC_descrizione_" + lingua_abbr + " NTEXT NULL; " + vbCrLf + _
			  " " + vbCrLf + _
			  " ALTER TABLE vtb_listini ADD " + vbCrLf + _
			  " 	lis_nome_" + lingua_abbr + " NVARCHAR(100) NULL, " + vbCrLf + _
			  " 	lis_condizioni_" + lingua_abbr + " NTEXT NULL; " + vbCrLf + _
			  " " + vbCrLf + _
			  " ALTER TABLE vtb_prenotazioniStati ADD " + vbCrLf + _
			  " 	pst_nome_" + lingua_abbr + " NVARCHAR(50) NULL; " + vbCrLf + _
			  " " + vbCrLf + _
			  " ALTER TABLE vtb_strutture_impostazioni ADD " + vbCrLf + _
			  " 	i_str_richiesta_form_testo_" + lingua_abbr + " ntext NULL, " + vbCrLf + _
			  " 	i_str_richiesta_inviata_testo_" + lingua_abbr + " ntext NULL, " + vbCrLf + _
			  " 	i_str_richiesta_email_oggetto_" + lingua_abbr + " ntext NULL, " + vbCrLf + _
			  " 	i_str_richiesta_email_testo_" + lingua_abbr + " ntext NULL, " + vbCrLf + _
			  " 	i_str_pre_immediata_form_testo_" + lingua_abbr + " ntext NULL, " + vbCrLf + _
			  " 	i_str_pre_immediata_inviata_testo_" + lingua_abbr + " ntext NULL, " + vbCrLf + _
			  " 	i_str_pre_immediata_email_oggetto_" + lingua_abbr + " ntext NULL, " + vbCrLf + _
			  " 	i_str_pre_immediata_email_testo_" + lingua_abbr + " ntext NULL, " + vbCrLf + _
			  " 	i_str_pre_differita_form_testo_" + lingua_abbr + " ntext NULL, " + vbCrLf + _
			  " 	i_str_pre_differita_inviata_testo_" + lingua_abbr + " ntext NULL, " + vbCrLf + _
			  " 	i_str_pre_differita_email_oggetto_" + lingua_abbr + " ntext NULL, " + vbCrLf + _
			  " 	i_str_pre_differita_email_testo_" + lingua_abbr + " ntext NULL, " + vbCrLf + _
			  " 	i_str_pre_differita_accettazione_email_oggetto_" + lingua_abbr + " ntext NULL, " + vbCrLf + _
			  " 	i_str_pre_differita_accettazione_email_testo_" + lingua_abbr + " ntext NULL, " + vbCrLf + _
			  " 	i_str_pre_differita_rifiuto_email_oggetto_" + lingua_abbr + " ntext NULL, " + vbCrLf + _
			  " 	i_str_pre_differita_rifiuto_email_testo_" + lingua_abbr + " ntext NULL, " + vbCrLf + _
			  " 	i_str_pre_differita_alternativa_email_oggetto_" + lingua_abbr + " ntext NULL, " + vbCrLf + _
			  " 	i_str_pre_differita_alternativa_email_testo_" + lingua_abbr + " ntext NULL, " + vbCrLf + _
			  " 	i_str_pre_differita_login_testo_" + lingua_abbr + " ntext NULL, " + vbCrLf + _
			  " 	i_str_pre_differita_completata_form_testo_" + lingua_abbr + " ntext NULL, " + vbCrLf + _
			  " 	i_str_pre_differita_completata_inviata_testo_" + lingua_abbr + " ntext NULL, " + vbCrLf + _
			  " 	i_str_pre_differita_completata_email_oggetto_" + lingua_abbr + " ntext NULL, " + vbCrLf + _
			  " 	i_str_pre_differita_completata_email_testo_" + lingua_abbr + " ntext NULL, " + vbCrLf + _
			  " 	i_str_pre_immediata_definitiva_email_oggetto_" + lingua_abbr + " ntext NULL, " + vbCrLf + _
			  " 	i_str_pre_immediata_definitiva_email_testo_" + lingua_abbr + " ntext NULL, " + vbCrLf + _
			  " 	i_str_pre_differita_definitiva_email_oggetto_" + lingua_abbr + " ntext NULL, " + vbCrLf + _
			  " 	i_str_pre_differita_definitiva_email_testo_" + lingua_abbr + " ntext NULL, " + vbCrLf + _
			  " 	i_str_pre_term_testo_" + lingua_abbr + " ntext NULL, " + vbCrLf + _
			  " 	i_str_pre_maggiori_dettagli_testo_" + lingua_abbr + " ntext NULL; "
	else
		Aggiornamento__BOOKING3__39 = "SELECT * FROM aa_versione"
	end if	
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO BOOKING3 40
'...........................................................................................
'	Giacomo, 12/07/2012
'...........................................................................................
'  	aggiornamento per aggiungere campi riguardanti la flassibilità delle date
'...........................................................................................
function Aggiornamento__BOOKING3__40(conn)
	Aggiornamento__BOOKING3__40 = _
		" ALTER TABLE vtb_strutture ADD " & _
		"	str_attiva_date_flessibili bit NULL;" & _
		" ALTER TABLE vtb_prenotazioni ADD " & _
		"	pre_date_flessibili bit NULL; " & _
		" UPDATE vtb_strutture SET str_attiva_date_flessibili = 0 "
end function
'*******************************************************************************************




%>
<%
'...........................................................................................
'........................................................................................... 
'libreria di funzioni che contiene tutti gli aggiornamenti per il NEXT-tour
'...........................................................................................
'...........................................................................................

'...........................................................................................
'ATTENZIONE:
'il NEXT-tour estende il B2B, quindi anche gli aggiornamenti lavorano in "estensione"
'anche negli aggiornamenti.
'...........................................................................................


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR  1
'...........................................................................................
'aggiunge tabelle di gestione profili e clienti con dati agggiuntivi per la prenotazione
'...........................................................................................
function Aggiornamento_TOUR__1(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento_TOUR__1 = _
				" CREATE TABLE dbo.ttb_profili ( " + _
                "	pro_id int IDENTITY (1, 1) NOT NULL ," + vbCrLf + _
                "	pro_nome_it nvarchar (255) NULL ," + vbCrLf + _
				"	pro_nome_en nvarchar (255) NULL ," + vbCrLf + _
				"	pro_nome_fr nvarchar (255) NULL ," + vbCrLf + _
				"	pro_nome_es nvarchar (255) NULL ," + vbCrLf + _
				"	pro_nome_de nvarchar (255) NULL ," + vbCrLf + _
                "   pro_max_prenotazioni_giorno int NULL, " + vbCrLf + _
                "   pro_max_prenotazione_partecipanti int NULL, " + vbCrLf + _
                "   pro_note ntext NULL " + vbCrLf + _
				" ); " + vbCrLf + _
                " CREATE TABLE dbo.ttb_anagrafiche ( " + _
                "   ana_id INT NOT NULL, " + _
                "   ana_profilo_id INT NOT NULL " + _
                " ); " + vbCrLf + _
                " ALTER TABLE dbo.ttb_profili ADD CONSTRAINT PK_ttb_profili " + vbCrLf + _
				"	PRIMARY KEY  CLUSTERED ( pro_id)  " + vbCrLf + _
                " ALTER TABLE dbo.ttb_anagrafiche ADD CONSTRAINT PK_ttb_anagrafiche " + vbCrLf + _
				"	PRIMARY KEY  CLUSTERED ( ana_id),  " + vbCrLf + _
                "	CONSTRAINT FK_ttb_anagrafiche__gtb_rivenditori " + vbCrLf + _
				"		FOREIGN KEY (ana_id) " + vbCrLf + _
				"		REFERENCES dbo.gtb_rivenditori ( riv_id ) " + vbCrLf + _
                "       ON UPDATE CASCADE ON DELETE CASCADE, " + vbCrLF + _
                "   CONSTRAINT KF_ttb_anagrafiche__ttb_profili " + vbCrLF + _
                "       FOREIGN KEY (ana_profilo_id) " + vbCrLF + _
                "       REFERENCES dbo.ttb_profili (pro_id) " + vbCrLF + _
                "       ON UPDATE CASCADE ON DELETE CASCADE "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR  2
'...........................................................................................
'aggiunge vista per "nascondere" struttura database anagrafiche
'...........................................................................................
function Aggiornamento_TOUR__2(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento_TOUR__2 = _
				" CREATE VIEW dbo.tv_anagrafiche AS " + vbCrLf + _
                "   SELECT * FROM gtb_rivenditori INNER JOIN tb_utenti ON gtb_rivenditori.riv_id = tb_utenti.ut_id " + vbCrLF + _
                "       INNER JOIN tb_Indirizzario ON tb_Utenti.ut_NextCom_ID = tb_Indirizzario.IDElencoIndirizzi " + vbCrLF + _
                "       INNER JOIN gtb_valute ON gtb_rivenditori.riv_valuta_id = gtb_valute.valu_ID " + vbCrLF + _
                "       INNER JOIN gtb_listini ON gtb_rivenditori.riv_listino_id = gtb_listini.listino_id "+ vbCrLF + _
                "       INNER JOIN tb_cnt_lingue ON tb_Indirizzario.lingua=tb_cnt_lingue.lingua_codice " + vbCrLF + _
                "       INNER JOIN ttb_anagrafiche ON gtb_rivenditori.riv_id = ttb_anagrafiche.ana_id " + vbCrLF + _
                "       INNER JOIN ttb_profili ON ttb_anagrafiche.ana_profilo_id = ttb_profili.pro_id "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR  3
'...........................................................................................
'aggiunge campo a tabella profilo per gestione profilo pubblico prenotazione privati 
'...........................................................................................
function Aggiornamento_TOUR__3(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento_TOUR__3 = _
				" ALTER TABLE ttb_profili ADD pro_pubblico BIT NULL; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 4
'...........................................................................................
'aggiunge campo a tabella profilo per marcare "guide"
'...........................................................................................
function Aggiornamento_TOUR__4(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento_TOUR__4 = _
				" ALTER TABLE ttb_profili ADD pro_guide BIT NULL; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 5
'...........................................................................................
'aggiunge tabella gestione lingue delle guide con relativa relazione con anagrafiche
'...........................................................................................
function Aggiornamento_TOUR__5(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento_TOUR__5 = _
				" CREATE TABLE dbo.ttb_guide_lingue ( " + _
                "       lingua_id INT IDENTITY(1,1) NOT NULL, " + _
				"		lingua_codice nvarchar (2) NULL , " + _
				"		lingua_nome_IT nvarchar (20) NULL , " + _
				"		lingua_nome nvarchar (20) NULL,  " + _
                "    PRIMARY KEY  CLUSTERED (lingua_id) " + _
				" ); " + _
                " CREATE TABLE dbo.trel_guide_anagrafiche_lingue (" + _
                "       rgal_id INT IDENTITY(1,1) NOT NULL, " + _
                "       rgal_anagrafica_id INT NOT NULL, " + _
                "       rgal_lingua_id INT NOT NULL, " + _
                "   CONSTRAINT FK_trel_guide_anagrafiche_lingue__ttb_guide_lingue " + vbCrLF + _
                "       FOREIGN KEY (rgal_lingua_id) " + vbCrLF + _
                "       REFERENCES dbo.ttb_guide_lingue (lingua_id) " + vbCrLF + _
                "       ON UPDATE CASCADE ON DELETE CASCADE, " + _
                "   CONSTRAINT FK_trel_guide_anagrafiche_lingue__ttb_anagrafiche " + vbCrLF + _
                "       FOREIGN KEY (rgal_anagrafica_id) " + vbCrLF + _
                "       REFERENCES dbo.ttb_anagrafiche (ana_id) " + vbCrLF + _
                "       ON UPDATE CASCADE ON DELETE CASCADE " + _
                " ); "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 6
'...........................................................................................
'aggiunge tabella "luoghi di gestione disponibilit&agrave; delle guide" e relativa relazione
'con pacchetti e flag che indica se il pacchetto ha bisogno di guide o meno
'...........................................................................................
function Aggiornamento_TOUR__6(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento_TOUR__6 = _
				" CREATE TABLE dbo.ttb_guide_luoghi( " + _
                "   luo_id INT IDENTITY(1,1) NOT NULL, " + _
                "   luo_nome_it nvarchar(250) NULL, " + _
                "   luo_nome_en nvarchar(250) NULL, " + _
                "   luo_nome_fr nvarchar(250) NULL, " + _
                "   luo_nome_de nvarchar(250) NULL, " + _
                "   luo_nome_es nvarchar(250) NULL, " + _
                "   PRIMARY KEY CLUSTERED (luo_id) " + _
                " ) ; " + _
                " ALTER TABLE gtb_articoli ADD " + _
                "   art_guide_luogo_id INT NULL, " + _
                "   art_guida_necessaria BIT NULL, " + _
                "   CONSTRAINT FK_gtb_articoli__ttb_guide_luoghi " + vbCrLF + _
                "       FOREIGN KEY (art_guide_luogo_id) " + vbCrLF + _
                "       REFERENCES dbo.ttb_guide_luoghi (luo_id) " + vbCrLF + _
                "       ON UPDATE NO ACTION ON DELETE NO ACTION " + _
                "       NOT FOR REPLICATION; " + _
                " ALTER TABLE gtb_articoli NOCHECK CONSTRAINT FK_gtb_articoli__ttb_guide_luoghi; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 7
'...........................................................................................
'aggiunge informazioni di modifica dei record di: profili ed anagrafiche
'...........................................................................................
function Aggiornamento_TOUR__7(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento_TOUR__7 = _
				" ALTER TABLE ttb_anagrafiche ADD " + _
                "   ana_insData	datetime NULL, " + _
                "   ana_insAdmin_id	int	NULL, " + _
                "   ana_modData	datetime NULL, " + _
                "   ana_modAdmin_id	int NULL ; " + _
                " ALTER TABLE ttb_profili ADD " + _
                "   pro_insData	datetime NULL, " + _
                "   pro_insAdmin_id	int	NULL, " + _
                "   pro_modData	datetime NULL, " + _
                "   pro_modAdmin_id	int NULL ; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 8
'...........................................................................................
'aggiunge relazione tra pacchetti (articoli B2B) e guide (anagrafiche)
'...........................................................................................
function Aggiornamento_TOUR__8(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento_TOUR__8 = _
               " CREATE TABLE dbo.trel_guide_anagrafiche_pacchetti (" + _
                "       rgap_id INT IDENTITY(1,1) NOT NULL, " + _
                "       rgap_anagrafica_id INT NOT NULL, " + _
                "       rgap_art_pacchetto_id INT NOT NULL, " + _
                "   CONSTRAINT PK_trel_guide_anagrafiche_pacchetti PRIMARY KEY  CLUSTERED (rgap_id), " + _
                "   CONSTRAINT FK_trel_guide_anagrafiche_pacchetti__ttb_anagrafiche " + vbCrLF + _
                "       FOREIGN KEY (rgap_anagrafica_id) " + vbCrLF + _
                "       REFERENCES dbo.ttb_anagrafiche (ana_id) " + vbCrLF + _
                "       ON UPDATE CASCADE ON DELETE CASCADE, " + _
                "   CONSTRAINT FK_trel_guide_anagrafiche_pacchetti__gtb_articoli " + vbCrLF + _
                "       FOREIGN KEY (rgap_art_pacchetto_id) " + vbCrLF + _
                "       REFERENCES dbo.gtb_articoli (art_id) " + vbCrLF + _
                "       ON UPDATE CASCADE ON DELETE CASCADE " + _
                " ); "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 9
'...........................................................................................
'aggiunge chiave primaria tra guide (anagrafiche) e lingue
'...........................................................................................
function Aggiornamento_TOUR__9(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento_TOUR__9 = _
               " ALTER TABLE trel_guide_anagrafiche_lingue ADD " + _
               " CONSTRAINT PK_trel_guide_anagrafiche_lingue PRIMARY KEY (rgal_id) ; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 10
'...........................................................................................
'aggiunge vista per "incapsulamento" struttura database pacchetti
'...........................................................................................
function Aggiornamento_TOUR__10(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento_TOUR__10 = _
               " CREATE VIEW dbo.tv_pacchetti AS " + vbCrLf + _
               "    SELECT * FROM gtb_articoli INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id " + vbCrLF + _
               "        INNER JOIN gtb_marche ON gtb_articoli.art_marca_id = gtb_marche.mar_id " + vbCrLF + _
               "        LEFT JOIN ttb_guide_luoghi ON gtb_articoli.art_guide_luogo_id = ttb_guide_luoghi.luo_id "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 11
'...........................................................................................
'aggiunge campo per indicare il magazzino "disponibilita'" per ogni profilo con relativa 
'relazione con i magazzini del b2b
'...........................................................................................
function Aggiornamento_TOUR__11(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento_TOUR__11 = _
               " ALTER TABLE ttb_profili ADD " + _
               "    pro_magazzino_dispo_id INT NULL ; " + _
               " UPDATE ttb_profili SET pro_magazzino_dispo_id = (SELECT TOP 1 mag_id FROM gtb_magazzini) ; " + _
               " ALTER TABLE ttb_profili ALTER COLUMN pro_magazzino_dispo_id INT NOT NULL ; " + _
               " ALTER TABLE ttb_profili ADD " + _
               "    CONSTRAINT FK_ttb_profili__gtb_magazzini " + _
               "        FOREIGN KEY (pro_magazzino_dispo_id) " + _
                "       REFERENCES gtb_magazzini (mag_id) " + _
                "       ON UPDATE NO ACTION ON DELETE NO ACTION "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 12
'...........................................................................................
'aggiunge campo per registrare la "data/ora" di ogni variante dei pacchetti in "sintetico"
'...........................................................................................
function Aggiornamento_TOUR__12(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento_TOUR__12 = _
				" ALTER TABLE grel_art_Valori ADD " + _
				"	rel_tour_dataora SMALLDATETIME NULL ; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 13
'...........................................................................................
'aggiunge campo per registrare tipo di disponibilit&agrave; associata al pacchetto
'...........................................................................................
function Aggiornamento_TOUR__13(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento_TOUR__13 = _
				" ALTER TABLE gtb_articoli ADD " + _
                "   art_tour_dispo_per_profilo BIT NULL ; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 14
'...........................................................................................
'aggiunge campo per registrare tipo vendita del pacchetto: se ad orari o a giorno
'...........................................................................................
function Aggiornamento_TOUR__14(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento_TOUR__14 = _
				" ALTER TABLE gtb_articoli ADD " + _
                "   art_tour_dispo_giornaliera BIT NULL ; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 15
'...........................................................................................
'aggiunge struttura per registrazione disponibilita' guida
'...........................................................................................
function Aggiornamento_TOUR__15(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento_TOUR__15 = _
				" CREATE TABLE dbo.trel_guide_luoghi_dispo (" + _
                "       rgld_id INT IDENTITY(1,1) NOT NULL, " + _
                "       rgld_anagrafica_id INT NOT NULL, " + _
                "       rgld_luogo_id INT NOT NULL, " + _
                "       rgld_data_from SMALLDATETIME NULL, " + _
                "       rgld_data_to SMALLDATETIME NULL, " + _
                "   CONSTRAINT PK_trel_guide_luoghi_dispo PRIMARY KEY  CLUSTERED (rgld_id), " + _
                "   CONSTRAINT FK_trel_guide_luoghi_dispo__ttb_anagrafiche " + vbCrLF + _
                "       FOREIGN KEY (rgld_anagrafica_id) " + vbCrLF + _
                "       REFERENCES dbo.ttb_anagrafiche (ana_id) " + vbCrLF + _
                "       ON UPDATE CASCADE ON DELETE CASCADE, " + _
                "   CONSTRAINT FK_trel_guide_luoghi_dispo__ttb_guide_luoghi " + vbCrLF + _
                "       FOREIGN KEY (rgld_luogo_id) " + vbCrLF + _
                "       REFERENCES dbo.ttb_guide_luoghi (luo_id) " + vbCrLF + _
                "       ON UPDATE CASCADE ON DELETE CASCADE " + _
                " ); "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 16
'...........................................................................................
'aggiunge campo per registrazione durata pacchetto
'...........................................................................................
function Aggiornamento_TOUR__16(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento_TOUR__16 = _
				" ALTER TABLE gtb_articoli ADD " + _
                "   art_tour_durata INT NULL ; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 17
'...........................................................................................
'aggiunge campo per registrazione home page area riservata profilo
'...........................................................................................
function Aggiornamento_TOUR__17(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento_TOUR__17 = _
				" ALTER TABLE ttb_profili ADD " + _
                "    pro_homepage_area_riservata INT NULL ; " + _
                " ALTER TABLE ttb_profili ADD " + _
                "    CONSTRAINT FK_ttb_profili__tb_pagineSito " + _
                "        FOREIGN KEY (pro_homepage_area_riservata) " + _
                "       REFERENCES tb_paginesito (id_pagineSito) " + _
                "       ON UPDATE NO ACTION ON DELETE NO ACTION ; " + _
                " ALTER TABLE ttb_profili NOCHECK CONSTRAINT FK_ttb_profili__tb_pagineSito; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 18
'...........................................................................................
'aggiunge struttura per registrazione prenotazioni della guida
'...........................................................................................
function Aggiornamento_TOUR__18(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento_TOUR__18 = _
				" CREATE TABLE dbo.trel_guide_prenotazioni (" + _
                "       rgp_id INT IDENTITY(1,1) NOT NULL, " + _
                "       rgp_anagrafica_id INT NOT NULL, " + _
                "       rgp_det_ord_id INT NOT NULL, " + _
                "       rgp_data_from SMALLDATETIME NULL, " + _
                "       rgp_data_to SMALLDATETIME NULL, " + _
                "   CONSTRAINT PK_trel_guide_prenotazioni PRIMARY KEY  CLUSTERED (rgp_id), " + _
                "   CONSTRAINT FK_trel_guide_prenotazioni__ttb_anagrafiche " + vbCrLF + _
                "       FOREIGN KEY (rgp_anagrafica_id) " + vbCrLF + _
                "       REFERENCES dbo.ttb_anagrafiche (ana_id) " + vbCrLF + _
                "       ON UPDATE CASCADE ON DELETE CASCADE, " + _
                "   CONSTRAINT FK_PK_trel_guide_prenotazioni__gtb_dettagli_ord " + vbCrLF + _
                "       FOREIGN KEY (rgp_det_ord_id) " + vbCrLF + _
                "       REFERENCES dbo.gtb_dettagli_ord (det_id) " + vbCrLF + _
                "       ON UPDATE NO ACTION ON DELETE NO ACTION " + _
                " ); "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 19
'...........................................................................................
'aggiunge campo per registrazione guida associata al dettagli d'ordine
'...........................................................................................
function Aggiornamento_TOUR__19(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento_TOUR__19 = _
				" ALTER TABLE gtb_dett_cart ADD " + _
                "    dett_tour_anagrafica_id INT NULL ; " + _
                " ALTER TABLE gtb_dett_cart ADD " + _
                "    CONSTRAINT FK_gtb_dett_cart__ttb_anagrafiche " + _
                "       FOREIGN KEY (dett_tour_anagrafica_id) " + _
                "       REFERENCES ttb_anagrafiche (ana_id) " + _
                "       ON UPDATE NO ACTION ON DELETE NO ACTION ; " + _
                " ALTER TABLE gtb_dett_cart NOCHECK CONSTRAINT FK_gtb_dett_cart__ttb_anagrafiche; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 20
'...........................................................................................
'aggiunge struttura per registrazione limiti divendita per categoria
'...........................................................................................
function Aggiornamento_TOUR__20(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento_TOUR__20 = _
                " CREATE TABLE " + SQL_Dbo(Conn) + "trel_vincoli_profili_tipologie ( " + _
                "	rpt_id " + SQL_PrimaryKey(conn, "trel_vincoli_profili_tipologie") + "," + vbCrLf + _
		        "	rpt_pro_id INT NOT NULL," + vbCrLf + _
                "   rpt_tip_id INT NOT NULL, " + vbCrLf + _
                "   rpt_max_prenotazioni_giorno int NULL, " + vbCrLf + _
                "   rpt_max_prenotazione_partecipanti int NULL " + vbCrLf + _
                " ) ; " + _
                " ALTER TABLE trel_vincoli_profili_tipologie ADD " + _
                "    CONSTRAINT FK_trel_vincoli_profili_tipologie__ttb_profili " + _
                "       FOREIGN KEY (rpt_pro_id) " + _
                "       REFERENCES ttb_profili (pro_id) " + _
                "       ON UPDATE NO ACTION ON DELETE NO ACTION, " + _
                "    CONSTRAINT FK_trel_vincoli_profili_tipologie__gtb_tipologie " + _
                "       FOREIGN KEY (rpt_tip_id) " + _
                "       REFERENCES gtb_tipologie (tip_id) " + _
                "       ON UPDATE NO ACTION ON DELETE NO ACTION " + _
                "; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 21
'...........................................................................................
'aggiornamento ai profili e alle anagrafiche
'...........................................................................................
function Aggiornamento_TOUR__21(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento_TOUR__21 = _
                "ALTER TABLE ttb_profili ADD" + vbCrLf + _
				"	pro_guide_obbligatorie BIT NULL," + vbCrLf + _
				"	pro_lastMinute BIT NULL," + vbCrLf + _
				"	pro_referente_obbligatorio BIT NULL;" + vbCrLf + _
				"ALTER TABLE trel_vincoli_profili_tipologie ADD" + vbCrLf + _
				"	rpt_max_fasce_prenotabili INT NULL;" + vbCrLf + _
				"ALTER TABLE ttb_anagrafiche ADD" + vbCrLf + _
				"	ana_homepage_area_riservata INT NULL;" + vbCrLf + _
				"ALTER TABLE gtb_ordini ADD" + vbCrLf + _
				" 	ord_referente NVARCHAR(255) NULL;"
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 22
'...........................................................................................
'aggiunge campo per registrazione guida associata al dettagli d'ordine
'...........................................................................................
function Aggiornamento_TOUR__22(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento_TOUR__22 = _
				" ALTER TABLE gtb_dettagli_ord ADD " + _
                "    det_tour_anagrafica_id INT NULL ; " + _
                " ALTER TABLE gtb_dettagli_ord ADD " + _
                "    CONSTRAINT FK_gtb_dettagli_ord__ttb_anagrafiche " + _
                "       FOREIGN KEY (det_tour_anagrafica_id) " + _
                "       REFERENCES ttb_anagrafiche (ana_id) " + _
                "       ON UPDATE NO ACTION ON DELETE NO ACTION ; " + _
                " ALTER TABLE gtb_dettagli_ord NOCHECK CONSTRAINT FK_gtb_dettagli_ord__ttb_anagrafiche; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 23
'...........................................................................................
'rimuove campi di opzioni associati al profilo
'...........................................................................................
function Aggiornamento_TOUR__23(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento_TOUR__23 = _
				SQL_RemoveForeignKey(conn, "ttb_profili", "", "", false, "FK_ttb_profili__gtb_magazzini") + _
				" ALTER TABLE ttb_profili DROP COLUMN pro_magazzino_dispo_id; " + _
				" ALTER TABLE ttb_profili DROP COLUMN pro_max_prenotazioni_giorno; " + _
				" ALTER TABLE ttb_profili DROP COLUMN pro_max_prenotazione_partecipanti; " + _
				" ALTER TABLE ttb_profili DROP COLUMN pro_guide_obbligatorie; " + _
				" ALTER TABLE ttb_profili DROP COLUMN pro_lastMinute; " + _
				" ALTER TABLE trel_vincoli_profili_tipologie ADD " + _
				"	rpt_magazzino_dispo_id int NULL, " + _
				" 	rpt_guide_obbligatorie bit NULL, " + _
				"	rpt_limite_ore_prenotazione INT NULL, " + _
				"   rpt_limite_ore_prenotazione_mod INT NULL " + _
				" ; " + _
				SQL_AddForeignKey(conn, "trel_vincoli_profili_tipologie", "rpt_magazzino_dispo_id", "gtb_magazzini", "mag_id", true, "")
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 24
'...........................................................................................
'correge relazioni tra vincoli e profili
'...........................................................................................
function Aggiornamento_TOUR__24(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento_TOUR__24 = _
				SQL_RemoveForeignKey(conn, "trel_vincoli_profili_tipologie", "", "", false, "FK_trel_vincoli_profili_tipologie__ttb_profili") + _
				SQL_RemoveForeignKey(conn, "trel_vincoli_profili_tipologie", "", "", false, "FK_trel_vincoli_profili_tipologie__gtb_tipologie") + _
				SQL_AddForeignKey(conn, "trel_vincoli_profili_tipologie", "rpt_pro_id", "ttb_profili", "pro_id", true, "") + _
				SQL_AddForeignKey(conn, "trel_vincoli_profili_tipologie", "rpt_tip_id", "gtb_tipologie", "tip_id", true, "")
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 25
'...........................................................................................
'aggiunge il tipo procuratoria ai profili
'...........................................................................................
function Aggiornamento_TOUR__25(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento_TOUR__25 = _
				" ALTER TABLE ttb_profili ADD" + vbCrLf + _
				"	pro_procuratoria BIT NULL;"
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 26
'...........................................................................................
'aggiunge il lock alle anagrafiche
'...........................................................................................
function AggiornamentoSpeciale__TOUR__26(DB, version)
	'esegue un aggiornamento fasullo per aumentare la versione
	sql = "SELECT * FROM AA_Versione"
	CALL DB.Execute(sql, version)
	if DB.last_update_executed then
		dim rs, OBJ_contatto
		set rs = server.createobject("adodb.recordset")
		rs.open "SELECT idElencoIndirizzi FROM tv_anagrafiche", conn, adOpenStatic, adLockOptimistic
		set OBJ_contatto = new IndirizzarioLock
		OBJ_contatto.conn = conn
		
		while not rs.eof
			CALL OBJ_Contatto.LockContact(rs("idElencoIndirizzi"), NEXTTOUR)
			rs.movenext
		wend
		
		rs.close
		set rs = nothing
		set OBJ_contatto = nothing
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 27
'...........................................................................................
'	Nicola, 10/10/2009
'...........................................................................................
'   aggiunge campi per collegamento google maps ai pacchetti
'...........................................................................................
function Aggiornamento_TOUR__27(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento_TOUR__27 = _
				" ALTER TABLE gtb_articoli ADD" + vbCrLf + _
				"	art_tour_google_maps_latitudine FLOAT NULL, " + _
				"	art_tour_google_maps_longitudine FLOAT NULL; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 28
'...........................................................................................
'	Matteo, 26/01/2010
'...........................................................................................
'   aggiunge tabella per le destinazioni transfer
'...........................................................................................
function Aggiornamento_TOUR__28(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento_TOUR__28 = _
				" CREATE TABLE dbo.ttb_destinazioni ( " + _
                "	dest_id int IDENTITY (1, 1) NOT NULL ," + vbCrLf + _
                "	dest_nome_it nvarchar (255) NULL ," + vbCrLf + _
				"	dest_nome_en nvarchar (255) NULL ," + vbCrLf + _
				"	dest_nome_fr nvarchar (255) NULL ," + vbCrLf + _
				"	dest_nome_es nvarchar (255) NULL ," + vbCrLf + _
				"	dest_nome_de nvarchar (255) NULL " + vbCrLf + _
				" ); " + vbCrLf + _
                " ALTER TABLE dbo.ttb_destinazioni ADD CONSTRAINT PK_ttb_destinazioni " + vbCrLf + _
				"	PRIMARY KEY  CLUSTERED ( dest_id)  " + vbCrLf + _
				" ALTER TABLE gtb_articoli ADD" + vbCrLf + _
				"	art_tour_dest_partenza_id INT NULL, " + _
				"	art_tour_dest_arrivo_id INT NULL; " + _
				SQL_AddForeignKeyExtended(conn, "gtb_articoli", "art_tour_dest_partenza_id", "ttb_destinazioni", "dest_id", false, false, "FK_gtb_articoli__ttb_destinazioni_partenza") + _
				SQL_AddForeignKeyExtended(conn, "gtb_articoli", "art_tour_dest_arrivo_id", "ttb_destinazioni", "dest_id", false, false, "FK_gtb_articoli__ttb_destinazioni_arrivo")
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 29
'...........................................................................................
'	Matteo, 03/02/2010
'...........................................................................................
'   aggiunge stato di lavorazione dell'ordine per le prenotazioni transfer
'...........................................................................................
function Aggiornamento_TOUR__29(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento_TOUR__29 = _
				" INSERT INTO gtb_stati_ordine (so_nome_it, so_ordine, so_stato_ordini, so_internet) VALUES ('prenotazione transfer', 0, 2, 0)"
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 30
'...........................................................................................
'	Nicola, 25/03/2010
'...........................................................................................
'   
'...........................................................................................
function Aggiornamento_TOUR__30(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento_TOUR__30 = _
				" ALTER TABLE ttb_profili ADD" + vbCrLf + _
				"	pro_diritto_prenotazione FLOAT NULL; "
	end select
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 31
'...........................................................................................
'	Giacomo, 25/03/2010
'...........................................................................................
'   aggiunge descrittore sito 
'...........................................................................................
function Aggiornamento_TOUR__31(conn)
	Aggiornamento_TOUR__31 = " SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale_TOUR__31(conn)
	CALL AddParametroSito(conn, "GRUPPO_LAVORO_DEFAULT", _
								null, _
								"GRUPPO DI LAVORO DEFAULT", _
								"", _
								adNumeric, _
								0, _
								"", _
								1, _
								1, _
								NEXTTOUR, _
								GetValueList(conn, NULL, "SELECT TOP 1 id_gruppo FROM tb_gruppi"), null, null, null, null)
	
	AggiornamentoSpeciale_TOUR__31 = " SELECT * FROM AA_Versione "
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 32
'...........................................................................................
'	Giacomo, 30/03/2010
'...........................................................................................
'   aggiunge parametro
'...........................................................................................
function Aggiornamento_TOUR__32(conn)
	Aggiornamento_TOUR__32 = " SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale_TOUR__32(conn)
	CALL AddParametroSito(conn, "VOUCHER_PAGE", _
								null, _
								"ID pagina sito di conferma prenotazione", _
								"", _
								adGUID, _
								0, _
								"", _
								1, _
								1, _
								NEXTTOUR, _
								null, null, null, null, null)
end function
'*******************************************************************************************




'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 33
'...........................................................................................
'	Giacomo, 31/03/2010
'...........................................................................................
'   aggiunge parametro
'...........................................................................................
function Aggiornamento_TOUR__33(conn)
	Aggiornamento_TOUR__33 = " SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale_TOUR__33(conn)
	CALL AddParametroSito(conn, "GUIDA_BADGE", _
								null, _
								"Pagina generazione badge della guida", _
								"", _
								adGUID, _
								0, _
								"", _
								1, _
								1, _
								NEXTTOUR, _
								null, null, null, null, null)
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 34
'...........................................................................................
'	Giacomo, 01/04/2010
'...........................................................................................
function Aggiornamento_TOUR__34(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento_TOUR__34 = _
				" ALTER TABLE ttb_profili DROP COLUMN pro_diritto_prenotazione; " + _
				" ALTER TABLE ttb_profili ADD " + _
				"   pro_metodo_pagamento_id INT NULL, " + _
				"   pro_listino_default_id INT NULL, " + _
				"   pro_max_accompagnatori INT NULL; " + _
				SQL_AddForeignKey(conn, "ttb_profili", "pro_metodo_pagamento_id", "gtb_modipagamento", "mosp_id", true, "") & _
				SQL_AddForeignKey(conn, "ttb_profili", "pro_listino_default_id", "gtb_listini", "listino_id", false, "") & _
				" UPDATE ttb_profili SET pro_metodo_pagamento_id = (SELECT TOP 1 mosp_id FROM gtb_modipagamento) ; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 35
'...........................................................................................
'	Giacomo, 07/04/2010
'...........................................................................................
function Aggiornamento_TOUR__35(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento_TOUR__35 = _
				" UPDATE trel_vincoli_profili_tipologie SET rpt_limite_ore_prenotazione = (rpt_limite_ore_prenotazione * 60) ; " + _
				" UPDATE trel_vincoli_profili_tipologie SET rpt_limite_ore_prenotazione_mod = (rpt_limite_ore_prenotazione_mod * 60) ; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 36
'...........................................................................................
'	Giacomo, 31/03/2010
'...........................................................................................
'   aggiunge parametro
'...........................................................................................
function Aggiornamento_TOUR__36(conn)
	Aggiornamento_TOUR__36 = " SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale_TOUR__36(conn)
	CALL AddParametroSito(conn, "CATEGORIA_TRANSFER", _
								null, _
								"id corrispondente alla categoria transfer ", _
								"", _
								adNumeric, _
								0, _
								"", _
								1, _
								1, _
								NEXTTOUR, _
								null, null, null, null, null)
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 37
'...........................................................................................
'	Matteo, 10/06/2010
'...........................................................................................
'   aggiunge campo per il link al contratto relativo al profilo
'...........................................................................................
function Aggiornamento_TOUR__37(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento_TOUR__37 = _
				" ALTER TABLE ttb_profili ADD" + vbCrLf + _
				"	pro_contratto_url nvarchar (255) NULL; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 38
'...........................................................................................
'	Matteo, 10/06/2010
'...........................................................................................
'   corregge l'aggiunta campo per il link al contratto relativo al profilo (multilingua)
'...........................................................................................
function Aggiornamento_TOUR__38(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento_TOUR__38 = _
				" ALTER TABLE ttb_profili DROP COLUMN pro_contratto_url; " + _
				" ALTER TABLE trel_vincoli_profili_tipologie ADD " + _
				SQL_MultiLanguageField("	pro_contratto_url_<lingua> " + SQL_CharField(Conn, 255)) + " ; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 39
'...........................................................................................
'	Matteo, 10/06/2010
'...........................................................................................
'   corregge l'aggiunta campo per il link al contratto relativo al profilo (multilingua)
'...........................................................................................
function Aggiornamento_TOUR__39(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento_TOUR__39 = _
				" ALTER TABLE trel_vincoli_profili_tipologie DROP COLUMN pro_contratto_url_it; " + _
				" ALTER TABLE trel_vincoli_profili_tipologie DROP COLUMN pro_contratto_url_en; " + _
				" ALTER TABLE trel_vincoli_profili_tipologie DROP COLUMN pro_contratto_url_de; " + _
				" ALTER TABLE trel_vincoli_profili_tipologie DROP COLUMN pro_contratto_url_fr; " + _
				" ALTER TABLE trel_vincoli_profili_tipologie DROP COLUMN pro_contratto_url_es; " + _
				" ALTER TABLE ttb_profili ADD " + _
				SQL_MultiLanguageField("	pro_contratto_url_<lingua> " + SQL_CharField(Conn, 255)) + " ; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 40
'...........................................................................................
'	Matteo, 19/07/2010
'...........................................................................................
'   aggiunge il diritto di prenotazione (money) ai vincoli di profilo
'...........................................................................................
function Aggiornamento_TOUR__40(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento_TOUR__40 = _
				" ALTER TABLE trel_vincoli_profili_tipologie ADD" + vbCrLf + _
				"	rpt_diritto_prenotazione MONEY NULL;"
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 41
'...........................................................................................
'	Andrea, 19/07/2010
'...........................................................................................
'   calcola i valori dei campi totali e totali iva per gli ordini
'...........................................................................................
function Aggiornamento_TOUR__41(conn)				
		dim rsnew, sql, sql1
		
		set rsnew = Server.CreateObject("ADODB.RecordSet")

		sql = " SELECT * FROM gtb_ordini" + _
		  " INNER JOIN gtb_dettagli_ord ON gtb_ordini.ord_id=gtb_dettagli_ord.det_ord_id"
		  
		rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdtext
		
		sql1=""
		while not rs.eof			
			sql1 = sql1 & " UPDATE gtb_ordini SET ord_totale=" + _
			ParseSQL(GetTotaleDettaglioPrenotazione(conn,rsnew,rs), adNumeric) +_
			" WHERE ord_id=" & rs("ord_id") & ";"		
			rs.movenext
		wend
		rs.close
		
		Aggiornamento_TOUR__41 = _
		" UPDATE gtb_ordini SET ord_totale_iva=0;" & sql1
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 42
'...........................................................................................
'	Matteo, 19/07/2010
'...........................................................................................
'   aggiunge il vincolo per il tipo di fatturazione
'...........................................................................................
function Aggiornamento_TOUR__42(conn)
	Aggiornamento_TOUR__42 = _
		" ALTER TABLE trel_vincoli_profili_tipologie ADD" + vbCrLf + _
		"	rpt_fatturazione_id INT NULL;" + _
		SQL_AddForeignKey(conn, "trel_vincoli_profili_tipologie", "rpt_fatturazione_id", "gtb_fatturazioni", "fatt_id", false, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 43
'...........................................................................................
'	Giacomo, 26/08/2010
'...........................................................................................
'   aggiunge tabella di relazione tra gli admin e gli organizzatori 
'...........................................................................................
function Aggiornamento_TOUR__43(conn)
	Aggiornamento_TOUR__43 = _
		"CREATE TABLE dbo.trel_admin_organizzatori (" + vbCrLf + _
		"	rao_admin_id " & SQL_PrimaryKeyInt(conn, "trel_admin_organizzatori") & ", " + vbCrLf + _
		"	rao_organizzatore_id int NULL ," + vbCrLf + _
		" ); " + vbCrLf + _
		SQL_AddForeignKey(conn, "trel_admin_organizzatori", "rao_admin_id", "tb_admin", "id_admin", true, "") & _
		SQL_AddForeignKey(conn, "trel_admin_organizzatori", "rao_organizzatore_id", "gtb_marche", "mar_id", false, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 44
'...........................................................................................
'	Giacomo, 26/08/2010
'...........................................................................................
'   modifica i permessi del tour
'...........................................................................................
function Aggiornamento_TOUR__44(conn)
	Aggiornamento_TOUR__44 = _
		"UPDATE rel_admin_sito SET rel_as_permesso = 1 WHERE sito_id = " & NEXTTOUR & ";" & _
		"UPDATE tb_siti SET sito_p1 = 'TOUR_USER', sito_p2 = 'TOUR_ORGANIZZATORE', sito_p3 = '', sito_prmEsterni_admin = '../nextTour/PassportAdmin.asp', " & _
		"	sito_prmEsterni_sito = '../nextTour/PassportSito.asp'  WHERE id_sito = " & NEXTTOUR & ";"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 45
'...........................................................................................
'	Giacomo, 28/10/2010
'...........................................................................................
'   aggiunta campo tour_lingua
'...........................................................................................
function Aggiornamento_TOUR__45(conn)
	Aggiornamento_TOUR__45 = _
		" ALTER TABLE grel_art_valori ADD" + vbCrLf + _
		"	rel_tour_lingua " & SQL_CharField(Conn, 50) & ";"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 46
'...........................................................................................
'	Giacomo, 28/10/2010
'...........................................................................................
'   inserisce la variante lingua con i relativi valori per 5 lingue
'...........................................................................................
function Aggiornamento_TOUR__46(conn)
	Aggiornamento_TOUR__46 = _
		" INSERT INTO gtb_varianti(var_nome_it, var_nome_en, var_ordine, var_descr_it) " + vbCrLf + _
		" VALUES('Lingua', 'Language', 3, 'Lingue dei tour');" + vbCrLf + _
		" INSERT INTO gtb_valori(val_nome_it, val_nome_en, val_icona, val_cod_int, val_ordine, val_var_id) " + vbCrLf + _
		" SELECT 'Italiano' AS a, 'Italian' AS b, '/interfaccia_1024/tour/bandiere/flag_IT.jpg' AS c, 'it' AS d, 1 AS e, var_id " + vbCrLf + _
		" FROM gtb_varianti WHERE var_nome_it LIKE 'Lingua' ;" + _
		" INSERT INTO gtb_valori(val_nome_it, val_nome_en, val_icona, val_cod_int, val_ordine, val_var_id) " + vbCrLf + _
		" SELECT 'Inglese' AS a, 'English' AS b, '/interfaccia_1024/tour/bandiere/flag_EN.jpg' AS c, 'en' AS d, 2 AS e, var_id " + vbCrLf + _
		" FROM gtb_varianti WHERE var_nome_it LIKE 'Lingua' ;" + _
		" INSERT INTO gtb_valori(val_nome_it, val_nome_en, val_icona, val_cod_int, val_ordine, val_var_id) " + vbCrLf + _
		" SELECT 'Francese' AS a, 'French' AS b, '/interfaccia_1024/tour/bandiere/flag_FR.jpg' AS c, 'fr' AS d, 3 AS e, var_id " + vbCrLf + _
		" FROM gtb_varianti WHERE var_nome_it LIKE 'Lingua' ;" + _
		" INSERT INTO gtb_valori(val_nome_it, val_nome_en, val_icona, val_cod_int, val_ordine, val_var_id) " + vbCrLf + _
		" SELECT 'Tedesco' AS a, 'German' AS b, '/interfaccia_1024/tour/bandiere/flag_DE.jpg' AS c, 'de' AS d, 4 AS e, var_id " + vbCrLf + _
		" FROM gtb_varianti WHERE var_nome_it LIKE 'Lingua' ;" + _
		" INSERT INTO gtb_valori(val_nome_it, val_nome_en, val_icona, val_cod_int, val_ordine, val_var_id) " + vbCrLf + _
		" SELECT 'Spagnolo' AS a, 'Spanish' AS b, '/interfaccia_1024/tour/bandiere/flag_ES.jpg' AS c, 'es' AS d, 5 AS e, var_id " + vbCrLf + _
		" FROM gtb_varianti WHERE var_nome_it LIKE 'Lingua' ;"
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 47
'...........................................................................................
'	Giacomo, 28/10/2010
'...........................................................................................
'   aggiunge paramento per l'id della variante lingua
'...........................................................................................
function Aggiornamento_TOUR__47(conn)
	Aggiornamento_TOUR__47 = " SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale_TOUR__47(conn)
	CALL AddParametroSito(conn, "VARIANTE_LINGUA", _
								null, _
								"id della variante lingua", _
								"", _
								adNumeric, _
								0, _
								"", _
								1, _
								1, _
								NEXTTOUR, _
								GetValueList(conn, NULL, "SELECT var_id FROM gtb_varianti WHERE var_nome_it LIKE 'Lingua'"), null, null, null, null)
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 48
'...........................................................................................
'	Nicola, 29/10/2010
'...........................................................................................
'   aggiunte tabella per tipologia 
'...........................................................................................
function Aggiornamento_TOUR__48(conn)
	Aggiornamento_TOUR__48 = _
		"CREATE TABLE dbo.ttb_pacchetti_tipo (" + vbCrLf + _
		"	tp_id " & SQL_PrimaryKeyInt(conn, "ttb_pacchetti_tipo") & ", " + vbCrLf + _
		"	tp_nome nvarchar(250) NULL ," + vbCrLf + _
		" ); " + vbCrLf + _
		" ALTER TABLE gtb_articoli ADD" + vbCrLf + _
		"	art_tour_tipo_id int NULL ; " + vbCrLf + _
		SQL_AddForeignKey(conn, "gtb_articoli", "art_tour_tipo_id", "ttb_pacchetti_tipo", "tp_id", false, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 49
'...........................................................................................
'	Matteo, 02/11/2010
'...........................................................................................
'   aggiunte tabella orari disponibilità tour
'...........................................................................................
function Aggiornamento_TOUR__49(conn)
	Aggiornamento_TOUR__49 = _
		"CREATE TABLE " & SQL_dbo(conn) & "ttb_pacchetti_dispo_orari (" + vbCrLf + _
		"	odp_id " & SQL_PrimaryKey(conn, "ttb_pacchetti_dispo_orari") & ", " + vbCrLf + _
		"	odp_lingua NVARCHAR(50) NULL , " + vbCrLf + _
		"	odp_se_lunedi BIT NULL , " + vbCrLf + _
		"	odp_se_martedi BIT NULL , " + vbCrLf + _
		"	odp_se_mercoledi BIT NULL , " + vbCrLf + _
		"	odp_se_giovedi BIT NULL , " + vbCrLf + _
		"	odp_se_venerdi BIT NULL , " + vbCrLf + _
		"	odp_se_sabato BIT NULL , " + vbCrLf + _
		"	odp_se_domenica BIT NULL , " + vbCrLf + _
		"	odp_orario NVARCHAR(50) NULL " + vbCrLf + _
		" ); "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 50
'...........................................................................................
'	Matteo, 02/11/2010
'...........................................................................................
'   aggiunte tabella periodi disponibilità tour
'...........................................................................................
function Aggiornamento_TOUR__50(conn)
	Aggiornamento_TOUR__50 = _
		"CREATE TABLE " & SQL_dbo(conn) & "ttb_pacchetti_dispo_periodi (" + vbCrLf + _
		"	pdp_id " & SQL_PrimaryKey(conn, "ttb_pacchetti_dispo_periodi") & ", " + vbCrLf + _
		"	pdp_tour_id int NULL , " + vbCrLf + _
		"	pdp_orario_id int NULL , " + vbCrLf + _
		"	pdp_data_inizio datetime NULL , " + vbCrLf + _
		"	pdp_data_fine datetime NULL , " + vbCrLf + _
		"	pdp_periodo_escluso BIT NULL " + vbCrLf + _
		" ); " + vbCrLf + _
		SQL_AddForeignKey(conn, "ttb_pacchetti_dispo_periodi", "pdp_tour_id", "gtb_articoli", "art_id", false, "") + _
		SQL_AddForeignKey(conn, "ttb_pacchetti_dispo_periodi", "pdp_orario_id", "ttb_pacchetti_dispo_orari", "odp_id", false, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 51
'...........................................................................................
'	Nicola, 04/11/2010
'...........................................................................................
'   aggiunge tabella dei giorni
'...........................................................................................
function Aggiornamento_TOUR__51(conn)
	Aggiornamento_TOUR__51 = _
		"CREATE TABLE " & SQL_dbo(conn) & "ttb_giorni_anno (" + vbCrLf + _
		"	g_id " & SQL_PrimaryKey(conn, "ttb_giorni_anno") & ", " + vbCrLf + _
		"	g_date datetime NULL " + vbCrLf + _
		" ); " + vbCrLf
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 52
'...........................................................................................
'	Nicola, 09/11/2010
'...........................................................................................
'   aggiunge trigger per generazione giorni
'...........................................................................................
function Aggiornamento_TOUR__52(conn)
	Aggiornamento_TOUR__52 = _
		" CREATE TRIGGER [dbo].[ttb_pacchetti_dispo_periodi_insert_update] " + vbCrLF + _
		"	ON [dbo].[ttb_pacchetti_dispo_periodi] " + vbCrLF + _
		"	AFTER INSERT, UPDATE " + vbCrLF + _
		" AS  " + vbCrLF + _
		" 	 DECLARE @MinDate SMALLDATETIME " + vbCrLF + _
		"	 DECLARE @MaxDate SMALLDATETIME " + vbCrLF + _
		"	 DECLARE @Date SMALLDATETIME " + vbCrLF + _
		"	 DECLARE @Index INT " + vbCrLF + _
		"	 /*apre recordset con ordini inseriti*/  " + vbCrLF + _
		"	 DECLARE rs CURSOR local FAST_FORWARD FOR   " + vbCrLF + _
		"		SELECT DISTINCT pdp_data_inizio, pdp_data_fine FROM inserted " + vbCrLF + _
		"	 OPEN rs  " + vbCrLF + _
		"	 FETCH NEXT FROM rs INTO @MinDate, @MaxDate  " + vbCrLF + _
		"	 WHILE @@FETCH_STATUS = 0  " + vbCrLF + _
		"	 BEGIN  " + vbCrLF + _
		"		SELECT @MinDate, @MaxDate, DateDiff(day, @MinDate, @MaxDate) " + vbCrLF + _
		"		SET @Index = 0 " + vbCrLF + _
		"	 	WHILE (@Index <= DateDiff(day, @MinDate, @MaxDate)) " + vbCrLF + _
		"		BEGIN " + vbCrLF + _
		"			SET @Date = DateAdd(day, @Index, @MinDate) " + vbCrLF + _
		"			IF NOT EXISTS(SELECT * FROM ttb_giorni_anno WHERE g_date = @Date) " + vbCrLF + _
		"				BEGIN " + vbCrLF + _
		"					INSERT INTO ttb_giorni_anno (g_date) VALUES (@Date) " + vbCrLF + _
		"				END " + vbCrLF + _
		"			SET @Index = @Index + 1 " + vbCrLF + _
		"		END " + vbCrLF + _
		"	 	FETCH NEXT FROM rs INTO @MinDate, @MaxDate  " + vbCrLF + _
		"	 END " + vbCrLF + _
		" ;" 
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 53
'...........................................................................................
'	Nicola, 09/11/2010
'...........................................................................................
'   imposta tipi pacchetti
'...........................................................................................
function Aggiornamento_TOUR__53(conn)
	Aggiornamento_TOUR__53 = _
		" INSERT INTO ttb_pacchetti_tipo (tp_id, tp_nome) " + _
		" 	   VALUES (1, 'Pacchetti con disponibilità giornaliera'); " + _
		" INSERT INTO ttb_pacchetti_tipo (tp_id, tp_nome) " + _
		" 	   VALUES (2, 'Pacchetti con disponibilità a periodo'); " + _
		" INSERT INTO ttb_pacchetti_tipo (tp_id, tp_nome) " + _
		" 	   VALUES (3, 'Pacchetti su richiesta'); "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 54
'...........................................................................................
'	Nicola, 09/11/2010
'...........................................................................................
'   imposta tipi pacchetti
'...........................................................................................
function Aggiornamento_TOUR__54(conn)
	Aggiornamento_TOUR__54 = _
		" UPDATE gtb_articoli " + _
		"  	 SET art_tour_tipo_id = 1; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 55
'...........................................................................................
'	Giacomo, 17/11/2010
'...........................................................................................
'   aggiunta campi su ttb_destinazioni e creazione tabella ttb_destinazioni_interne
'...........................................................................................
function Aggiornamento_TOUR__55(conn)
	Aggiornamento_TOUR__55 = _
		" ALTER TABLE ttb_destinazioni ADD " & _
			SQL_MultiLanguageField("dest_descrizione_<lingua> " & SQL_CharField(Conn, 0)) & "NULL, " & _
		" 	dest_foto_thumb " & SQL_CharField(Conn, 255) & " NULL, " & _
		" 	dest_foto_zoom " & SQL_CharField(Conn, 255) & " NULL, " & _
		" 	dest_ordine int NULL, " & _
		" 	dest_codice " & SQL_CharField(Conn, 50) & " NULL; " & _
		"CREATE TABLE " & SQL_dbo(conn) & "ttb_destinazioni_interne (" + vbCrLf + _
		"	destint_id " & SQL_PrimaryKey(conn, "ttb_destinazioni_interne") & ", " + vbCrLf + _
			SQL_MultiLanguageField("destint_nome_<lingua> " & SQL_CharField(Conn, 0)) & "NULL, " & _
			SQL_MultiLanguageField("destint_descrizione_<lingua> " & SQL_CharField(Conn, 0)) & "NULL, " & _
		" 	destint_foto_thumb " & SQL_CharField(Conn, 255) & " NULL, " & _
		" 	destint_foto_zoom " & SQL_CharField(Conn, 255) & " NULL, " & _
		" 	destint_ordine int NULL, " & _
		" 	destint_codice " & SQL_CharField(Conn, 50) & " NULL, " & _
		" 	destint_destinazione_id int NOT NULL " & _
		" ); " & _
		SQL_AddForeignKey(conn, "ttb_destinazioni_interne", "destint_destinazione_id", "ttb_destinazioni", "dest_id", true, "") & ";"
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 56
'...........................................................................................
'	Giacomo, 26/11/2010
'...........................................................................................
'   aggiunta e modifica tabelle per gestione transfer venetoinside
'...........................................................................................
function Aggiornamento_TOUR__56(conn)
	Aggiornamento_TOUR__56 = _
		" DROP TABLE ttb_destinazioni_interne ; " & _
		" ALTER TABLE ttb_destinazioni DROP COLUMN dest_foto_thumb, dest_foto_zoom, " & _
		SQL_MultiLanguageField("dest_descrizione_<lingua>") & _
		" ALTER TABLE ttb_destinazioni ADD " & _
			SQL_MultiLanguageField("dest_descr_<lingua> " & SQL_CharField(Conn, 0)) & "NULL, " & _
		" 	dest_logo " & SQL_CharField(Conn, 255) & " NULL, " & _
		" 	dest_foto " & SQL_CharField(Conn, 255) & " NULL, " & _
		" 	dest_foglia BIT NULL, " & _
		" 	dest_livello int NULL, " & _
		" 	dest_padre_id int NULL, " & _
		" 	dest_ordine_assoluto " & SQL_CharField(Conn, 255) & " NULL, " & _
		" 	dest_external_id " & SQL_CharField(Conn, 50) & " NULL, " & _
		" 	dest_tipologia_padre_base int NULL, " & _
		" 	dest_visibile BIT NULL, " & _
		" 	dest_albero_visibile BIT NULL, " & _
		" 	dest_tipologie_padre_lista " & SQL_CharField(Conn, 255) & " NULL; " & vbCrLf & _
		" UPDATE ttb_destinazioni SET dest_visibile = 1, dest_albero_visibile = 1; " & vbCrLf & _
		" ALTER TABLE ttb_destinazioni ALTER COLUMN dest_visibile BIT NOT NULL; " & vbCrLf & _
		" ALTER TABLE ttb_destinazioni ALTER COLUMN dest_albero_visibile BIT NOT NULL; " & vbCrLf & _
		SQL_AddForeignKey(conn, "ttb_destinazioni", "dest_padre_id", "ttb_destinazioni", "dest_id", false, "") & vbCrLf & _
		SQL_AddForeignKey(conn, "ttb_destinazioni", "dest_tipologia_padre_base", "ttb_destinazioni", "dest_id", false, "") & vbCrLf & _
		" CREATE TABLE " & SQL_dbo(conn) & "trel_partenze_arrivi (" & _
		"	rpa_id " & SQL_PrimaryKey(conn, "trel_partenze_arrivi") & ", " & _
		"	rpa_partenza_id int NOT NULL, " & _
		"	rpa_arrivo_id int NOT NULL, " & _
		"	rpa_transfer_id int NOT NULL ); " & vbCrLf & _
		SQL_AddForeignKey(conn, "trel_partenze_arrivi", "rpa_partenza_id", "ttb_destinazioni", "dest_id", true, "p") & vbCrLf & _
		SQL_AddForeignKeyExtended(conn, "trel_partenze_arrivi", "rpa_arrivo_id", "ttb_destinazioni", "dest_id", true, false, "a") & vbCrLf & _
		SQL_AddForeignKey(conn, "trel_partenze_arrivi", "rpa_transfer_id", "gtb_articoli", "art_id", true, "") & ";"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 57
'...........................................................................................
'	Matteo, 01/12/2010
'...........................................................................................
'   elimina chiave esterna dai periodi agli orari
'   aggiunge alla tabella degli orari i campi di fine orario, variazione prezzo in euro e percentuale
'   mantiene le relazioni di dati tra orari e periodi dopo il cambio di chiave esterna
'   elimina dalla tabella dei periodi il campo per la vecchia chiave esterna
'	aggiunge chiave esterna dagli orari ai periodi
'...........................................................................................
function Aggiornamento_TOUR__57(conn)
	Aggiornamento_TOUR__57 = _
		SQL_RemoveForeignKey(conn, "ttb_pacchetti_dispo_periodi", "", "", false, "FK_ttb_pacchetti_dispo_periodi__ttb_pacchetti_dispo_orari") + vbCrLf + _
		" ALTER TABLE ttb_pacchetti_dispo_orari ADD " + vbCrLf + _
			"   odp_orario_fine NVARCHAR(50) NULL, " + vbCrLf + _
			"   odp_prezzo_var_sconto REAL NULL, " + vbCrLf + _
			"   odp_prezzo_var_euro MONEY NULL, " + vbCrLf + _
			"	odp_periodo_id INT NULL;" + vbCrLf + _
		" UPDATE ttb_pacchetti_dispo_orari " + vbCrLf + _
		"    SET odp_periodo_id = (SELECT pdp_id " + vbCrLf + _
								   " FROM ttb_pacchetti_dispo_periodi " + vbCrLf + _
								  " WHERE pdp_orario_id = odp_id);" + vbCrLf + _
		" ALTER TABLE ttb_pacchetti_dispo_periodi DROP COLUMN pdp_orario_id; " + vbCrLf + _
		SQL_AddForeignKey(conn, "ttb_pacchetti_dispo_orari", "odp_periodo_id", "ttb_pacchetti_dispo_periodi", "pdp_id", true, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 58
'...........................................................................................
'	Nicola 04/03/2011
'...........................................................................................
'   aggiunge su vincoli per singola tipologia di prenotazione la pagina che permette di gestire
'	i comandi per la prenotazione
'...........................................................................................
function Aggiornamento_TOUR__58(conn)
	Aggiornamento_TOUR__58 = _
		" ALTER TABLE trel_vincoli_profili_tipologie ADD " + vbCrLf + _
		"	rpt_pagina_comandi_prenotazione INT NULL ; " + _
		SQL_AddForeignKey(conn, "trel_vincoli_profili_tipologie", "rpt_pagina_comandi_prenotazione", "tb_paginesito", "id_paginesito", false, "")
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 59
'...........................................................................................
'	Giacomo, 24/05/2011 - Aggiunta campi nuova lingua per il Tour
'...........................................................................................
function Aggiornamento_TOUR__59(conn, lingua_abbr)
	Aggiornamento_TOUR__59 = _
		  " ALTER TABLE ttb_destinazioni ADD " + vbCrLf + _
		  " 	dest_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + vbCrLf + _
		  " 	dest_descr_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + vbCrLf + _
		  " ALTER TABLE ttb_guide_luoghi ADD " + vbCrLf + _
		  " 	luo_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 250) + " NULL;" + vbCrLf + _
		  " ALTER TABLE ttb_profili ADD " + vbCrLf + _
		  " 	pro_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + vbCrLf + _
		  " 	pro_contratto_url_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 60
'..........................................................................................................................
'Giacomo - 24/05/2011
'..........................................................................................................................
'creazione viste tv_pacchetti divise per lingua
'..........................................................................................................................
function Aggiornamento_TOUR__60(conn)

	DropObject conn,"tv_pacchetti_it","VIEW"
	DropObject conn,"tv_pacchetti_en","VIEW"
	DropObject conn,"tv_pacchetti_fr","VIEW"
	DropObject conn,"tv_pacchetti_es","VIEW"
	DropObject conn,"tv_pacchetti_de","VIEW"
	DropObject conn,"tv_pacchetti_cn","VIEW"
	DropObject conn,"tv_pacchetti_ru","VIEW"
	DropObject conn,"tv_pacchetti_pt","VIEW"

	Dim Agg,Agg_it,Agg_en,Agg_es,Agg_fr,Agg_de,Agg_cn,Agg_ru,Agg_pt
	
	Agg = _
        " SELECT art_id, art_nome_it, art_nome_en, art_cod_int, art_cod_pro, art_cod_alt, art_prezzo_base, art_scontoQ_id, " + vbCrLF + _
		"		art_giacenza_min, art_lotto_riordino, art_qta_min_ord, art_NovenSingola, art_se_accessorio, art_ha_accessori, art_se_bundle, " + vbCrLF + _
		"		art_se_confezione, art_in_bundle, art_in_confezione, art_varianti, art_disabilitato, art_tipologia_id, art_marca_id, art_iva_id, " + vbCrLF + _
		"		art_external_id, art_raggruppamento_id, art_accessori_note_it, art_accessori_note_en, art_composizione_note_it, " + vbCrLF + _
		"		art_composizione_note_en, art_descr_it, art_descr_en, art_note, art_guide_luogo_id,  " + vbCrLF + _
		"		art_guida_necessaria, art_insData, art_insAdmin_id, art_modData, art_modAdmin_id, art_tour_dispo_per_profilo, art_tour_dispo_giornaliera,  " + vbCrLF + _
		"		art_tour_durata, art_non_vendibile, art_applicativo_id, art_unico, art_descr_riassunto_it, art_descr_riassunto_en, " + vbCrLF + _
		"		art_descr_prezzo_it, art_descr_prezzo_en, art_spedizione_id, art_tour_google_maps_latitudine,  " + vbCrLF + _
		"		art_tour_google_maps_longitudine, art_url_it, art_url_en, art_tour_dest_partenza_id, art_tour_dest_arrivo_id, art_ordine,  " + vbCrLF + _
		"		art_dettagli_ord_tipo_id, art_tour_tipo_id, art_qta_max_ord, tip_id, tip_nome_it, tip_nome_en, tip_logo, tip_foto,  " + vbCrLF + _
		"		tip_codice, tip_foglia, tip_livello, tip_padre_id, tip_ordine, tip_ordine_assoluto, tip_external_id, tip_tipologia_padre_base,  " + vbCrLF + _
		"		tip_visibile, tip_albero_visibile, tip_descr_it, tip_descr_en, tip_tipologie_padre_lista, mar_id, mar_nome_it, mar_nome_en,  " + vbCrLF + _
		"		mar_logo, mar_sito, mar_codice, mar_generica, mar_descr_it, mar_descr_en, mar_anagrafica_id, luo_id,  " + vbCrLF + _
		"		luo_nome_it, luo_nome_en " + vbCrLF + _
		" FROM gtb_articoli " + vbCrLF + _
		"		INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id " + vbCrLF + _
        "       INNER JOIN gtb_marche ON gtb_articoli.art_marca_id = gtb_marche.mar_id " + vbCrLF + _
        "       LEFT JOIN ttb_guide_luoghi ON gtb_articoli.art_guide_luogo_id = ttb_guide_luoghi.luo_id " + vbCrLF + _
		";"
		
	Agg_it=" CREATE VIEW " & SQL_Dbo(conn) & "tv_pacchetti_it AS " + vbCrLf + Agg
	Agg_en=" CREATE VIEW " & SQL_Dbo(conn) & "tv_pacchetti_en AS " + vbCrLf + Agg
	Agg_cn=	_
		" CREATE VIEW " & SQL_Dbo(conn) & "tv_pacchetti_cn AS " + vbCrLf + _
        " SELECT art_id, art_nome_it, art_nome_en, art_nome_cn, art_cod_int, art_cod_pro, art_cod_alt, art_prezzo_base, art_scontoQ_id, " + vbCrLF + _
		"		art_giacenza_min, art_lotto_riordino, art_qta_min_ord, art_NovenSingola, art_se_accessorio, art_ha_accessori, art_se_bundle, " + vbCrLF + _
		"		art_se_confezione, art_in_bundle, art_in_confezione, art_varianti, art_disabilitato, art_tipologia_id, art_marca_id, art_iva_id, " + vbCrLF + _
		"		art_external_id, art_raggruppamento_id, art_accessori_note_it, art_accessori_note_en, art_accessori_note_cn, art_composizione_note_it, " + vbCrLF + _
		"		art_composizione_note_en, art_composizione_note_cn, art_descr_it, art_descr_en, art_descr_cn, art_note, art_guide_luogo_id,  " + vbCrLF + _
		"		art_guida_necessaria, art_insData, art_insAdmin_id, art_modData, art_modAdmin_id, art_tour_dispo_per_profilo, art_tour_dispo_giornaliera,  " + vbCrLF + _
		"		art_tour_durata, art_non_vendibile, art_applicativo_id, art_unico, art_descr_riassunto_it, art_descr_riassunto_en, art_descr_riassunto_cn,  " + vbCrLF + _
		"		art_descr_prezzo_it, art_descr_prezzo_en, art_descr_prezzo_cn, art_spedizione_id, art_tour_google_maps_latitudine,  " + vbCrLF + _
		"		art_tour_google_maps_longitudine, art_url_it, art_url_en, art_url_cn, art_tour_dest_partenza_id, art_tour_dest_arrivo_id, art_ordine,  " + vbCrLF + _
		"		art_dettagli_ord_tipo_id, art_tour_tipo_id, art_qta_max_ord, tip_id, tip_nome_it, tip_nome_en, tip_nome_cn, tip_logo, tip_foto,  " + vbCrLF + _
		"		tip_codice, tip_foglia, tip_livello, tip_padre_id, tip_ordine, tip_ordine_assoluto, tip_external_id, tip_tipologia_padre_base,  " + vbCrLF + _
		"		tip_visibile, tip_albero_visibile, tip_descr_it, tip_descr_en, tip_descr_cn, tip_tipologie_padre_lista, mar_id, mar_nome_it, mar_nome_en,  " + vbCrLF + _
		"		mar_nome_cn, mar_logo, mar_sito, mar_codice, mar_generica, mar_descr_it, mar_descr_en, mar_descr_cn, mar_anagrafica_id, luo_id,  " + vbCrLF + _
		"		luo_nome_it, luo_nome_en, luo_nome_cn " + vbCrLF + _
		" FROM gtb_articoli " + vbCrLF + _
		"		INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id " + vbCrLF + _
        "       INNER JOIN gtb_marche ON gtb_articoli.art_marca_id = gtb_marche.mar_id " + vbCrLF + _
        "       LEFT JOIN ttb_guide_luoghi ON gtb_articoli.art_guide_luogo_id = ttb_guide_luoghi.luo_id " + vbCrLF + _
		";"
		
	Agg_es = Replace(Agg_cn,"_cn","_es")
	Agg_fr = Replace(Agg_cn,"_cn","_fr")
	Agg_ru = Replace(Agg_cn,"_cn","_ru")
	Agg_pt = Replace(Agg_cn,"_cn","_pt")
	Agg_de = Replace(Agg_cn,"_cn","_de")

	Aggiornamento_TOUR__60 = _
		DropObject(conn,"tv_pacchetti_it","VIEW") + _
		DropObject(conn,"tv_pacchetti_en","VIEW") + _
		DropObject(conn,"tv_pacchetti_fr","VIEW") + _
		DropObject(conn,"tv_pacchetti_de","VIEW") + _
		DropObject(conn,"tv_pacchetti_es","VIEW") + _
		DropObject(conn,"tv_pacchetti_ru","VIEW") + _
		DropObject(conn,"tv_pacchetti_pt","VIEW") + _
		DropObject(conn,"tv_pacchetti_cn","VIEW") + _
		Agg_it + Agg_en  + Agg_fr + Agg_de + Agg_es + Agg_ru + Agg_pt + Agg_cn
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TOUR 61
'...........................................................................................
'	Nicola, 10/05/2013
'...........................................................................................
'   aggiunge campi in lingua per note sul voucher
'...........................................................................................
function Aggiornamento_TOUR__61(conn)
	Aggiornamento_TOUR__61 = _
		" ALTER TABLE ttb_pacchetti_dispo_orari ADD " & _
			SQL_MultiLanguageField("odp_orari_note_<lingua> " & SQL_CharField(Conn, 255) & " NULL" )
end function
'*******************************************************************************************













%>